"""Word 97-2003 (.doc) ikili biçimini, harici bir programa ihtiyaç duymadan okur.

Yalnızca bu uygulamanın ihtiyaç duyduğu kadarını çözer: ana metin, paragraflar,
tablolar (hücre/satır sınırları, yatay ve dikey birleştirmeler).
Biçim tanımı: [MS-DOC] Word (.doc) Binary File Format.
"""
from __future__ import annotations

import struct
from bisect import bisect_right
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import olefile

from .model import Cell, Document, Table, sembol_karakteri

# Paragraf özellikleri (sprm) kodları
SPRM_P_FINTABLE = 0x2416
SPRM_P_FTTP = 0x2417
SPRM_P_ITAP = 0x6649
SPRM_P_DTAP = 0x664A
SPRM_P_FINNER_TABLE_CELL = 0x244B
SPRM_P_FINNER_TTP = 0x244C
SPRM_T_DEF_TABLE = 0xD608
SPRM_P_CHG_TABS = 0xC615
SPRM_C_SYMBOL = 0x6A09


class DocReadError(Exception):
    pass


@dataclass
class _ParaProps:
    in_table: bool = False
    ttp: bool = False
    itap: Optional[int] = None
    inner_cell: bool = False
    inner_ttp: bool = False
    tdef: Optional[Tuple[List[int], List[int]]] = None  # (rgdxaCenter, tcgrf listesi)

    @property
    def depth(self) -> int:
        if self.itap is not None:
            return self.itap
        return 1 if self.in_table else 0


@dataclass
class _Piece:
    cp_start: int
    cp_end: int
    fc: int
    compressed: bool


def _operand_size(sprm: int, data: bytes, pos: int) -> Tuple[int, int]:
    """(operand başlangıcı, operand uzunluğu) döndürür."""
    spra = sprm >> 13
    if spra in (0, 1):
        return pos, 1
    if spra in (2, 4, 5):
        return pos, 2
    if spra == 3:
        return pos, 4
    if spra == 7:
        return pos, 3
    # spra == 6: değişken uzunluk
    if sprm == SPRM_T_DEF_TABLE:
        cb = struct.unpack_from("<H", data, pos)[0]
        return pos + 2, max(cb - 1, 0)
    if sprm == SPRM_P_CHG_TABS:
        cb = data[pos]
        if cb == 255:
            # Nadiren görülür: uzunluğu hesaplamak için iç yapıyı oku
            p = pos + 1
            cdel = data[p]
            p += 1 + cdel * 4
            cadd = data[p]
            p += 1 + cadd * 3
            return pos + 1, p - pos - 1
        return pos + 1, cb
    cb = data[pos]
    return pos + 1, cb


def _parse_grpprl(grpprl: bytes, props: _ParaProps) -> None:
    pos = 0
    n = len(grpprl)
    while pos + 2 <= n:
        sprm = struct.unpack_from("<H", grpprl, pos)[0]
        pos += 2
        if pos >= n and (sprm >> 13) != 0:
            break
        try:
            start, size = _operand_size(sprm, grpprl, pos)
        except (IndexError, struct.error):
            break
        operand = grpprl[start:start + size]
        if sprm == SPRM_P_FINTABLE and operand:
            props.in_table = operand[0] != 0
        elif sprm == SPRM_P_FTTP and operand:
            props.ttp = operand[0] != 0
        elif sprm == SPRM_P_ITAP and len(operand) >= 4:
            props.itap = struct.unpack_from("<i", operand)[0]
        elif sprm == SPRM_P_DTAP and len(operand) >= 4:
            props.itap = (props.itap if props.itap is not None else (1 if props.in_table else 0)) + \
                struct.unpack_from("<i", operand)[0]
        elif sprm == SPRM_P_FINNER_TABLE_CELL and operand:
            props.inner_cell = operand[0] != 0
        elif sprm == SPRM_P_FINNER_TTP and operand:
            props.inner_ttp = operand[0] != 0
        elif sprm == SPRM_T_DEF_TABLE and len(operand) >= 1:
            itc = operand[0]
            centers = []
            p = 1
            for _ in range(itc + 1):
                if p + 2 > len(operand):
                    break
                centers.append(struct.unpack_from("<h", operand, p)[0])
                p += 2
            flags = []
            for _ in range(itc):
                if p + 2 > len(operand):
                    flags.append(0)
                else:
                    flags.append(struct.unpack_from("<H", operand, p)[0])
                p += 20
            if len(centers) == itc + 1:
                props.tdef = (centers, flags)
        pos = start + size


class _DocParser:
    def __init__(self, path: str):
        if not olefile.isOleFile(path):
            raise DocReadError("Dosya Word 97-2003 (.doc) biçiminde değil.")
        ole = olefile.OleFileIO(path)
        try:
            if not ole.exists("WordDocument"):
                raise DocReadError("WordDocument akışı bulunamadı.")
            self.wd = ole.openstream("WordDocument").read()
            ident, = struct.unpack_from("<H", self.wd, 0)
            if ident != 0xA5EC:
                raise DocReadError("Geçersiz Word belgesi imzası.")
            flags, = struct.unpack_from("<H", self.wd, 0x0A)
            if flags & 0x0100:
                raise DocReadError("Belge şifreli; açılamıyor.")
            table_name = "1Table" if flags & 0x0200 else "0Table"
            if not ole.exists(table_name):
                raise DocReadError("Tablo akışı bulunamadı.")
            self.tbl = ole.openstream(table_name).read()
        finally:
            ole.close()

        wd = self.wd
        csw, = struct.unpack_from("<H", wd, 32)
        p = 34 + csw * 2
        cslw, = struct.unpack_from("<H", wd, p)
        rglw = p + 2
        self.ccp_text, = struct.unpack_from("<i", wd, rglw + 3 * 4)
        p = rglw + cslw * 4
        cb_rg, = struct.unpack_from("<H", wd, p)
        blob = p + 2

        def fclcb(index: int) -> Tuple[int, int]:
            if index >= cb_rg:
                return 0, 0
            return struct.unpack_from("<II", wd, blob + index * 8)

        self.fc_bte_chpx, self.lcb_bte_chpx = fclcb(12)
        self.fc_bte_papx, self.lcb_bte_papx = fclcb(13)
        self.fc_clx, self.lcb_clx = fclcb(33)
        self.pieces = self._read_pieces()
        self._bte = self._read_bte(self.fc_bte_papx, self.lcb_bte_papx)
        self._bte_chp = self._read_bte(self.fc_bte_chpx, self.lcb_bte_chpx)
        self._fkp_cache: Dict[int, Tuple[List[int], List[bytes]]] = {}
        self._chp_cache: Dict[int, Tuple[List[int], List[bytes]]] = {}

    # ----- Parça tablosu (piece table) -----
    def _read_pieces(self) -> List[_Piece]:
        clx = self.tbl[self.fc_clx:self.fc_clx + self.lcb_clx]
        pos = 0
        while pos < len(clx) and clx[pos] == 0x01:
            cb, = struct.unpack_from("<h", clx, pos + 1)
            pos += 3 + cb
        if pos >= len(clx) or clx[pos] != 0x02:
            raise DocReadError("Metin parça tablosu okunamadı.")
        lcb, = struct.unpack_from("<I", clx, pos + 1)
        plc = clx[pos + 5:pos + 5 + lcb]
        n = (lcb - 4) // 12
        cps = struct.unpack_from("<%dI" % (n + 1), plc, 0)
        pieces = []
        for i in range(n):
            raw, = struct.unpack_from("<I", plc, 4 * (n + 1) + i * 8 + 2)
            compressed = bool(raw & 0x40000000)
            fc = raw & 0x3FFFFFFF
            if compressed:
                fc //= 2
            pieces.append(_Piece(cps[i], cps[i + 1], fc, compressed))
        return pieces

    def _piece_text(self, piece: _Piece) -> str:
        count = piece.cp_end - piece.cp_start
        if piece.compressed:
            raw = self.wd[piece.fc:piece.fc + count]
            return raw.decode("cp1252", errors="replace")
        raw = self.wd[piece.fc:piece.fc + 2 * count]
        return raw.decode("utf-16-le", errors="replace")

    # ----- Paragraf özellikleri -----
    def _read_bte(self, fc: int, lcb: int) -> Tuple[List[int], List[int]]:
        if not lcb:
            return [], []
        data = self.tbl[fc:fc + lcb]
        n = (len(data) - 4) // 8
        fcs = list(struct.unpack_from("<%dI" % (n + 1), data, 0))
        pns = [x & 0x3FFFFF for x in struct.unpack_from("<%dI" % n, data, 4 * (n + 1))]
        return fcs, pns

    def _fkp(self, pn: int) -> Tuple[List[int], List[bytes]]:
        if pn in self._fkp_cache:
            return self._fkp_cache[pn]
        page = self.wd[pn * 512:pn * 512 + 512]
        crun = page[511]
        rgfc = list(struct.unpack_from("<%dI" % (crun + 1), page, 0))
        grpprls = []
        base = 4 * (crun + 1)
        for i in range(crun):
            b_off = page[base + i * 13]
            if b_off == 0:
                grpprls.append(b"")
                continue
            o = b_off * 2
            cb = page[o]
            if cb == 0:
                cb2 = page[o + 1]
                body = page[o + 2:o + 2 + 2 * cb2]
            else:
                body = page[o + 1:o + 1 + 2 * cb - 1]
            grpprls.append(body[2:])  # ilk 2 bayt: istd (stil)
        self._fkp_cache[pn] = (rgfc, grpprls)
        return rgfc, grpprls

    def _chp_fkp(self, pn: int) -> Tuple[List[int], List[bytes]]:
        if pn in self._chp_cache:
            return self._chp_cache[pn]
        page = self.wd[pn * 512:pn * 512 + 512]
        crun = page[511]
        rgfc = list(struct.unpack_from("<%dI" % (crun + 1), page, 0))
        grpprls = []
        base = 4 * (crun + 1)
        for i in range(crun):
            o = page[base + i] * 2
            if o == 0:
                grpprls.append(b"")
                continue
            cb = page[o]
            grpprls.append(page[o + 1:o + 1 + cb])
        self._chp_cache[pn] = (rgfc, grpprls)
        return rgfc, grpprls

    def symbol_at(self, fc: int) -> Optional[int]:
        """Karakterin sprmCSymbol ile verilmiş simge kodu (varsa)."""
        fcs, pns = self._bte_chp
        if not pns:
            return None
        j = bisect_right(fcs, fc) - 1
        if j < 0 or j >= len(pns):
            return None
        try:
            rgfc, grpprls = self._chp_fkp(pns[j])
        except struct.error:
            return None
        i = bisect_right(rgfc, fc) - 1
        if not 0 <= i < len(grpprls):
            return None
        g = grpprls[i]
        pos = 0
        while pos + 2 <= len(g):
            sprm = struct.unpack_from("<H", g, pos)[0]
            pos += 2
            try:
                start, size = _operand_size(sprm, g, pos)
            except (IndexError, struct.error):
                return None
            if sprm == SPRM_C_SYMBOL and size >= 4 and start + 4 <= len(g):
                return struct.unpack_from("<H", g, start + 2)[0]
            pos = start + size
        return None

    def para_props(self, fc: int) -> _ParaProps:
        props = _ParaProps()
        fcs, pns = self._bte
        if not pns:
            return props
        j = bisect_right(fcs, fc) - 1
        if j < 0 or j >= len(pns):
            return props
        rgfc, grpprls = self._fkp(pns[j])
        i = bisect_right(rgfc, fc) - 1
        if 0 <= i < len(grpprls):
            _parse_grpprl(grpprls[i], props)
        return props

    # ----- Metni paragraflara bölme -----
    def paragraphs(self):
        """(metin, bitiş karakteri, özellikler) üçlüleri üretir."""
        buf: List[str] = []
        field_stack: List[str] = []
        limit = self.ccp_text
        for piece in self.pieces:
            if piece.cp_start >= limit:
                break
            text = self._piece_text(piece)
            end = min(piece.cp_end, limit) - piece.cp_start
            for k in range(end):
                ch = text[k]
                if ch in "\r\x07":
                    fc = piece.fc + (k if piece.compressed else 2 * k)
                    yield "".join(buf), ch, self.para_props(fc)
                    buf = []
                    continue
                if ch == "\x13":
                    field_stack.append("instr")
                    continue
                if ch == "\x14":
                    if field_stack:
                        field_stack[-1] = "result"
                    continue
                if ch == "\x15":
                    if field_stack:
                        field_stack.pop()
                    continue
                if any(s == "instr" for s in field_stack):
                    continue
                if ch == "(":
                    # 0x28 + sprmCSymbol: Ekle > Simge ile eklenmiş karakter
                    kod = self.symbol_at(piece.fc + (k if piece.compressed else 2 * k))
                    buf.append(sembol_karakteri(kod) if kod is not None else ch)
                elif "\uf000" <= ch <= "\uf0ff":
                    buf.append(sembol_karakteri(ord(ch)))
                elif ch in "\x0b\x0c":
                    buf.append("\n")
                elif ch == "\x1e":
                    buf.append("-")
                elif ch == "\xa0":
                    buf.append(" ")
                elif ch == "\t" or ord(ch) >= 0x20:
                    buf.append(ch)
                # diğer kontrol karakterleri (resim, dipnot vb.) atlanır
        if buf:
            yield "".join(buf), "\r", _ParaProps()


def _finish_table(rows_raw) -> Table:
    """Ham satırları (hücre metinleri + tdef bilgisi) ızgaraya oturtur."""
    boundaries = set()
    for cells, tdef in rows_raw:
        if tdef and len(tdef[0]) == len(cells) + 1:
            boundaries.update(tdef[0])
    grid = sorted(boundaries)
    table = Table()
    for cells, tdef in rows_raw:
        row: List[Cell] = []
        usable = tdef and len(tdef[0]) == len(cells) + 1 and grid
        for idx, text in enumerate(cells):
            span = 1
            vcont = False
            if usable:
                left, right = tdef[0][idx], tdef[0][idx + 1]
                span = max(1, sum(1 for g in grid if left <= g < right))
                flags = tdef[1][idx] if idx < len(tdef[1]) else 0
                if flags & 0x0002 and row:  # fMerged: yatay birleştirmenin devamı
                    row[-1].span += span
                    if text:
                        row[-1].text = (row[-1].text + "\n" + text).strip("\n")
                    continue
                vcont = bool(flags & 0x0020) and not bool(flags & 0x0040)
            row.append(Cell(text=text, span=span, vcontinue=vcont))
        table.rows.append(row)
    return table


def read_doc(path: str) -> Document:
    parser = _DocParser(path)
    doc = Document()
    rows_raw: List[Tuple[List[str], Optional[Tuple[List[int], List[int]]]]] = []
    row_cells: List[str] = []
    cell_paras: List[str] = []

    def flush_table():
        nonlocal rows_raw, row_cells, cell_paras
        if cell_paras:
            row_cells.append("\n".join(cell_paras))
            cell_paras = []
        if row_cells:
            rows_raw.append((row_cells, None))
            row_cells = []
        if rows_raw:
            doc.blocks.append(_finish_table(rows_raw))
            rows_raw = []

    for text, mark, props in parser.paragraphs():
        depth = props.depth
        if depth <= 0:
            flush_table()
            doc.blocks.append(text)
            continue
        if depth >= 2:
            # İç içe tablo: içeriği dıştaki hücrenin metnine ekle
            if not props.inner_ttp and text.strip():
                cell_paras.append(text)
            continue
        if mark == "\x07" and props.ttp:
            if cell_paras:
                row_cells.append("\n".join(cell_paras))
                cell_paras = []
            rows_raw.append((row_cells, props.tdef))
            row_cells = []
        elif mark == "\x07":
            cell_paras.append(text)
            row_cells.append("\n".join(cell_paras))
            cell_paras = []
        else:
            cell_paras.append(text)
    flush_table()
    return doc
