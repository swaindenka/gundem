"""Gündem şablonunu (BASLIK_TEST.docx) çıkarılan verilerle doldurur."""
from __future__ import annotations

import copy
import datetime as dt
import re
from typing import List, Optional, Sequence, Tuple

import docx
from docx.oxml.ns import qn

from . import duzen as dz
from .extract import BASLIK, UYE, CaseData, norm, parse_alt_bolum

AYLAR = ["OCAK", "ŞUBAT", "MART", "NİSAN", "MAYIS", "HAZİRAN", "TEMMUZ",
         "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK"]
GUNLER = ["PAZARTESİ", "SALI", "ÇARŞAMBA", "PERŞEMBE", "CUMA", "CUMARTESİ", "PAZAR"]

VML_NS = "urn:schemas-microsoft-com:vml"
O_NS = "urn:schemas-microsoft-com:office:office"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"


class TemplateError(Exception):
    pass


def gundem_basligi(tarih: dt.date) -> str:
    return "%d %s %d %s GÜNÜ GÜNDEMİ" % (
        tarih.day, AYLAR[tarih.month - 1], tarih.year, GUNLER[tarih.weekday()])


# --------------------------------------------------------------------------
# Düşük seviye XML yardımcıları
# --------------------------------------------------------------------------
def _children(el, tag: str):
    return [c for c in el if c.tag == qn(tag)]


def _runs(p):
    return [r for r in p.iter(qn("w:r"))]


def _run_text(r) -> str:
    out = []
    for el in r:
        if el.tag == qn("w:t"):
            out.append(el.text or "")
        elif el.tag == qn("w:tab"):
            out.append("\t")
    return "".join(out)


def _format_run(paragraphs) -> Optional[object]:
    """Biçimi örnek alınacak run: metni boş olmayan ilk run (yoksa ilk run)."""
    first = None
    for p in paragraphs:
        for r in _runs(p):
            if first is None:
                first = r
            if _run_text(r).strip():
                return r
    return first


def _make_run(proto_run, text: str, bold: Optional[bool] = None):
    r = docx.oxml.OxmlElement("w:r")
    if proto_run is not None:
        rpr = proto_run.find(qn("w:rPr"))
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
    if bold is not None:
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = docx.oxml.OxmlElement("w:rPr")
            r.insert(0, rpr)
        for tag in ("w:b", "w:bCs"):
            for old in rpr.findall(qn(tag)):
                rpr.remove(old)
        if bold:
            b = docx.oxml.OxmlElement("w:b")
            # rPr içinde w:b, rFonts'tan sonra gelmeli
            fonts = rpr.find(qn("w:rFonts"))
            idx = list(rpr).index(fonts) + 1 if fonts is not None else 0
            rpr.insert(idx, b)
    parts = re.split(r"(\t)", text)
    for part in parts:
        if part == "\t":
            r.append(docx.oxml.OxmlElement("w:tab"))
        elif part:
            t = docx.oxml.OxmlElement("w:t")
            t.text = part
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r.append(t)
    return r


def _set_paragraph_text(p, proto_run, text: str, bold: Optional[bool] = None):
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)
    if text:
        p.append(_make_run(proto_run, text, bold))


def set_cell_lines(tc, text: str, bold: Optional[bool] = None):
    """Hücrenin metnini değiştirir; paragraf ve yazı biçimini korur. Her satır ayrı paragraf olur."""
    paras = _children(tc, "w:p")
    if not paras:
        p = docx.oxml.OxmlElement("w:p")
        tc.append(p)
        paras = [p]
    proto_run = _format_run(paras)
    proto_run = copy.deepcopy(proto_run) if proto_run is not None else None
    proto_p = paras[0]
    lines = text.split("\n") if text else [""]
    for p in paras[1:]:
        tc.remove(p)
    _set_paragraph_text(proto_p, proto_run, lines[0], bold)
    anchor = proto_p
    for line in lines[1:]:
        newp = copy.deepcopy(proto_p)
        _set_paragraph_text(newp, proto_run, line, bold)
        anchor.addnext(newp)
        anchor = newp


def _tcs(tr):
    return _children(tr, "w:tc")


def _el_text(el) -> str:
    return "".join(_run_text(r) for r in el.iter(qn("w:r")))


class _IdFixer:
    """Kopyalanan öğelerdeki VML şekil kimliklerini benzersiz yapar."""

    def __init__(self, body):
        self.n = 5000
        used = set()
        for el in body.iter():
            if isinstance(el.tag, str) and el.tag.startswith("{%s}" % VML_NS):
                v = el.get("id")
                if v:
                    used.add(v)
        self.used = used

    def _next(self) -> str:
        while True:
            self.n += 1
            v = "_x0000_s%d" % self.n
            if v not in self.used:
                self.used.add(v)
                return v

    def fix(self, root):
        for el in root.iter():
            if not isinstance(el.tag, str):
                continue
            if el.tag.startswith("{%s}" % VML_NS) and el.get("id"):
                el.set("id", self._next())
                if el.get("{%s}spid" % O_NS):
                    el.set("{%s}spid" % O_NS, el.get("id"))
            for attr in ("{%s}paraId" % W14_NS, "{%s}textId" % W14_NS):
                if attr in el.attrib:
                    del el.attrib[attr]
        for bm in list(root.iter(qn("w:bookmarkStart"))) + list(root.iter(qn("w:bookmarkEnd"))):
            bm.getparent().remove(bm)


# --------------------------------------------------------------------------
# Şablon yapısı
# --------------------------------------------------------------------------
class _Template:
    def __init__(self, path: str):
        self.document = docx.Document(path)
        body = self.document.element.body
        self.body = body
        children = [c for c in body if c.tag != qn("w:sectPr")]
        tables = [c for c in children if c.tag == qn("w:tbl")]
        self.header_tbl = next((t for t in tables if "GÜNDEM" in _el_text(t).upper()), None)
        self.case_tbl = next((t for t in tables if "SIRA" in norm(_el_text(t))), None)
        if self.case_tbl is None:
            raise TemplateError("Şablonda 'SIRA NO' başlıklı tablo bulunamadı.")
        idx = children.index(self.case_tbl)
        # Dava tablosu ve ondan sonra gelen (oylama sonucu) paragrafları: her dava için tekrarlanır
        self.case_block = children[idx:]
        rows = _children(self.case_tbl, "w:tr")
        if len(rows) < 4:
            raise TemplateError("Şablondaki dava tablosu beklenen yapıda değil.")
        self.row_ust, self.row_alt = rows[1], rows[2]
        if len(_tcs(self.row_ust)) < 4 or len(_tcs(self.row_alt)) < 3:
            raise TemplateError("Şablondaki ESAS/MAHKEME satırları beklenen yapıda değil.")
        variable = rows[3:]
        # sondaki boş (ızgara) satırları koru
        tail = []
        while variable and not _el_text(variable[-1]).strip():
            tail.insert(0, variable.pop())
        self.variable_rows = variable
        self.proto_baslik = next((r for r in variable if len(_tcs(r)) == 1), None)
        uye = [r for r in variable if len(_tcs(r)) >= 2 and
               re.match(r"\s*(BASKAN|UYE)", norm(_el_text(_tcs(r)[0])))]
        tek = [r for r in variable if len(_tcs(r)) >= 2 and r not in uye]
        self.proto_uye = uye[0] if uye else (tek[0] if tek else None)
        self.proto_tek = tek[0] if tek else self.proto_uye
        if self.proto_uye is None:
            raise TemplateError("Şablonda hakim satırı (BAŞKAN/ÜYE) bulunamadı.")
        if self.proto_baslik is None:
            self.proto_baslik = self.proto_uye
        self.proto_baslik = copy.deepcopy(self.proto_baslik)
        self.proto_uye = copy.deepcopy(self.proto_uye)
        self.proto_tek = copy.deepcopy(self.proto_tek)
        self.olculer = self._olculer(rows, variable)

    def _olculer(self, rows, variable) -> "dz.SablonOlculeri":
        """Sayfaya sığdırma tahmini için şablonun ölçüleri."""
        def tr_h(tr, varsayilan):
            trpr = tr.find(qn("w:trPr"))
            if trpr is not None and trpr.trHeight_val is not None:
                return int(trpr.trHeight_val.twips)
            return varsayilan

        ic = next((t for t in self.proto_uye.iter(qn("w:tbl"))), None)
        deger_satirlari = ic.findall(qn("w:tr")) if ic is not None else []
        deger_satir = tr_h(deger_satirlari[0], 291) if deger_satirlari else 291
        deger_min = sum(max(tr_h(tr, 0), 240) for tr in deger_satirlari) + 60 if deger_satirlari else 600
        grid = tuple(int(g.get(qn("w:w"))) for g in self.case_tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol")))
        # Değişmeyen kısımlar: başlık tablosu, boşluk, dava tablosunun ilk 3 satırı, oylama bölümü.
        # Değer, yerleşik şablon için LibreOffice çıktısıyla ölçülerek ayarlanmıştır.
        oylama = [el for el in self.case_block[1:] if el.tag == qn("w:p")]
        sabit = 2150 + sum(tr_h(r, 400) for r in rows[:3]) + 300 * len(oylama)
        return dz.SablonOlculeri(
            sabit=sabit, oylama_bosluk=300, yazi=11.0, satir_min=tr_h(self.proto_uye, 700),
            baslik_min=tr_h(self.proto_baslik, 250), deger_min=deger_min, deger_satir=deger_satir,
            grid=grid,
        )


def _fill_header(tpl: _Template, tarih: dt.date, tetkik_hakimi: str):
    if tpl.header_tbl is None:
        return
    for p in tpl.header_tbl.iter(qn("w:p")):
        txt = _el_text(p)
        up = txt.upper()
        if "GÜNDEM" in up:
            proto = copy.deepcopy(_format_run([p]))
            _set_paragraph_text(p, proto, gundem_basligi(tarih))
        elif "TETK" in norm(txt) and ":" in txt:
            runs = _runs(p)
            # ':' içeren run'a kadar olan kısmı koru, sonrasını isimle değiştir
            keep_until = next((i for i, r in enumerate(runs) if ":" in _run_text(r)), None)
            if keep_until is None:
                continue
            name_proto = copy.deepcopy(runs[keep_until + 1]) if keep_until + 1 < len(runs) else copy.deepcopy(runs[keep_until])
            colon_run = runs[keep_until]
            ctext = _run_text(colon_run)
            head = ctext[:ctext.index(":") + 1] + " "
            for t in colon_run.findall(qn("w:t")):
                colon_run.remove(t)
            t = docx.oxml.OxmlElement("w:t")
            t.text = head
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            colon_run.append(t)
            for r in runs[keep_until + 1:]:
                r.getparent().remove(r)
            if tetkik_hakimi.strip():
                colon_run.addnext(_make_run(name_proto, tetkik_hakimi.strip()))


def _oylama_bosluklarini_azalt(block, adet: int) -> None:
    """Oylama bölümündeki (kenarlıksız) boş paragraflardan `adet` kadarını siler."""
    if adet <= 0:
        return
    paragraflar = [el for el in block[1:] if el.tag == qn("w:p")]
    bas = next((i for i, p in enumerate(paragraflar) if "OYLAMA" in _el_text(p).upper()), None)
    if bas is None:
        return
    adaylar = []
    for p in paragraflar[bas + 1:]:
        if "OYLAMA" in _el_text(p).upper():
            break
        ppr = p.find(qn("w:pPr"))
        kenarlik = ppr is not None and ppr.find(qn("w:pBdr")) is not None
        if not _el_text(p).strip() and not kenarlik:
            adaylar.append(p)
    for p in adaylar[:adet]:
        p.getparent().remove(p) if p.getparent() is not None else block.remove(p)
        if p in block:
            block.remove(p)


def _build_case(tpl: _Template, case: CaseData, fixer: _IdFixer, duzen: "dz.Duzen" = dz.KADEMELER[0]):
    block = [copy.deepcopy(el) for el in tpl.case_block]
    tbl = block[0]
    rows = _children(tbl, "w:tr")
    ust, alt = _tcs(rows[1]), _tcs(rows[2])
    set_cell_lines(ust[0], case.sira_no)
    set_cell_lines(ust[1], case.esas)
    set_cell_lines(ust[2], case.daire)
    set_cell_lines(ust[3], case.konu)
    set_cell_lines(alt[1], case.mahkeme)
    set_cell_lines(alt[2], case.daire_karar)

    # Değişken (hakim) satırlarını kaldır, yerine çıkarılan satırları koy
    var_count = len(tpl.variable_rows)
    old_rows = rows[3:3 + var_count]
    anchor = rows[2]
    for r in old_rows:
        tbl.remove(r)
    kalemler = parse_alt_bolum(case.alt_bolum)
    grid = tpl.olculer.grid
    yeni_satirlar = []
    if duzen.kompakt:
        for tur, metin, sol, birlestirme in dz.kalemleri_grupla(kalemler):
            if tur == BASLIK:
                row = copy.deepcopy(tpl.proto_baslik)
                set_cell_lines(_tcs(row)[0], metin, bold=True)
            elif tur == UYE:
                row = dz.kompakt_satir(tpl.proto_uye, sol, metin, birlestirme, grid, duzen, set_cell_lines)
            else:
                row = copy.deepcopy(tpl.proto_tek)
                set_cell_lines(_tcs(row)[0], metin)
            yeni_satirlar.append((tur, row))
    else:
        for kind, text in kalemler:
            if kind == BASLIK:
                row = copy.deepcopy(tpl.proto_baslik)
                set_cell_lines(_tcs(row)[0], text, bold=True)
            elif kind == UYE:
                row = copy.deepcopy(tpl.proto_uye)
                set_cell_lines(_tcs(row)[0], text)
            else:
                row = copy.deepcopy(tpl.proto_tek)
                set_cell_lines(_tcs(row)[0], text)
            yeni_satirlar.append((kind, row))
    for tur, row in yeni_satirlar:
        if tur != BASLIK:
            dz.satiri_kucult(row, duzen)
        elif duzen.yazi is not None:
            dz.yazi_boyutu_ayarla(row, max(duzen.yazi, 8))
        anchor.addnext(row)
        anchor = row
    _oylama_bosluklarini_azalt(block, duzen.oylama_bosluk_azalt)
    for el in block:
        fixer.fix(el)
    return block


def duzen_belirle(tpl: _Template, case: CaseData, zorla: Optional[str] = None):
    """Davanın tek sayfaya sığması için kullanılacak düzeni seçer: (Duzen, sığıyor_mu)."""
    ust = tpl.olculer.grid
    ust_metinler = [(case.esas + "\n" + case.mahkeme, sum(ust[1:3]) if len(ust) >= 3 else 3399),
                    (case.daire + "\n" + case.daire_karar, sum(ust[3:6]) if len(ust) >= 6 else 3719),
                    (case.konu, ust[6] if len(ust) >= 7 else 1951)]
    return dz.duzen_sec(tpl.olculer, parse_alt_bolum(case.alt_bolum), ust_metinler, zorla)


def _add_page_break(p):
    """Paragrafın sonuna sayfa sonu ekler (yeni boş paragraf oluşturmadan)."""
    r = docx.oxml.OxmlElement("w:r")
    br = docx.oxml.OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    p.append(r)


def _new_page_break_paragraph():
    p = docx.oxml.OxmlElement("w:p")
    _add_page_break(p)
    return p


def build_gundem(template_path: str, cases: Sequence[CaseData], tarih: dt.date,
                 output_path: str, tetkik_hakimi: str = "", duzen: str = "otomatik") -> List[Tuple[str, bool]]:
    """Gündemi oluşturur. Her dava için (kullanılan düzen adı, tek sayfaya sığıyor_mu) döndürür.

    duzen: 'otomatik' (sığana kadar sıkıştır) ya da duzen.KADEMELER içindeki bir ad.
    """
    if not cases:
        raise ValueError("En az bir dosya gerekli.")
    tpl = _Template(template_path)
    _fill_header(tpl, tarih, tetkik_hakimi)
    fixer = _IdFixer(tpl.body)
    sect = tpl.body.find(qn("w:sectPr"))

    # Her sayfanın başında tekrarlanacak kısım: başlık tablosu (tarih + tetkik hakimi)
    # ve onunla dava tablosu arasındaki boşluk paragrafları
    children = [c for c in tpl.body if c.tag != qn("w:sectPr")]
    case_idx = children.index(tpl.case_tbl)
    if tpl.header_tbl is not None and tpl.header_tbl in children[:case_idx]:
        page_head = children[children.index(tpl.header_tbl):case_idx]
    else:
        page_head = []

    built: List = []
    secimler: List[Tuple[str, bool]] = []
    for i, case in enumerate(cases):
        if i > 0:
            # Sayfa sonunu önceki bloğun son paragrafına ekle; yoksa yeni paragraf
            last = built[-1] if built else None
            if last is not None and last.tag == qn("w:p"):
                _add_page_break(last)
            else:
                built.append(_new_page_break_paragraph())
            head = [copy.deepcopy(el) for el in page_head]
            for el in head:
                fixer.fix(el)
            built.extend(head)
        secim, sigiyor = duzen_belirle(tpl, case, None if duzen == "otomatik" else duzen)
        secimler.append((secim.ad, sigiyor))
        built.extend(_build_case(tpl, case, fixer, secim))
    for el in tpl.case_block:
        tpl.body.remove(el)
    for el in built:
        if sect is not None:
            sect.addprevious(el)
        else:
            tpl.body.append(el)
    # Kişisel bilgileri belge özelliklerinden temizle
    cp = tpl.document.core_properties
    cp.author = ""
    cp.last_modified_by = ""
    tpl.document.save(output_path)
    return secimler
