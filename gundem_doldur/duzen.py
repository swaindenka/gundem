"""Bir davanın tek sayfaya sığması için düzen seçimi.

Hakim sayısı arttıkça dava sayfası ikinci sayfaya taşar. Bu modül davanın yüksekliğini tahmin
eder ve sığana kadar kademeli olarak sıkıştırır:

  0. klasik   : şablonun kendi düzeni (başlıklar tam genişlikte ayrı satır)
  1. kompakt  : "... Hâkimleri" başlığı sol sütunda, o heyetin hakimlerini dikey olarak kapsar
  2-4.        : kompakt + daha küçük yazı, daha alçak satırlar, en son oylama boşlukları azaltılır
"""
from __future__ import annotations

import copy
import re
from dataclasses import dataclass
from typing import List, Optional, Sequence, Tuple

import docx
from docx.oxml.ns import qn
from docx.shared import Pt, Twips

from .extract import BASLIK, UYE


@dataclass(frozen=True)
class Duzen:
    ad: str
    kompakt: bool = False
    yazi: Optional[float] = None        # hakim satırı yazı boyutu (pt); None = şablondaki
    sol_yazi: float = 8.0               # kompakt düzende sol sütundaki başlık yazısı (pt)
    satir_min: Optional[int] = None     # hakim satırı en az yüksekliği (twip); None = şablondaki
    deger_yazi: Optional[float] = None  # değerlendirme kutusu yazı boyutu (pt)
    deger_satir: Optional[int] = None   # değerlendirme kutusu satır yüksekliği (twip)
    oylama_bosluk_azalt: int = 0        # oylama bölümünden çıkarılacak boş satır sayısı


KADEMELER: Tuple[Duzen, ...] = (
    Duzen("klasik"),
    Duzen("kompakt", kompakt=True),
    Duzen("kompakt-10", kompakt=True, yazi=10, sol_yazi=7.5, satir_min=620, deger_yazi=7.5, deger_satir=250),
    Duzen("kompakt-9", kompakt=True, yazi=9, sol_yazi=7, satir_min=540, deger_yazi=7, deger_satir=230,
          oylama_bosluk_azalt=2),
    Duzen("kompakt-8", kompakt=True, yazi=8, sol_yazi=6.5, satir_min=520, deger_yazi=6.5, deger_satir=210,
          oylama_bosluk_azalt=3),
)
DUZEN_ADLARI = {d.ad: d for d in KADEMELER}


# --------------------------------------------------------------------------
# XML yardımcıları
# --------------------------------------------------------------------------
def _tcs(tr):
    return [c for c in tr if c.tag == qn("w:tc")]


def yazi_boyutu_ayarla(el, pt: float) -> None:
    """`el` altındaki tüm run'ların (ve paragraf işaretlerinin) yazı boyutunu ayarlar."""
    yarim = int(round(pt * 2))
    for r in list(el.iter(qn("w:r"))):
        if r.find(qn("w:rPr")) is None:
            r.insert(0, docx.oxml.OxmlElement("w:rPr"))
    for rpr in el.iter(qn("w:rPr")):
        rpr.sz_val = Pt(pt)
        sz = rpr.find(qn("w:sz"))
        cs = rpr.find(qn("w:szCs"))
        if cs is None:
            cs = docx.oxml.OxmlElement("w:szCs")
            sz.addnext(cs)
        cs.set(qn("w:val"), str(yarim))


def satir_yuksekligi(tr, twip: int) -> None:
    trpr = tr.get_or_add_trPr()
    trpr.trHeight_val = Twips(twip)


def _grid(tbl) -> List[int]:
    grid = tbl.find(qn("w:tblGrid"))
    return [int(g.get(qn("w:w"))) for g in grid.findall(qn("w:gridCol"))] if grid is not None else []


def kompakt_satir(proto_uye, sol_metin: Optional[str], orta_metin: str, birlestirme: Optional[str],
                  grid: Sequence[int], duzen: Duzen, set_cell_lines):
    """Kompakt hakim satırı: [sol: heyet başlığı] [orta: hakim] [değerlendirme kutusu].

    birlestirme: 'restart' (heyetin ilk satırı, başlık burada), 'continue' (heyetin devamı) ya da None.
    """
    row = copy.deepcopy(proto_uye)
    tcs = _tcs(row)
    orta = tcs[0]
    sol = copy.deepcopy(orta)
    orta.addprevious(sol)
    span = orta.tcPr.grid_span if orta.tcPr is not None else 1
    sol_gen = grid[0] if grid else 1280
    orta_gen = sum(grid[1:span]) if grid and len(grid) >= span else 3399
    sol.tcPr.grid_span = 1
    sol.tcPr.width = Twips(sol_gen)
    orta.tcPr.grid_span = max(span - 1, 1)
    orta.tcPr.width = Twips(orta_gen)

    set_cell_lines(sol, sol_metin or "", bold=True)
    for p in sol.findall(qn("w:p")):
        ppr = p.get_or_add_pPr()
        ppr.jc_val = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            ppr.remove(ind)
    yazi_boyutu_ayarla(sol, duzen.sol_yazi)
    if birlestirme:
        sol.tcPr.vMerge_val = birlestirme
        sol.tcPr.vAlign_val = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_lines(orta, orta_metin)
    # Dar hücrede iki yana yaslama 'BAŞKAN      : Ad' gibi boşluklar açar: sola yasla
    for p in orta.findall(qn("w:p")):
        p.get_or_add_pPr().jc_val = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
    return row


def satiri_kucult(row, duzen: Duzen) -> None:
    """Hakim satırına kademenin yazı boyutu / satır yüksekliği ayarlarını uygular."""
    tcs = _tcs(row)
    if duzen.satir_min is not None:
        satir_yuksekligi(row, duzen.satir_min)
    metin_hucreleri = [tc for tc in tcs if tc.find(qn("w:tbl")) is None]
    if duzen.yazi is not None:
        for tc in metin_hucreleri:
            # kompakt düzendeki sol sütun kendi boyutunu korur
            if tc.tcPr is not None and tc.tcPr.vMerge_val is not None:
                continue
            if duzen.kompakt and tc is tcs[0] and len(tcs) == 3:
                continue
            yazi_boyutu_ayarla(tc, duzen.yazi)
    for tc in tcs:
        ic = tc.find(qn("w:tbl"))
        if ic is None:
            continue
        if duzen.deger_yazi is not None:
            yazi_boyutu_ayarla(ic, duzen.deger_yazi)
        if duzen.deger_satir is not None:
            for tr in ic.findall(qn("w:tr")):
                if tr.trPr is not None and tr.trPr.trHeight_val is not None:
                    satir_yuksekligi(tr, duzen.deger_satir)


def kalemleri_grupla(kalemler: Sequence[Tuple[str, str]]):
    """Kompakt düzen için: (tür, metin, sol_metin, birlestirme) dizisi üretir.

    Ardından üye satırları gelen başlık sol sütuna taşınır; üyesiz başlık tam genişlikte kalır.
    """
    out = []
    i = 0
    while i < len(kalemler):
        tur, metin = kalemler[i]
        if tur == BASLIK and i + 1 < len(kalemler) and kalemler[i + 1][0] == UYE:
            j = i + 1
            uyeler = []
            while j < len(kalemler) and kalemler[j][0] == UYE:
                uyeler.append(kalemler[j][1])
                j += 1
            baslik = re.sub(r"\s*:\s*$", "", metin)
            for k, uye in enumerate(uyeler):
                out.append((UYE, uye, baslik if k == 0 else "", "restart" if k == 0 else "continue"))
            i = j
            continue
        if tur == UYE:
            out.append((UYE, metin, "", None))
        else:
            out.append((tur, metin, None, None))
        i += 1
    return out


# --------------------------------------------------------------------------
# Yükseklik tahmini (twip; 1 pt = 20 twip)
# --------------------------------------------------------------------------
SAYFA_KULLANILABILIR = 16838 - 2 * 1417   # A4, şablondaki üst/alt kenar boşlukları
GUVENLIK = 0.95                            # Word ile tahmin arasındaki farklar için pay


def _genislik(metin: str, pt: float, kalin: bool = False) -> float:
    """Times New Roman'da metnin yaklaşık genişliği (twip)."""
    w = 0.0
    for c in metin:
        if c == "\t":
            w += 0.9
        elif c.isspace():
            w += 0.25
        elif c.isupper():
            w += 0.68
        elif c.isdigit():
            w += 0.5
        elif c.isalpha():
            w += 0.45
        else:
            w += 0.33
    return w * pt * 20 * (1.06 if kalin else 1.0)


def _satir_sayisi(metin: str, pt: float, hucre_gen: int, kalin: bool = False) -> int:
    kullanilabilir = max(hucre_gen - 216, 200)
    toplam = 0
    for satir in (metin.split("\n") if metin else [""]):
        # kelime kaydırmayı kabaca hesaba kat: her satır sonunda yarım kelime kadar boşluk kalır
        g = _genislik(satir, pt, kalin)
        n = max(1, int(-(-g // (kullanilabilir * 0.92))))
        toplam += n
    return toplam


def _satir_h(pt: float) -> float:
    return pt * 20 * 1.15


@dataclass
class SablonOlculeri:
    sabit: int            # başlık tablosu + boşluk + dava başlık satırları + oylama bölümü (twip)
    oylama_bosluk: int    # oylama bölümündeki bir boş satırın yüksekliği
    yazi: float           # şablondaki hakim satırı yazı boyutu (pt)
    satir_min: int        # şablondaki hakim satırı en az yüksekliği
    baslik_min: int       # şablondaki başlık satırı en az yüksekliği
    deger_min: int        # değerlendirme kutusunun yüksekliği
    deger_satir: int      # değerlendirme kutusundaki bir satırın şablondaki yüksekliği
    grid: Tuple[int, ...]


def _baslik_satiri(o: SablonOlculeri, metin: str, genislik: int) -> float:
    # Tam genişlik başlık satırı: paragraf satır aralığı 1,15 ve paragraf sonrası boşluk (ölçüldü)
    return max(o.baslik_min, _satir_sayisi(metin, 11, genislik, True) * _satir_h(11) * 1.15 + 200)


def tahmini_yukseklik(o: SablonOlculeri, kalemler: Sequence[Tuple[str, str]], d: Duzen,
                      ust_metinler: Sequence[Tuple[str, int]] = ()) -> int:
    """Davanın sayfada kaplayacağı yaklaşık yükseklik (twip)."""
    yazi = d.yazi or o.yazi
    satir_min = d.satir_min if d.satir_min is not None else o.satir_min
    deger = o.deger_min
    if d.deger_satir is not None:
        deger = int(o.deger_min * d.deger_satir / max(o.deger_satir, 1))
    tam = sum(o.grid[:7]) if len(o.grid) >= 7 else 10349
    sol_gen = o.grid[0] if o.grid else 1280
    uc = sum(o.grid[:3]) if len(o.grid) >= 3 else 4679
    orta_gen = uc - sol_gen

    h = o.sabit - d.oylama_bosluk_azalt * o.oylama_bosluk
    # esas / mahkeme / daire / konu hücrelerindeki uzun metinler
    for metin, gen in ust_metinler:
        fazla = _satir_sayisi(metin, 11, gen, True) - 2
        if fazla > 0:
            h += fazla * _satir_h(11)

    def uye_satiri(metin: str, gen: int) -> float:
        return max(satir_min, deger, _satir_sayisi(metin, yazi, gen) * _satir_h(yazi) + 40)

    if not d.kompakt:
        for tur, metin in kalemler:
            if tur == BASLIK:
                h += _baslik_satiri(o, metin, tam)
            else:
                h += uye_satiri(metin, uc)
        return int(h)

    grup_h = 0.0
    grup_baslik = ""
    for tur, metin, sol, birlestirme in kalemleri_grupla(kalemler):
        if birlestirme == "restart":
            if grup_baslik:
                h += max(grup_h, _satir_sayisi(grup_baslik, d.sol_yazi, sol_gen, True) * _satir_h(d.sol_yazi))
            grup_baslik, grup_h = sol or "", 0.0
        if birlestirme:
            grup_h += uye_satiri(metin, orta_gen)
            continue
        if grup_baslik:
            h += max(grup_h, _satir_sayisi(grup_baslik, d.sol_yazi, sol_gen, True) * _satir_h(d.sol_yazi))
            grup_baslik, grup_h = "", 0.0
        if tur == BASLIK:
            h += _baslik_satiri(o, metin, tam)
        elif tur == UYE:
            h += uye_satiri(metin, orta_gen)
        else:
            h += uye_satiri(metin, uc)
    if grup_baslik:
        h += max(grup_h, _satir_sayisi(grup_baslik, d.sol_yazi, sol_gen, True) * _satir_h(d.sol_yazi))
    return int(h)


def duzen_sec(o: SablonOlculeri, kalemler, ust_metinler=(), zorla: Optional[str] = None) -> Tuple[Duzen, bool]:
    """Sayfaya sığan ilk kademeyi seçer. (düzen, sığıyor_mu) döndürür."""
    if zorla and zorla in DUZEN_ADLARI:
        d = DUZEN_ADLARI[zorla]
        return d, tahmini_yukseklik(o, kalemler, d, ust_metinler) <= SAYFA_KULLANILABILIR * GUVENLIK
    sinir = SAYFA_KULLANILABILIR * GUVENLIK
    for d in KADEMELER:
        if tahmini_yukseklik(o, kalemler, d, ust_metinler) <= sinir:
            return d, True
    return KADEMELER[-1], False
