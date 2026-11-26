"""Çok hakimli davaların tek sayfaya sığdırılması (kompakt düzen ve kademeli küçültme)."""
import datetime as dt
import os
import re
import shutil
import subprocess
import sys

import docx
import pytest
from docx.oxml.ns import qn

KOK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, KOK)

from gundem_doldur import duzen as dz  # noqa: E402
from gundem_doldur.extract import CaseData, parse_alt_bolum  # noqa: E402
from gundem_doldur.writer import _Template, build_gundem  # noqa: E402

SABLON = os.path.join(KOK, "sablon", "BASLIK_TEST.docx")


def _dava(heyet_sayisi: int, sira="1") -> CaseData:
    satirlar = ["Not Verilecek Hakim: İlk Derece Asliye Hukuk Hâkimi Mustafa Enes Yılmazoğlu (195555)"]
    for g in range(heyet_sayisi):
        satirlar += [
            "%d. Kararı Veren Samsun Bölge Adliye Mahkemesi 6. Hukuk Dairesi Hâkimleri:" % (g + 1),
            "BAŞKAN : Musa Enes Çakmak (5240)",
            "ÜYE : Ali Veli Karadenizli (66665)",
            "ÜYE : Ayşe Kaya Demirtaş (77775)",
        ]
    return CaseData(sira_no=sira, esas="2025/1234", mahkeme="Bakırköy 55. Asliye Hukuk Mahkemesi",
                    daire="3. Hukuk Dairesi", daire_karar="13.12.2024 tarihli ve\n2024/9000 E. 2024/60000 K.",
                    konu="Alacak davası", alt_bolum="\n".join(satirlar))


def _olustur(tmp_path, davalar, duzen="otomatik", ad="g.docx"):
    out = str(tmp_path / ad)
    secimler = build_gundem(SABLON, davalar, dt.date(2026, 12, 2), out, "", duzen)
    return out, secimler


def _dava_tablolari(path):
    d = docx.Document(path)
    return [t for t in d.tables if "SIRA" in t.rows[0].cells[0].text]


def test_az_hakim_klasik_kalir(tmp_path):
    _, secimler = _olustur(tmp_path, [_dava(1)])
    assert secimler == [("klasik", True)]


@pytest.mark.parametrize("heyet", [3, 4])
def test_cok_hakim_sikistirilir(tmp_path, heyet):
    _, secimler = _olustur(tmp_path, [_dava(heyet)])
    ad, sigar = secimler[0]
    assert ad != "klasik" and sigar


def test_asiri_hakimda_uyari(tmp_path):
    _, secimler = _olustur(tmp_path, [_dava(8)])
    assert secimler == [("kompakt-8", False)]


def test_kademeler_sirayla_kuculur():
    o = _Template(SABLON).olculer
    kalemler = parse_alt_bolum(_dava(4).alt_bolum)
    yukseklikler = [dz.tahmini_yukseklik(o, kalemler, d) for d in dz.KADEMELER]
    assert yukseklikler == sorted(yukseklikler, reverse=True)


def test_kompakt_duzen_yapisi(tmp_path):
    """Heyet başlığı sol sütunda, heyetin hakimlerini dikey birleştirme ile kapsar."""
    out, _ = _olustur(tmp_path, [_dava(2)], "kompakt")
    tbl = _dava_tablolari(out)[0]._tbl
    satirlar = tbl.findall(qn("w:tr"))
    birlesik = []
    for tr in satirlar:
        tcs = tr.findall(qn("w:tc"))
        if len(tcs) == 3 and tcs[0].tcPr is not None and tcs[0].tcPr.vMerge_val is not None:
            metin = "".join(t.text or "" for t in tcs[0].iter(qn("w:t")))
            orta = "".join(t.text or "" for t in tcs[1].iter(qn("w:t")))
            birlesik.append((tcs[0].tcPr.vMerge_val, metin, orta))
    assert [b[0] for b in birlesik] == ["restart", "continue", "continue"] * 2
    assert birlesik[0][1] == "1. Kararı Veren Samsun Bölge Adliye Mahkemesi 6. Hukuk Dairesi Hâkimleri"
    assert birlesik[1][1] == "" and birlesik[0][2].startswith("BAŞKAN")
    tum = "".join(t.text or "" for t in tbl.iter(qn("w:t")))
    # başlık artık tam genişlik ayrı satır değil
    assert tum.count("Hukuk Dairesi Hâkimleri") == 2


def test_kucuk_yazi_ve_oylama(tmp_path):
    out_k, _ = _olustur(tmp_path, [_dava(2)], "klasik", "k.docx")
    out_9, _ = _olustur(tmp_path, [_dava(2)], "kompakt-9", "n.docx")
    tbl = _dava_tablolari(out_9)[0]._tbl
    boyutlar = {sz.get(qn("w:val")) for tr in tbl.findall(qn("w:tr"))[3:]
                for tc in tr.findall(qn("w:tc"))[1:2] for sz in tc.iter(qn("w:sz"))}
    assert "18" in boyutlar  # 9 pt
    p_k = len(docx.Document(out_k).paragraphs)
    p_9 = len(docx.Document(out_9).paragraphs)
    assert p_9 == p_k - 2  # oylama bölümünden 2 boş satır çıkarıldı
    assert sum("OYLAMA" in p.text for p in docx.Document(out_9).paragraphs) == 2


def test_heyetsiz_uyeler_ve_tek_satirlar(tmp_path):
    c = CaseData(sira_no="1", alt_bolum="Direnme kararını veren hâkim: Ali Veli-1111111\n"
                                        "BAŞKAN : Musa Enes Çakmak (5240)\nÜYE : Ayşe Kaya (7777)")
    out, _ = _olustur(tmp_path, [c], "kompakt")
    metin = "\n".join(c.text for t in _dava_tablolari(out) for r in t.rows for c in r.cells)
    assert "Ali Veli-1111111" in metin and "Musa Enes Çakmak" in metin and "Ayşe Kaya" in metin


def _sayfa_sayisi(path):
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(path), path],
                   capture_output=True, timeout=180)
    pdf = os.path.splitext(path)[0] + ".pdf"
    bilgi = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout
    return int(re.search(r"Pages:\s+(\d+)", bilgi).group(1))


@pytest.mark.skipif(not (shutil.which("soffice") and shutil.which("pdfinfo")), reason="LibreOffice yok")
def test_her_dava_tek_sayfa_libreoffice(tmp_path):
    """Otomatik düzende 1-4 heyetli (4-13 hakimli) davaların her biri tek sayfada kalmalı."""
    davalar = [_dava(h, str(i + 1)) for i, h in enumerate([1, 2, 3, 4])]
    out, secimler = _olustur(tmp_path, davalar)
    assert all(s for _, s in secimler)
    assert _sayfa_sayisi(out) == len(davalar)
    # karşılaştırma: klasik düzende taşar
    out_k, _ = _olustur(tmp_path, davalar, "klasik", "klasik.docx")
    assert _sayfa_sayisi(out_k) > len(davalar)
