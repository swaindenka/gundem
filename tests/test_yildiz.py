"""Yıldız (***) işaretinin farklı biçimlerde tanınması ve tetkik hakimi imzası."""
import os
import shutil
import subprocess
import sys

import docx
import pytest
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

KOK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, KOK)

from gundem_doldur.extract import UYE, extract_case, parse_alt_bolum  # noqa: E402
from gundem_doldur.reader import read_document  # noqa: E402

LISTE = [
    "İlk kararı veren ilk derece mahkemesi (Bursa 2. AHM) hakimi Hakan Yıldız 445566",
    "Direnme kararını veren Bursa BAM 3. Hukuk Dairesi hakimleri:",
    "BAŞKAN : Ali Veli (12345)",
    "ÜYE : Ayşe Kaya (23456)",
    "ÜYE : Mehmet Demir (34567)",
]
UYELER = ["BAŞKAN\t: Ali Veli  (12345)", "ÜYE\t: Ayşe Kaya  (23456)", "ÜYE\t: Mehmet Demir  (34567)"]


def _belge(path, yildiz_ekle, sonra=LISTE, once=()):
    d = docx.Document()
    t = d.add_table(rows=4, cols=3)
    for i, row in enumerate([("ESAS NO", ":", "2025/9"), ("MAHKEMESİ", ":", "Bursa 2. AHM"),
                             ("DAİRESİ", ":", "7. Hukuk Dairesi"), ("KONUSU", ":", "Alacak")]):
        for j, v in enumerate(row):
            t.cell(i, j).text = v
    d.add_paragraph("Gövde metni. Sonuç olarak direnme kararının bozulması gerekir.")
    for l in once:
        d.add_paragraph(l)
    yildiz_ekle(d)
    for l in sonra:
        d.add_paragraph(l)
    d.save(path)
    return path


def _koyu(metin):
    def ekle(d):
        d.add_paragraph().add_run(metin).bold = True
    return ekle


def _uyeler(case):
    return [t for k, t in parse_alt_bolum(case.alt_bolum) if k == UYE]


@pytest.mark.parametrize("isaret", ["***", "*******", "* * *", "∗∗∗", "＊＊＊", "★★★", " ***​"])
def test_yildiz_cesitleri(tmp_path, isaret):
    case = extract_case(read_document(_belge(str(tmp_path / "a.docx"), _koyu(isaret))))
    assert case.yildiz_var
    assert case.alt_kaynak == "yildiz"
    assert _uyeler(case) == UYELER
    assert not any("tanınamadı" in u or "bulunamadı" in u for u in case.uyarilar)


def _sembol_yildiz(d):
    """Ekle > Simge ile Symbol yazı tipinden üç yıldız (w:sym F02A)."""
    p = d.add_paragraph()
    for _ in range(3):
        r = p.add_run()
        r.bold = True
        sym = OxmlElement("w:sym")
        sym.set(qn("w:font"), "Symbol")
        sym.set(qn("w:char"), "F02A")
        r._r.append(sym)


def test_simge_ile_eklenmis_yildiz_docx(tmp_path):
    case = extract_case(read_document(_belge(str(tmp_path / "s.docx"), _sembol_yildiz)))
    assert case.yildiz_var and case.alt_kaynak == "yildiz"
    assert _uyeler(case) == UYELER


@pytest.mark.skipif(not shutil.which("soffice"), reason="LibreOffice yok")
@pytest.mark.parametrize("ekle", [_sembol_yildiz, _koyu("***")], ids=["simge", "koyu"])
def test_yildiz_doc(tmp_path, ekle):
    p = _belge(str(tmp_path / "s.docx"), ekle)
    subprocess.run(["soffice", "--headless", "--convert-to", "doc:MS Word 97", "--outdir", str(tmp_path), p],
                   capture_output=True, timeout=120)
    case = extract_case(read_document(str(tmp_path / "s.doc")))
    assert case.yildiz_var and case.alt_kaynak == "yildiz"
    assert _uyeler(case) == UYELER


def test_yildiz_metinle_ayni_satirda(tmp_path):
    def ekle(d):
        d.add_paragraph("*** " + LISTE[0])
    case = extract_case(read_document(_belge(str(tmp_path / "b.docx"), ekle, sonra=LISTE[1:])))
    assert case.alt_kaynak == "yildiz"
    assert "Hakan Yıldız" in case.alt_bolum and _uyeler(case) == UYELER


def test_yildizin_hemen_altinda_tetkik_hakimi(tmp_path):
    """Kullanıcının dosyası: '***' altında önce tetkik hakimi, sonra hakim listesi."""
    sonra = ["Tetkik Hakimi", "Rapor Yazan", ""] + LISTE
    case = extract_case(read_document(_belge(str(tmp_path / "t.docx"), _koyu("***"), sonra=sonra)))
    assert case.alt_kaynak == "yildiz"
    assert _uyeler(case) == UYELER
    assert "Rapor Yazan" not in case.alt_bolum and "Tetkik" not in case.alt_bolum


def test_tetkik_imzasi_isim_ustte(tmp_path):
    sonra = LISTE + ["", "Rapor Yazan", "Tetkik Hakimi"]
    case = extract_case(read_document(_belge(str(tmp_path / "u.docx"), _koyu("***"), sonra=sonra)))
    assert _uyeler(case) == UYELER
    assert "Rapor Yazan" not in case.alt_bolum


def test_uyari_yildiz_yok(tmp_path):
    case = extract_case(read_document(_belge(str(tmp_path / "y.docx"), lambda d: None)))
    assert not case.yildiz_var
    assert any("yıldız (***) işareti bulunamadı" in u for u in case.uyarilar)


def test_uyari_yildiz_var_ama_hakim_tanınamadı(tmp_path):
    p = _belge(str(tmp_path / "z.docx"), _koyu("***"), sonra=["Saygılarımla arz olunur."], once=LISTE)
    case = extract_case(read_document(p))
    assert case.yildiz_var
    assert _uyeler(case) == UYELER
    assert any("altında hakim bilgisi tanınamadı" in u for u in case.uyarilar)
