"""Hakim bilgilerinin farklı yazım biçimlerinden doğru çıkarılması."""
import os
import sys

import docx
import pytest

KOK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, KOK)

from gundem_doldur.extract import BASLIK, TEK, UYE, extract_case, parse_alt_bolum  # noqa: E402
from gundem_doldur.reader import read_document  # noqa: E402

BEKLENEN = [
    (BASLIK, "İlk BAM kararını veren İstanbul BAM 5. Hukuk Dairesi hakimleri:"),
    (UYE, "BAŞKAN\t: Ali Veli  (12345)"),
    (UYE, "ÜYE\t: Ayşe Kaya  (23456)"),
    (UYE, "ÜYE\t: Mehmet Demir  (34567)"),
]
BASLIK_SATIRI = "İlk BAM kararını veren İstanbul BAM 5. Hukuk Dairesi hakimleri:"


@pytest.mark.parametrize("satirlar", [
    # standart
    ["BAŞKAN : Ali Veli (12345)", "ÜYE : Ayşe Kaya (23456)", "ÜYE : Mehmet Demir (34567)"],
    # iki nokta yok, sicil parantezsiz
    ["Başkan Ali Veli 12345", "Üye Ayşe Kaya 23456", "Üye Mehmet Demir 34567"],
    # numaralı / madde işaretli
    ["1- Başkan: Ali Veli 12345", "2- Üye: Ayşe Kaya 23456", "3- Üye: Mehmet Demir 34567"],
    ["• BAŞKAN - Ali Veli - 12345", "• ÜYE - Ayşe Kaya - 23456", "• ÜYE - Mehmet Demir - 34567"],
    # "Sicil No:" yazımı
    ["Başkan : Ali Veli Sicil No: 12345", "Üye : Ayşe Kaya Sicil No: 23456", "Üye : Mehmet Demir Sicil: 34567"],
    # görev isimden sonra
    ["Ali Veli (Başkan) 12345", "Ayşe Kaya (Üye) 23456", "Mehmet Demir (Üye) 34567"],
    ["Ali Veli - Başkan - 12345", "Ayşe Kaya - Üye - 23456", "Mehmet Demir - Üye - 34567"],
    # görev, isim ve sicil ayrı satırlarda
    ["BAŞKAN", "Ali Veli", "12345", "ÜYE", "Ayşe Kaya", "23456", "ÜYE :", "Mehmet Demir", "(34567)"],
    # isim + sicil aynı satır, görev ayrı satır
    ["Başkan:", "Ali Veli 12345", "Üye:", "Ayşe Kaya 23456", "Üye:", "Mehmet Demir 34567"],
])
def test_uye_bicimleri(satirlar):
    assert parse_alt_bolum("\n".join([BASLIK_SATIRI] + satirlar)) == BEKLENEN


def test_baslik_iki_noktasiz():
    """'hakimleri:' yazmayan daire başlığı, altında üyeler varsa başlık sayılır."""
    items = parse_alt_bolum("İstanbul BAM 5. Hukuk Dairesi\nBAŞKAN : Ali Veli 12345\nÜYE : Ayşe Kaya 23456")
    assert [k for k, _ in items] == [BASLIK, UYE, UYE]


def test_ilk_derece_hakimi_ayri_satirlarda():
    items = parse_alt_bolum("İlk kararı veren ilk derece mahkemesi (Bakırköy 5. AHM) hakimi\nTest Hakim\n9999999")
    assert items == [(TEK, "İlk kararı veren ilk derece mahkemesi (Bakırköy 5. AHM) hakimi Test Hakim 9999999")]


def test_degerlendirme_kutulari_atlanir():
    items = parse_alt_bolum("BAŞKAN : Ali Veli 12345\nÇok iyi İyi Orta Zayıf Notsuz\nAÇIKLAMA:")
    assert items == [(UYE, "BAŞKAN\t: Ali Veli  (12345)")]


def _belge(path, ekle):
    d = docx.Document()
    t = d.add_table(rows=4, cols=3)
    for i, row in enumerate([("ESAS NO", ":", "2025/9"), ("MAHKEMESİ", ":", "Bursa 2. AHM"),
                             ("DAİRESİ", ":", "7. Hukuk Dairesi"), ("KONUSU", ":", "Alacak")]):
        for j, v in enumerate(row):
            t.cell(i, j).text = v
    d.add_paragraph("Gövde metni.")
    d.add_paragraph("*******")
    ekle(d)
    d.save(path)


@pytest.mark.parametrize("uzanti", ["docx", "doc"])
def test_tablodaki_hakim_listesi(tmp_path, uzanti):
    """Hakim listesi yıldızın altında tablo içinde: her hücre ayrı (görev | : | isim | sicil)."""
    def ekle(d):
        d.add_paragraph(BASLIK_SATIRI)
        t = d.add_table(rows=3, cols=4)
        for i, (rol, ad, sicil) in enumerate([("BAŞKAN", "Ali Veli", "12345"), ("ÜYE", "Ayşe Kaya", "23456"),
                                              ("ÜYE", "Mehmet Demir", "34567")]):
            t.cell(i, 0).text, t.cell(i, 1).text, t.cell(i, 2).text, t.cell(i, 3).text = rol, ":", ad, sicil
    p = str(tmp_path / "t.docx")
    _belge(p, ekle)
    if uzanti == "doc":
        import shutil
        import subprocess
        if not shutil.which("soffice"):
            pytest.skip("LibreOffice yok")
        subprocess.run(["soffice", "--headless", "--convert-to", "doc:MS Word 97", "--outdir", str(tmp_path), p],
                       capture_output=True, timeout=120)
        p = str(tmp_path / "t.doc")
    case = extract_case(read_document(p))
    assert case.alt_kaynak == "yildiz"
    assert parse_alt_bolum(case.alt_bolum) == BEKLENEN


def test_tabloda_sutun_duzeni(tmp_path):
    """Tek satırlık tablo: görevler bir hücrede alt alta, isimler diğerinde, siciller üçüncüde."""
    def ekle(d):
        d.add_paragraph(BASLIK_SATIRI)
        t = d.add_table(rows=1, cols=3)
        t.cell(0, 0).text = "BAŞKAN\nÜYE\nÜYE"
        t.cell(0, 1).text = "Ali Veli\nAyşe Kaya\nMehmet Demir"
        t.cell(0, 2).text = "12345\n23456\n34567"
    p = str(tmp_path / "s.docx")
    _belge(p, ekle)
    assert parse_alt_bolum(extract_case(read_document(p)).alt_bolum) == BEKLENEN


def test_tetkik_hakimi_imzasi_kesilir(tmp_path):
    def ekle(d):
        d.add_paragraph(BASLIK_SATIRI)
        d.add_paragraph("BAŞKAN : Ali Veli (12345)")
        d.add_paragraph("ÜYE : Ayşe Kaya (23456)")
        d.add_paragraph("ÜYE : Mehmet Demir (34567)")
        d.add_paragraph("")
        d.add_paragraph("Tetkik Hakimi")
        d.add_paragraph("Başka Biri")
    p = str(tmp_path / "k.docx")
    _belge(p, ekle)
    assert parse_alt_bolum(extract_case(read_document(p)).alt_bolum) == BEKLENEN


def test_yildiz_altinda_yalniz_tetkik_varsa_emsal_sonrasi(tmp_path):
    p = str(tmp_path / "e.docx")
    d = docx.Document()
    d.add_paragraph("Gövde. Direnme kararını veren mahkeme hakimi gerekçesinde şunu belirtmiştir.")
    d.add_paragraph("EMSAL KARARLAR:")
    d.add_paragraph("Yargıtay HGK 2019/1 E. 2020/2 K.")
    d.add_paragraph(BASLIK_SATIRI)
    d.add_paragraph("Başkan Ali Veli 12345")
    d.add_paragraph("Üye Ayşe Kaya 23456")
    d.add_paragraph("Üye Mehmet Demir 34567")
    d.add_paragraph("Sonuç ve öneri metni burada yer alır.")
    d.add_paragraph("*******")
    d.add_paragraph("Tetkik Hakimi Biri")
    d.save(p)
    case = extract_case(read_document(p))
    assert case.alt_kaynak == "emsal"
    assert parse_alt_bolum(case.alt_bolum) == BEKLENEN


def test_govde_cumlesi_hakim_sayilmaz(tmp_path):
    p = str(tmp_path / "g.docx")
    d = docx.Document()
    d.add_paragraph("Direnme kararını veren mahkeme hakimi gerekçesinde şunu belirtmiştir.")
    d.save(p)
    assert extract_case(read_document(p)).alt_bolum == ""


def test_metin_kutusu_iki_kez_okunmaz():
    from lxml import etree
    from gundem_doldur.reader import _para_text
    xml = (
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<w:r><w:t>BAŞKAN</w:t><w:tab/><w:t>: Ali Veli</w:t></w:r>'
        '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><w:txbxContent>'
        '<w:p><w:r><w:t>12345</w:t></w:r></w:p></w:txbxContent></w:drawing></mc:Choice>'
        '<mc:Fallback><w:pict><w:txbxContent><w:p><w:r><w:t>12345</w:t></w:r></w:p>'
        '</w:txbxContent></w:pict></mc:Fallback></mc:AlternateContent></w:r>'
        '<w:r><w:delText>silinen</w:delText></w:r></w:p>'
    )
    metin = _para_text(etree.fromstring(xml))
    assert metin.count("12345") == 1
    assert "silinen" not in metin
    assert metin.startswith("BAŞKAN\t: Ali Veli")


@pytest.mark.parametrize("ham", [
    "İlk kararı veren ilk derece mahkemesi (X) hakimi\nHakan Yıldız\nSicil No: 445566\n"
    "İstanbul BAM 3. HD\nBaşkan\tAli Veli\t12345\nÜye\tAyşe Kaya\t23456",
    "Direnme kararını veren daire hakimleri:\nAli Veli (Başkan) 12345\nÜYE\nAyşe Kaya\n23456",
])
def test_duzenleme_tekrar_uygulanabilir(ham):
    from gundem_doldur.extract import alt_bolum_duzenle
    bir = alt_bolum_duzenle(ham)
    assert alt_bolum_duzenle(bir) == bir
    assert parse_alt_bolum(bir) == parse_alt_bolum(ham)
    assert "BAŞKAN\t: Ali Veli  (12345)" in bir
