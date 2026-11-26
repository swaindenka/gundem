"""Kullanıcının gerçek belgelerinden görülen hakim yazım biçimleri.

Ekran görüntüsü:
    Direnme Kararını Veren İlk Derece Hâkimi; Mustafa Enes sicil:3111547,
    Esastan Ret Kararını Veren Bölge Adliye Mahkemesi Hâkimleri;
    (altında 3 hakim — önceki sürüm bunları almıyordu)
Mesaj:
    ***********
    Yetkisizlik kararını veren hâkim: isim soyisim-sicil
    Birinci kararı veren hâkim : isim soyisim - sicil
    İkinci ve direnme kararını veren hakim : isim soyisim sicil
"""
import os
import sys

import docx
import pytest

KOK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, KOK)

from gundem_doldur.extract import BASLIK, TEK, UYE, extract_case, parse_alt_bolum  # noqa: E402
from gundem_doldur.reader import read_document  # noqa: E402

ILK = "Direnme Kararını Veren İlk Derece Hâkimi; Mustafa Enes sicil:3111547,"
BAM = "Esastan Ret Kararını Veren Bölge Adliye Mahkemesi Hâkimleri;"
BEKLENEN = [
    (TEK, "Direnme Kararını Veren İlk Derece Hâkimi; Mustafa Enes sicil:3111547"),
    (BASLIK, "Esastan Ret Kararını Veren Bölge Adliye Mahkemesi Hâkimleri:"),
    (UYE, "Ali Veli  (1111111)"),
    (UYE, "Ayşe Kaya  (2222222)"),
    (UYE, "Mehmet Demir  (3333333)"),
]
ROLLU = BEKLENEN[:2] + [
    (UYE, "BAŞKAN\t: Ali Veli  (1111111)"),
    (UYE, "ÜYE\t: Ayşe Kaya  (2222222)"),
    (UYE, "ÜYE\t: Mehmet Demir  (3333333)"),
]

UC_HAKIM_BICIMLERI = {
    "virgullu_tek_satir": ["Ali Veli sicil:1111111, Ayşe Kaya sicil:2222222, Mehmet Demir sicil:3333333."],
    "ve_ile": ["Ali Veli sicil:1111111, Ayşe Kaya sicil:2222222 ve Mehmet Demir sicil:3333333."],
    "ayri_satirlar": ["Ali Veli sicil:1111111,", "Ayşe Kaya sicil:2222222,", "Mehmet Demir sicil:3333333."],
    "sicil_kelimesiz": ["Ali Veli 1111111", "Ayşe Kaya 1111111".replace("1111111", "2222222"),
                        "Mehmet Demir 3333333"],
    "tireli": ["Ali Veli-1111111", "Ayşe Kaya - 2222222", "Mehmet Demir -3333333"],
    "sicil_no": ["Ali Veli Sicil No: 1111111", "Ayşe Kaya Sicil No: 2222222", "Mehmet Demir Sicil No: 3333333"],
    "isim_sicil_ayri": ["Ali Veli", "sicil:1111111", "Ayşe Kaya", "sicil:2222222", "Mehmet Demir", "sicil:3333333"],
}


@pytest.mark.parametrize("ad", sorted(UC_HAKIM_BICIMLERI))
def test_bam_hakimleri_alt_satirlarda(ad):
    assert parse_alt_bolum("\n".join([ILK, BAM] + UC_HAKIM_BICIMLERI[ad])) == BEKLENEN


def test_bam_hakimleri_baslikla_ayni_satirda():
    satir = BAM + " Ali Veli sicil:1111111, Ayşe Kaya sicil:2222222, Mehmet Demir sicil:3333333."
    assert parse_alt_bolum("\n".join([ILK, satir])) == BEKLENEN


@pytest.mark.parametrize("satirlar", [
    ["Başkan Ali Veli sicil:1111111, Üye Ayşe Kaya sicil:2222222, Üye Mehmet Demir sicil:3333333."],
    ["Başkan: Ali Veli sicil:1111111,", "Üye: Ayşe Kaya sicil:2222222,", "Üye: Mehmet Demir sicil:3333333."],
    ["Ali Veli (Başkan) sicil:1111111,", "Ayşe Kaya (Üye) sicil:2222222,", "Mehmet Demir (Üye) sicil:3333333."],
])
def test_bam_hakimleri_gorevli(satirlar):
    assert parse_alt_bolum("\n".join([ILK, BAM] + satirlar)) == ROLLU


def test_mesajdaki_bicim():
    metin = "\n".join([
        "Yetkisizlik kararını veren hâkim: Ali Veli-1111111",
        "Birinci kararı veren hâkim : Ayşe Kaya - 2222222",
        "İkinci ve direnme kararını veren hakim : Mehmet Demir 3333333",
    ])
    assert parse_alt_bolum(metin) == [
        (TEK, "Yetkisizlik kararını veren hâkim: Ali Veli-1111111"),
        (TEK, "Birinci kararı veren hâkim : Ayşe Kaya - 2222222"),
        (TEK, "İkinci ve direnme kararını veren hakim : Mehmet Demir 3333333"),
    ]


def _belge(path, govde_sonu, yildiz_sonrasi, tablo=None):
    d = docx.Document()
    t = d.add_table(rows=4, cols=3)
    for i, row in enumerate([("ESAS NO", ":", "2025/9"), ("MAHKEMESİ", ":", "Bursa 2. AHM"),
                             ("DAİRESİ", ":", "7. Hukuk Dairesi"), ("KONUSU", ":", "Alacak")]):
        for j, v in enumerate(row):
            t.cell(i, j).text = v
    for i in range(150):
        d.add_paragraph("Gövde paragrafı %d. Direnme kararını veren hâkim, gerekçesinde önceki kararda "
                        "direnilmesine karar vermiştir." % i)
    d.add_paragraph("EMSAL KARARLAR:")
    d.add_paragraph("Yargıtay HGK 2019/1 E. 2020/2 K.")
    for l in govde_sonu:
        d.add_paragraph(l)
    if tablo:
        tablo(d)
    d.add_paragraph().add_run("***************").bold = True
    for l in yildiz_sonrasi:
        d.add_paragraph(l)
    d.save(path)
    return path


def test_ekran_goruntusundeki_dosya(tmp_path):
    """Hakimler yıldızdan önce, yıldızın altında başka metin: 3 BAM hakimi de alınmalı."""
    p = _belge(str(tmp_path / "a.docx"),
               [ILK, BAM, "Ali Veli sicil:1111111, Ayşe Kaya sicil:2222222, Mehmet Demir sicil:3333333."],
               ["Saygıyla arz olunur."])
    case = extract_case(read_document(p))
    assert parse_alt_bolum(case.alt_bolum) == BEKLENEN
    assert "Gövde" not in case.alt_bolum


def test_mesajdaki_dosya_yildiz_altinda(tmp_path):
    p = _belge(str(tmp_path / "b.docx"), [], [
        "Yetkisizlik kararını veren hâkim: Ali Veli-1111111",
        "Birinci kararı veren hâkim : Ayşe Kaya - 2222222",
        "İkinci ve direnme kararını veren hakim : Mehmet Demir 3333333",
    ])
    case = extract_case(read_document(p))
    assert case.alt_kaynak == "yildiz"
    assert [k for k, _ in parse_alt_bolum(case.alt_bolum)] == [TEK, TEK, TEK]
    assert "Mehmet Demir 3333333" in case.alt_bolum
    assert not any("tanınamadı" in u for u in case.uyarilar)


def test_imza_tablosu_hakim_basina_hucre(tmp_path):
    """BAM hakimleri tabloda yan yana: her hücrede 'görev / isim / sicil' alt alta."""
    def tablo(d):
        t = d.add_table(rows=1, cols=3)
        t.cell(0, 0).text = "Başkan\nAli Veli\n1111111"
        t.cell(0, 1).text = "Üye\nAyşe Kaya\n2222222"
        t.cell(0, 2).text = "Üye\nMehmet Demir\n3333333"
    p = _belge(str(tmp_path / "c.docx"), [], [ILK, BAM])
    # tabloyu yıldızın altına, başlığın arkasına ekle
    d = docx.Document(p)
    tablo(d)
    d.save(p)
    case = extract_case(read_document(p))
    assert parse_alt_bolum(case.alt_bolum) == ROLLU


def test_imza_tablosu_satir_satir(tmp_path):
    """Görevler bir satırda, isimler altındaki satırda, siciller en altta (3x3 tablo)."""
    p = _belge(str(tmp_path / "d.docx"), [], [ILK, BAM])
    d = docx.Document(p)
    t = d.add_table(rows=3, cols=3)
    for i, satir in enumerate([("Başkan", "Üye", "Üye"), ("Ali Veli", "Ayşe Kaya", "Mehmet Demir"),
                               ("1111111", "2222222", "3333333")]):
        for j, v in enumerate(satir):
            t.cell(i, j).text = v
    d.save(p)
    case = extract_case(read_document(p))
    assert parse_alt_bolum(case.alt_bolum) == ROLLU


def test_isimsiz_hakim_cumlesi_sayilmaz(tmp_path):
    """'Direnme kararını veren hâkim, gerekçesinde ...' gibi gövde cümleleri hakim listesi değildir."""
    d = docx.Document()
    for i in range(20):
        d.add_paragraph("Direnme kararını veren hâkim, gerekçesinde önceki kararda direnilmesine karar vermiştir.")
    p = str(tmp_path / "e.docx")
    d.save(p)
    assert extract_case(read_document(p)).alt_bolum == ""


HEYET = [
    (BASLIK, "Direnme kararı veren Bölge Adliye Mah. Heyeti:"),
    (UYE, "BAŞKAN\t: Musa Enes Çakmak  (5555)"),
    (UYE, "ÜYE\t: Ali Veli  (6666)"),
    (UYE, "ÜYE\t: Ayşe Kaya  (7777)"),
]


@pytest.mark.parametrize("satirlar", [
    # kullanıcının bildirdiği: başlık ile Başkan aynı satırda, üyeler altta
    ["Direnme kararı veren Bölge Adliye Mah. Heyeti : Başkan : Musa Enes Çakmak (5555)",
     "Üye : Ali Veli (6666)", "Üye : Ayşe Kaya (7777)"],
    # hepsi tek satırda, ayraçsız
    ["Direnme kararı veren Bölge Adliye Mah. Heyeti : Başkan : Musa Enes Çakmak (5555) "
     "Üye : Ali Veli (6666) Üye : Ayşe Kaya (7777)"],
    # hepsi tek satırda, virgüllü, tireli
    ["Direnme kararı veren Bölge Adliye Mah. Heyeti; Başkan - Musa Enes Çakmak - 5555, "
     "Üye - Ali Veli - 6666, Üye - Ayşe Kaya - 7777."],
    # başlık ayrı satırda, üyeler tek satırda
    ["Direnme kararı veren Bölge Adliye Mah. Heyeti :",
     "Başkan : Musa Enes Çakmak (5555), Üye : Ali Veli (6666), Üye : Ayşe Kaya (7777)"],
])
def test_heyet_basligi_ile_baskan_ayni_satirda(satirlar):
    assert parse_alt_bolum("\n".join(satirlar)) == HEYET


def test_heyet_belgede_yildizdan_once(tmp_path):
    """Aynı biçim yıldızdan önce (gövdede arama) da bulunmalı."""
    p = _belge(str(tmp_path / "h.docx"),
               ["Direnme kararı veren Bölge Adliye Mah. Heyeti : Başkan : Musa Enes Çakmak (5555)",
                "Üye : Ali Veli (6666)", "Üye : Ayşe Kaya (7777)"],
               ["Saygıyla arz olunur."])
    case = extract_case(read_document(p))
    assert parse_alt_bolum(case.alt_bolum) == HEYET


@pytest.mark.parametrize("satir", [
    "Direnme Kararını Veren İlk Derece Hâkimi; Mustafa Enes sicil:3111547",
    "Yetkisizlik kararını veren hâkim: Ali Veli-1111111",
    "İlk kararı veren ilk derece mahkemesi (Bakırköy 55. Asliye Hukuk Mahkemesi) hakimi Ali Veli 9999999",
])
def test_tek_hakim_satiri_bolunmez(satir):
    assert parse_alt_bolum(satir) == [(TEK, satir)]
