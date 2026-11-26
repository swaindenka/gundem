import datetime as dt
import os
import sys

import docx
import pytest

KOK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, KOK)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import ornekler  # noqa: E402
from gundem_doldur.doc_binary import read_doc  # noqa: E402
from gundem_doldur.extract import BASLIK, TEK, UYE, extract_case, parse_alt_bolum, split_daire  # noqa: E402
from gundem_doldur.reader import read_document  # noqa: E402
from gundem_doldur.writer import build_gundem, gundem_basligi  # noqa: E402

VERI = os.path.join(KOK, "tests", "veri")
SABLON = os.path.join(KOK, "sablon", "BASLIK_TEST.docx")


def _kontrol_baslik_duzeni(case):
    assert case.esas == "2025/1234"
    assert case.mahkeme == "Bakırköy 55. Asliye Hukuk Mahkemesi"
    assert case.daire == "3. Hukuk Dairesi"
    assert case.daire_karar == "13.12.2024 tarihli ve\n2024/9000 E. 2024/60000 K."
    assert case.konu == "Alacak davası"
    items = parse_alt_bolum(case.alt_bolum)
    assert [k for k, _ in items] == [TEK, BASLIK, UYE, UYE, UYE, BASLIK, UYE, UYE, UYE]
    assert items[2][1] == "BAŞKAN\t: Test1isim Test1Soyisim (999898)"


def test_docx_baslik_duzeni(tmp_path):
    p = str(tmp_path / "a.docx")
    ornekler.baslik_duzeni(p)
    _kontrol_baslik_duzeni(extract_case(read_document(p)))


def test_doc_baslik_duzeni():
    _kontrol_baslik_duzeni(extract_case(read_document(os.path.join(VERI, "baslik_duzeni.doc"))))


@pytest.mark.parametrize("dosya", ["etiket_duzeni.doc", None])
def test_etiket_duzeni(tmp_path, dosya):
    if dosya:
        p = os.path.join(VERI, dosya)
    else:
        p = str(tmp_path / "b.docx")
        ornekler.etiket_duzeni(p)
    case = extract_case(read_document(p))
    assert case.esas == "2025/77"
    assert case.mahkeme == "Ankara 3. İş Mahkemesi"
    assert case.daire == "9. Hukuk Dairesi"
    assert case.daire_karar.startswith("01.02.2024 tarihli ve\n2023/1 E.")
    assert case.konu == "İşçilik alacağı"
    items = parse_alt_bolum(case.alt_bolum)
    assert items[0] == (TEK, "İlk kararı veren ilk derece mahkemesi (Ankara 3. İş Mahkemesi) hakimi Ali Veli 123456")
    assert items[2] == (UYE, "BAŞKAN\t: Ayşe Yılmaz (111)")
    assert items[3] == (UYE, "ÜYE\t: Mehmet Kaya (222)")


def test_doc_karmasik_tablo():
    """İç içe tablo ve birleştirilmiş hücreler içeren .doc doğru okunmalı."""
    doc = read_doc(os.path.join(VERI, "sablon_karmasik.doc"))
    t = doc.tables[1]
    assert [c.text for c in t.rows[0]][:4] == ["SIRA NO", "ESAS NO-MAHKEME", "DAİRE KARARI", "KONUSU"]
    assert [c.span for c in t.rows[0]][:4] == [1, 2, 3, 1]
    assert t.rows[2][3].vcontinue
    assert t.rows[3][0].text.strip().startswith("İlk kararı veren")


def test_yildiz_yoksa_hakimler_metinden_bulunur(tmp_path):
    p = str(tmp_path / "c.docx")
    ornekler.baslik_duzeni(p, yildiz="")
    case = extract_case(read_document(p))
    assert len(parse_alt_bolum(case.alt_bolum)) == 9
    assert "gövde" not in case.alt_bolum.lower()
    assert any("kontrol edin" in u.lower() for u in case.uyarilar)


def test_hakim_yoksa_uyari(tmp_path):
    p = str(tmp_path / "d.docx")
    ornekler.baslik_duzeni(p, alt=[], yildiz="")
    case = extract_case(read_document(p))
    assert case.alt_bolum == ""
    assert any("bulunamadı" in u for u in case.uyarilar)


def test_split_daire():
    assert split_daire("3. HD'nin 13.12.2024 tarihli ve 2024/1 E. 2024/2 K.") == (
        "3. HD", "13.12.2024 tarihli ve\n2024/1 E. 2024/2 K.")
    assert split_daire("3. Hukuk Dairesi") == ("3. Hukuk Dairesi", "")


def test_gundem_basligi():
    assert gundem_basligi(dt.date(2026, 12, 2)) == "2 ARALIK 2026 ÇARŞAMBA GÜNÜ GÜNDEMİ"


def test_gundem_olustur(tmp_path):
    a = str(tmp_path / "a.docx")
    ornekler.baslik_duzeni(a)
    c1 = extract_case(read_document(a))
    c1.sira_no = "1"
    c2 = extract_case(read_document(os.path.join(VERI, "etiket_duzeni.doc")))
    c2.sira_no = "2"
    out = str(tmp_path / "gundem.docx")
    build_gundem(SABLON, [c1, c2], dt.date(2026, 12, 2), out, "")

    d = docx.Document(out)
    assert len(d.tables) == 4
    assert d.tables[0].rows[0].cells[0].text.startswith("2 ARALIK 2026 ÇARŞAMBA GÜNÜ GÜNDEMİ")
    assert d.tables[0].rows[1].cells[0].text.strip() == "TETKİK  HAKİMİ :"
    t1 = d.tables[1]
    assert t1.rows[1].cells[0].text == "1"
    assert t1.rows[1].cells[1].text == "2025/1234"
    assert t1.rows[2].cells[1].text == "Bakırköy 55. Asliye Hukuk Mahkemesi"
    assert t1.rows[2].cells[3].text == "13.12.2024 tarihli ve\n2024/9000 E. 2024/60000 K."
    metin = "\n".join(c.text for r in t1.rows for c in r.cells)
    assert "Test9isim" in metin and "Direnme kararını veren" in metin
    t2 = d.tables[3]
    assert t2.rows[1].cells[0].text == "2"
    assert t2.rows[1].cells[1].text == "2025/77"
    # yer tutucu metinler kalmamalı
    tum = "\n".join(c.text for t in d.tables for r in t.rows for c in r.cells)
    assert "Test4isim" not in tum
    hucreler = {c.text for t in d.tables for r in t.rows for c in r.cells}
    assert not hucreler & {"ESAS", "MAHKEME", "DAİRE", "KONU"}
    oylama = [p.text for p in d.paragraphs if "OYLAMA" in p.text]
    assert len(oylama) == 4


def test_tetkik_hakimi_yazilir(tmp_path):
    a = str(tmp_path / "a.docx")
    ornekler.baslik_duzeni(a)
    out = str(tmp_path / "g.docx")
    build_gundem(SABLON, [extract_case(read_document(a))], dt.date(2026, 12, 2), out, "Ad Soyad")
    d = docx.Document(out)
    assert d.tables[0].rows[1].cells[0].text.strip() == "TETKİK  HAKİMİ : Ad Soyad"


@pytest.mark.parametrize("yildiz_sonrasi", [("Tetkik Hakimi",), ()])
def test_hakimler_emsalden_sonra(tmp_path, yildiz_sonrasi):
    p = str(tmp_path / "e.docx")
    ornekler.emsal_duzeni(p, yildiz_sonrasi)
    case = extract_case(read_document(p))
    items = parse_alt_bolum(case.alt_bolum)
    assert [k for k, _ in items] == [TEK, BASLIK, UYE, UYE, UYE, BASLIK, UYE, UYE, UYE]
    assert "Sonuç" not in case.alt_bolum and "Yargıtay HGK" not in case.alt_bolum
    assert any("Emsal" in u for u in case.uyarilar)


def test_her_sayfada_baslik(tmp_path):
    a = str(tmp_path / "a.docx")
    ornekler.baslik_duzeni(a)
    cases = []
    for i in range(3):
        c = extract_case(read_document(a))
        c.sira_no = str(i + 1)
        cases.append(c)
    out = str(tmp_path / "g.docx")
    build_gundem(SABLON, cases, dt.date(2026, 12, 2), out, "Ad Soyad")
    d = docx.Document(out)
    basliklar = [t for t in d.tables if "GÜNDEMİ" in t.rows[0].cells[0].text]
    assert len(basliklar) == 3
    for t in basliklar:
        assert t.rows[0].cells[0].text.startswith("2 ARALIK 2026 ÇARŞAMBA GÜNÜ GÜNDEMİ")
        assert t.rows[1].cells[0].text.strip() == "TETKİK  HAKİMİ : Ad Soyad"
    # sıralama: başlık, dava, başlık, dava, başlık, dava
    sira = ["B" if "GÜNDEMİ" in t.rows[0].cells[0].text else "D" for t in d.tables]
    assert sira == ["B", "D", "B", "D", "B", "D"]
    # sayfa sonu sayısı = dosya sayısı - 1
    assert d.element.body.xml.count('w:type="page"') == 2
