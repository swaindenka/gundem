import os
import sys

import pytest

KOK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, KOK)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import ornekler  # noqa: E402
from gundem_doldur.bolumler import Dava, Parca, birlestir, bolum_bilgisi  # noqa: E402
from gundem_doldur.extract import BASLIK, TEK, UYE, extract_case, parse_alt_bolum  # noqa: E402
from gundem_doldur.reader import read_document  # noqa: E402


@pytest.mark.parametrize("ad,no", [
    ("2025-123 birinci bölüm.docx", 1),
    ("2025-123 İKİNCİ BÖLÜM.doc", 2),
    ("2025-123 üçüncü bölüm.docx", 3),
    ("2025-123 Dördüncü Bölüm.doc", 4),
    ("2025-123 1. bölüm.docx", 1),
    ("2025-123 2.bölüm.doc", 2),
    ("2025-123_3_bolum.docx", 3),
    ("2025-123 bölüm 4.docx", 4),
    ("2025-123 II. Bölüm.docx", 2),
    ("2025-123 1'inci bölüm.docx", 1),
    ("2025-123 (Birinci Kısım).docx", 1),
])
def test_bolum_bilgisi(ad, no):
    grup, bulunan = bolum_bilgisi(ad)
    assert bulunan == no
    assert grup == "2025 123"


@pytest.mark.parametrize("ad", ["2025-123.docx", "HGK 2025-41 rapor.docx", "bölümler hakkında.docx"])
def test_bolumsuz_dosya(ad):
    assert bolum_bilgisi(ad) == ("", None)


def _parcalar(yollar):
    out = []
    for y in yollar:
        grup, no = bolum_bilgisi(y)
        out.append((grup, Parca(no=no, yol=y, veri=extract_case(read_document(y), kaynak=y))))
    return out


def test_birlestir(tmp_path):
    yollar = ornekler.bolumlu_dava(str(tmp_path))
    parcalar = _parcalar(yollar)
    assert len({g for g, _ in parcalar}) == 1
    # sıralama karışık gelse de sonuç aynı olmalı
    sonuc = birlestir([parcalar[2][1], parcalar[0][1], parcalar[1][1]])
    assert sonuc.kaynak == yollar[0]
    assert sonuc.esas == "2025/123"
    assert sonuc.mahkeme == "Bursa 1. Sulh Hukuk Mahkemesi"
    assert sonuc.daire == "7. Hukuk Dairesi"
    assert sonuc.daire_karar == "05.05.2024 tarihli ve\n2024/10 E. 2024/20 K."
    assert sonuc.konu == "Kira alacağı"
    # hakimler 2. bölümdeki yanıltıcı cümleden değil, 3. bölümün yıldız altından alınmalı
    assert sonuc.alt_kaynak == "yildiz"
    assert [k for k, _ in parse_alt_bolum(sonuc.alt_bolum)] == [TEK, BASLIK, UYE, UYE, UYE, BASLIK, UYE, UYE, UYE]
    assert any("3 bölüm birleştirildi" in u for u in sonuc.uyarilar)
    assert not any("Bulunamayan" in u or "bulunamadı" in u for u in sonuc.uyarilar)


def test_eksik_bolum_uyarisi(tmp_path):
    yollar = ornekler.bolumlu_dava(str(tmp_path))
    parcalar = _parcalar([yollar[0], yollar[2]])
    sonuc = birlestir([p for _, p in parcalar])
    assert any("Eksik bölüm" in u and "2." in u for u in sonuc.uyarilar)


def test_dava_sonradan_bolum_ekleme_elle_duzeltmeyi_korur(tmp_path):
    yollar = ornekler.bolumlu_dava(str(tmp_path))
    parcalar = _parcalar(yollar)
    grup, ilk = parcalar[0]
    dava = Dava.bolumden(grup, ilk)
    assert dava.veri.esas == "2025/123" and dava.veri.alt_bolum == ""
    dava.veri.sira_no = "5"           # program/kullanıcı tarafından verilen sıra no
    dava.veri.konu = "Elle düzeltilmiş konu"
    dava.parca_ekle(parcalar[2][1])
    dava.parca_ekle(parcalar[1][1])
    assert dava.veri.sira_no == "5"
    assert dava.veri.konu == "Elle düzeltilmiş konu"
    assert dava.veri.esas == "2025/123"
    assert dava.veri.alt_kaynak == "yildiz" and "Test9isim" in dava.veri.alt_bolum
    assert dava.etiket() == "2025-123 birinci bölüm.docx  [bölüm 1, 2, 3]"
    assert dava.dosyalar == yollar
