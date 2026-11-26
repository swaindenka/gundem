"""Uzun belgelerde hakim listesinin ortadan değil, sondan alınması."""
import os
import sys

import docx
import pytest

KOK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, KOK)

from gundem_doldur.extract import UYE, extract_case, parse_alt_bolum  # noqa: E402
from gundem_doldur.reader import read_document  # noqa: E402

GERCEK = [
    "İlk kararı veren ilk derece mahkemesi (Bursa 2. AHM) hakimi Hakan Yıldız 445566",
    "Direnme kararını veren Bursa BAM 3. Hukuk Dairesi hakimleri:",
    "BAŞKAN : Ali Veli (12345)",
    "ÜYE : Ayşe Kaya (23456)",
    "ÜYE : Mehmet Demir (34567)",
]
# Metnin ortasında alıntılanan eski bir karar (yanlış aday)
ALINTI = [
    "Bozma öncesi kararı veren Bölge Adliye Mahkemesi 9. Hukuk Dairesi hakimleri:",
    "BAŞKAN : Eski Başkan (99999)",
    "ÜYE : Eski Üye (88888)",
]


def _govde(d, n, etiket):
    for i in range(n):
        d.add_paragraph("%s gövde paragrafı %d. Davacı vekili dava dilekçesinde alacağın tahsilini "
                        "istemiş, mahkemece davanın kabulüne karar verilmiştir." % (etiket, i))


def _uret(path, orta_ayrac, orta, son_ayrac, son, emsal_basta=False):
    d = docx.Document()
    t = d.add_table(rows=4, cols=3)
    for i, row in enumerate([("ESAS NO", ":", "2025/9"), ("MAHKEMESİ", ":", "Bursa 2. AHM"),
                             ("DAİRESİ", ":", "7. Hukuk Dairesi"), ("KONUSU", ":", "Alacak")]):
        for j, v in enumerate(row):
            t.cell(i, j).text = v
    if emsal_basta:
        d.add_paragraph("EMSAL KARARLAR:")
        d.add_paragraph("Yargıtay HGK 2019/1 E. 2020/2 K.")
    _govde(d, 300, "Birinci")          # ~45 sayfa
    if orta_ayrac:
        d.add_paragraph(orta_ayrac)
    for l in orta:
        d.add_paragraph(l)
    _govde(d, 200, "İkinci")           # ~30 sayfa
    if son_ayrac:
        d.add_paragraph(son_ayrac)
    for l in son:
        d.add_paragraph(l)
    d.save(path)


def _uyeler(case):
    return [t for k, t in parse_alt_bolum(case.alt_bolum) if k == UYE]


SENARYOLAR = {
    # ortada kısa '***' ayracı + alıntı, sonda ******* + gerçek liste
    "kisa_ayrac": ("***", ALINTI, "*******", GERCEK),
    # ortada ******* + alıntı, sonda yıldızsız gerçek liste
    "sonda_yildiz_yok": ("*******", ALINTI, "", GERCEK),
    # ortada ve sonda *******
    "iki_yildiz": ("*******", ALINTI, "*******", GERCEK),
    # yıldız hiç yok: ortada alıntı, sonda gerçek liste
    "yildiz_hic_yok": ("", ALINTI, "", GERCEK),
    # sondaki listenin ardından tetkik hakimi imzası
    "tetkik_imzasi": ("*******", ALINTI, "*******", GERCEK + ["", "Tetkik Hakimi", "Rapor Yazan Hakim"]),
}


@pytest.mark.parametrize("ad", sorted(SENARYOLAR))
def test_hakimler_sondan_alinir(tmp_path, ad):
    orta_ayrac, orta, son_ayrac, son = SENARYOLAR[ad]
    p = str(tmp_path / (ad + ".docx"))
    _uret(p, orta_ayrac, orta, son_ayrac, son)
    case = extract_case(read_document(p))
    assert _uyeler(case) == ["BAŞKAN\t: Ali Veli  (12345)", "ÜYE\t: Ayşe Kaya  (23456)",
                             "ÜYE\t: Mehmet Demir  (34567)"], case.alt_bolum[:300]
    assert "Eski" not in case.alt_bolum
    assert "gövde paragrafı" not in case.alt_bolum
    assert "Hakan Yıldız" in case.alt_bolum
    assert "Rapor Yazan" not in case.alt_bolum


def test_emsal_basta_alinti_ortada(tmp_path):
    """Emsal listesi başta; Emsal'den sonraki ilk liste değil, sondaki liste alınmalı."""
    p = str(tmp_path / "e.docx")
    _uret(p, "", ALINTI, "", GERCEK, emsal_basta=True)
    case = extract_case(read_document(p))
    assert "Eski" not in case.alt_bolum
    assert "Ali Veli" in case.alt_bolum


def test_ilk_derece_ile_bam_arasinda_bos_satirlar(tmp_path):
    """Liste içindeki boş/ara satırlar listeyi bölmemeli (ilk derece hakimi kaybolmamalı)."""
    p = str(tmp_path / "b.docx")
    son = [GERCEK[0], "", "Sayın Genel Kurul", ""] + GERCEK[1:]
    _uret(p, "", [], "", son)
    case = extract_case(read_document(p))
    assert "Hakan Yıldız" in case.alt_bolum and "Mehmet Demir" in case.alt_bolum


def test_bolumlerde_son_bolum_oncelikli(tmp_path):
    from gundem_doldur.bolumler import Parca, birlestir
    from gundem_doldur.extract import CaseData
    a = CaseData(esas="1", alt_bolum="BAŞKAN\t: Eski Başkan  (99999)", alt_kaynak="yildiz")
    b = CaseData(alt_bolum="BAŞKAN\t: Ali Veli  (12345)", alt_kaynak="genel")
    sonuc = birlestir([Parca(no=1, yol="x 1. bölüm.docx", veri=a), Parca(no=2, yol="x 2. bölüm.docx", veri=b)])
    assert "Ali Veli" in sonuc.alt_bolum


def test_ortadan_alinirsa_uyari(tmp_path):
    """Sonda hiçbir hakim bilgisi yoksa ortadaki liste alınır ama kullanıcı uyarılır."""
    p = str(tmp_path / "u.docx")
    _uret(p, "*******", GERCEK, "", [])
    case = extract_case(read_document(p))
    assert "Ali Veli" in case.alt_bolum
    assert any("ortalarından" in u for u in case.uyarilar)


def test_sondan_alinirsa_uyari_yok(tmp_path):
    p = str(tmp_path / "s.docx")
    _uret(p, "*******", ALINTI, "*******", GERCEK)
    case = extract_case(read_document(p))
    assert not any("ortalarından" in u for u in case.uyarilar)
