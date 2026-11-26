"""Testler için örnek kaynak belgeler üretir (gerçek veri içermez)."""
from __future__ import annotations

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT

ALT_BOLUM = [
    "İlk kararı veren ilk derece mahkemesi (Bakırköy 55. Asliye Hukuk Mahkemesi) hakimi  Testisim TestSoyisim  9999999",
    "İlk BAM kararını veren İstanbul Bölge Adliye Mahkemesi 555. Hukuk Dairesi hakimleri:",
    "BAŞKAN\t: Test1isim Test1Soyisim  (999898)",
    "ÜYE\t: Test2isim Test2Soyisim  (4564123)",
    "ÜYE\t: Test3isim Test3Soyisim  (34524524)",
    "Direnme kararını veren İstanbul Bölge Adliye Mahkemesi 59741. Hukuk Dairesi hakimleri:",
    "BAŞKAN\t: Test7isim Test7Soyisim  (746545641)",
    "ÜYE\t: Test8isim Test8Soyisim  (4654112314)",
    "ÜYE\t: Test9isim Test9Soyisim  (86446515)",
]


def baslik_duzeni(path: str, alt=ALT_BOLUM, yildiz="*******"):
    """Şablondakiyle aynı düzende üst tablo (SIRA NO sütunu olmadan)."""
    d = docx.Document()
    t = d.add_table(rows=4, cols=3)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.cell(0, 0).text = "ESAS NO-MAHKEME"
    t.cell(0, 1).text = "DAİRE KARARI"
    t.cell(0, 2).text = "KONUSU"
    t.cell(1, 0).text = "2025/1234"
    t.cell(1, 1).text = "3. Hukuk Dairesi"
    konu = t.cell(1, 2).merge(t.cell(3, 2))
    konu.text = "Alacak davası"
    t.cell(2, 0).text = "Bakırköy 55. Asliye Hukuk Mahkemesi"
    t.cell(2, 1).text = "13.12.2024 tarihli ve 2024/9000 E. 2024/60000 K."
    d.add_paragraph("")
    d.add_paragraph("Raporun gövde metni burada yer alır. Davacı vekili dava dilekçesinde ...")
    d.add_paragraph("Yerel mahkemece verilen karar ... bozulmuştur.")
    d.add_paragraph(yildiz)
    for line in alt:
        d.add_paragraph(line)
    d.save(path)


def etiket_duzeni(path: str):
    """'ESAS NO | : | değer' biçiminde 4x3 tablo."""
    d = docx.Document()
    t = d.add_table(rows=4, cols=3)
    t.style = "Table Grid"
    rows = [
        ("ESAS NO", ":", "2025/77"),
        ("MAHKEMESİ", ":", "Ankara 3. İş Mahkemesi"),
        ("DAİRESİ", ":", "9. Hukuk Dairesinin 01.02.2024 tarihli ve 2023/1 E. 2024/2 K. sayılı kararı"),
        ("KONUSU", ":", "İşçilik alacağı"),
    ]
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            t.cell(i, j).text = v
    d.add_paragraph("Gövde")
    d.add_paragraph("* * * * * * *")
    d.add_paragraph("İlk kararı veren ilk derece mahkemesi (Ankara 3. İş Mahkemesi) hakimi")
    d.add_paragraph("Ali Veli 123456")
    d.add_paragraph("Direnme kararını veren Ankara Bölge Adliye Mahkemesi 5. Hukuk Dairesi hakimleri:")
    d.add_paragraph("Başkan : Ayşe Yılmaz (111)")
    d.add_paragraph("Üye : Mehmet Kaya (222)")
    d.add_paragraph("Üye : Fatma Demir (333)")
    d.save(path)


def emsal_duzeni(path: str, yildiz_sonrasi=("Tetkik Hakimi",)):
    """Hakim bilgileri yıldızdan önce, 'Emsal' bölümünden sonra yazılmış belge."""
    d = docx.Document()
    t = d.add_table(rows=4, cols=3)
    for i, row in enumerate([("ESAS NO", ":", "2025/5"), ("MAHKEMESİ", ":", "İzmir 2. Asliye Hukuk Mahkemesi"),
                             ("DAİRESİ", ":", "4. Hukuk Dairesi"), ("KONUSU", ":", "Tazminat")]):
        for j, v in enumerate(row):
            t.cell(i, j).text = v
    d.add_paragraph("Gövde metni. Direnme kararı veren mahkeme hakimi gerekçesinde ... belirtmiştir.")
    d.add_paragraph("EMSAL KARARLAR:")
    d.add_paragraph("Yargıtay HGK 2019/1 E. 2020/2 K.")
    d.add_paragraph("Yargıtay 4. HD 2018/3 E. 2018/4 K.")
    d.add_paragraph("")
    for line in ALT_BOLUM:
        d.add_paragraph(line)
    d.add_paragraph("Sonuç ve öneri metni.")
    d.add_paragraph("*******")
    for line in yildiz_sonrasi:
        d.add_paragraph(line)
    d.save(path)


def bolumlu_dava(klasor: str, adlar=("2025-123 birinci bölüm.docx", "2025-123 ikinci bölüm.docx",
                                    "2025-123 üçüncü bölüm.docx")):
    """Uzun bir dosyanın üç bölüme ayrılmış hali: tablo 1. bölümde, hakimler son bölümde."""
    import os
    yollar = [os.path.join(klasor, a) for a in adlar]
    # 1. bölüm: üst tablo + gövde
    d = docx.Document()
    t = d.add_table(rows=4, cols=3)
    t.cell(0, 0).text, t.cell(0, 1).text, t.cell(0, 2).text = "ESAS NO-MAHKEME", "DAİRE KARARI", "KONUSU"
    t.cell(1, 0).text, t.cell(1, 1).text, t.cell(1, 2).text = "2025/123", "7. Hukuk Dairesi", "Kira alacağı"
    t.cell(2, 0).text = "Bursa 1. Sulh Hukuk Mahkemesi"
    t.cell(2, 1).text = "05.05.2024 tarihli ve 2024/10 E. 2024/20 K."
    d.add_paragraph("Birinci bölüm gövde metni.")
    d.save(yollar[0])
    # 2. bölüm: yalnızca gövde (yanıltıcı 'kararı veren ... hakimi' ifadesi içerir)
    d = docx.Document()
    d.add_paragraph("İkinci bölüm. Direnme kararını veren mahkeme hakimi gerekçesinde şunu belirtmiştir.")
    d.add_paragraph("Devam eden açıklamalar " * 20)
    d.save(yollar[1])
    # 3. bölüm: gövde + yıldız + hakimler
    d = docx.Document()
    d.add_paragraph("Üçüncü bölüm, sonuç kısmı.")
    d.add_paragraph("*******")
    for line in ALT_BOLUM:
        d.add_paragraph(line)
    d.save(yollar[2])
    return yollar
