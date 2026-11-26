# Gündem Doldurucu

Kaynak Word belgelerindeki bilgileri okuyup **gündem tablosunu** (`sablon/BASLIK_TEST.docx`) otomatik dolduran Windows programı.

## İndirme

[Releases](../../releases/latest) sayfasından **GundemDoldurucu.exe** dosyasını indirin, çift tıklayıp çalıştırın. Kurulum gerekmez.

> Windows "Bilinmeyen yayımcı" uyarısı verirse **Ek bilgi → Yine de çalıştır** deyin (program imzasız olduğu için).

## Kullanım

1. **Dosya Ekle...** ile bir veya birden çok kaynak belge seçin (`.docx` ya da Word 97-2003 `.doc`).
2. Program her belgeden şunları okur:
   - **Üstteki tablodan:** Esas No, Mahkeme, Daire, Daire Kararı (tarih ve E./K. numarası), Konusu
   - **`*******` işaretinin altından:** ilk kararı veren hakim, BAM / direnme kararını veren daire hakimleri (BAŞKAN / ÜYE).
     Yıldızın altında hakim bilgisi yoksa **"Emsal" bölümünden sonrasına**, orada da yoksa belgenin geri
     kalanına bakılır; bu durumda program "kontrol edin" uyarısı gösterir.
3. Sağdaki alanlarda okunan bilgileri kontrol edin; gerekirse düzeltin.
4. Üstten **gündem tarihini** (gün / ay / yıl) seçin. Haftanın günü otomatik yazılır
   (ör. `2 ARALIK 2026 ÇARŞAMBA GÜNÜ GÜNDEMİ`).
5. **Tetkik hakimi** kutusu isteğe bağlıdır; boş bırakılırsa gündemde boş kalır.
6. **Gündemi Oluştur** ile `.docx` olarak kaydedin. Birden çok dosya eklediyseniz her biri
   ayrı sayfada, sıra numarasıyla yer alır (sırayı ▲ ▼ ile değiştirebilirsiniz). Her sayfanın
   başında tarih ve tetkik hakimi başlığı tekrarlanır.

### Bölümlere ayrılmış dosyalar

Uzun dosyalar parçalara ayrılıp dosya adında bölüm belirtildiyse program bunları **tek dava** olarak
birleştirir (tek sıra numarası, tek sayfa). Tanınan adlandırmalar:

`birinci bölüm`, `İKİNCİ BÖLÜM`, `üçüncü kısım`, `1. bölüm`, `2.bölüm`, `1'inci bölüm`, `bölüm 3`, `II. bölüm`, `part 2`

Bölüm ifadesi çıkarıldığında dosya adının geri kalanı aynı olan dosyalar birleştirilir
(ör. `2025-123 birinci bölüm.docx` + `2025-123 ikinci bölüm.doc`).

- Esas / Mahkeme / Daire / Konu: numarası en küçük bölümden (genellikle 1. bölüm)
- Hakimler: `*******` altında hakim bilgisi olan en son bölümden
- Bölümler aynı anda ya da sonradan eklenebilir; sonradan eklenen bölüm mevcut davaya katılır,
  elle düzelttiğiniz alanlar korunur.
- Araya bir bölüm eksik kaldıysa (ör. 1. ve 3. eklendi, 2. yok) program uyarır.

### Tanınan hakim yazımları

Program hakim satırlarını şablondaki biçime getirir: `BAŞKAN : Ad Soyad  (sicil)`. Tanınan yazımlar:

- `BAŞKAN : Ad Soyad (12345)`, `Başkan Ad Soyad 12345`, `1- Üye: Ad Soyad 12345`, `• ÜYE - Ad Soyad - 12345`
- `Üye : Ad Soyad Sicil No: 12345`
- Görev isimden sonra: `Ad Soyad (Başkan) 12345`, `Ad Soyad - Üye - 12345`
- Görev, isim ve sicil ayrı satırlarda ya da **tablo hücrelerinde** (`BAŞKAN | : | Ad Soyad | 12345`)
- Görevler bir sütunda alt alta, isimler ve siciller yan sütunlarda
- Görevler: BAŞKAN, BAŞKAN V., BAŞKANVEKİLİ, KIDEMLİ ÜYE, ÜYE, HAKİM

Yıldızın altındaki "Tetkik Hakimi ..." imzası hakim listesine alınmaz. Hakimler kutusunda
düzenlenmiş hali görünür; görev / isim / sicil doğru ayrılmamışsa kaydetmeden önce düzeltebilirsiniz.

### Hakimler kutusu nasıl yorumlanır?

Kutudaki her satır tabloda ayrı bir satır olur:

| Satır | Tabloda |
|---|---|
| `BAŞKAN : ...` / `ÜYE : ...` ile başlayan | Değerlendirme kutulu (Çok iyi / İyi / ... / AÇIKLAMA) hakim satırı |
| `:` ile biten (ör. `... Hukuk Dairesi hakimleri:`) | Tam genişlikte kalın başlık satırı |
| Diğerleri (ör. `İlk kararı veren ilk derece mahkemesi (...) hakimi ...`) | Değerlendirme kutulu tek satır |

## .doc dosyaları

`.doc` (Word 97-2003) dosyaları programın kendi okuyucusuyla okunur. Okunamazsa bilgisayarda
Microsoft Word veya LibreOffice varsa onunla otomatik dönüştürülerek tekrar denenir.

## Şablon

Program, gündem şablonunu içinde taşır. Farklı bir şablon kullanmak için **Şablon Seç...**
düğmesiyle aynı yapıdaki başka bir `.docx` seçilebilir.

## Geliştirme

```bash
pip install -r requirements.txt pytest
python -m pytest tests
python main.py
```

`main` dalına yapılan her birleştirmede GitHub Actions `.exe` dosyasını derler; `gundem_doldur/__init__.py`
içindeki sürüm numarası yeniyse otomatik olarak yeni bir release yayınlar.
