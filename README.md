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
