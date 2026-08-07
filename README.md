# Banka Komisyon Karşılaştırma

Garanti BBVA, İş Bankası, Akbank ve Yapı Kredi bankalarının ürün ve hizmet ücretlerini her gün otomatik olarak kendi sitelerinden çekip karşılaştırmalı bir Excel dosyası oluşturan ve değişiklik olduğunda mail bildirimi gönderen otomasyon sistemi.

## Özellikler
- 4 bankanın komisyon ücretlerini yan yana karşılaştırma
- Mobil / Şube kanal ayrımı
- Günlük otomatik çalışma (her sabah 08:00)
- Değişiklik tespiti ve mail bildirimi
- Excel eki ile detaylı raporlama

## Kurulum

### Secrets (Settings → Secrets → Actions)
| Secret | Açıklama |
|--------|----------|
| `MAIL_USER` | Gönderici Gmail adresi |
| `MAIL_PASS` | Gmail uygulama şifresi |
| `MAIL_TO` | Alıcı mail adresi |

## Dosya Yapısı
```
models.py                  # UcretSatiri dataclass
scraper_garanti.py         # Garanti BBVA scraper
scraper_isbank.py          # İş Bankası scraper
scraper_akbank.py          # Akbank scraper
scraper_yapikredi.py       # Yapı Kredi scraper
karsilastirma_excel.py     # Karşılaştırmalı Excel yazıcı
main.py                    # Ana script
notify.py                  # Mail bildirimi
requirements.txt
.github/workflows/daily-scrape.yml
```
