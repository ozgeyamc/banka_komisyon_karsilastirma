"""
Banka komisyon karşılaştırma botu - ana script.
"""

import sys
from karsilastirma_excel import karsilastirma_excel_yaz, EXCEL_DOSYA_ADI

BANKA_SIRASI = [
    ("GARANTİ",   "scraper_garanti",   "scrape_garanti_bbva"),
    ("İŞBANKASI", "scraper_isbank",    "scrape_isbank"),
    ("AKBANK",    "scraper_akbank",    "scrape_akbank"),
    ("YAPIKREDI", "scraper_yapikredi", "scrape_yapikredi"),
]


def main() -> int:
    print("=== Banka Komisyon Karşılaştırma Botu ===")
    banka_verileri = {}

    for banka_adi, module, func in BANKA_SIRASI:
        print(f"\n--- {banka_adi} çekiliyor ---")
        try:
            mod = __import__(module)
            fn = getattr(mod, func)
            satirlar = fn()
            if satirlar:
                print(f"{banka_adi}: {len(satirlar)} satır bulundu.")
                banka_verileri[banka_adi] = satirlar
        except ImportError:
            print(f"[UYARI] {banka_adi} modülü bulunamadı, atlanıyor.", file=sys.stderr)
        except Exception as exc:
            print(f"[HATA] {banka_adi}: {exc}", file=sys.stderr)

    if not banka_verileri:
        print("[HATA] Hiçbir bankadan veri çekilemedi!", file=sys.stderr)
        return 1

    try:
        karsilastirma_excel_yaz(banka_verileri, EXCEL_DOSYA_ADI)
    except Exception as exc:
        print(f"[HATA] Excel yazılırken hata: {exc}", file=sys.stderr)
        return 1

    print(f"\nTamamlandı. Excel: {EXCEL_DOSYA_ADI}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
