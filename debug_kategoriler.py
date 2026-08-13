"""
Her bankadan gelen kategori ve masraf adlarını yazdırır.
python debug_kategoriler.py
"""
import sys
from scraper_garanti   import scrape_garanti_bbva
from scraper_isbank    import scrape_isbank
from scraper_akbank    import scrape_akbank
from scraper_yapikredi import scrape_yapikredi

BANKALAR = {
    "GARANTİ":   scrape_garanti_bbva,
    "İŞBANKASI": scrape_isbank,
    "AKBANK":    scrape_akbank,
    "YAPIKREDI": scrape_yapikredi,
}

for banka, fn in BANKALAR.items():
    print(f"\n{'='*60}")
    print(f"  {banka}")
    print(f"{'='*60}")
    try:
        satirlar = fn()
        kategoriler = {}
        for s in satirlar:
            kategoriler.setdefault(s.kategori, []).append(s.masraf)
        for kat, masraflar in kategoriler.items():
            print(f"\n  [{kat}]")
            for m in masraflar[:5]:   # ilk 5 masraf
                print(f"    - {m}")
            if len(masraflar) > 5:
                print(f"    ... ({len(masraflar)} toplam)")
    except Exception as e:
        print(f"  HATA: {e}", file=sys.stderr)
