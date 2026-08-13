"""
Her bankadan gelen kategori ve masraf adlarını debug_output.txt dosyasına yazar.
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

lines = []
for banka, fn in BANKALAR.items():
    lines.append(f"\n{'='*60}")
    lines.append(f"  {banka}")
    lines.append(f"{'='*60}")
    try:
        satirlar = fn()
        kategoriler = {}
        for s in satirlar:
            kategoriler.setdefault(s.kategori, []).append(s.masraf)
        for kat, masraflar in kategoriler.items():
            lines.append(f"\n  [{kat}]")
            for m in masraflar[:5]:
                lines.append(f"    - {m}")
            if len(masraflar) > 5:
                lines.append(f"    ... ({len(masraflar)} toplam)")
    except Exception as e:
        lines.append(f"  HATA: {e}")

output = "\n".join(lines)
print(output)

with open("debug_output.txt", "w", encoding="utf-8") as f:
    f.write(output)

print("\n[debug] debug_output.txt kaydedildi.", file=sys.stderr)
