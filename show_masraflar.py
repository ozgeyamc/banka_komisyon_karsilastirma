"""
Tüm bankaların masraf adlarını görmek için debug script.
Bunu çalıştırmanız gerekmez - kod direkt olarak karşılaştırma yapıyor.
"""

from scraper_garanti import scrape_garanti_bbva
from scraper_isbank import scrape_isbank
from scraper_akbank import scrape_akbank

print('=== GARANTİ (ilk 10 masraf) ===')
for i, s in enumerate(scrape_garanti_bbva()[:10]):
    print(f'{i+1}. {s.masraf[:80]}')

print('\n=== İŞBANKASI (ilk 10 masraf) ===')
for i, s in enumerate(scrape_isbank()[:10]):
    print(f'{i+1}. {s.masraf[:80]}')

print('\n=== AKBANK (ilk 10 masraf) ===')
for i, s in enumerate(scrape_akbank()[:10]):
    print(f'{i+1}. {s.masraf[:80]}')
