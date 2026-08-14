from scraper_yapikredi import scrape_yapikredi

satirlar = scrape_yapikredi()
print('\nİlk 10 masraf adı:')
for i, s in enumerate(satirlar[:10]):
    print(f'{i+1}. Masraf: {s.masraf[:100]}')
    print(f'   Asgari Tutar: {s.asgari_tutar}')
    print(f'   Kanal: {s.kanal}')
    print()
