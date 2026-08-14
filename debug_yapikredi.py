"""
Yapı Kredi sayfasının HTML yapısını analiz eden debug script.
Ne olduğunu anlamak için sayfayı indir ve yapıyı yazdır.
"""
import sys
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

YAPIKREDI_URL = "https://www.yapikredi.com.tr/bireysel-bankacilik/hesaplama-araclari/bireysel-urun-ve-hizmet-ucretleri"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36"}

def debug_yapikredi():
    print("[debug] Yapı Kredi sayfası analiz ediliyor...", file=sys.stderr)
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            ctx = browser.new_context(
                user_agent=HEADERS["User-Agent"],
                viewport={"width": 1280, "height": 800},
                locale="tr-TR",
                extra_http_headers={"Accept-Language": "tr-TR,tr;q=0.9"}
            )
            page = ctx.new_page()
            page.goto(YAPIKREDI_URL, timeout=120000, wait_until="domcontentloaded")
            page.wait_for_timeout(5000)
            
            # Çerez banner'ı kapat
            for sel in ["#onetrust-accept-btn-handler", "button:has-text('Tümünü Kabul Et')"]:
                try:
                    page.click(sel, timeout=3000)
                    page.wait_for_timeout(500)
                    break
                except:
                    pass
            
            # Accordion'ları aç
            print("[debug] Accordion buttonları bulunuyor...", file=sys.stderr)
            buttons = page.query_selector_all("button[aria-expanded='false']")
            print(f"[debug] {len(buttons)} tane closed button bulundu", file=sys.stderr)
            
            for i, btn in enumerate(buttons[:20]):  # İlk 20'yi aç
                try:
                    btn.click(timeout=1000)
                    page.wait_for_timeout(300)
                    print(f"[debug] Button {i+1} tıklandı", file=sys.stderr)
                except Exception as e:
                    print(f"[debug] Button {i+1} tıklama hatası: {e}", file=sys.stderr)
            
            page.wait_for_timeout(3000)
            html = page.content()
        finally:
            browser.close()
    
    soup = BeautifulSoup(html, "lxml")
    
    # 1. Kaç tabel var
    tables = soup.find_all("table")
    print(f"\n=== TABLOLAR ===")
    print(f"Toplam tabel sayısı: {len(tables)}\n")
    
    # 2. Her tabloyu analiz et
    for idx, table in enumerate(tables[:5]):  # İlk 5 tabloyu analiz et
        print(f"\n--- TABEL {idx+1} ---")
        
        # Tablo yukarısındaki heading bulma
        parent = table.parent
        for _ in range(10):
            if parent is None:
                break
            heading = parent.find("h1") or parent.find("h2") or parent.find("h3")
            if heading:
                print(f"Kategori: {heading.get_text()[:80]}")
                break
            parent = parent.parent
        
        # Tablo yapısı
        thead = table.find("thead")
        tbody = table.find("tbody")
        
        if thead:
            headers = [c.get_text().strip()[:30] for c in thead.find_all(["th", "td"])]
            print(f"Headers: {headers}")
        
        if tbody:
            rows = tbody.find_all("tr")
            print(f"Satır sayısı: {len(rows)}")
            if rows:
                first_row = [c.get_text().strip()[:30] for c in rows[0].find_all(["th", "td"])]
                print(f"İlk satır: {first_row}")
        else:
            all_rows = table.find_all("tr")
            print(f"Satır sayısı (tbody yok): {len(all_rows)}")
            if all_rows and len(all_rows) > 1:
                first_row = [c.get_text().strip()[:30] for c in all_rows[1].find_all(["th", "td"])]
                print(f"İlk veri satırı: {first_row}")

if __name__ == "__main__":
    debug_yapikredi()
