import re
import sys
from typing import List
from models import UcretSatiri

YAPIKREDI_URL = "https://www.yapikredi.com.tr/bireysel-bankacilik/hesaplama-araclari/bireysel-urun-ve-hizmet-ucretleri"
DATE_PATTERN = re.compile(r"G[üu]ncellenme\s*Tarihi\s*:\s*(\d{2}\.\d{2}\.\d{4}(?:\s+\d{2}:\d{2})?)", re.IGNORECASE)
DATE_PATTERN_TR = re.compile(
    r"(?:son\s+)?g[üu]ncellenme\s+tarihi\s*:?\s*(\d{1,2})\s+(ocak|şubat|mart|nisan|mayıs|haziran|temmuz|ağustos|eylül|ekim|kasım|aralık)\s+(\d{4})",
    re.IGNORECASE,
)
TURKCE_AYLAR = {"ocak":"01","şubat":"02","mart":"03","nisan":"04","mayıs":"05","haziran":"06",
                "temmuz":"07","ağustos":"08","eylül":"09","ekim":"10","kasım":"11","aralık":"12"}
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36"}


class ScraperError(Exception):
    pass


def _normalize(val): 
    return str(val).strip().replace("\xa0"," ").replace("\u200b","").strip()

def _parse_aciklama(raw):
    raw = raw.strip()
    m = DATE_PATTERN.search(raw)
    if m: return DATE_PATTERN.sub("", raw).strip(" ."), m.group(1)
    m = DATE_PATTERN_TR.search(raw)
    if m:
        g,ay,y = m.group(1).zfill(2), TURKCE_AYLAR.get(m.group(2).lower(),"00"), m.group(3)
        return DATE_PATTERN_TR.sub("", raw).strip(" ."), f"{g}.{ay}.{y}"
    return raw, ""

def _find_cat(el, fallback):
    """
    Kategori bilgisini bulmak için HTML yapısında yukarıya doğru arama yapıyor.
    Tablonun parent'ı olan div/section'ın üstündeki heading'leri kontrol ediyor.
    """
    p = el.parent
    for _ in range(20):
        if p is None: 
            break
        
        # Tablonun direkt üstündeki h2/h3 bul
        for sib in p.find_all_previous(["h1", "h2", "h3", "h4", "h5", "div"], limit=10):
            if sib.name in ["h1", "h2", "h3", "h4", "h5"]:
                t = _normalize(sib.get_text())
                if len(t) > 3 and t not in ["Müşteri Ol","Ara","Kapat","Menü","Ana Sayfa"]: 
                    print(f"[yapikredi-debug] Kategori bulundu: {t}", file=sys.stderr)
                    return t
            elif sib.name == "div" and ("class" in sib.attrs):
                # div'in içindeki h2/h3 ara
                for h in sib.find_all(["h1", "h2", "h3", "h4", "h5"], limit=1):
                    t = _normalize(h.get_text())
                    if len(t) > 3 and t not in ["Müşteri Ol","Ara","Kapat","Menü","Ana Sayfa"]:
                        print(f"[yapikredi-debug] Kategori bulundu (div içinde): {t}", file=sys.stderr)
                        return t
        
        p = p.parent
    
    print(f"[yapikredi-debug] Kategori bulunamadı, fallback kullanılıyor: {fallback}", file=sys.stderr)
    return fallback

def _kanal_tespit(kategori: str, masraf: str) -> str:
    """
    Kanal bilgisi çoğu zaman masraf metninde değil, tablo başlığı/kategori
    metninde geçer. Her ikisi birlikte kontrol edilir.
    """
    kaynak = f"{kategori} {masraf}".lower()
    if any(k in kaynak for k in ["mobil", "internet şube", "i̇nternet şube", "işcep", "iscep", "online", "dijital", "app"]):
        return "mobil"
    if any(k in kaynak for k in ["şube", "sube", "çözüm merkezi", "cozum merkezi", "gişe", "gise"]):
        return "sube"
    if "atm" in kaynak:
        return "mobil"
    return ""

def _extract(table, kat):
    satirlar = []
    thead = table.find("thead")
    tbody = table.find("tbody")
    headers = []
    
    if thead:
        hr = thead.find("tr")
        if hr: 
            headers = [_normalize(c.get_text(strip=True)).lower() for c in hr.find_all(["th","td"])]
    
    if tbody:
        rows = tbody.find_all("tr")
        if not headers and rows:
            headers = [_normalize(c.get_text(strip=True)).lower() for c in rows[0].find_all(["th","td"])]
            rows = rows[1:]
    else:
        all_rows = table.find_all("tr")
        if not all_rows: 
            return satirlar
        if not headers:
            headers = [_normalize(c.get_text(strip=True)).lower() for c in all_rows[0].find_all(["th","td"])]
        rows = all_rows[1:]
    
    # Boş header varsa atla
    if not headers or len(headers) < 2:
        return satirlar
    
    print(f"[yapikredi-debug] Tablo headers: {headers[:3]}", file=sys.stderr)
    
    def fc(keys):
        for i,h in enumerate(headers):
            if all(k in h for k in keys): 
                return i
        return -1
    
    cm = fc(["masraf"])
    ca1 = fc(["asgari","tutar"])
    ca2 = fc(["asgari","oran"])
    cz1 = fc(["azami","tutar"])
    cz2 = fc(["azami","oran"])
    cac = fc(["açıklama"]) if fc(["açıklama"]) >= 0 else fc(["aciklama"])
    ct = fc(["güncelleme"]) if fc(["güncelleme"]) >= 0 else fc(["guncelleme"])
    
    if cm == -1: 
        cm, ca1, ca2, cz1, cz2 = 0, 1, 2, 3, 4
    
    print(f"[yapikredi-debug] Kolon indeksleri - Masraf:{cm} AsgariT:{ca1} AsgariO:{ca2} AzamiT:{cz1} AzamiO:{cz2}", file=sys.stderr)
    
    satir_sayisi = 0
    for row in rows:
        cells = row.find_all(["th","td"])
        if len(cells) < 2: 
            continue
        
        v = [_normalize(c.get_text(strip=True)) for c in cells]
        def g(i): 
            return v[i] if 0 <= i < len(v) else ""
        
        masraf = g(cm)
        
        # Boş satırları ve başlık satırlarını atla
        if not masraf or masraf in ["Masraf","masraf","-","–",""] or len(masraf) < 2:
            continue
        
        tarih = g(ct) if ct >= 0 else ""
        aciklama, at = _parse_aciklama(g(cac))
        if not tarih: 
            tarih = at
        
        kanal = _kanal_tespit(kat, masraf)
        tam_masraf = f"{kat} | {masraf}"
        
        satirlar.append(UcretSatiri(
            kategori=kat, 
            masraf=tam_masraf, 
            asgari_tutar=g(ca1), 
            asgari_oran=g(ca2),
            azami_tutar=g(cz1), 
            azami_oran=g(cz2), 
            aciklama=aciklama,
            site_guncelleme_tarihi=tarih, 
            kanal=kanal
        ))
        satir_sayisi += 1
    
    print(f"[yapikredi-debug] Tablo: {satir_sayisi} satır çekildi, kategori: {kat}", file=sys.stderr)
    return satirlar

def _cerez_kapat(page):
    """
    Sayfanın altında beliren çerez izni banner'ı kapatıyor.
    """
    for sel in [
        "#onetrust-accept-btn-handler",
        "button:has-text('Tümünü Kabul Et')",
        "text=Tümünü Kabul Et",
        "button:has-text('Kabul Et')",
        ".cookie-accept",
        "[class*='cookie'] button",
    ]:
        try:
            page.click(sel, timeout=3000)
            page.wait_for_timeout(500)
            print("[yapikredi] çerez banner'ı kapatıldı.", file=sys.stderr)
            return
        except:
            pass
    print("[yapikredi] çerez banner'ı bulunamadı/kapatılamadı (devam ediliyor).", file=sys.stderr)

def _scroll_aciklama_basliklari(page):
    """
    Accordion/collapsible başlıkları açmak için daha sağlam yöntem.
    Sayfa üzerinde scroll yaparak tüm elementleri render et.
    """
    print("[yapikredi] Sayfa scroll ediliyor ve tüm accordion'lar açılıyor...", file=sys.stderr)
    
    # Tüm olası accordion/tab seçicileri
    selectors = [
        "button[aria-expanded='false']",
        "button[data-bs-toggle='collapse']",
        "[aria-expanded='false']",
        ".accordion-button.collapsed",
        ".card-header button",
        "[class*='accordion'] button",
        "[class*='collapse']",
        "a[data-toggle='collapse']",
    ]
    
    for sel in selectors:
        elements = page.query_selector_all(sel)
        print(f"[yapikredi-debug] Seçici '{sel}' - {len(elements)} element bulundu", file=sys.stderr)
        
        for el in elements:
            try:
                el.click(timeout=1000)
                page.wait_for_timeout(200)
            except:
                pass
    
    # Sayfayı aşağıya doğru scroll et
    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    page.wait_for_timeout(2000)
    
    # Tekrar yukarıya
    page.evaluate("window.scrollTo(0, 0)")
    page.wait_for_timeout(1000)

def scrape_yapikredi(url=YAPIKREDI_URL) -> List[UcretSatiri]:
    from playwright.sync_api import sync_playwright
    from bs4 import BeautifulSoup
    
    print(f"[yapikredi] {url} çekiliyor...", file=sys.stderr)
    
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
            page.goto(url, timeout=120000, wait_until="domcontentloaded")
            page.wait_for_timeout(5000)

            # Çerez banner'ını kapat
            _cerez_kapat(page)
            
            # Accordion/tab'ları aç
            _scroll_aciklama_basliklari(page)
            
            page.wait_for_timeout(5000)
            html = page.content()
        finally:
            browser.close()
    
    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")
    
    print(f"[yapikredi-debug] Toplam {len(tables)} tablo bulundu", file=sys.stderr)
    
    if not tables: 
        raise ScraperError("Yapı Kredi: tablo bulunamadı.")
    
    result = []
    for i, t in enumerate(tables):
        print(f"[yapikredi-debug] Tablo {i+1} işleniyor...", file=sys.stderr)
        result.extend(_extract(t, _find_cat(t, f"Genel-{i+1}")))
    
    if not result: 
        raise ScraperError("Yapı Kredi: veri satırı çekilemedi.")
    
    print(f"[yapikredi] {len(result)} satır.", file=sys.stderr)
    return result
