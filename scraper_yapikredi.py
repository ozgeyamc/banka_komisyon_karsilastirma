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
    """
    p = el.parent
    for _ in range(20):
        if p is None: 
            break
        
        for sib in p.find_all_previous(["h1", "h2", "h3", "h4", "h5"], limit=10):
            t = _normalize(sib.get_text())
            if len(t) > 3 and t not in ["Müşteri Ol","Ara","Kapat","Menü","Ana Sayfa"]: 
                return t
        
        p = p.parent
    
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

def _is_header_row(cells):
    """
    Satırın başlık satırı mı yoksa veri satırı mı olduğunu kontrol et.
    Başlık satırları genellikle text içerir, veri satırları sayı/oran içerir.
    """
    if len(cells) < 2:
        return True
    
    first_cell = _normalize(cells[0].get_text()).lower()
    
    # Başlık satırı göstergeleri
    if any(k in first_cell for k in ["asgari", "azami", "mobil", "şube", "tutar", "oran", "ücreti", "bedel"]):
        return True
    
    # Çok kısa veya çok uzun metinler
    if len(first_cell) < 2 or len(first_cell) > 200:
        return True
    
    return False

def _extract(table, kat):
    satirlar = []
    tbody = table.find("tbody")
    all_rows = table.find_all("tr")
    
    if not all_rows:
        return satirlar
    
    # Başlık satırını atla (ilk satır)
    rows = all_rows[1:] if tbody is None else tbody.find_all("tr")
    
    if not rows:
        return satirlar
    
    # İlk veri satırından headers belirle
    headers = []
    for i, row in enumerate(rows):
        cells = row.find_all(["th","td"])
        if not _is_header_row(cells):
            # İlk veri satırı bulundu, bunu header olarak kullan
            headers = [_normalize(c.get_text(strip=True)).lower() for c in cells]
            rows = rows[i+1:]  # Kalan satırları işle
            break
    
    if not headers:
        return satirlar
    
    print(f"[yapikredi-debug] Tablo headers: {headers[:4]}", file=sys.stderr)
    
    def fc(keys):
        for i,h in enumerate(headers):
            if all(k in h for k in keys): 
                return i
        return -1
    
    cm = fc(["işlem", "masraf", "gönderim", "döviz", "havale"])  # Masraf sütununu bul
    if cm == -1:
        cm = 0  # İlk sütun masraf
    
    ca1 = fc(["asgari","tutar"])
    ca2 = fc(["asgari","oran"])
    cz1 = fc(["azami","tutar"])
    cz2 = fc(["azami","oran"])
    cac = fc(["açıklama"]) if fc(["açıklama"]) >= 0 else fc(["aciklama"])
    ct = fc(["güncelleme"]) if fc(["güncelleme"]) >= 0 else fc(["guncelleme"])
    
    print(f"[yapikredi-debug] Kolon indeksleri - Masraf:{cm} AsgariT:{ca1} AsgariO:{ca2}", file=sys.stderr)
    
    satir_sayisi = 0
    for row in rows:
        cells = row.find_all(["th","td"])
        
        # Başlık satırını atla
        if _is_header_row(cells):
            continue
        
        if len(cells) < 2: 
            continue
        
        v = [_normalize(c.get_text(strip=True)) for c in cells]
        def g(i): 
            return v[i] if 0 <= i < len(v) else ""
        
        masraf = g(cm)
        
        # Boş satırları atla
        if not masraf or len(masraf) < 2:
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
    Accordion/collapsible başlıkları açmak için sağlam yöntem.
    """
    print("[yapikredi] Sayfa scroll ediliyor ve tüm accordion'lar açılıyor...", file=sys.stderr)
    
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
    
    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    page.wait_for_timeout(2000)
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

            _cerez_kapat(page)
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
