import re
import sys
from typing import List
from models import UcretSatiri

AKBANK_URL = "https://www.akbank.com/urun-ve-hizmet-ucretleri"
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


def _normalize(val): return str(val).strip().replace("\xa0"," ").replace("\u200b","").strip()

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
    p = el.parent
    for _ in range(10):
        if p is None: break
        for sib in p.find_all_previous(["h1","h2","h3","h4","h5"], limit=3):
            t = _normalize(sib.get_text())
            if len(t) > 5 and t not in ["Müşteri Ol","Ara","Kapat","Menü","Ana Sayfa"]: return t
        p = p.parent
    return fallback

def _kanal_tespit(kategori: str, masraf: str) -> str:
    """
    Kanal bilgisi çoğu zaman masraf metninde değil, tablo başlığı/kategori
    metninde geçer. Her ikisi birlikte kontrol edilir.
    """
    kaynak = f"{kategori} {masraf}".lower()
    if any(k in kaynak for k in ["mobil", "internet şube", "i̇nternet şube", "işcep", "iscep", "online", "dijital"]):
        return "mobil"
    if any(k in kaynak for k in ["şube", "sube", "çözüm merkezi", "cozum merkezi", "gişe", "gise"]):
        return "sube"
    if "atm" in kaynak:
        return "mobil"
    return ""

def _extract(table, kat):
    satirlar = []
    thead = table.find("thead"); tbody = table.find("tbody")
    headers = []
    if thead:
        hr = thead.find("tr")
        if hr: headers = [_normalize(c.get_text(strip=True)).lower() for c in hr.find_all(["th","td"])]
    if tbody:
        rows = tbody.find_all("tr")
        if not headers and rows:
            headers = [_normalize(c.get_text(strip=True)).lower() for c in rows[0].find_all(["th","td"])]
            rows = rows[1:]
    else:
        all_rows = table.find_all("tr")
        if not all_rows: return satirlar
        if not headers:
            headers = [_normalize(c.get_text(strip=True)).lower() for c in all_rows[0].find_all(["th","td"])]
        rows = all_rows[1:]
    def fc(keys):
        for i,h in enumerate(headers):
            if all(k in h for k in keys): return i
        return -1
    cm=fc(["masraf"]); ca1=fc(["asgari","tutar"]); ca2=fc(["asgari","oran"])
    cz1=fc(["azami","tutar"]); cz2=fc(["azami","oran"])
    cac=fc(["açıklama"]) if fc(["açıklama"])>=0 else fc(["aciklama"])
    ct=fc(["güncelleme"]) if fc(["güncelleme"])>=0 else fc(["guncelleme"])
    if cm==-1: cm,ca1,ca2,cz1,cz2,cac=0,1,2,3,4,5
    for row in rows:
        cells=row.find_all(["th","td"])
        if len(cells)<2: continue
        v=[_normalize(c.get_text(strip=True)) for c in cells]
        def g(i): return v[i] if 0<=i<len(v) else ""
        masraf=g(cm)
        if not masraf: continue
        tarih=g(ct) if ct>=0 else ""
        aciklama,at=_parse_aciklama(g(cac))
        if not tarih: tarih=at
        kanal = _kanal_tespit(kat, masraf)
        satirlar.append(UcretSatiri(kategori=kat,masraf=masraf,asgari_tutar=g(ca1),asgari_oran=g(ca2),
                                    azami_tutar=g(cz1),azami_oran=g(cz2),aciklama=aciklama,
                                    site_guncelleme_tarihi=tarih,kanal=kanal))
    return satirlar

def scrape_akbank(url=AKBANK_URL) -> List[UcretSatiri]:
    from playwright.sync_api import sync_playwright
    from bs4 import BeautifulSoup
    print(f"[akbank] {url} çekiliyor...", file=sys.stderr)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            page = browser.new_page(user_agent=HEADERS["User-Agent"])
            page.goto(url, timeout=90000, wait_until="domcontentloaded")
            page.wait_for_timeout(3000)
            for sel in ["button[aria-expanded='false']",".accordion-button.collapsed","[data-bs-toggle='collapse']",
                        ".card-header button","[class*='accordion']","[class*='tab-']"]:
                for el in page.query_selector_all(sel):
                    try: el.click(timeout=1500); page.wait_for_timeout(200)
                    except: pass
            page.wait_for_timeout(3000)
            html = page.content()
        finally:
            browser.close()
    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")
    if not tables: raise ScraperError("Akbank: tablo bulunamadı.")
    result = []
    for t in tables:
        result.extend(_extract(t, _find_cat(t, "Genel")))
    if not result: raise ScraperError("Akbank: veri satırı çekilemedi.")
    print(f"[akbank] {len(result)} satır.", file=sys.stderr)
    return result
