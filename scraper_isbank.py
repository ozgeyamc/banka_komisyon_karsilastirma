import sys
from typing import List
from models import UcretSatiri

ISBANK_URL = "https://www.isbank.com.tr/urun-ve-hizmet-ucretleri"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36"}


class ScraperError(Exception):
    pass


def _normalize(val) -> str:
    if val is None: return ""
    return str(val).strip().replace("\xa0", " ").replace("\u200b", "").strip()


def _meta(el, cls="UHU_icerik_meta"):
    span = el.find("span", class_=cls) if el else None
    return _normalize(span.get_text()) if span else ""


def _kanal_tespit(kategori: str, masraf: str) -> str:
    """
    Kanal bilgisi çoğunlukla `masraf` metninde değil, tablo/kategori başlığında
    geçer (örn: 'Şube-Çözüm Merkezi', 'İnternet Şube ve Mobil Bankacılık Kanalları',
    'ATM Kanalı'). Bu yüzden hem kategori hem masraf birlikte kontrol edilir.
    """
    kaynak = f"{kategori} {masraf}".lower()
    if any(k in kaynak for k in ["mobil", "internet şube", "i̇nternet şube", "işcep", "iscep", "online", "dijital"]):
        return "mobil"
    if any(k in kaynak for k in ["şube", "sube", "çözüm merkezi", "cozum merkezi", "gişe", "gise"]):
        return "sube"
    if "atm" in kaynak:
        return "mobil"
    return ""


def scrape_isbank(url=ISBANK_URL) -> List[UcretSatiri]:
    from playwright.sync_api import sync_playwright
    from bs4 import BeautifulSoup
    print(f"[isbank] {url} çekiliyor...", file=sys.stderr)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            ctx = browser.new_context(user_agent=HEADERS["User-Agent"], viewport={"width":1280,"height":800},
                                       locale="tr-TR", timezone_id="Europe/Istanbul")
            ctx.add_init_script("Object.defineProperty(navigator,'webdriver',{get:()=>undefined})")
            page = ctx.new_page()
            page.goto(url, timeout=90000, wait_until="domcontentloaded")
            page.wait_for_timeout(15000)
            html = page.content()
        finally:
            browser.close()

    soup = BeautifulSoup(html, "lxml")
    result = []
    for hi in range(1, 20):
        grup = soup.find(id=f"h{hi}")
        if not grup: break
        ana_kat_el = grup.find(class_="UHU_group_header")
        ana_kat = _normalize(ana_kat_el.get_text()) if ana_kat_el else f"Grup {hi}"
        for item_el in grup.find_all(class_="UHU_item_header"):
            alt_kat = _normalize(item_el.get_text())
            isc = item_el.find_next_sibling(id="UHU_itemSubCover") or item_el.parent
            sub_headers = isc.find_all(class_="UHU_itemSub_header") if isc else []
            for sub_el in (sub_headers if sub_headers else [None]):
                if sub_el:
                    sub_kat = _normalize(sub_el.get_text())
                    tam_kat = f"{ana_kat} - {alt_kat} - {sub_kat}"
                    igc = sub_el.find_next_sibling(id="UHU_item_icerik_GC") or sub_el.parent
                else:
                    tam_kat = f"{ana_kat} - {alt_kat}"
                    igc = isc
                if not igc: continue
                for blok in igc.find_all(class_="UHU_item_icerikC"):
                    masraf_el = blok.find(class_="UHU_item_icerikH")
                    masraf = _normalize(masraf_el.get_text()) if masraf_el else ""
                    if not masraf: continue
                    kanal = _kanal_tespit(tam_kat, masraf)
                    result.append(UcretSatiri(
                        kategori=tam_kat, masraf=masraf,
                        asgari_tutar=_meta(blok.find(class_="UHU_item_icerik1")),
                        asgari_oran=_meta(blok.find(class_="UHU_item_icerik2")),
                        azami_tutar=_meta(blok.find(class_="UHU_item_icerik3")),
                        azami_oran=_meta(blok.find(class_="UHU_item_icerik4")),
                        site_guncelleme_tarihi=_meta(blok.find(class_="UHU_item_icerik5"), cls="UHU_icerik_meta2"),
                        aciklama=_normalize(blok.find(class_="UHU_item_icerikF").get_text()) if blok.find(class_="UHU_item_icerikF") else "",
                        kanal=kanal,
                    ))
    if not result:
        raise ScraperError("İş Bankası: veri çekilemedi.")
    print(f"[isbank] {len(result)} satır.", file=sys.stderr)
    return result
