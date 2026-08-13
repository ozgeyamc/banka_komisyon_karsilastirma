"""
Banka komisyon karşılaştırma - Excel çıktısı.
Masraf adları standart geldiği için olduğu gibi kullanılır.
Kategori eşleştirmesi kategori adından yapılır.
"""

import os
import re
from datetime import datetime, timezone, timedelta
from typing import Dict, List, Tuple, Optional
from models import UcretSatiri

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

EXCEL_DOSYA_ADI = "komisyon_karsilastirma.xlsx"
SHEET_ADI = "karşılaştırma"

BANKALAR = ["GARANTİ", "İŞBANKASI", "AKBANK", "YAPIKREDI"]
BANKA_TAMAD = {
    "GARANTİ":   "Garanti BBVA",
    "İŞBANKASI": "İş Bankası",
    "AKBANK":    "Akbank",
    "YAPIKREDI": "Yapı ve Kredi Bankası",
}
BANKA_RENKLER = {
    "GARANTİ":   {"bg": "00B050", "fg": "FFFFFF"},
    "İŞBANKASI": {"bg": "012169", "fg": "FFFFFF"},
    "AKBANK":    {"bg": "FF0000", "fg": "FFFFFF"},
    "YAPIKREDI": {"bg": "003087", "fg": "FFD700"},
}

thin = Side(style="thin", color="CCCCCC")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# Standart kategori adları ve sırası
KATEGORI_SIRA = [
    "EFT Gönderimi",
    "Havale Gönderimi",
    "FAST",
    "Kiralık Kasa",
    "Kıymetli Maden Teslimleri",
    "Kredi Risk Raporu",
    "Fatura Ödemeleri - Kart",
    "SGK Prim Ödemeleri",
    "HGS Etiket Bedeli",
    "Şans Oyunu Ödemeleri Aracılık",
    "Aidat Ödemeleri Aracılık",
    "Özel Okul Ödeme",
    "Telefon Ödemeleri Aracılık",
    "Vergi Tahsilat Aracılık",
    "Arşiv Araştırma Ücreti",
    "Mevduat Araştırma",
    "Bakiye Sorma - Yurtiçi - ATM",
    "Bakiye Sorma - Yurtdışı - ATM",
    "Çek Defteri ve Çek Düzenleme Ücreti",
    "Çek İade Ücreti",
    "Çek Tahsilat Ücreti",
    "Çek Belgelendirme ve Düzeltme Ücreti",
    "Senet İade Ücreti",
    "Senet Protesto İşlemleri Ücreti",
    "Senet Tahsile Alma Ücreti",
]

# Scraper'dan gelen kategori adı → standart kategori adı eşleştirmesi
# Her bankanın farklı kategori adlarını standart ada map'le
KATEGORI_MAP = {
    # EFT
    "elektronik fon transferi (eft) ücreti": "EFT Gönderimi",
    "eft ücreti": "EFT Gönderimi",
    "eft-fast ücreti": "EFT Gönderimi",
    "eft- fast  ücreti": "EFT Gönderimi",
    "eft gönderimi": "EFT Gönderimi",
    "fonların anlık ve sürekli transferi (fast)": "FAST",
    "fast": "FAST",
    # Havale
    "havale ücreti": "Havale Gönderimi",
    "havale": "Havale Gönderimi",
    "havale gönderimi": "Havale Gönderimi",
    "atm kartlı/kartsız havale": "Havale Gönderimi",
    # Kıymetli Maden
    "kıymetli maden transfer ücreti": "Kıymetli Maden Teslimleri",
    "kıymetli maden transferi - altın transfer sistemi ücreti": "Kıymetli Maden Teslimleri",
    "kıymetli maden teslimleri -": "Kıymetli Maden Teslimleri",
    "kıymetli maden teslimleri": "Kıymetli Maden Teslimleri",
    # Kiralık Kasa
    "kiralık kasa ücreti": "Kiralık Kasa",
    "kiralık kasa": "Kiralık Kasa",
    "kasa24 ücretleri": "Kiralık Kasa",
    # Kredi Risk
    "üçüncü kişi ve kuruluşlardan temin edilecek rapor ücretleri - kredi risk raporu": "Kredi Risk Raporu",
    "kredi risk raporu": "Kredi Risk Raporu",
    "findeks paket ücretleri": "Kredi Risk Raporu",
    # Fatura
    "fatura ödemeleri - kart": "Fatura Ödemeleri - Kart",
    "kredi kartı işlem ücretleri": "Fatura Ödemeleri - Kart",
    "kurum tahsilatları": "Fatura Ödemeleri - Kart",
    # SGK
    "sgk prim ödemeleri": "SGK Prim Ödemeleri",
    # HGS
    "geçiş ürünleri ücretleri": "HGS Etiket Bedeli",
    "hgs etiket bedeli": "HGS Etiket Bedeli",
    "hgs": "HGS Etiket Bedeli",
    # Şans Oyunu
    "şans oyunu ödemeleri aracılık": "Şans Oyunu Ödemeleri Aracılık",
    # Aidat
    "aidat ödemeleri aracılık": "Aidat Ödemeleri Aracılık",
    # Özel Okul
    "özel okul ödeme": "Özel Okul Ödeme",
    # Telefon
    "telefon ödemeleri aracılık": "Telefon Ödemeleri Aracılık",
    # Vergi
    "vergi tahsilat aracılık": "Vergi Tahsilat Aracılık",
    # Arşiv
    "belge ve bilgilendirme ücreti": "Arşiv Araştırma Ücreti",
    "arşiv araştırma ücreti": "Arşiv Araştırma Ücreti",
    # Mevduat Araştırma
    "mevduat araştırma": "Mevduat Araştırma",
    "referans mektubu": "Mevduat Araştırma",
    "mutabakat / teyit yazıları": "Mevduat Araştırma",
    # Bakiye ATM
    "bakiye sorma - yurtiçi - atm": "Bakiye Sorma - Yurtiçi - ATM",
    "bakiye sorma - yurtdışı - atm": "Bakiye Sorma - Yurtdışı - ATM",
    "ortak atm kullanımı": "Bakiye Sorma - Yurtiçi - ATM",
    # Çek
    "çek defteri ve çek düzenleme ücreti": "Çek Defteri ve Çek Düzenleme Ücreti",
    "çekler": "Çek Defteri ve Çek Düzenleme Ücreti",
    "çek iade ücreti": "Çek İade Ücreti",
    "çek tahsilat ücreti": "Çek Tahsilat Ücreti",
    "çek belgelendirme ve düzeltme ücreti": "Çek Belgelendirme ve Düzeltme Ücreti",
    "çek belgelendirme ve düzeltme i̇şlemleri ücreti": "Çek Belgelendirme ve Düzeltme Ücreti",
    # Senet
    "senet iade ücreti": "Senet İade Ücreti",
    "senetler": "Senet İade Ücreti",
    "senet protesto i̇şlemleri ücreti": "Senet Protesto İşlemleri Ücreti",
    "senet protesto işlemleri ücreti": "Senet Protesto İşlemleri Ücreti",
    "senet tahsile alma ücreti": "Senet Tahsile Alma Ücreti",
}

# Masraf adı → standart kategori (kategori eşleşmezse masraf adından bul)
MASRAF_KATEGORI_MAP = {
    "arşiv araştırma ücreti": "Arşiv Araştırma Ücreti",
    "bakiye sorma - yurtiçi - atm": "Bakiye Sorma - Yurtiçi - ATM",
    "bakiye sorma - yurtdışı - atm": "Bakiye Sorma - Yurtdışı - ATM",
    "hgs etiket ücreti": "HGS Etiket Bedeli",
    "hgs kart ücreti": "HGS Etiket Bedeli",
    "şans oyunu ödemeleri aracılık": "Şans Oyunu Ödemeleri Aracılık",
    "senet iade ücreti": "Senet İade Ücreti",
    "senet protesto -": "Senet Protesto İşlemleri Ücreti",
    "senet protesto kaldırma -": "Senet Protesto İşlemleri Ücreti",
    "aynı banka senet tahsili -": "Senet Tahsile Alma Ücreti",
    "muhabir banka senet tahsili -": "Senet Tahsile Alma Ücreti",
    "çek iade ücreti": "Çek İade Ücreti",
    "çek defteri (yaprak başı) -": "Çek Defteri ve Çek Düzenleme Ücreti",
    "çek düzenleme -": "Çek Defteri ve Çek Düzenleme Ücreti",
    "özel nitelikli çek düzenleme -": "Çek Defteri ve Çek Düzenleme Ücreti",
    "aynı banka çeki -": "Çek Tahsilat Ücreti",
    "diğer banka çeki -": "Çek Tahsilat Ücreti",
    "döviz çekleri tahsilatı (diğer banka) -": "Çek Tahsilat Ücreti",
    "karşılıksız çek belgelendirme -": "Çek Belgelendirme ve Düzeltme Ücreti",
    "çek düzeltme hakkı -": "Çek Belgelendirme ve Düzeltme Ücreti",
    "üçüncü kişi ve kuruluşlardan temin edilecek rapor ücretleri - kredi risk raporu": "Kredi Risk Raporu",
}

# Masraf adı → standart kanal (masraf adında kanal bilgisi varsa)
def _kanal_bul(masraf: str, scraper_kanal: str) -> str:
    if scraper_kanal in ("mobil", "sube"):
        return scraper_kanal
    ml = masraf.lower()
    if any(k in ml for k in ["mobil", "internet", "iscep", "online", "dijital", "e-"]):
        return "mobil"
    if any(k in ml for k in ["şube", "gişe", "çözüm merkezi", "müşteri iletişim"]):
        return "sube"
    if "atm" in ml:
        return "mobil"
    return "mobil"


def _standart_kategori(kat: str, masraf: str) -> Optional[str]:
    """Kategori adından standart kategori bul."""
    k = kat.lower().strip()
    # Direkt eşleşme
    if k in KATEGORI_MAP:
        return KATEGORI_MAP[k]
    # Kısmi eşleşme
    for anahtar, standart in KATEGORI_MAP.items():
        if anahtar in k or k in anahtar:
            return standart
    # Masraf adından bul
    m = masraf.lower().strip()
    if m in MASRAF_KATEGORI_MAP:
        return MASRAF_KATEGORI_MAP[m]
    for anahtar, standart in MASRAF_KATEGORI_MAP.items():
        if anahtar in m:
            return standart
    return None


def _deger(s: UcretSatiri) -> str:
    t = (s.asgari_tutar or "").strip()
    o = (s.asgari_oran or "").strip()
    if t in ("-", ""):
        t = ""
    if o in ("-", ""):
        o = ""
    parts = [x for x in [t, o] if x]
    return " / ".join(parts)


def _bugun() -> str:
    tz = timezone(timedelta(hours=3))
    return datetime.now(tz).strftime("%d.%m.%Y %H:%M")


def karsilastirma_excel_yaz(
    banka_verileri: Dict[str, List[UcretSatiri]],
    dosya_yolu: str = EXCEL_DOSYA_ADI,
) -> None:
    if os.path.exists(dosya_yolu):
        os.remove(dosya_yolu)

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_ADI
    MAX_COL = 9

    def sb(cell):
        cell.border = BORDER

    # ── Başlık ──
    ws.merge_cells("A1:A2")
    c = ws["A1"]
    c.fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    c.font = Font(color="FFFFFF", bold=True, size=11)
    c.alignment = Alignment(horizontal="center", vertical="center")
    sb(c)

    col = 2
    for banka in BANKALAR:
        r = BANKA_RENKLER[banka]
        fill = PatternFill(start_color=r["bg"], end_color=r["bg"], fill_type="solid")
        l1, l2 = get_column_letter(col), get_column_letter(col + 1)
        ws.merge_cells(f"{l1}1:{l2}1")
        c1 = ws[f"{l1}1"]
        c1.value = BANKA_TAMAD[banka]
        c1.fill = fill
        c1.font = Font(color=r["fg"], bold=True, size=11)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        sb(c1)
        for j, kanal_adi in enumerate(["Mobil", "Şube"]):
            c2 = ws.cell(row=2, column=col + j)
            c2.value = kanal_adi
            c2.fill = fill
            c2.font = Font(color=r["fg"], bold=True, size=10)
            c2.alignment = Alignment(horizontal="center", vertical="center")
            sb(c2)
        col += 2

    ws.column_dimensions["A"].width = 42
    for i in range(2, MAX_COL + 1):
        ws.column_dimensions[get_column_letter(i)].width = 17
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 18
    ws.freeze_panes = "A3"

    # ── Veri işleme ──
    # kat_satirlar: standart_kat → [masraf_adi sırası]
    kat_satirlar: Dict[str, List[str]] = {}
    # veri: (standart_kat, masraf_adi) → {banka: {kanal: deger}}
    veri: Dict[Tuple[str, str], Dict[str, Dict[str, str]]] = {}

    for banka in BANKALAR:
        for s in banka_verileri.get(banka, []):
            masraf = s.masraf or ""
            kategori = s.kategori or ""

            # Yapıkredi: "KategoriAdı | MasrafAdı" formatı
            if " | " in masraf:
                parca = masraf.split(" | ", 1)
                kategori = parca[0]
                masraf = parca[1]

            masraf = masraf.strip()
            if not masraf:
                continue

            # Standart kategori bul
            std_kat = _standart_kategori(kategori, masraf)
            if std_kat is None:
                continue

            # Kanal bul
            kanal = _kanal_bul(masraf, s.kanal or "")

            # Değer
            d = _deger(s)
            if not d:
                continue

            key = (std_kat, masraf)
            kat_satirlar.setdefault(std_kat, [])
            if masraf not in kat_satirlar[std_kat]:
                kat_satirlar[std_kat].append(masraf)

            veri.setdefault(key, {})
            veri[key].setdefault(banka, {"mobil": "", "sube": ""})
            if d:
                veri[key][banka][kanal] = d

    # ── Excel yazma ──
    KAT_FONT = Font(bold=True, size=10)
    DATA_FONT = Font(size=10)
    MASRAF_FONT = Font(size=10)

    yazilan = set()
    row = 3
    toplam = 0

    def yaz_kat(kat: str):
        nonlocal row, toplam
        if kat not in kat_satirlar:
            return
        yazilan.add(kat)

        ws.merge_cells(f"A{row}:{get_column_letter(MAX_COL)}{row}")
        cc = ws[f"A{row}"]
        cc.value = kat
        cc.font = KAT_FONT
        cc.alignment = Alignment(horizontal="center", vertical="center")
        sb(cc)
        ws.row_dimensions[row].height = 16
        row += 1

        for masraf in kat_satirlar[kat]:
            key = (kat, masraf)
            mc = ws.cell(row=row, column=1, value=masraf)
            mc.font = MASRAF_FONT
            mc.alignment = Alignment(horizontal="left", vertical="center",
                                     wrap_text=True, indent=1)
            sb(mc)

            col = 2
            for banka in BANKALAR:
                bd = veri.get(key, {}).get(banka, {"mobil": "", "sube": ""})
                for kanal in ("mobil", "sube"):
                    d = bd.get(kanal, "")
                    c = ws.cell(row=row, column=col, value=d)
                    c.font = DATA_FONT
                    c.alignment = Alignment(horizontal="center", vertical="center",
                                            wrap_text=True)
                    sb(c)
                    col += 1

            ws.row_dimensions[row].height = 20
            row += 1
            toplam += 1

        row += 1

    for kat in KATEGORI_SIRA:
        yaz_kat(kat)

    for kat in list(kat_satirlar.keys()):
        if kat not in yazilan:
            yaz_kat(kat)

    ws.cell(row=row, column=1,
            value=f"Son güncelleme: {_bugun()}").font = Font(size=9, color="888888")

    wb.save(dosya_yolu)
    print(f"[excel] {dosya_yolu} kaydedildi. {toplam} satır yazıldı.")
