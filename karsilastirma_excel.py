"""
Bankalardan çekilen verileri karşılaştırmalı Excel formatında yazan modül.
Eşleştirme: işlem tipi + tutar aralığı standardizasyonu ile yapılır.
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
YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


def _bugun() -> str:
    tz = timezone(timedelta(hours=3))
    return datetime.now(tz).strftime("%d.%m.%Y %H:%M")


def _deger(s: UcretSatiri) -> str:
    t = (s.asgari_tutar or "").strip()
    o = (s.asgari_oran  or "").strip()
    az_t = (s.azami_tutar or "").strip()
    az_o = (s.azami_oran  or "").strip()
    parts = []
    if t: parts.append(t)
    if o: parts.append(o)
    if az_t and az_t not in parts: parts.append(az_t)
    if az_o and az_o not in parts: parts.append(az_o)
    return " / ".join(parts) if parts else ""


def _norm(m: str) -> str:
    m = m.lower().strip()
    for a, b in [("ı","i"),("ğ","g"),("ü","u"),("ş","s"),("ö","o"),("ç","c"),
                 ("â","a"),("î","i"),("û","u"),("i̇","i")]:
        m = m.replace(a, b)
    return re.sub(r"\s+", " ", m).strip()


def _extract_tutar(m: str) -> Optional[str]:
    """Masraf adından tutar aralığını çıkar: '0-8300', '8300-399000', '399000+' """
    m = re.sub(r"[.,\s]", "", m)  # nokta/virgül/boşluk kaldır
    patterns = [
        (r"0[-–]8300",        "0-8300"),
        (r"8300[-–]399000",   "8300-399000"),
        (r"399000",           "399000+"),
        (r"0[-–]100",         "0-100"),
        (r"100[-–]",          "100+"),
        (r"1[-–]10g",         "1-10gr"),
        (r"11[-–]100g",       "11-100gr"),
        (r"buyuk|büyük",      "büyük"),
        (r"orta",             "orta"),
        (r"kucuk|küçük",      "küçük"),
    ]
    for pat, label in patterns:
        if re.search(pat, m, re.IGNORECASE):
            return label
    return None


def _standart_anahtar(masraf: str) -> Tuple[str, str, str]:
    """
    Masraf adından (kategori, standart_isim, kanal) döndür.
    standart_isim = Excel'in A sütununa yazılacak temiz isim.
    """
    n = _norm(masraf)
    tutar = _extract_tutar(re.sub(r"\s", "", masraf))

    def kanal():
        if any(k in n for k in ["mobil","internet","iscep","online","dijital","e-"]):
            return "mobil"
        if any(k in n for k in ["sube","şube","gise","gişe","cozum","çözüm","iletisim merkezi"]):
            return "sube"
        if "atm" in n:
            return "mobil"
        return ""

    # EFT
    if "eft" in n and "swift" not in n and "uluslararasi" not in n:
        if "duzenli" in n or "süpürme" in n or "supurme" in n:
            label = f"Düzenli EFT - {tutar} TRY" if tutar else "Düzenli EFT"
            return "EFT Gönderimi", label, kanal()
        if "odenmesi" in n or "ödenmesi" in n or "isime gelen" in n or "isme gelen" in n:
            return "EFT Gönderimi", "EFT İsme Gelen", "sube"
        label = f"EFT - {tutar} TRY" if tutar else "EFT Gönderimi"
        return "EFT Gönderimi", label, kanal()

    # Havale
    if "havale" in n and "uluslararasi" not in n and "swift" not in n:
        if "duzenli" in n or "supurme" in n:
            label = f"Düzenli Havale - {tutar} TRY" if tutar else "Düzenli Havale"
            return "Havale Gönderimi", label, kanal()
        if "cebe para" in n:
            return "Havale Gönderimi", "Cebe Para Gönderme", kanal()
        label = f"Havale - {tutar} TRY" if tutar else "Havale Gönderimi"
        return "Havale Gönderimi", label, kanal()

    # FAST
    if "fast" in n:
        label = f"FAST - {tutar} TRY" if tutar else "FAST"
        return "FAST", label, kanal()

    # Kiralık Kasa
    if "kasa" in n:
        if "depozito" in n:
            return "Kiralık Kasa", f"Kasa Depozito - {tutar or 'genel'}", ""
        return "Kiralık Kasa", f"Yıllık Kasa Ücreti - {tutar or 'genel'}", ""

    # Altın / Kıymetli Maden
    if "altin" in n or "kiymetli maden" in n or "ats" in n:
        label = f"Altın Transfer - {tutar}" if tutar else "Altın Transfer"
        return "Kıymetli Maden", label, kanal()

    # HGS
    if "hgs" in n:
        if "etiket" in n:
            return "HGS Etiket Bedeli", "HGS Etiket Bedeli", ""
        return "HGS", "HGS Kart Ücreti", ""

    # Şans Oyunu
    if "sans oyunu" in n or "piyango" in n:
        return "Şans Oyunu Ödemeleri Aracılık", "Şans Oyunu Ödemeleri Aracılık", kanal()

    # Fatura
    if "fatura" in n and "kredi" not in n and "ekstre" not in n:
        return "Fatura Ödemeleri - Kart", "Fatura Ödemeleri", kanal()

    # SGK
    if "sgk" in n:
        label = f"SGK - {tutar} TRY" if tutar else "SGK Prim Ödemeleri"
        return "SGK Prim Ödemeleri", label, kanal()

    # Vergi
    if "vergi" in n and "kredi" not in n:
        label = f"Vergi - {tutar} TRY" if tutar else "Vergi Tahsilat"
        return "Vergi Tahsilat Aracılık", label, kanal()

    # Aidat
    if "aidat" in n:
        label = f"Aidat - {tutar} TRY" if tutar else "Aidat Ödemeleri"
        return "Aidat Ödemeleri Aracılık", label, kanal()

    # Özel Okul
    if "okul" in n:
        label = f"Özel Okul - {tutar} TRY" if tutar else "Özel Okul Ödeme"
        return "Özel Okul Ödeme", label, kanal()

    # Telefon
    if "telefon" in n:
        label = f"Telefon - {tutar} TRY" if tutar else "Telefon Ödemeleri"
        return "Telefon Ödemeleri Aracılık", label, kanal()

    # Arşiv
    if "arsiv" in n or "arşiv" in n:
        return "Arşiv Araştırma Ücreti", "Arşiv Araştırma Ücreti", kanal()

    # Mevduat araştırma / referans mektubu
    if "referans" in n or "itibar" in n or "niyet" in n:
        return "Mevduat Araştırma", "Referans Mektubu -", kanal()
    if "hesap ozeti" in n or "hesap özeti" in n:
        return "Mevduat Araştırma", "Hesap Özeti Verilmesi -", kanal()
    if "hesap arastirma" in n or "hesap araştırma" in n:
        return "Mevduat Araştırma", "Hesap Araştırma Talebi -", kanal()
    if "borcu yok" in n:
        return "Mevduat Araştırma", "Borcu Yoktur Yazısı", kanal()
    if "vize" in n and "okul" in n:
        return "Mevduat Araştırma", "Vize ve Özel Okullar için Düzenlenen Mektuplar -", kanal()

    # Bakiye sorma ATM
    if "bakiye" in n and "atm" in n:
        if "yurtici" in n or "yurt ici" in n or ("yurt" in n and "dis" not in n):
            return "Bakiye Sorma - Yurtiçi - ATM", "Bakiye Sorma - Yurtiçi - ATM", ""
        return "Bakiye Sorma - Yurtdışı - ATM", "Bakiye Sorma - Yurtdışı - ATM", ""

    # KKB / Kredi Risk
    if "kkb" in n or ("kredi" in n and "risk" in n) or ("ucuncu" in n and "rapor" in n):
        return "Üçüncü Kişi ve Kuruluşlardan Temin Edilecek Rapor Ücretleri - Kredi Risk Raporu",
               "Üçüncü Kişi ve Kuruluşlardan Temin Edilecek Rapor Ücretleri - Kredi Risk Raporu", kanal()

    # Çek Defteri
    if ("cek" in n or "çek" in n) and ("defteri" in n or "yaprak" in n or "teslim" in n):
        return "Çek Defteri ve Çek Düzenleme Ücreti", "Çek Defteri (Yaprak Başı) -", ""
    if ("cek" in n or "çek" in n) and "duzenleme" in n:
        if "ozel" in n or "özel" in n or "nitelik" in n:
            return "Çek Defteri ve Çek Düzenleme Ücreti", "Özel Nitelikli Çek Düzenleme -", ""
        return "Çek Defteri ve Çek Düzenleme Ücreti", "Çek Düzenleme -", ""

    # Çek İade
    if ("cek" in n or "çek" in n) and "iade" in n:
        return "Çek İade Ücreti", "Çek İade Ücreti", ""

    # Çek Tahsilat
    if ("cek" in n or "çek" in n) and ("tahsil" in n or "odeme" in n or "ödeme" in n):
        if "ayni banka" in n or "aynı banka" in n:
            return "Çek Tahsilat Ücreti", "Aynı Banka Çeki -", ""
        if "diger banka" in n or "diğer banka" in n or "baska banka" in n or "başka banka" in n:
            return "Çek Tahsilat Ücreti", "Diğer Banka Çeki -", ""
        if "doviz" in n or "döviz" in n or "yp" in n:
            return "Çek Tahsilat Ücreti", "Döviz Çekleri Tahsilatı (Diğer Banka) -", ""
        return "Çek Tahsilat Ücreti", "Çek Tahsilat", ""

    # Çek Belgelendirme
    if ("cek" in n or "çek" in n) and ("belgelend" in n or "karsiliksiz" in n or "karşılıksız" in n or "duzeltme" in n):
        if "karsiliksiz" in n or "karşılıksız" in n:
            return "Çek Belgelendirme ve Düzeltme Ücreti", "Karşılıksız Çek Belgelendirme -", ""
        return "Çek Belgelendirme ve Düzeltme Ücreti", "Çek Düzeltme Hakkı -", ""

    # Senet İade
    if "senet" in n and "iade" in n:
        return "Senet İade Ücreti", "Senet İade Ücreti", ""

    # Senet Protesto
    if "senet" in n and "protesto" in n:
        if "kaldir" in n or "kaldır" in n:
            return "Senet Protesto İşlemleri Ücreti", "Senet Protesto Kaldırma -", ""
        return "Senet Protesto İşlemleri Ücreti", "Senet Protesto -", ""

    # Senet Tahsil
    if "senet" in n and ("tahsil" in n or "tahsile" in n):
        if "ayni" in n or "aynı" in n:
            return "Senet Tahsile Alma Ücreti", "Aynı Banka Senet Tahsili -", ""
        return "Senet Tahsile Alma Ücreti", "Muhabir Banka Senet Tahsili -", ""

    return None, None, ""


KATEGORI_SIRA = [
    "EFT Gönderimi",
    "Havale Gönderimi",
    "FAST",
    "Kiralık Kasa",
    "Kıymetli Maden",
    "Üçüncü Kişi ve Kuruluşlardan Temin Edilecek Rapor Ücretleri - Kredi Risk Raporu",
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

    # ── Başlık satırları ──
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
        for j, kanal in enumerate(["Mobil", "Şube"]):
            c2 = ws.cell(row=2, column=col + j)
            c2.value = kanal
            c2.fill = fill
            c2.font = Font(color=r["fg"], bold=True, size=10)
            c2.alignment = Alignment(horizontal="center", vertical="center")
            sb(c2)
        col += 2

    ws.column_dimensions["A"].width = 42
    for i in range(2, MAX_COL + 1):
        ws.column_dimensions[get_column_letter(i)].width = 18
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 18
    ws.freeze_panes = "A3"

    # ── Veriyi işle ──
    # kat → [(standart_isim, kanal_adi)] sırası
    kat_satirlar: Dict[str, List[str]] = {}
    # (kat, standart_isim) → {banka: {kanal: deger}}
    veri: Dict[Tuple[str,str], Dict[str, Dict[str,str]]] = {}

    for banka in BANKALAR:
        for s in banka_verileri.get(banka, []):
            kat, isim, kanal = _standart_anahtar(s.masraf)
            if kat is None:
                continue
            if kanal not in ("mobil", "sube"):
                # scraper'dan gelen kanal bilgisine bak
                k2 = (s.kanal or "").lower()
                kanal = k2 if k2 in ("mobil", "sube") else "mobil"

            key = (kat, isim)
            if kat not in kat_satirlar:
                kat_satirlar[kat] = []
            if isim not in kat_satirlar[kat]:
                kat_satirlar[kat].append(isim)

            if key not in veri:
                veri[key] = {}
            if banka not in veri[key]:
                veri[key][banka] = {"mobil": "", "sube": ""}

            d = _deger(s)
            if d:
                veri[key][banka][kanal] = d

    # ── Satırları yaz ──
    KAT_FONT = Font(bold=True, size=10)
    DATA_FONT = Font(size=10)
    MASRAF_FONT = Font(size=10)

    yazilan = set()
    row = 3
    toplam = 0

    def yaz_kategori_blogu(kat):
        nonlocal row, toplam
        if kat not in kat_satirlar:
            return
        yazilan.add(kat)

        # Kategori başlık
        ws.merge_cells(f"A{row}:{get_column_letter(MAX_COL)}{row}")
        cc = ws[f"A{row}"]
        cc.value = kat
        cc.font = KAT_FONT
        cc.alignment = Alignment(horizontal="center", vertical="center")
        sb(cc)
        ws.row_dimensions[row].height = 16
        row += 1

        for isim in kat_satirlar[kat]:
            key = (kat, isim)
            mc = ws.cell(row=row, column=1, value=isim)
            mc.font = MASRAF_FONT
            mc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
            sb(mc)

            col = 2
            degerler = []
            for banka in BANKALAR:
                bd = veri.get(key, {}).get(banka, {"mobil": "", "sube": ""})
                degerler.append(bd.get("mobil", ""))
                degerler.append(bd.get("sube", ""))

            # Sarı renk: aynı kanalda bankalar arasında fark var mı?
            mobil_degerler = [degerler[i] for i in range(0, 8, 2) if degerler[i]]
            sube_degerler  = [degerler[i] for i in range(1, 8, 2) if degerler[i]]
            mobil_fark = len(set(mobil_degerler)) > 1
            sube_fark  = len(set(sube_degerler)) > 1

            for i, d in enumerate(degerler):
                c = ws.cell(row=row, column=col, value=d)
                c.font = DATA_FONT
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                sb(c)
                # Sarı: o sütun farklıysa
                if i % 2 == 0 and mobil_fark and d:
                    c.fill = YELLOW
                    c.font = Font(size=10, color="FF0000", bold=True)
                elif i % 2 == 1 and sube_fark and d:
                    c.fill = YELLOW
                    c.font = Font(size=10, color="FF0000", bold=True)
                col += 1

            ws.row_dimensions[row].height = 20
            row += 1
            toplam += 1

        row += 1  # kategori arası boş satır

    for kat in KATEGORI_SIRA:
        yaz_kategori_blogu(kat)

    for kat in list(kat_satirlar.keys()):
        if kat not in yazilan:
            yaz_kategori_blogu(kat)

    ws.cell(row=row, column=1,
            value=f"Son güncelleme: {_bugun()}").font = Font(size=9, color="888888")

    wb.save(dosya_yolu)
    print(f"[excel] {dosya_yolu} kaydedildi. {toplam} satır yazıldı.")
