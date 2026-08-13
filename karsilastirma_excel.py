"""
Bankalardan çekilen verileri karşılaştırmalı Excel formatında yazan modül.
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
    # Sadece gerçek değer varsa ekle, "-" veya boş ise ekleme
    parts = []
    if t and t != "-":
        parts.append(t)
    if o and o != "-":
        parts.append(o)
    return " / ".join(parts) if parts else ""


def _norm(m: str) -> str:
    m = m.lower().strip()
    for a, b in [("ı","i"),("ğ","g"),("ü","u"),("ş","s"),("ö","o"),("ç","c"),
                 ("â","a"),("î","i"),("û","u"),("i̇","i")]:
        m = m.replace(a, b)
    return re.sub(r"\s+", " ", m).strip()


def _extract_tutar(m: str) -> Optional[str]:
    m2 = m.replace("–", "-").replace("—", "-")
    m2 = re.sub(r"[.,\s]", "", m2)
    patterns = [
        (r"0[-]8300",        "0-8300"),
        (r"8300[-]399000",   "8300-399000"),
        (r"399000",          "399000+"),
        (r"buyuk|büyük",     "büyük"),
        (r"orta(?!k)",       "orta"),
        (r"kucuk|küçük",     "küçük"),
        (r"1[-]10g",         "1-10gr"),
        (r"11[-]100g",       "11-100gr"),
    ]
    for pat, label in patterns:
        if re.search(pat, m2, re.IGNORECASE):
            return label
    return None


def _norm_deger(d: str) -> str:
    """Karşılaştırma için değeri normalize et."""
    if not d:
        return ""
    d = d.strip().upper()
    # Türkçe ondalık virgül → nokta (sadece rakamlar arasında)
    d = re.sub(r"(\d),(\d)", r"\1.\2", d)
    # Boşlukları kaldır
    d = d.replace(" ", "")
    # TL → TRY
    d = re.sub(r"\bTL\b", "TRY", d)
    d = d.replace("TL", "TRY")
    # Sadece anlamlı karakterler
    d = re.sub(r"[^0-9A-Z%./]", "", d)
    # Float normalize: 7.97 == 7.970 == 7,97
    try:
        num = re.search(r"[\d.]+", d)
        if num:
            val = float(num.group())
            d = d[:num.start()] + f"{val:.2f}" + d[num.end():]
    except Exception:
        pass
    return d


def _standart_anahtar(masraf: str) -> Tuple[str, str, str]:
    yk_kat = ""
    if " | " in masraf:
        parts = masraf.split(" | ", 1)
        yk_kat = _norm(parts[0])
        masraf = parts[1]

    n = _norm(masraf)
    tutar = _extract_tutar(masraf)

    def kanal():
        if any(k in n for k in ["mobil","internet","iscep","online","dijital"]):
            return "mobil"
        if any(k in n for k in ["sube","şube","gise","gişe","cozum","çözüm","iletisim"]):
            return "sube"
        if "atm" in n:
            return "mobil"
        return ""

    if yk_kat:
        if "eft" in yk_kat or "fast" in yk_kat:
            if "fast" in yk_kat:
                label = f"FAST - {tutar} TRY" if tutar else "FAST"
                return "FAST", label, kanal()
            label = f"EFT - {tutar} TRY" if tutar else "EFT Gönderimi"
            return "EFT Gönderimi", label, kanal()
        if "havale" in yk_kat and "atm" not in yk_kat:
            if "duzenli" in n or "talimat" in n:
                label = f"Düzenli Havale - {tutar} TRY" if tutar else "Düzenli Havale"
                return "Havale Gönderimi", label, kanal()
            label = f"Havale - {tutar} TRY" if tutar else "Havale Gönderimi"
            return "Havale Gönderimi", label, kanal()
        if "atm" in yk_kat and "havale" in yk_kat:
            label = f"Havale - {tutar} TRY" if tutar else "Havale Gönderimi"
            return "Havale Gönderimi", label, "mobil"
        if "kiralik kasa" in yk_kat or "kiralık kasa" in yk_kat:
            if "depozito" in n:
                return "Kiralık Kasa", f"Kasa Depozito - {tutar or n}", ""
            return "Kiralık Kasa", f"Yıllık Kasa Ücreti - {tutar or n}", ""
        if "senet" in yk_kat:
            if "iade" in n:
                return "Senet İade Ücreti", "Senet İade Ücreti", ""
            if "protesto" in n and "kaldir" in n:
                return "Senet Protesto İşlemleri Ücreti", "Senet Protesto Kaldırma -", ""
            if "protesto" in n:
                return "Senet Protesto İşlemleri Ücreti", "Senet Protesto -", ""
            if "bankamizda" in n or "bankamızda" in n:
                return "Senet Tahsile Alma Ücreti", "Aynı Banka Senet Tahsili -", ""
            return "Senet Tahsile Alma Ücreti", "Muhabir Banka Senet Tahsili -", ""
        if "cek" in yk_kat or "çek" in yk_kat:
            if "ayni sube" in n or "aynı şube" in n:
                return "Çek Tahsilat Ücreti", "Aynı Banka Çeki -", ""
            if "baska sube" in n or "başka şube" in n:
                return "Çek Tahsilat Ücreti", "Diğer Banka Çeki -", ""
            if "karsiliksiz" in n:
                return "Çek Belgelendirme ve Düzeltme Ücreti", "Karşılıksız Çek Belgelendirme -", ""
            if "iade" in n:
                return "Çek İade Ücreti", "Çek İade Ücreti", ""
            if "yaprak" in n or "defteri" in n:
                return "Çek Defteri ve Çek Düzenleme Ücreti", "Çek Defteri (Yaprak Başı) -", ""
            if "duzenleme" in n:
                return "Çek Defteri ve Çek Düzenleme Ücreti", "Çek Düzenleme -", ""
            return "Çek Tahsilat Ücreti", "Çek Tahsilat", ""
        if "ortak atm" in yk_kat:
            if "bakiye" in n:
                return "Bakiye Sorma - Yurtiçi - ATM", "Bakiye Sorma - Yurtiçi - ATM", ""
        if "kurum tahsilat" in yk_kat or "kredi karti islem" in yk_kat:
            if "fatura" in n:
                return "Fatura Ödemeleri - Kart", "Fatura Ödemeleri", kanal()
            if "sgk" in n or "prim" in n:
                return "SGK Prim Ödemeleri", "SGK Prim Ödemeleri", kanal()
        if "findeks" in yk_kat or ("kredi" in yk_kat and "risk" in yk_kat):
            return ("Üçüncü Kişi ve Kuruluşlardan Temin Edilecek Rapor Ücretleri - Kredi Risk Raporu",
                    "Üçüncü Kişi ve Kuruluşlardan Temin Edilecek Rapor Ücretleri - Kredi Risk Raporu",
                    kanal())
        if "referans" in yk_kat:
            return "Mevduat Araştırma", "Referans Mektubu -", kanal()
        return None, None, ""

    if "eft" in n and "swift" not in n and "uluslararasi" not in n:
        if "duzenli" in n or "supurme" in n:
            label = f"Düzenli EFT - {tutar} TRY" if tutar else "Düzenli EFT"
            return "EFT Gönderimi", label, kanal()
        if "odenmesi" in n or "isme gelen" in n:
            return "EFT Gönderimi", "EFT İsme Gelen", "sube"
        label = f"EFT - {tutar} TRY" if tutar else "EFT Gönderimi"
        return "EFT Gönderimi", label, kanal()

    if "havale" in n and "uluslararasi" not in n and "swift" not in n:
        if "duzenli" in n or "supurme" in n:
            label = f"Düzenli Havale - {tutar} TRY" if tutar else "Düzenli Havale"
            return "Havale Gönderimi", label, kanal()
        if "cebe para" in n:
            return "Havale Gönderimi", "Cebe Para Gönderme", kanal()
        label = f"Havale - {tutar} TRY" if tutar else "Havale Gönderimi"
        return "Havale Gönderimi", label, kanal()

    if "fast" in n:
        label = f"FAST - {tutar} TRY" if tutar else "FAST"
        return "FAST", label, kanal()

    if "kasa" in n:
        if "depozito" in n:
            return "Kiralık Kasa", f"Kasa Depozito - {tutar or 'genel'}", ""
        return "Kiralık Kasa", f"Yıllık Kasa Ücreti - {tutar or 'genel'}", ""

    if "altin" in n or "kiymetli maden" in n or "ats" in n:
        label = f"Altın Transfer - {tutar}" if tutar else "Altın Transfer"
        return "Kıymetli Maden", label, kanal()

    if "hgs" in n:
        if "etiket" in n:
            return "HGS Etiket Bedeli", "HGS Etiket Bedeli", ""
        return "HGS", "HGS Kart Ücreti", ""

    if "sans oyunu" in n or "piyango" in n:
        return "Şans Oyunu Ödemeleri Aracılık", "Şans Oyunu Ödemeleri Aracılık", kanal()

    if "fatura" in n and "kredi" not in n and "ekstre" not in n:
        return "Fatura Ödemeleri - Kart", "Fatura Ödemeleri", kanal()

    if "sgk" in n:
        label = f"SGK - {tutar} TRY" if tutar else "SGK Prim Ödemeleri"
        return "SGK Prim Ödemeleri", label, kanal()

    if "vergi" in n and "kredi" not in n:
        label = f"Vergi - {tutar} TRY" if tutar else "Vergi Tahsilat"
        return "Vergi Tahsilat Aracılık", label, kanal()

    if "aidat" in n:
        label = f"Aidat - {tutar} TRY" if tutar else "Aidat Ödemeleri"
        return "Aidat Ödemeleri Aracılık", label, kanal()

    if "okul" in n:
        label = f"Özel Okul - {tutar} TRY" if tutar else "Özel Okul Ödeme"
        return "Özel Okul Ödeme", label, kanal()

    if "telefon" in n:
        label = f"Telefon - {tutar} TRY" if tutar else "Telefon Ödemeleri"
        return "Telefon Ödemeleri Aracılık", label, kanal()

    if "arsiv" in n:
        return "Arşiv Araştırma Ücreti", "Arşiv Araştırma Ücreti", kanal()

    if "referans" in n or "itibar" in n or "niyet" in n:
        return "Mevduat Araştırma", "Referans Mektubu -", kanal()
    if "hesap ozeti" in n:
        return "Mevduat Araştırma", "Hesap Özeti Verilmesi -", kanal()
    if "hesap arastirma" in n:
        return "Mevduat Araştırma", "Hesap Araştırma Talebi -", kanal()
    if "borcu yok" in n:
        return "Mevduat Araştırma", "Borcu Yoktur Yazısı", kanal()
    if "vize" in n and "okul" in n:
        return "Mevduat Araştırma", "Vize ve Özel Okullar için Düzenlenen Mektuplar -", kanal()

    if "bakiye" in n and "atm" in n:
        if "yurtici" in n or ("yurt" in n and "dis" not in n):
            return "Bakiye Sorma - Yurtiçi - ATM", "Bakiye Sorma - Yurtiçi - ATM", ""
        return "Bakiye Sorma - Yurtdışı - ATM", "Bakiye Sorma - Yurtdışı - ATM", ""

    if "kkb" in n or ("kredi" in n and "risk" in n) or ("ucuncu" in n and "rapor" in n):
        return ("Üçüncü Kişi ve Kuruluşlardan Temin Edilecek Rapor Ücretleri - Kredi Risk Raporu",
                "Üçüncü Kişi ve Kuruluşlardan Temin Edilecek Rapor Ücretleri - Kredi Risk Raporu",
                kanal())

    if ("cek" in n or "çek" in n) and ("defteri" in n or "yaprak" in n or "teslim" in n):
        return "Çek Defteri ve Çek Düzenleme Ücreti", "Çek Defteri (Yaprak Başı) -", ""
    if ("cek" in n or "çek" in n) and "duzenleme" in n:
        if "ozel" in n or "nitelik" in n:
            return "Çek Defteri ve Çek Düzenleme Ücreti", "Özel Nitelikli Çek Düzenleme -", ""
        return "Çek Defteri ve Çek Düzenleme Ücreti", "Çek Düzenleme -", ""

    if ("cek" in n or "çek" in n) and "iade" in n:
        return "Çek İade Ücreti", "Çek İade Ücreti", ""

    if ("cek" in n or "çek" in n) and ("tahsil" in n or "odeme" in n):
        if "ayni banka" in n or "aynı banka" in n:
            return "Çek Tahsilat Ücreti", "Aynı Banka Çeki -", ""
        if "diger banka" in n or "baska banka" in n:
            return "Çek Tahsilat Ücreti", "Diğer Banka Çeki -", ""
        if "doviz" in n or "yp" in n:
            return "Çek Tahsilat Ücreti", "Döviz Çekleri Tahsilatı (Diğer Banka) -", ""
        return "Çek Tahsilat Ücreti", "Çek Tahsilat", ""

    if ("cek" in n or "çek" in n) and ("belgelend" in n or "karsiliksiz" in n or "duzeltme" in n):
        if "karsiliksiz" in n:
            return "Çek Belgelendirme ve Düzeltme Ücreti", "Karşılıksız Çek Belgelendirme -", ""
        return "Çek Belgelendirme ve Düzeltme Ücreti", "Çek Düzeltme Hakkı -", ""

    if "senet" in n and "iade" in n:
        return "Senet İade Ücreti", "Senet İade Ücreti", ""

    if "senet" in n and "protesto" in n:
        if "kaldir" in n:
            return "Senet Protesto İşlemleri Ücreti", "Senet Protesto Kaldırma -", ""
        return "Senet Protesto İşlemleri Ücreti", "Senet Protesto -", ""

    if "senet" in n and ("tahsil" in n or "tahsile" in n):
        if "ayni" in n:
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
        ws.column_dimensions[get_column_letter(i)].width = 18
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 18
    ws.freeze_panes = "A3"

    kat_satirlar: Dict[str, List[str]] = {}
    veri: Dict[Tuple[str, str], Dict[str, Dict[str, str]]] = {}

    for banka in BANKALAR:
        for s in banka_verileri.get(banka, []):
            kat, isim, kanal = _standart_anahtar(s.masraf)
            if kat is None:
                continue
            if kanal not in ("mobil", "sube"):
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

    KAT_FONT = Font(bold=True, size=10)
    DATA_FONT = Font(size=10)
    RED_FONT = Font(size=10, color="FF0000", bold=True)
    MASRAF_FONT = Font(size=10)

    yazilan = set()
    row = 3
    toplam = 0

    def yaz_kategori_blogu(kat):
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

            # Sadece dolu değerleri karşılaştır
            mobil_norm = [_norm_deger(degerler[i]) for i in range(0, 8, 2) if degerler[i]]
            sube_norm  = [_norm_deger(degerler[i]) for i in range(1, 8, 2) if degerler[i]]
            mobil_fark = len(set(mobil_norm)) > 1
            sube_fark  = len(set(sube_norm)) > 1

            for i, d in enumerate(degerler):
                c = ws.cell(row=row, column=col, value=d)
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                sb(c)
                # Sarı YOK — sadece kırmızı yazı farklı değerlerde
                if i % 2 == 0 and mobil_fark and d:
                    c.font = RED_FONT
                elif i % 2 == 1 and sube_fark and d:
                    c.font = RED_FONT
                else:
                    c.font = DATA_FONT
                col += 1

            ws.row_dimensions[row].height = 20
            row += 1
            toplam += 1

        row += 1

    for kat in KATEGORI_SIRA:
        yaz_kategori_blogu(kat)

    for kat in list(kat_satirlar.keys()):
        if kat not in yazilan:
            yaz_kategori_blogu(kat)

    ws.cell(row=row, column=1,
            value=f"Son güncelleme: {_bugun()}").font = Font(size=9, color="888888")

    wb.save(dosya_yolu)
    print(f"[excel] {dosya_yolu} kaydedildi. {toplam} satır yazıldı.")
