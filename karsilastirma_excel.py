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

BANKALAR = ["GARANTİ", "İŞBANKASI", "AKBANK"]
BANKA_TAMAD = {
    "GARANTİ":   "Garanti BBVA",
    "İŞBANKASI": "İş Bankası",
    "AKBANK":    "Akbank",
}

# Başlık (satır 1) renkleri
BANKA_RENKLER = {
    "GARANTİ":   {"bg": "00B050", "fg": "FFFFFF"},
    "İŞBANKASI": {"bg": "012169", "fg": "FFFFFF"},
    "AKBANK":    {"bg": "FF0000", "fg": "FFFFFF"},
}

# Alt başlık (Mobil/Şube - satır 2) renkleri
BANKA_ALT_RENKLER = {
    "GARANTİ":   {"bg": "C6E0B4", "fg": "FFFFFF"},
    "İŞBANKASI": {"bg": "BDD7EE", "fg": "FFFFFF"},
    "AKBANK":    {"bg": "F4B183", "fg": "FFFFFF"},
}

# Hücre veri metni rengi - bankaya göre
BANKA_VERI_RENK = {
    "GARANTİ":   "00B050",
    "İŞBANKASI": "0070C0",
    "AKBANK":    "FF0000",
}

thin = Side(style="thin", color="CCCCCC")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
NOT_KOLONU = 8  # H sütunu - notlar buraya yazılır


def _bugun() -> str:
    tz = timezone(timedelta(hours=3))
    return datetime.now(tz).strftime("%d.%m.%Y %H:%M")


def _deger(s: UcretSatiri) -> str:
    t = (s.asgari_tutar or "").strip()
    o = (s.asgari_oran  or "").strip()
    parts = []
    if t: parts.append(t)
    if o: parts.append(o)
    return " / ".join(parts) if parts else ""


def _norm(m: str) -> str:
    m = m.lower().strip()
    for a, b in [("ı","i"),("ğ","g"),("ü","u"),("ş","s"),("ö","o"),("ç","c"),
                 ("â","a"),("î","i"),("û","u"),("i̇","i")]:
        m = m.replace(a, b)
    return re.sub(r"\s+", " ", m).strip()


def _extract_tutar(m: str) -> Optional[str]:
    """
    Metindeki tutar aralığını standart bir etikete çevirir.
    """
    t = m.replace("–", "-").replace("—", "-").lower()
    tn = _norm(m)

    for pat, label in [
        (r"buyuk|büyük", "büyük"),
        (r"orta(?!k)", "orta"),
        (r"kucuk|küçük", "küçük"),
        (r"1[-\s]*10\s*g", "1-10gr"),
        (r"11[-\s]*100\s*g", "11-100gr"),
    ]:
        if re.search(pat, t, re.IGNORECASE):
            return label

    nums = []
    for raw in re.findall(r"\d{1,3}(?:\.\d{3})*(?:,\d+)?", t):
        cleaned = re.sub(r"\.(?=\d{3}(\D|$))", "", raw)
        cleaned = re.sub(r",\d+$", "", cleaned)
        if cleaned.isdigit():
            nums.append(int(cleaned))
    if not nums:
        return None

    ust_ifade = any(k in tn for k in ["ve ustu", "ve uzeri", "uzeri", "ustu"])
    alt_ifade = any(k in tn for k in ["ve alti", "altinda", "ve altinda"])

    if ust_ifade:
        return "399000+"
    if alt_ifade:
        return "0-8300"

    if len(nums) >= 2:
        ilk = nums[0]
        if ilk <= 0:
            return "0-8300"
        if ilk >= 300000:
            return "399000+"
        return "8300-399000"

    ilk = nums[0]
    if ilk <= 0:
        return "0-8300"
    if ilk >= 300000:
        return "399000+"
    if ilk <= 8300:
        return "0-8300"
    return "8300-399000"


def _norm_deger(d: str) -> str:
    if not d:
        return ""
    d = d.strip().upper()
    d = re.sub(r"(\d),(\d)", r"\1.\2", d)
    d = d.replace(" ", "").replace("TL", "TRY")
    d = re.sub(r"[^0-9A-Z%./]", "", d)
    try:
        num = re.search(r"[\d.]+", d)
        if num:
            val = float(num.group())
            d = d[:num.start()] + f"{val:.2f}" + d[num.end():]
    except Exception:
        pass
    return d


def _standart_anahtar(masraf: str) -> Tuple[str, str, str]:
    """
    Masraf adını standard kategorilere ve alt kategorilere eşleştir.
    Tüm bankalar arasında tutarlı karşılaştırma yapabilmek için.
    """
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

    # Kategori + Kanal + Tutar ile başlık oluştur
    
    # EFT / FAST
    if "eft" in n or "fast" in n:
        if "fast" in n:
            label = f"FAST - {tutar} TRY" if tutar else "FAST"
            return "FAST", label, kanal()
        if "duzenli" in n or "supurme" in n:
            label = f"Düzenli EFT - {tutar} TRY" if tutar else "Düzenli EFT"
            return "EFT Gönderimi", label, kanal()
        label = f"EFT - {tutar} TRY" if tutar else "EFT Gönderimi"
        return "EFT Gönderimi", label, kanal()

    # HAVALE
    if "havale" in n and "swift" not in n and "uluslararasi" not in n:
        if "duzenli" in n or "supurme" in n:
            label = f"Düzenli Havale - {tutar} TRY" if tutar else "Düzenli Havale"
            return "Havale Gönderimi", label, kanal()
        label = f"Havale - {tutar} TRY" if tutar else "Havale Gönderimi"
        return "Havale Gönderimi", label, kanal()

    # SWIFT / Uluslararası
    if "swift" in n or "uluslararasi" in n:
        label = f"Uluslararası Transfer - {tutar}" if tutar else "Uluslararası Transfer"
        return "Uluslararası Transfer", label, kanal()

    # KIYMETLİ MADEN / ALTIN
    if "altin" in n or "kiymetli maden" in n or "ats" in n:
        label = f"Altın Transfer - {tutar}" if tutar else "Altın Transfer"
        return "Kıymetli Maden", label, kanal()

    # KIRALıK KASA
    if "kasa" in n:
        if "depozito" in n:
            return "Kiralık Kasa", f"Kasa Depozito - {tutar or 'genel'}", ""
        return "Kiralık Kasa", f"Yıllık Kasa Ücreti - {tutar or 'genel'}", ""

    # HGS
    if "hgs" in n:
        if "etiket" in n:
            return "HGS Etiket Bedeli", "HGS Etiket Bedeli", ""
        return "HGS", "HGS Kart Ücreti", ""

    # ÇEK
    if "cek" in n or "çek" in n:
        if "defteri" in n or "yaprak" in n or "teslim" in n:
            return "Çek Defteri ve Çek Düzenleme Ücreti", "Çek Defteri (Yaprak Başı) -", ""
        if "duzenleme" in n:
            return "Çek Defteri ve Çek Düzenleme Ücreti", "Çek Düzenleme -", ""
        if "iade" in n:
            return "Çek İade Ücreti", "Çek İade Ücreti", ""
        if "tahsil" in n or "odeme" in n:
            if "ayni" in n or "aynı" in n:
                return "Çek Tahsilat Ücreti", "Aynı Banka Çeki -", ""
            if "diger" in n or "başka" in n or "baska" in n:
                return "Çek Tahsilat Ücreti", "Diğer Banka Çeki -", ""
            return "Çek Tahsilat Ücreti", "Çek Tahsilat", ""
        if "belgelend" in n or "karsiliksiz" in n:
            return "Çek Belgelendirme ve Düzeltme Ücreti", "Çek Belgelendirme -", ""
        return "Çek Tahsilat Ücreti", "Çek Tahsilat", ""

    # SENET
    if "senet" in n:
        if "iade" in n:
            return "Senet İade Ücreti", "Senet İade Ücreti", ""
        if "protesto" in n:
            if "kaldir" in n:
                return "Senet Protesto İşlemleri Ücreti", "Senet Protesto Kaldırma -", ""
            return "Senet Protesto İşlemleri Ücreti", "Senet Protesto -", ""
        if "tahsil" in n or "tahsile" in n:
            return "Senet Tahsile Alma Ücreti", "Senet Tahsile Alma -", ""
        return "Senet İade Ücreti", "Senet İade Ücreti", ""

    # FATURA / SGK / VERGİ / AIDAT
    if "fatura" in n and "kredi" not in n:
        return "Fatura Ödemeleri", "Fatura Ödemeleri", kanal()
    
    if "sgk" in n or "prim" in n:
        return "SGK Prim Ödemeleri", "SGK Prim Ödemeleri", kanal()
    
    if "vergi" in n:
        return "Vergi Tahsilat", "Vergi Tahsilat", kanal()
    
    if "aidat" in n:
        return "Aidat Ödemeleri", "Aidat Ödemeleri", kanal()

    # ŞANS OYUNU
    if "sans oyunu" in n or "piyango" in n:
        return "Şans Oyunu Ödemeleri", "Şans Oyunu Ödemeleri", kanal()

    # TELEFON
    if "telefon" in n:
        return "Telefon Ödemeleri", "Telefon Ödemeleri", kanal()

    # OKUL
    if "okul" in n:
        return "Özel Okul Ödeme", "Özel Okul Ödeme", kanal()

    # BAKIYE SORMA
    if "bakiye" in n and "atm" in n:
        if "yurtici" in n or ("yurt" in n and "dis" not in n):
            return "Bakiye Sorma - Yurtiçi - ATM", "Bakiye Sorma - Yurtiçi - ATM", ""
        return "Bakiye Sorma - Yurtdışı - ATM", "Bakiye Sorma - Yurtdışı - ATM", ""

    # KREDİ RISK RAPORU / KKB
    if "kkb" in n or ("kredi" in n and "risk" in n):
        return "Kredi Risk Raporu", "Kredi Risk Raporu", kanal()

    # REFERANS / ARŞIV
    if "referans" in n or "itibar" in n or "niyet" in n:
        return "Referans Mektubu", "Referans Mektubu", kanal()
    
    if "arsiv" in n:
        return "Arşiv Araştırma", "Arşiv Araştırma", kanal()
    
    if "hesap" in n and "ozeti" in n:
        return "Hesap Özeti", "Hesap Özeti", kanal()

    # Eşleştirilemeyenler None dön
    return None, None, ""


KATEGORI_SIRA = [
    "EFT Gönderimi",
    "Havale Gönderimi",
    "FAST",
    "Uluslararası Transfer",
    "Kıymetli Maden",
    "Kiralık Kasa",
    "HGS Etiket Bedeli",
    "Çek Defteri ve Çek Düzenleme Ücreti",
    "Çek İade Ücreti",
    "Çek Tahsilat Ücreti",
    "Çek Belgelendirme ve Düzeltme Ücreti",
    "Senet İade Ücreti",
    "Senet Protesto İşlemleri Ücreti",
    "Senet Tahsile Alma Ücreti",
    "Fatura Ödemeleri",
    "SGK Prim Ödemeleri",
    "Vergi Tahsilat",
    "Aidat Ödemeleri",
    "Şans Oyunu Ödemeleri",
    "Telefon Ödemeleri",
    "Özel Okul Ödeme",
    "Bakiye Sorma - Yurtiçi - ATM",
    "Bakiye Sorma - Yurtdışı - ATM",
    "Kredi Risk Raporu",
    "Referans Mektubu",
    "Arşiv Araştırma",
    "Hesap Özeti",
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
    MAX_COL = 7  # A..G (A=masraf, B-G=3 banka x 2 kanal)

    def sb(cell):
        cell.border = BORDER

    # ---- A1 boş üst-sol köşe ----
    ws.merge_cells("A1:A2")
    c = ws["A1"]
    c.fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    c.font = Font(color="FFFFFF", bold=True, size=11)
    c.alignment = Alignment(horizontal="center", vertical="center")
    sb(c)

    # ---- Banka başlıkları (satır 1) + Mobil/Şube alt başlıkları (satır 2, italik) ----
    col = 2
    for banka in BANKALAR:
        r = BANKA_RENKLER[banka]
        ar = BANKA_ALT_RENKLER[banka]
        fill = PatternFill(start_color=r["bg"], end_color=r["bg"], fill_type="solid")
        alt_fill = PatternFill(start_color=ar["bg"], end_color=ar["bg"], fill_type="solid")
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
            c2.fill = alt_fill
            c2.font = Font(color=ar["fg"], bold=True, italic=True, size=10)
            c2.alignment = Alignment(horizontal="center", vertical="center")
            sb(c2)
        col += 2

    ws.column_dimensions["A"].width = 45
    for i in range(2, MAX_COL + 1):
        ws.column_dimensions[get_column_letter(i)].width = 14
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
    MASRAF_FONT = Font(size=10)

    yazilan = set()
    row = 3
    toplam = 0

    def yaz_kategori_blogu(kat):
        nonlocal row, toplam

        if kat not in kat_satirlar:
            return
        yazilan.add(kat)

        # Kategori başlığı
        cc = ws.cell(row=row, column=1, value=kat)
        cc.font = KAT_FONT
        cc.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 18
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
                degerler.append((banka, "mobil", bd.get("mobil", "")))
                degerler.append((banka, "sube",  bd.get("sube", "")))

            mobil_norm = [_norm_deger(v.split("\n")[0]) for (_, k, v) in degerler if k == "mobil" and v]
            sube_norm  = [_norm_deger(v.split("\n")[0]) for (_, k, v) in degerler if k == "sube"  and v]
            mobil_fark = len(set(mobil_norm)) > 1
            sube_fark  = len(set(sube_norm)) > 1

            for i, (banka, kanal, d) in enumerate(degerler):
                cell_value = d if d else ""
                c = ws.cell(row=row, column=col, value=cell_value)
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                sb(c)

                fark = mobil_fark if kanal == "mobil" else sube_fark
                if fark and d:
                    c.fill = YELLOW
                    c.font = Font(size=10, color="FF0000", bold=True)
                else:
                    c.font = Font(size=10, color=BANKA_VERI_RENK[banka],
                                  bold=(kanal == "mobil"))
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
