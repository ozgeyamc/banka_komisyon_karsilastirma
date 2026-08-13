"""
Banka komisyon karşılaştırma - Excel çıktısı.
Her masraf için kategori+tutar aralığı anahtarıyla eşleştirme yapılır.
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

# Tutar aralığı → Excel'de görünecek isim
TUTAR_ETIKET = {
    "0-8300":      "0 - 8.300 TL",
    "8300-399000": "8.300,01 - 399.000 TL",
    "399000+":     "399.000,01 TL ve Üzeri",
    "buyuk":       "Büyük Boy",
    "orta":        "Orta Boy",
    "kucuk":       "Küçük Boy",
    "1-10gr":      "1 - 10 gr",
    "11-100gr":    "11 - 100 gr",
}

# Ana kategori sırası
KATEGORI_SIRA = [
    "EFT Gönderimi",
    "Havale Gönderimi",
    "FAST",
    "Kiralık Kasa",
    "Kıymetli Maden Transferi",
    "Kredi Risk Raporu",
    "Fatura Ödemeleri",
    "SGK Prim Ödemeleri",
    "HGS",
    "Şans Oyunu Ödemeleri",
    "Aidat Ödemeleri",
    "Özel Okul Ödemeleri",
    "Telefon Ödemeleri",
    "Vergi Tahsilat",
    "Arşiv Araştırma",
    "Mevduat Araştırma",
    "Bakiye Sorgulama - ATM",
    "Çek Defteri ve Düzenleme",
    "Çek İade",
    "Çek Tahsilat",
    "Çek Belgelendirme",
    "Senet İade",
    "Senet Protesto",
    "Senet Tahsil",
]


def _bugun() -> str:
    tz = timezone(timedelta(hours=3))
    return datetime.now(tz).strftime("%d.%m.%Y %H:%M")


def _deger(s: UcretSatiri) -> str:
    t = (s.asgari_tutar or "").strip()
    o = (s.asgari_oran or "").strip()
    if t == "-": t = ""
    if o == "-": o = ""
    parts = [x for x in [t, o] if x]
    return " / ".join(parts)


def _norm(m: str) -> str:
    """Türkçe karakterleri ASCII'ye çevir, küçük harf yap."""
    m = m.lower().strip()
    for a, b in [("ı","i"),("ğ","g"),("ü","u"),("ş","s"),("ö","o"),("ç","c"),
                 ("â","a"),("î","i"),("û","u")]:
        m = m.replace(a, b)
    m = m.replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", m).strip()


def _tutar_kodu(metin: str) -> Optional[str]:
    """Metinden tutar aralığı kodunu çıkar."""
    # Nokta/virgül/boşluk/tire-tipi temizle
    m = _norm(metin)
    m2 = re.sub(r"[.,\s]", "", m)

    if re.search(r"0[-]8300|8300tl?vealt|8\.?300tl?vealt|1try[-–]8", m2):
        return "0-8300"
    if re.search(r"8300[-]399000|8300[-–]399", m2):
        return "8300-399000"
    if re.search(r"399000|399\.000|399000tl?veuzeri|399000tl?uzer", m2):
        return "399000+"
    if "buyuk" in m2 or "büyük" in m or "b tipi" in m or "e tipi" in m:
        return "buyuk"
    if re.search(r"orta(?!k)", m2):
        return "orta"
    if "kucuk" in m2 or "küçük" in m or "a tipi" in m:
        return "kucuk"
    if re.search(r"1[-]10g", m2):
        return "1-10gr"
    if re.search(r"11[-]100g", m2):
        return "11-100gr"
    return None


def _ana_kategori_ve_tip(kategori_norm: str, masraf_norm: str) -> Optional[Tuple[str, str, str]]:
    """
    (ana_kategori, satir_isim, kanal) döndür.
    Hiçbir kurala uymuyorsa None döndür.
    kanal: "mobil" veya "sube"
    """
    k = kategori_norm
    m = masraf_norm

    # Kanal tespiti - masraf adından
    def kanal_bul(metin: str) -> str:
        if any(x in metin for x in ["mobil", "internet", "iscep", "dijital", "online"]):
            return "mobil"
        if any(x in metin for x in ["sube", "gise", "cozum merkezi", "iletisim merkezi", "telefon bankaciligi"]):
            return "sube"
        if "atm" in metin:
            return "mobil"
        return "mobil"  # varsayılan

    kanal = kanal_bul(m)

    # Tutar kodu - önce masraf adından, yoksa kategori adından
    tutar = _tutar_kodu(m) or _tutar_kodu(k)
    t_etiket = TUTAR_ETIKET.get(tutar, "") if tutar else ""

    # ── EFT ──
    if "eft" in k or ("eft" in m and "swift" not in m and "uluslararasi" not in m):
        if "fast" in k or "fast" in m:
            if not tutar:
                return None
            return "FAST", t_etiket, kanal
        if "isime gelen" in m or "isme gelen" in m or "odenmesi" in m:
            return "EFT Gönderimi", "İsme Gelen EFT", "sube"
        if "duzenli" in k or "duzenli" in m or "supurme" in m or "talimat" in m:
            if not tutar:
                return None
            return "EFT Gönderimi", f"Düzenli EFT - {t_etiket}", kanal
        if not tutar:
            return None
        return "EFT Gönderimi", t_etiket, kanal

    # ── FAST (kategori bazlı) ──
    if "fast" in k:
        if not tutar:
            return None
        return "FAST", t_etiket, "mobil"

    # ── Havale ──
    if "havale" in k or ("havale" in m and "uluslararasi" not in m and "swift" not in m):
        if "cebe para" in m:
            return "Havale Gönderimi", "Cebe Para Gönderme", kanal
        if "duzenli" in k or "duzenli" in m or "supurme" in m or "talimat" in m:
            if not tutar:
                return None
            return "Havale Gönderimi", f"Düzenli Havale - {t_etiket}", kanal
        if not tutar:
            return None
        return "Havale Gönderimi", t_etiket, kanal

    # ── Kiralık Kasa ──
    if "kasa" in k or ("kasa" in m and "kart" not in m):
        if not tutar:
            return None
        if "depozito" in m or "depozito" in k:
            return "Kiralık Kasa", f"Depozito - {t_etiket}", ""
        return "Kiralık Kasa", t_etiket, ""

    # ── Kıymetli Maden / Altın ──
    if "altin" in k or "kiymetli maden" in k or "ats" in k or \
       "altin" in m or ("kiymetli" in m and "maden" in m):
        if not tutar:
            return None
        return "Kıymetli Maden Transferi", t_etiket, kanal

    # ── HGS ──
    if "hgs" in k or "hgs" in m:
        return "HGS", "HGS Etiket/Kart Bedeli", ""

    # ── Şans Oyunu ──
    if "sans oyunu" in k or "sans oyunu" in m or "piyango" in m:
        return "Şans Oyunu Ödemeleri", "Şans Oyunu Ödemeleri", kanal

    # ── Fatura ──
    if ("fatura" in k or "fatura" in m) and "kredi" not in m and "ekstre" not in m:
        label = t_etiket if t_etiket else "Fatura Ödemeleri"
        return "Fatura Ödemeleri", label, kanal

    # ── SGK ──
    if "sgk" in k or "sgk" in m or ("prim" in m and "sigorta" not in m):
        label = t_etiket if t_etiket else "SGK Prim Ödemeleri"
        return "SGK Prim Ödemeleri", label, kanal

    # ── Vergi ──
    if ("vergi" in k or "vergi" in m) and "kredi" not in m:
        label = t_etiket if t_etiket else "Vergi Tahsilat"
        return "Vergi Tahsilat", label, kanal

    # ── Aidat ──
    if "aidat" in k or "aidat" in m:
        label = t_etiket if t_etiket else "Aidat Ödemeleri"
        return "Aidat Ödemeleri", label, kanal

    # ── Özel Okul ──
    if "okul" in k or "okul" in m:
        label = t_etiket if t_etiket else "Özel Okul Ödemeleri"
        return "Özel Okul Ödemeleri", label, kanal

    # ── Telefon ──
    if "telefon" in k or "telefon" in m:
        label = t_etiket if t_etiket else "Telefon Ödemeleri"
        return "Telefon Ödemeleri", label, kanal

    # ── Arşiv ──
    if "arsiv" in k or "arsiv" in m:
        return "Arşiv Araştırma", "Arşiv Araştırma Ücreti", kanal

    # ── Kredi Risk / KKB / Findeks ──
    if "kkb" in m or "findeks" in k or ("kredi" in m and "risk" in m) or \
       ("ucuncu" in m and "rapor" in m) or ("kredi risk" in k):
        return "Kredi Risk Raporu", "Kredi Risk Raporu", kanal

    # ── Mevduat Araştırma ──
    if "referans" in m or "itibar" in m or "niyet" in m:
        return "Mevduat Araştırma", "Referans Mektubu", kanal
    if "hesap ozeti" in m:
        return "Mevduat Araştırma", "Hesap Özeti", kanal
    if "hesap arastirma" in m:
        return "Mevduat Araştırma", "Hesap Araştırma", kanal
    if "borcu yok" in m:
        return "Mevduat Araştırma", "Borcu Yoktur Yazısı", kanal
    if "vize" in m and "okul" in m:
        return "Mevduat Araştırma", "Vize/Okul Mektubu", kanal
    if "mutabakat" in m or "teyit" in m:
        return "Mevduat Araştırma", "Mutabakat/Teyit Yazısı", kanal

    # ── Bakiye Sorgulama ATM ──
    if ("bakiye" in m or "limit sorgulama" in m) and "atm" in m:
        if "yurtdisi" in m or "yurt disi" in m:
            return "Bakiye Sorgulama - ATM", "Yurtdışı ATM Bakiye Sorgulama", ""
        return "Bakiye Sorgulama - ATM", "Yurtiçi ATM Bakiye Sorgulama", ""

    # ── Çek ──
    if "cek" in k or "cek" in m:
        if "defteri" in m or "yaprak" in m or "teslim" in m:
            return "Çek Defteri ve Düzenleme", "Çek Defteri (Yaprak Başı)", ""
        if "duzenleme" in m:
            if "ozel" in m or "nitelik" in m:
                return "Çek Defteri ve Düzenleme", "Özel Nitelikli Çek Düzenleme", ""
            return "Çek Defteri ve Düzenleme", "Çek Düzenleme", ""
        if "iade" in m:
            return "Çek İade", "Çek İade Ücreti", ""
        if "karsiliksiz" in m:
            return "Çek Belgelendirme", "Karşılıksız Çek Belgelendirme", ""
        if "duzeltme" in m or "belgelend" in m:
            return "Çek Belgelendirme", "Çek Düzeltme Hakkı", ""
        if "tahsil" in m or "odeme" in m:
            if "ayni" in m:
                return "Çek Tahsilat", "Aynı Banka Çeki", ""
            if "baska" in m or "diger" in m:
                return "Çek Tahsilat", "Diğer Banka Çeki", ""
            if "doviz" in m or "yp" in m:
                return "Çek Tahsilat", "Döviz Çeki Tahsilatı", ""
            return "Çek Tahsilat", "Çek Tahsilat", ""

    # ── Senet ──
    if "senet" in k or "senet" in m:
        if "iade" in m or "protestosuz" in m:
            return "Senet İade", "Senet İade Ücreti", ""
        if "protesto" in m:
            if "kaldir" in m:
                return "Senet Protesto", "Senet Protesto Kaldırma", ""
            return "Senet Protesto", "Senet Protesto Ücreti", ""
        if "tahsil" in m:
            if "bankamizda" in m or "ayni" in m:
                return "Senet Tahsil", "Aynı Banka Senet Tahsili", ""
            return "Senet Tahsil", "Muhabir Banka Senet Tahsili", ""

    return None


def _masrafi_isle(s: UcretSatiri) -> Optional[Tuple[str, str, str]]:
    """
    UcretSatiri'nden (ana_kategori, satir_isim, kanal) üret.
    Yapıkredi: masraf "KategoriAdı | MasrafAdı" formatında gelir.
    """
    masraf = s.masraf or ""
    kategori = s.kategori or ""

    # Yapıkredi prefix ayır
    if " | " in masraf:
        parca = masraf.split(" | ", 1)
        kategori = parca[0]
        masraf = parca[1]

    k = _norm(kategori)
    m = _norm(masraf)

    return _ana_kategori_ve_tip(k, m)


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
        for j, kanal_adi in enumerate(["Mobil", "Şube"]):
            c2 = ws.cell(row=2, column=col + j)
            c2.value = kanal_adi
            c2.fill = fill
            c2.font = Font(color=r["fg"], bold=True, size=10)
            c2.alignment = Alignment(horizontal="center", vertical="center")
            sb(c2)
        col += 2

    ws.column_dimensions["A"].width = 38
    for i in range(2, MAX_COL + 1):
        ws.column_dimensions[get_column_letter(i)].width = 16
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 18
    ws.freeze_panes = "A3"

    # ── Veriyi işle ──
    # kat_satirlar: kategori → [satir_isim sırası]
    kat_satirlar: Dict[str, List[str]] = {}
    # veri: (kategori, satir_isim) → {banka: {kanal: deger}}
    veri: Dict[Tuple[str, str], Dict[str, Dict[str, str]]] = {}

    for banka in BANKALAR:
        for s in banka_verileri.get(banka, []):
            sonuc = _masrafi_isle(s)
            if sonuc is None:
                continue
            kat, isim, kanal = sonuc
            if not isim:
                continue
            if kanal not in ("mobil", "sube"):
                k2 = (s.kanal or "").lower()
                kanal = k2 if k2 in ("mobil", "sube") else "mobil"

            key = (kat, isim)
            kat_satirlar.setdefault(kat, [])
            if isim not in kat_satirlar[kat]:
                kat_satirlar[kat].append(isim)

            veri.setdefault(key, {})
            veri[key].setdefault(banka, {"mobil": "", "sube": ""})

            d = _deger(s)
            if d:
                veri[key][banka][kanal] = d

    # ── Excel'e yaz ──
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

        # Kategori başlık satırı
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

        row += 1  # kategori sonrası boş satır

    for kat in KATEGORI_SIRA:
        yaz_kat(kat)

    # Listede olmayan kategoriler
    for kat in list(kat_satirlar.keys()):
        if kat not in yazilan:
            yaz_kat(kat)

    ws.cell(row=row, column=1,
            value=f"Son güncelleme: {_bugun()}").font = Font(size=9, color="888888")

    wb.save(dosya_yolu)
    print(f"[excel] {dosya_yolu} kaydedildi. {toplam} satır yazıldı.")
