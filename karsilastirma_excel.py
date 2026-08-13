"""
Bankalardan çekilen verileri karşılaştırmalı Excel formatında yazan modül.
"""

import os
import re
from datetime import datetime, timezone, timedelta
from typing import Dict, List, Tuple
from models import UcretSatiri

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

EXCEL_DOSYA_ADI = "komisyon_karsilastirma.xlsx"
SHEET_ADI = "KARŞILAŞTIRMA"

BANKALAR = ["GARANTİ", "İŞBANKASI", "AKBANK", "YAPIKREDI"]

BANKA_RENKLER = {
    "GARANTİ":   {"bg": "00B050", "fg": "FFFFFF"},
    "İŞBANKASI": {"bg": "012169", "fg": "FFFFFF"},
    "AKBANK":    {"bg": "FF0000", "fg": "FFFFFF"},
    "YAPIKREDI": {"bg": "003087", "fg": "FFD700"},
}

# Kategori eşleştirme — her bankadan gelen kategori adlarını
# ortak bir üst başlığa bağlar
KATEGORI_MAP = {
    "eft": "EFT Gönderimi",
    "elektronik fon": "EFT Gönderimi",
    "havale": "Havale Gönderimi",
    "fast": "FAST",
    "kiralik kasa": "Kiralık Kasa",
    "kiralık kasa": "Kiralık Kasa",
    "kıymetli maden": "Kıymetli Maden Teslimleri",
    "kiymetli maden": "Kıymetli Maden Teslimleri",
    "kredi risk": "Kredi Risk Raporu",
    "fatura": "Fatura Ödemeleri",
    "sgk": "SGK Prim Ödemeleri",
    "hgs": "HGS Etiket Bedeli",
    "sans oyunu": "Şans Oyunu Ödemeleri",
    "şans oyunu": "Şans Oyunu Ödemeleri",
    "aidat": "Aidat Ödemeleri",
    "okul": "Özel Okul Ödeme",
    "telefon": "Telefon Ödemeleri",
    "vergi": "Vergi Tahsilat",
    "arsiv": "Arşiv Araştırma",
    "arşiv": "Arşiv Araştırma",
    "mevduat": "Mevduat Araştırma",
    "bakiye sorma": "Bakiye Sorma - ATM",
    "cek defteri": "Çek Defteri ve Düzenleme",
    "çek defteri": "Çek Defteri ve Düzenleme",
    "cek duzenleme": "Çek Defteri ve Düzenleme",
    "çek düzenleme": "Çek Defteri ve Düzenleme",
    "cek tahsil": "Çek Tahsilat",
    "çek tahsil": "Çek Tahsilat",
    "cek belgelen": "Çek Belgelendirme",
    "çek belgelen": "Çek Belgelendirme",
    "senet iade": "Senet İade",
    "senet protesto": "Senet Protesto",
    "senet tahsil": "Senet Tahsile Alma",
}

KATEGORI_SIRA = [
    "EFT Gönderimi",
    "Havale Gönderimi",
    "FAST",
    "Kiralık Kasa",
    "Kıymetli Maden Teslimleri",
    "Kredi Risk Raporu",
    "Fatura Ödemeleri",
    "SGK Prim Ödemeleri",
    "HGS Etiket Bedeli",
    "Şans Oyunu Ödemeleri",
    "Aidat Ödemeleri",
    "Özel Okul Ödeme",
    "Telefon Ödemeleri",
    "Vergi Tahsilat",
    "Arşiv Araştırma",
    "Mevduat Araştırma",
    "Bakiye Sorma - ATM",
    "Çek Defteri ve Düzenleme",
    "Çek Tahsilat",
    "Çek Belgelendirme",
    "Senet İade",
    "Senet Protesto",
    "Senet Tahsile Alma",
    "Diğer",
]


def _bugun() -> str:
    tz = timezone(timedelta(hours=3))
    return datetime.now(tz).strftime("%d.%m.%Y %H:%M")


def _deger(s: UcretSatiri) -> str:
    """Tek hücreye yazılacak kısa değer."""
    t = (s.asgari_tutar or "").strip()
    o = (s.asgari_oran or "").strip()
    if t and o:
        return f"{t} / {o}"
    return t or o or ""


def _norm(m: str) -> str:
    m = m.lower().strip()
    m = re.sub(r"\s+", " ", m)
    for a, b in [("i̇","i"),("ı","i"),("ğ","g"),("ü","u"),("ş","s"),("ö","o"),("ç","c")]:
        m = m.replace(a, b)
    return m


def _kategori_bul(kategori_adi: str) -> str:
    k = _norm(kategori_adi)
    for anahtar, ortak in KATEGORI_MAP.items():
        if anahtar in k:
            return ortak
    return "Diğer"


def _kanal_ayir(satirlar: List[UcretSatiri]) -> List[Tuple[str, str, str]]:
    """[(masraf_adi, mobil_deger, sube_deger), ...]"""
    seen: Dict[str, dict] = {}
    order: List[str] = []
    for s in satirlar:
        key = s.masraf.strip()
        if key not in seen:
            seen[key] = {"mobil": "", "sube": "", "count": 0}
            order.append(key)
        deger = _deger(s)
        kanal = (s.kanal or "").lower()
        seen[key]["count"] += 1
        if kanal == "mobil":
            seen[key]["mobil"] = deger
        elif kanal == "sube":
            seen[key]["sube"] = deger
        else:
            if seen[key]["count"] == 1:
                seen[key]["mobil"] = deger
            else:
                seen[key]["sube"] = deger
    return [(k, seen[k]["mobil"], seen[k]["sube"]) for k in order]


def karsilastirma_excel_yaz(
    banka_verileri: Dict[str, List[UcretSatiri]],
    dosya_yolu: str = EXCEL_DOSYA_ADI,
) -> None:
    if os.path.exists(dosya_yolu):
        os.remove(dosya_yolu)

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_ADI

    # Sütunlar: A=masraf, B-C=GAR, D-E=IS, F-G=AK, H-I=YK
    MAX_COL = 9

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def set_border(cell):
        cell.border = border

    # ── Satır 1: Banka başlıkları ──
    ws.merge_cells("A1:A2")
    c = ws["A1"]
    c.value = "MASRAF"
    c.fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    c.font = Font(color="FFFFFF", bold=True, size=11)
    c.alignment = Alignment(horizontal="center", vertical="center")
    set_border(c)

    col = 2
    for banka in BANKALAR:
        r = BANKA_RENKLER[banka]
        fill = PatternFill(start_color=r["bg"], end_color=r["bg"], fill_type="solid")
        font_h = Font(color=r["fg"], bold=True, size=11)
        font_s = Font(color=r["fg"], bold=True, size=10)

        l1 = get_column_letter(col)
        l2 = get_column_letter(col + 1)
        ws.merge_cells(f"{l1}1:{l2}1")
        c1 = ws[f"{l1}1"]
        c1.value = banka
        c1.fill = fill
        c1.font = font_h
        c1.alignment = Alignment(horizontal="center", vertical="center")
        set_border(c1)

        for j, kanal in enumerate(["Mobil", "Şube"]):
            c2 = ws.cell(row=2, column=col + j)
            c2.value = kanal
            c2.fill = fill
            c2.font = font_s
            c2.alignment = Alignment(horizontal="center", vertical="center")
            set_border(c2)
        col += 2

    # ── Sütun genişlikleri ──
    ws.column_dimensions["A"].width = 42
    for i in range(2, MAX_COL + 1):
        ws.column_dimensions[get_column_letter(i)].width = 18
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 18
    ws.freeze_panes = "A3"

    # ── Veriyi işle ──
    # Her banka → kategori → [(masraf, mobil, sube)]
    banka_kat_data: Dict[str, Dict[str, List[Tuple[str,str,str]]]] = {}
    for banka in BANKALAR:
        satirlar = banka_verileri.get(banka, [])
        kat_gruplar: Dict[str, List[UcretSatiri]] = {}
        for s in satirlar:
            kat = _kategori_bul(s.kategori)
            kat_gruplar.setdefault(kat, []).append(s)
        banka_kat_data[banka] = {}
        for kat, sat_list in kat_gruplar.items():
            banka_kat_data[banka][kat] = _kanal_ayir(sat_list)

    # Tüm kategorilerdeki tüm masraf adlarını norm→display eşleştir
    # Her kategori için: norm_masraf → {banka: (mobil, sube)}
    kat_masraf_banka: Dict[str, Dict[str, Dict[str, Tuple[str,str]]]] = {}
    norm_display: Dict[str, str] = {}

    for banka in BANKALAR:
        for kat, triplets in banka_kat_data[banka].items():
            if kat not in kat_masraf_banka:
                kat_masraf_banka[kat] = {}
            for masraf, mob, sub in triplets:
                norm = _norm(masraf)
                norm_display.setdefault(norm, masraf)
                if norm not in kat_masraf_banka[kat]:
                    kat_masraf_banka[kat][norm] = {}
                kat_masraf_banka[kat][norm][banka] = (mob, sub)

    # ── Satırları yaz ──
    KATEGORI_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    KATEGORI_FONT = Font(bold=True, size=10, color="1F3864")
    DATA_FONT     = Font(size=10)
    MASRAF_FONT   = Font(size=10)

    # Kategori sırasına göre yaz
    yazilan_katlar = set()
    row = 3

    for kat in KATEGORI_SIRA:
        if kat not in kat_masraf_banka:
            continue
        yazilan_katlar.add(kat)

        # Kategori başlık satırı
        ws.merge_cells(f"A{row}:{get_column_letter(MAX_COL)}{row}")
        cc = ws[f"A{row}"]
        cc.value = kat
        cc.fill = KATEGORI_FILL
        cc.font = KATEGORI_FONT
        cc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        set_border(cc)
        ws.row_dimensions[row].height = 16
        row += 1

        for norm, banka_dict in kat_masraf_banka[kat].items():
            display = norm_display.get(norm, norm)
            mc = ws.cell(row=row, column=1, value=display)
            mc.font = MASRAF_FONT
            mc.alignment = Alignment(horizontal="left", vertical="center",
                                     wrap_text=True, indent=2)
            set_border(mc)

            col = 2
            for banka in BANKALAR:
                mob, sub = banka_dict.get(banka, ("", ""))
                for val in [mob, sub]:
                    dc = ws.cell(row=row, column=col, value=val)
                    dc.font = DATA_FONT
                    dc.alignment = Alignment(horizontal="center", vertical="center",
                                             wrap_text=True)
                    set_border(dc)
                    col += 1

            ws.row_dimensions[row].height = 28
            row += 1

    # Sırada olmayan kategoriler sona ekle
    for kat, masraf_dict in kat_masraf_banka.items():
        if kat in yazilan_katlar:
            continue
        ws.merge_cells(f"A{row}:{get_column_letter(MAX_COL)}{row}")
        cc = ws[f"A{row}"]
        cc.value = kat
        cc.fill = KATEGORI_FILL
        cc.font = KATEGORI_FONT
        cc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        set_border(cc)
        ws.row_dimensions[row].height = 16
        row += 1

        for norm, banka_dict in masraf_dict.items():
            display = norm_display.get(norm, norm)
            mc = ws.cell(row=row, column=1, value=display)
            mc.font = MASRAF_FONT
            mc.alignment = Alignment(horizontal="left", vertical="center",
                                     wrap_text=True, indent=2)
            set_border(mc)

            col = 2
            for banka in BANKALAR:
                mob, sub = banka_dict.get(banka, ("", ""))
                for val in [mob, sub]:
                    dc = ws.cell(row=row, column=col, value=val)
                    dc.font = DATA_FONT
                    dc.alignment = Alignment(horizontal="center", vertical="center",
                                             wrap_text=True)
                    set_border(dc)
                    col += 1

            ws.row_dimensions[row].height = 28
            row += 1

    ws.cell(row=row + 1, column=1,
            value=f"Son güncelleme: {_bugun()}").font = Font(size=9, color="888888")

    wb.save(dosya_yolu)
    print(f"[excel] {dosya_yolu} kaydedildi. {row - 3} satır yazıldı.")
