"""
Bankalardan çekilen verileri karşılaştırmalı Excel formatında yazan modül.
Her banka kendi kategorilerini getirir; masraf adı eşleşen satırlar
aynı satıra, eşleşmeyenler kendi bankasının altına yazılır.
"""

import os
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

BASLIK_FILL   = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
BASLIK_FONT   = Font(color="FFFFFF", bold=True, size=11)
KATEGORI_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
KATEGORI_FONT = Font(bold=True, size=10, color="1F3864")
INCE_BORDER   = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)


def _bugun() -> str:
    tz = timezone(timedelta(hours=3))
    return datetime.now(tz).strftime("%d.%m.%Y %H:%M")


def _deger(s: UcretSatiri) -> str:
    t  = (s.asgari_tutar or "").strip()
    o  = (s.asgari_oran  or "").strip()
    zt = (s.azami_tutar  or "").strip()
    zo = (s.azami_oran   or "").strip()

    ana = f"{t} / {o}" if (t and o) else (t or o)
    maks = ""
    if zt or zo:
        maks = "maks: " + (f"{zt} / {zo}" if (zt and zo) else (zt or zo))

    parts = [p for p in [ana, maks] if p]
    return "\n".join(parts) if parts else ""


def _normalize_masraf(m: str) -> str:
    """Küçük harf, boşluk normalize — eşleştirme için."""
    import re
    m = m.lower().strip()
    m = re.sub(r"\s+", " ", m)
    m = m.replace("i̇", "i").replace("ı", "i").replace("ğ", "g")
    m = m.replace("ü", "u").replace("ş", "s").replace("ö", "o").replace("ç", "c")
    return m


def _kanal_ayir(satirlar: List[UcretSatiri]) -> List[Tuple[str, str, str]]:
    """
    [(masraf_adi, mobil_deger, sube_deger), ...] döner.
    Kanal bilgisi varsa kullan, yoksa occurrence sırasına göre ilk=mobil ikinci=sube.
    """
    seen: Dict[str, Dict] = {}
    order: List[str] = []

    for s in satirlar:
        key = s.masraf.strip()
        if key not in seen:
            seen[key] = {"mobil": "", "sube": "", "count": 0, "display": key}
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
            elif seen[key]["count"] == 2:
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

    # Toplam 9 sütun: A + 2*4 banka
    # A=masraf, B=GAR mob, C=GAR sub, D=IS mob, E=IS sub,
    # F=AK mob, G=AK sub, H=YK mob, I=YK sub
    MAX_COL = 9

    # ── Satır 1: Banka başlıkları ──
    ws.merge_cells("A1:A2")
    c = ws["A1"]
    c.value = "MASRAF / KATEGORİ"
    c.fill = BASLIK_FILL
    c.font = BASLIK_FONT
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = INCE_BORDER

    col = 2
    for banka in BANKALAR:
        renk = BANKA_RENKLER[banka]
        fill = PatternFill(start_color=renk["bg"], end_color=renk["bg"], fill_type="solid")
        font_h = Font(color=renk["fg"], bold=True, size=11)
        font_s = Font(color=renk["fg"], bold=True, size=10)

        l1 = get_column_letter(col)
        l2 = get_column_letter(col + 1)
        ws.merge_cells(f"{l1}1:{l2}1")
        c1 = ws[f"{l1}1"]
        c1.value = banka
        c1.fill = fill
        c1.font = font_h
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.border = INCE_BORDER

        for j, kanal in enumerate(["Mobil", "Şube"]):
            c2 = ws.cell(row=2, column=col + j)
            c2.value = kanal
            c2.fill = fill
            c2.font = font_s
            c2.alignment = Alignment(horizontal="center", vertical="center")
            c2.border = INCE_BORDER

        col += 2

    # ── Sütun genişlikleri ──
    ws.column_dimensions["A"].width = 48
    for i in range(2, MAX_COL + 1):
        ws.column_dimensions[get_column_letter(i)].width = 20

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 18
    ws.freeze_panes = "A3"

    # ── Veriyi işle ──
    # Her banka için {normalize_masraf: (masraf_gercek, mobil, sube)}
    banka_data: Dict[str, Dict[str, Tuple[str, str, str]]] = {}
    banka_kategori: Dict[str, Dict[str, List[str]]] = {}  # {banka: {kat: [norm_masraf]}}

    for banka in BANKALAR:
        satirlar = banka_verileri.get(banka, [])
        banka_data[banka] = {}
        banka_kategori[banka] = {}

        # Kategori bazlı grupla
        kat_gruplar: Dict[str, List[UcretSatiri]] = {}
        for s in satirlar:
            kat = s.kategori.strip()
            if kat not in kat_gruplar:
                kat_gruplar[kat] = []
            kat_gruplar[kat].append(s)

        for kat, sat_list in kat_gruplar.items():
            banka_kategori[banka][kat] = []
            for masraf, mob, sub in _kanal_ayir(sat_list):
                norm = _normalize_masraf(masraf)
                banka_data[banka][norm] = (masraf, mob, sub)
                banka_kategori[banka][kat].append(norm)

    # Tüm kategorileri ve masraf adlarını topla (tüm bankalar birleşik)
    tum_kategoriler: Dict[str, List[str]] = {}  # {kat: [norm_masraf sıralı]}
    seen_norms: set = set()

    for banka in BANKALAR:
        for kat, norm_list in banka_kategori[banka].items():
            if kat not in tum_kategoriler:
                tum_kategoriler[kat] = []
            for norm in norm_list:
                if norm not in seen_norms:
                    tum_kategoriler[kat].append(norm)
                    seen_norms.add(norm)

    # norm → display adı (ilk bulunan bankadan al)
    norm_display: Dict[str, str] = {}
    for banka in BANKALAR:
        for norm, (display, _, _) in banka_data[banka].items():
            if norm not in norm_display:
                norm_display[norm] = display

    # ── Satırları yaz ──
    row = 3
    for kategori, norm_list in tum_kategoriler.items():
        # Kategori başlık satırı
        ws.merge_cells(f"A{row}:{get_column_letter(MAX_COL)}{row}")
        cat_c = ws[f"A{row}"]
        cat_c.value = kategori
        cat_c.fill = KATEGORI_FILL
        cat_c.font = KATEGORI_FONT
        cat_c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        cat_c.border = INCE_BORDER
        ws.row_dimensions[row].height = 16
        row += 1

        for norm in norm_list:
            display = norm_display.get(norm, norm)
            ws.cell(row=row, column=1, value=display).alignment = Alignment(
                horizontal="left", vertical="center", wrap_text=True, indent=2
            )
            ws.cell(row=row, column=1).border = INCE_BORDER

            col = 2
            for banka in BANKALAR:
                _, mob, sub = banka_data[banka].get(norm, ("", "", ""))
                for val in [mob, sub]:
                    c = ws.cell(row=row, column=col, value=val)
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    c.border = INCE_BORDER
                    col += 1

            ws.row_dimensions[row].height = 32
            row += 1

    # Son not
    ws.cell(row=row + 1, column=1, value=f"Son güncelleme: {_bugun()}")

    wb.save(dosya_yolu)
    print(f"[excel] {dosya_yolu} kaydedildi. {row - 3} satır yazıldı.")
