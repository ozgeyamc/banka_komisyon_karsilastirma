"""
Bankalardan çekilen verileri karşılaştırmalı Excel formatında yazan modül.
Format: Satır=masraf, Sütun=Banka/Kanal (Mobil/Şube)
"""

import os
from datetime import datetime, timezone, timedelta
from typing import Dict, List
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

BASLIK_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
BASLIK_FONT = Font(color="FFFFFF", bold=True, size=11)
KATEGORI_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
KATEGORI_FONT = Font(bold=True, size=10)
ALT_BASLIK_FONT = Font(color="FFFFFF", bold=True, size=10)

INCE_BORDER = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)


def _bugun() -> str:
    tz = timezone(timedelta(hours=3))
    return datetime.now(tz).strftime("%d.%m.%Y %H:%M")


def _deger(satir: UcretSatiri) -> str:
    """Asgari tutar veya oranı birleştirerek tek hücre değeri üretir."""
    parts = []
    t = (satir.asgari_tutar or "").strip()
    o = (satir.asgari_oran or "").strip()
    zt = (satir.azami_tutar or "").strip()
    zo = (satir.azami_oran or "").strip()

    if t and o:
        parts.append(f"{t} / {o}")
    elif t:
        parts.append(t)
    elif o:
        parts.append(o)

    if zt and zo:
        parts.append(f"maks: {zt} / {zo}")
    elif zt:
        parts.append(f"maks: {zt}")
    elif zo:
        parts.append(f"maks: {zo}")

    return "\n".join(parts) if parts else "N/A"


def _kanal_belirle(satirlar: List[UcretSatiri], banka: str) -> Dict[str, Dict[str, str]]:
    """
    {masraf_adi: {"mobil": deger, "sube": deger}} döner.
    Kanal bilgisi yoksa occurrence sırasına göre ilk=mobil, ikinci=sube.
    """
    sonuc: Dict[str, Dict[str, str]] = {}
    sayac: Dict[str, int] = {}

    for s in satirlar:
        masraf = s.masraf.strip()
        deger = _deger(s)
        kanal = (s.kanal or "").lower()

        if masraf not in sonuc:
            sonuc[masraf] = {"mobil": "", "sube": ""}
            sayac[masraf] = 0

        if kanal == "mobil":
            sonuc[masraf]["mobil"] = deger
        elif kanal == "sube":
            sonuc[masraf]["sube"] = deger
        else:
            sayac[masraf] += 1
            if sayac[masraf] == 1:
                sonuc[masraf]["mobil"] = deger
            else:
                sonuc[masraf]["sube"] = deger

    return sonuc


def karsilastirma_excel_yaz(
    banka_verileri: Dict[str, List[UcretSatiri]],
    dosya_yolu: str = EXCEL_DOSYA_ADI,
) -> None:
    if os.path.exists(dosya_yolu):
        os.remove(dosya_yolu)

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_ADI

    # --- Sütun düzeni ---
    # A: Masraf adı
    # B,C: GARANTİ Mobil/Şube
    # D,E: İŞBANKASI Mobil/Şube
    # F,G: AKBANK Mobil/Şube
    # H,I: YAPIKREDI Mobil/Şube

    # Satır 1: Banka başlıkları (merge)
    ws.merge_cells("A1:A2")
    ws["A1"] = "MASRAF / KATEGORİ"
    ws["A1"].fill = BASLIK_FILL
    ws["A1"].font = BASLIK_FONT
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    col = 2
    for banka in BANKALAR:
        renk = BANKA_RENKLER.get(banka, {"bg": "808080", "fg": "FFFFFF"})
        banka_fill = PatternFill(start_color=renk["bg"], end_color=renk["bg"], fill_type="solid")
        banka_font = Font(color=renk["fg"], bold=True, size=11)

        col_letter_1 = get_column_letter(col)
        col_letter_2 = get_column_letter(col + 1)
        ws.merge_cells(f"{col_letter_1}1:{col_letter_2}1")
        cell = ws[f"{col_letter_1}1"]
        cell.value = banka
        cell.fill = banka_fill
        cell.font = banka_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

        # Alt başlık: Mobil / Şube
        for i, kanal in enumerate(["Mobil", "Şube"]):
            c = ws.cell(row=2, column=col + i, value=kanal)
            c.fill = banka_fill
            c.font = ALT_BASLIK_FONT
            c.alignment = Alignment(horizontal="center", vertical="center")

        col += 2

    # Sütun genişlikleri
    ws.column_dimensions["A"].width = 45
    for i in range(2, 10):
        ws.column_dimensions[get_column_letter(i)].width = 22

    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 18
    ws.freeze_panes = "A3"

    # --- Veriyi işle ---
    # Her banka için {masraf: {mobil, sube}}
    banka_masraf: Dict[str, Dict[str, Dict[str, str]]] = {}
    for banka in BANKALAR:
        satirlar = banka_verileri.get(banka, [])
        banka_masraf[banka] = _kanal_belirle(satirlar, banka)

    # Tüm masraf adlarını sırayla topla (kategori gruplaması için)
    # Kategori başlıklarını da ekle
    kategori_masraf: Dict[str, List[str]] = {}
    for banka in BANKALAR:
        satirlar = banka_verileri.get(banka, [])
        for s in satirlar:
            kat = s.kategori.strip()
            if kat not in kategori_masraf:
                kategori_masraf[kat] = []
            if s.masraf.strip() not in kategori_masraf[kat]:
                kategori_masraf[kat].append(s.masraf.strip())

    # Satırları yaz
    row = 3
    for kategori, masraf_listesi in kategori_masraf.items():
        # Kategori başlık satırı
        ws.merge_cells(f"A{row}:I{row}")
        cat_cell = ws[f"A{row}"]
        cat_cell.value = kategori
        cat_cell.fill = KATEGORI_FILL
        cat_cell.font = KATEGORI_FONT
        cat_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 16
        row += 1

        for masraf in masraf_listesi:
            ws.cell(row=row, column=1, value=masraf).alignment = Alignment(
                horizontal="left", vertical="center", wrap_text=True, indent=2
            )

            col = 2
            for banka in BANKALAR:
                kanal_dict = banka_masraf[banka].get(masraf, {"mobil": "", "sube": ""})
                for kanal_key in ["mobil", "sube"]:
                    c = ws.cell(row=row, column=col, value=kanal_dict.get(kanal_key, ""))
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    col += 1

            ws.row_dimensions[row].height = 30
            row += 1

    # Border uygula
    for r in ws.iter_rows(min_row=1, max_row=row - 1, min_col=1, max_col=9):
        for cell in r:
            cell.border = INCE_BORDER

    # Son güncelleme notu
    ws.cell(row=row + 1, column=1, value=f"Son güncelleme: {_bugun()}")

    wb.save(dosya_yolu)
    print(f"[excel] {dosya_yolu} kaydedildi. {row - 3} satır yazıldı.")
