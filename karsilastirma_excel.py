"""
Bankalardan çekilen verileri karşılaştırmalı Excel formatında yazan modül.
Eşleştirme: masraf adı normalize edilerek yapılır.
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

KATEGORI_SIRA = [
    "EFT Gönderimi",
    "Havale Gönderimi",
    "FAST",
    "Kiralık Kasa",
    "Kıymetli Maden",
    "HGS",
    "Fatura Ödemeleri",
    "SGK",
    "Şans Oyunu",
    "Arşiv Araştırma",
    "Kredi Risk Raporu",
    "Çek Defteri",
    "Çek Tahsilat",
    "Çek İşlemleri",
    "Senet İşlemleri",
    "Diğer",
]


def _bugun() -> str:
    tz = timezone(timedelta(hours=3))
    return datetime.now(tz).strftime("%d.%m.%Y %H:%M")


def _deger(s: UcretSatiri) -> str:
    t = (s.asgari_tutar or "").strip()
    o = (s.asgari_oran or "").strip()
    if t and o:
        return f"{t} / {o}"
    return t or o or ""


def _norm_masraf(m: str) -> str:
    m = m.lower().strip()
    for a, b in [("ı","i"),("ğ","g"),("ü","u"),("ş","s"),("ö","o"),("ç","c"),
                 ("â","a"),("î","i"),("û","u"),("i̇","i")]:
        m = m.replace(a, b)
    m = re.sub(r"[,.\-/\\()\[\]]", " ", m)
    # Banka adlarını çıkar
    for banka in ["garanti bbva", "garanti", "akbank", "yapi kredi", "yapikredi",
                  "isbank", "is bankasi", "iscep", "mobil bankacilık", "internet subesi",
                  "internet sube", "sube cozum merkezi", "cozum merkezi",
                  "musteri iletisim merkezi"]:
        m = m.replace(banka, "")
    m = re.sub(r"\s+", " ", m).strip()
    return m


def _ust_kategori(masraf_norm: str) -> str:
    m = masraf_norm
    if "eft" in m and not "swift" in m:
        return "EFT Gönderimi"
    if "havale" in m:
        return "Havale Gönderimi"
    if "fast" in m:
        return "FAST"
    if "kiralik kasa" in m or "yillik kasa" in m or "kasa ucreti" in m or "kasa depozito" in m:
        return "Kiralık Kasa"
    if "altin" in m or "kiymetli maden" in m:
        return "Kıymetli Maden"
    if "hgs" in m:
        return "HGS"
    if "sans oyunu" in m or "piyango" in m:
        return "Şans Oyunu"
    if "fatura" in m and "kredi" not in m:
        return "Fatura Ödemeleri"
    if "sgk" in m:
        return "SGK"
    if "arsiv" in m:
        return "Arşiv Araştırma"
    if "kkb" in m or ("kredi" in m and "risk" in m):
        return "Kredi Risk Raporu"
    if "cek defteri" in m or "keside" in m:
        return "Çek Defteri"
    if "cek tahsil" in m or "cek odeme" in m:
        return "Çek Tahsilat"
    if "cek" in m:
        return "Çek İşlemleri"
    if "senet" in m:
        return "Senet İşlemleri"
    return "Diğer"


def _kanal_masraftan(masraf: str) -> str:
    m = masraf.lower()
    if any(k in m for k in ["mobil", "internet", "iscep", "dijital", "online", "e-"]):
        return "mobil"
    if "atm" in m:
        return "mobil"  # ATM = self-service, mobil sütununa
    if any(k in m for k in ["şube", "sube", "gişe", "gise", "çözüm", "cozum", "iletişim merkezi"]):
        return "sube"
    return ""


def karsilastirma_excel_yaz(
    banka_verileri: Dict[str, List[UcretSatiri]],
    dosya_yolu: str = EXCEL_DOSYA_ADI,
) -> None:
    if os.path.exists(dosya_yolu):
        os.remove(dosya_yolu)

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def sb(cell):
        cell.border = border

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_ADI

    MAX_COL = 9  # A + 4 banka x 2 kanal

    # ── Satır 1-2: Başlıklar ──
    ws.merge_cells("A1:A2")
    c = ws["A1"]
    c.value = "MASRAF"
    c.fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    c.font = Font(color="FFFFFF", bold=True, size=11)
    c.alignment = Alignment(horizontal="center", vertical="center")
    sb(c)

    col = 2
    for banka in BANKALAR:
        r = BANKA_RENKLER[banka]
        fill = PatternFill(start_color=r["bg"], end_color=r["bg"], fill_type="solid")
        l1 = get_column_letter(col)
        l2 = get_column_letter(col + 1)
        ws.merge_cells(f"{l1}1:{l2}1")
        c1 = ws[f"{l1}1"]
        c1.value = banka
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

    ws.column_dimensions["A"].width = 45
    for i in range(2, MAX_COL + 1):
        ws.column_dimensions[get_column_letter(i)].width = 18
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 18
    ws.freeze_panes = "A3"

    # ── Veriyi işle: norm_masraf → {banka: {kanal: deger}} ──
    # norm_masraf → display adı (ilk görülen)
    norm_display: Dict[str, str] = {}
    # norm_masraf → üst kategori
    norm_kat: Dict[str, str] = {}
    # kategori → [norm_masraf sırası]
    kat_masraflar: Dict[str, List[str]] = {}
    # norm_masraf → {banka: {"mobil": str, "sube": str}}
    masraf_banka: Dict[str, Dict[str, Dict[str, str]]] = {}

    for banka in BANKALAR:
        satirlar = banka_verileri.get(banka, [])
        for s in satirlar:
            norm = _norm_masraf(s.masraf)
            if not norm:
                continue
            # display adı
            norm_display.setdefault(norm, s.masraf.strip())
            # üst kategori
            if norm not in norm_kat:
                kat = _ust_kategori(norm)
                norm_kat[norm] = kat
                if kat not in kat_masraflar:
                    kat_masraflar[kat] = []
                if norm not in kat_masraflar[kat]:
                    kat_masraflar[kat].append(norm)
            # kanal
            kanal = (s.kanal or "").lower()
            if not kanal:
                kanal = _kanal_masraftan(s.masraf)
            if kanal not in ("mobil", "sube"):
                kanal = "mobil"
            # değer
            deger = _deger(s)
            if norm not in masraf_banka:
                masraf_banka[norm] = {}
            if banka not in masraf_banka[norm]:
                masraf_banka[norm][banka] = {"mobil": "", "sube": ""}
            if deger:
                masraf_banka[norm][banka][kanal] = deger

    # ── Satırları yaz ──
    KATEGORI_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    KATEGORI_FONT = Font(bold=True, size=10, color="1F3864")
    DATA_FONT = Font(size=10)
    MASRAF_FONT = Font(size=10)

    yazilan_katlar = set()
    row = 3
    toplam = 0

    for kat in KATEGORI_SIRA:
        if kat not in kat_masraflar:
            continue
        yazilan_katlar.add(kat)

        ws.merge_cells(f"A{row}:{get_column_letter(MAX_COL)}{row}")
        cc = ws[f"A{row}"]
        cc.value = kat
        cc.fill = KATEGORI_FILL
        cc.font = KATEGORI_FONT
        cc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        sb(cc)
        ws.row_dimensions[row].height = 16
        row += 1

        for norm in kat_masraflar[kat]:
            display = norm_display.get(norm, norm)
            mc = ws.cell(row=row, column=1, value=display)
            mc.font = MASRAF_FONT
            mc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=2)
            sb(mc)

            col = 2
            for banka in BANKALAR:
                bd = masraf_banka.get(norm, {}).get(banka, {"mobil": "", "sube": ""})
                for kanal in ["mobil", "sube"]:
                    dc = ws.cell(row=row, column=col, value=bd.get(kanal, ""))
                    dc.font = DATA_FONT
                    dc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    sb(dc)
                    col += 1

            ws.row_dimensions[row].height = 22
            row += 1
            toplam += 1

    # Kalan kategoriler
    for kat, normlar in kat_masraflar.items():
        if kat in yazilan_katlar:
            continue
        ws.merge_cells(f"A{row}:{get_column_letter(MAX_COL)}{row}")
        cc = ws[f"A{row}"]
        cc.value = kat
        cc.fill = KATEGORI_FILL
        cc.font = KATEGORI_FONT
        cc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        sb(cc)
        ws.row_dimensions[row].height = 16
        row += 1

        for norm in normlar:
            display = norm_display.get(norm, norm)
            mc = ws.cell(row=row, column=1, value=display)
            mc.font = MASRAF_FONT
            mc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=2)
            sb(mc)

            col = 2
            for banka in BANKALAR:
                bd = masraf_banka.get(norm, {}).get(banka, {"mobil": "", "sube": ""})
                for kanal in ["mobil", "sube"]:
                    dc = ws.cell(row=row, column=col, value=bd.get(kanal, ""))
                    dc.font = DATA_FONT
                    dc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    sb(dc)
                    col += 1

            ws.row_dimensions[row].height = 22
            row += 1
            toplam += 1

    ws.cell(row=row + 1, column=1,
            value=f"Son güncelleme: {_bugun()}").font = Font(size=9, color="888888")

    wb.save(dosya_yolu)
    print(f"[excel] {dosya_yolu} kaydedildi. {toplam} satır yazıldı.")
