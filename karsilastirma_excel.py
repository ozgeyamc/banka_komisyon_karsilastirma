"""
Banka komisyon karşılaştırma - Excel çıktısı.

Kategori eşleştirmesi ARTIK sabit string sözlüğüyle değil, anahtar kelime
tabanlı bir sınıflandırıcı ile yapılıyor. Sebep: bankaların kategori adları
birbirinden çok farklı ve bazen çok katmanlı (örn. İş Bankası:
"Para Aktarma - EFT / FAST - EFT-FAST Ücreti / Şube-Çözüm Merkezi").
Sabit sözlükle her varyasyonu yakalamak mümkün değil; kelime bazlı öncelikli
kurallar hem daha sağlam hem de kredi/kart/yatırım gibi alakasız binlerce
satırı otomatik eler (hiçbir kural eşleşmediği için).
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

# NOT: notify.py bu sabiti buradan import ediyor. Burada değiştirirsen
# notify.py otomatik senkron kalır - iki dosyada ayrı ayrı tutmayın.
SHEET_ADI = "KARŞILAŞTIRMA"

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

# ── Standart kategori adları ve Excel'deki sırası ──
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
    "Güvenli Araç Alım Satım",
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

# ── Anahtar kelime tabanlı sınıflandırma kuralları ──
# Sıra önemli: yukarıdan aşağıya ilk eşleşen kural kazanır.
# Her kural: (dışlama_kelimeleri, dahil_etme_kelimeleri (herhangi biri), standart_kategori)
# Metin = (kategori + " " + masraf) küçük harfe çevrilmiş hali üzerinde çalışır.
SINIFLANDIRMA_KURALLARI: List[Tuple[List[str], List[str], str]] = [
    # Kıymetli maden - havale/eft kelimelerinden önce kontrol edilmeli
    ([], ["kıymetli maden", "altın transfer", "ats ile altın", "fiziki altın teslim",
          "külçe altın"], "Kıymetli Maden Teslimleri"),

    # FAST - eft'den önce kontrol edilmeli
    ([], ["fast"], "FAST"),

    # EFT
    ([], ["eft"], "EFT Gönderimi"),

    # Havale (uluslararası/swift/western union hariç)
    (["swift", "uluslararası", "western union", "global fast"], ["havale"], "Havale Gönderimi"),

    # Kiralık Kasa
    ([], ["kiralık kasa", "kasa24", "kasa 24"], "Kiralık Kasa"),

    # Kredi Risk Raporu
    ([], ["risk raporu", "kkb", "findeks"], "Kredi Risk Raporu"),

    # HGS
    ([], ["hgs"], "HGS Etiket Bedeli"),

    # Şans Oyunu
    ([], ["şans oyun"], "Şans Oyunu Ödemeleri Aracılık"),

    # Güvenli Araç Alım Satım
    ([], ["güvenli araç"], "Güvenli Araç Alım Satım"),

    # Aidat
    ([], ["aidat ödeme"], "Aidat Ödemeleri Aracılık"),

    # Özel Okul
    ([], ["özel okul"], "Özel Okul Ödeme"),

    # Telefon operatör ödemeleri
    ([], ["telefon operatör", "tl/paket yükleme", "paket yükleme"], "Telefon Ödemeleri Aracılık"),

    # Vergi tahsilat
    ([], ["vergi tahsil"], "Vergi Tahsilat Aracılık"),

    # SGK
    ([], ["sgk"], "SGK Prim Ödemeleri"),

    # Fatura / Kurum ödemeleri (kart üzerinden)
    ([], ["fatura ödeme", "fatura/kurum", "kurum ödeme", "kurum tahsilat",
          "kurum/kurum"], "Fatura Ödemeleri - Kart"),

    # Arşiv Araştırma
    ([], ["arşiv araştırma", "belge ve bilgilendirme", "borcu yoktur"], "Arşiv Araştırma Ücreti"),

    # Mevduat Araştırma / Referans Mektubu / Mutabakat
    ([], ["referans mektubu", "mutabakat", "teyit yazı", "mevduat araştırma"], "Mevduat Araştırma"),

    # Çek - alt kategoriler (özelden genele)
    ([], ["çek iade", "çek muamelesiz iade"], "Çek İade Ücreti"),
    ([], ["çek belgelendirme", "karşılıksız çek", "çek düzeltme"], "Çek Belgelendirme ve Düzeltme Ücreti"),
    ([], ["çek defter", "çek düzenleme", "bloke çek", "karekodlu çek",
          "karşılıklı çek düzenleme", "hediye çeki", "seyahat çeki"], "Çek Defteri ve Çek Düzenleme Ücreti"),
    ([], ["çek tahsil", "çek ödeme", "çek takas", "dövizli çek"], "Çek Tahsilat Ücreti"),
    ([], ["çek"], "Çek Tahsilat Ücreti"),  # genel fallback

    # Senet - alt kategoriler
    ([], ["senet iade", "senet protestosuz"], "Senet İade Ücreti"),
    ([], ["senet protesto"], "Senet Protesto İşlemleri Ücreti"),
    ([], ["senet tahsil"], "Senet Tahsile Alma Ücreti"),
    ([], ["senet"], "Senet Tahsile Alma Ücreti"),  # genel fallback
]

# Bakiye sorma özel kural: yurt içi/yurt dışı ayrımı ayrı fonksiyonla yapılıyor.
_BAKIYE_ANAHTAR = ["bakiye sorgulama", "bakiye sorma", "vach bakiye", "limit sorgulama"]
_YURTDISI_ANAHTAR = ["yurtdışı", "yurt dışı", "yurt dişi"]


def _normalize_metin(kategori: str, masraf: str) -> str:
    return f"{kategori} {masraf}".lower().replace("i̇", "i")


def _standart_kategori(kategori: str, masraf: str) -> Optional[str]:
    """Anahtar kelime tabanlı sınıflandırma. İlk eşleşen kural kazanır."""
    metin = _normalize_metin(kategori, masraf)

    # Bakiye sorma önce kontrol edilsin (kendi içinde yurtiçi/yurtdışı ayrımı var)
    if any(k in metin for k in _BAKIYE_ANAHTAR):
        if any(k in metin for k in _YURTDISI_ANAHTAR):
            return "Bakiye Sorma - Yurtdışı - ATM"
        return "Bakiye Sorma - Yurtiçi - ATM"

    for disla, dahil, standart in SINIFLANDIRMA_KURALLARI:
        if disla and any(k in metin for k in disla):
            continue
        if any(k in metin for k in dahil):
            return standart

    return None


def _kanal_bul(kategori: str, masraf: str, scraper_kanal: str) -> str:
    """Kanal (mobil/şube) tespiti - hem kategori hem masraf metnine bakar."""
    if scraper_kanal in ("mobil", "sube"):
        return scraper_kanal
    metin = _normalize_metin(kategori, masraf)
    if any(k in metin for k in ["mobil", "internet", "iscep", "işcep", "online",
                                  "dijital", "e-", "e devlet", "asistan"]):
        return "mobil"
    if any(k in metin for k in ["şube", "gişe", "çözüm merkezi", "müşteri iletişim"]):
        return "sube"
    if "atm" in metin:
        return "mobil"
    return "mobil"


def _sayi_temizle(v: str) -> str:
    v = (v or "").strip()
    if v in ("-", "", "nan", "None"):
        return ""
    return v


def _deger(s: UcretSatiri) -> str:
    """Tutar/oran değerini biçimlendir. Asgari boşsa azami'ye düşer
    (bazı bankalarda tek kolonluk ücretler 'azami' başlığı altına düşebiliyor)."""
    t = _sayi_temizle(s.asgari_tutar) or _sayi_temizle(s.azami_tutar)
    o = _sayi_temizle(s.asgari_oran) or _sayi_temizle(s.azami_oran)
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
    kat_satirlar: Dict[str, List[str]] = {}
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

            std_kat = _standart_kategori(kategori, masraf)
            if std_kat is None:
                continue

            kanal = _kanal_bul(kategori, masraf, s.kanal or "")

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
