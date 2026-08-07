"""
Komisyon değişikliklerini tespit edip mail atan modül.
"""

import os
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

EXCEL_ADI = "komisyon_karsilastirma.xlsx"


def load_excel(path: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name="KARŞILAŞTIRMA", dtype=str).fillna("")
    except Exception:
        return pd.DataFrame()


def detect_changes(old_df: pd.DataFrame, new_df: pd.DataFrame) -> pd.DataFrame:
    if old_df.empty:
        return new_df

    try:
        old_df = old_df.reset_index(drop=True)
        new_df = new_df.reset_index(drop=True)

        common_cols = [c for c in new_df.columns if c in old_df.columns]
        if not common_cols:
            return new_df

        degisiklikler = []

        for i, new_row in new_df.iterrows():
            masraf = str(new_row.iloc[0]).strip()
            if not masraf or masraf.startswith("Son güncelleme"):
                continue

            # Eski Excel'de aynı masraf satırını bul
            old_match = old_df[old_df.iloc[:, 0].str.strip() == masraf]
            if old_match.empty:
                row = new_row.copy()
                row["DEĞİŞİKLİK"] = "YENİ EKLENDI"
                degisiklikler.append(row)
            else:
                old_row = old_match.iloc[0]
                farklar = []
                for col in common_cols[1:]:
                    ov = str(old_row.get(col, "")).strip()
                    nv = str(new_row.get(col, "")).strip()
                    if ov != nv:
                        farklar.append(f"{col}: {ov} → {nv}")
                if farklar:
                    row = new_row.copy()
                    row["DEĞİŞİKLİK"] = " | ".join(farklar)
                    degisiklikler.append(row)

        if not degisiklikler:
            return pd.DataFrame()

        return pd.DataFrame(degisiklikler).reset_index(drop=True)

    except Exception as e:
        print(f"[notify] Değişiklik tespiti hatası: {e}")
        return pd.DataFrame()


def build_html_table(df: pd.DataFrame) -> str:
    rows_html = ""
    for _, row in df.iterrows():
        degisiklik = str(row.get("DEĞİŞİKLİK", ""))
        renk = "#fff3cd" if "YENİ" in degisiklik else "#fde8e8"
        cells = "".join(
            f"<td style='padding:6px 10px;border:1px solid #ddd'>{row.get(c, '')}</td>"
            for c in df.columns
        )
        rows_html += f"<tr style='background:{renk}'>{cells}</tr>"

    headers = "".join(
        f"<th style='padding:8px 10px;background:#1a3c5e;color:white;border:1px solid #ddd'>{c}</th>"
        for c in df.columns
    )
    return f"""
    <table style='border-collapse:collapse;font-family:Arial,sans-serif;font-size:12px;width:100%'>
        <thead><tr>{headers}</tr></thead>
        <tbody>{rows_html}</tbody>
    </table>
    """


def send_mail(changes_df: pd.DataFrame, new_excel_path: str, test_mode: bool = False):
    mail_user = os.environ["MAIL_USER"]
    mail_pass = os.environ["MAIL_PASS"]
    mail_to = os.environ["MAIL_TO"].split(",")

    if test_mode:
        konu = "✅ TEST — Komisyon karşılaştırma sistemi çalışıyor"
        html_body = """
        <html><body style='font-family:Arial,sans-serif;color:#333'>
            <h2 style='color:#1a3c5e'>✅ Test Bildirimi</h2>
            <p>Banka komisyon karşılaştırma sistemi başarıyla çalışıyor.</p>
            <p style='color:#888;font-size:12px'>Bu bir test mailidir.</p>
        </body></html>
        """
    else:
        sayi = len(changes_df)
        konu = f"⚠️ Komisyon Değişikliği — {sayi} değişiklik tespit edildi"
        html_table = build_html_table(changes_df)
        html_body = f"""
        <html><body style='font-family:Arial,sans-serif;color:#333'>
            <h2 style='color:#1a3c5e'>Komisyon Değişiklik Bildirimi</h2>
            <p>Bugünkü güncellemede <strong>{sayi} değişiklik</strong> tespit edildi.</p>
            {html_table}
            <br>
            <p style='color:#888;font-size:12px'>Güncel karşılaştırma Excel'i ekte yer almaktadır.</p>
        </body></html>
        """

    msg = MIMEMultipart("mixed")
    msg["From"] = mail_user
    msg["To"] = ", ".join(mail_to)
    msg["Subject"] = konu
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    if os.path.exists(new_excel_path):
        with open(new_excel_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f'attachment; filename="{EXCEL_ADI}"')
            msg.attach(part)

    print(f"[notify] SMTP bağlantısı kuruluyor... {mail_user}")
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(mail_user, mail_pass)
        server.sendmail(mail_user, mail_to, msg.as_string())
    print(f"[notify] Mail gönderildi → {mail_to}")


def check_and_notify(old_excel: str, new_excel: str, test_mode: bool = False):
    print("[notify] Değişiklik kontrolü başlıyor...")

    if test_mode:
        print("[notify] TEST MODU")
        send_mail(pd.DataFrame(), new_excel, test_mode=True)
        return

    old_df = load_excel(old_excel)
    new_df = load_excel(new_excel)
    changes = detect_changes(old_df, new_df)

    if changes.empty:
        print("[notify] Değişiklik yok, mail gönderilmedi.")
        return

    print(f"[notify] {len(changes)} değişiklik bulundu, mail gönderiliyor...")
    send_mail(changes, new_excel)
