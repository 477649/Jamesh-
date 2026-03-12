import os
import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
REPORT_DIR = ROOT / "outputs" / "reports"


def load_latest_report():
    reports = sorted(
        REPORT_DIR.glob("RetailPro_Trading_Insight_Report_*.xlsx"),
        key=lambda x: x.stat().st_mtime,
        reverse=True,
    )

    if not reports:
        raise FileNotFoundError(f"No trading report found in {REPORT_DIR}")

    return reports[0]


def safe_round(df, cols, digits=2):
    df = df.copy()
    for col in cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").round(digits)
    return df


def format_percent_cols(df, cols):
    df = df.copy()
    for col in cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").map(
                lambda x: f"{x:.2f}%" if pd.notna(x) else ""
            )
    return df


def format_html_table(df, title):
    if df is None or df.empty:
        return f"<h3 style='margin-bottom:8px;'>{title}</h3><p>No data available.</p>"

    html = f"""
    <h3 style="margin-bottom:8px;">{title}</h3>
    <table border="1" cellpadding="6" cellspacing="0"
           style="border-collapse:collapse; font-family:Arial; font-size:13px; margin-bottom:18px;">
      <tr style="background-color:#d9eaf7;">
    """

    for col in df.columns:
        html += f"<th>{col}</th>"
    html += "</tr>"

    for _, row in df.iterrows():
        html += "<tr>"
        for val in row:
            html += f"<td>{'' if pd.isna(val) else val}</td>"
        html += "</tr>"

    html += "</table>"
    return html


def build_email_body(report_path):
    smart_money = pd.read_excel(report_path, sheet_name="Smart_Money")
    price_movers = pd.read_excel(report_path, sheet_name="Price_Movers")
    symbol_summary = pd.read_excel(report_path, sheet_name="Symbol_Summary")
    top_picks = pd.read_excel(report_path, sheet_name="Top_Picks")

    # -----------------------------
    # Best Smart Money Stocks
    # -----------------------------
    sm_cols = [
        "Symbol",
        "Sectors",
        "Last_Price",
        "VWAP",
        "Momentum",
        "SmartMoneyScore",
        "SmartMoneySignal",
    ]

    sm_1d = (
        smart_money[smart_money["Window"] == "1D"]
        .sort_values("SmartMoneyScore", ascending=False)
        .head(5)
        .copy()
    )
    sm_7d = (
        smart_money[smart_money["Window"] == "7D"]
        .sort_values("SmartMoneyScore", ascending=False)
        .head(5)
        .copy()
    )
    sm_15d = (
        smart_money[smart_money["Window"] == "15D"]
        .sort_values("SmartMoneyScore", ascending=False)
        .head(5)
        .copy()
    )

    sm_1d = sm_1d[[c for c in sm_cols if c in sm_1d.columns]]
    sm_7d = sm_7d[[c for c in sm_cols if c in sm_7d.columns]]
    sm_15d = sm_15d[[c for c in sm_cols if c in sm_15d.columns]]

    for df in [sm_1d, sm_7d, sm_15d]:
        df.rename(
            columns={
                "Sectors": "Sector",
                "Last_Price": "Last Price",
                "SmartMoneyScore": "Smart Money Score",
                "SmartMoneySignal": "Signal",
            },
            inplace=True,
        )

    sm_1d = safe_round(sm_1d, ["Last Price", "VWAP", "Smart Money Score"])
    sm_7d = safe_round(sm_7d, ["Last Price", "VWAP", "Smart Money Score"])
    sm_15d = safe_round(sm_15d, ["Last Price", "VWAP", "Smart Money Score"])

    sm_1d = format_percent_cols(sm_1d, ["Momentum"])
    sm_7d = format_percent_cols(sm_7d, ["Momentum"])
    sm_15d = format_percent_cols(sm_15d, ["Momentum"])

    # -----------------------------
    # Top Pick Stocks - 1 Day (Top 7)
    # Source: Top_Picks
    # -----------------------------
    top_picks_1d = (
        top_picks[top_picks["Window"] == "1D"]
        .sort_values("Total_Qty", ascending=False)
        .head(7)
        .copy()
    )

    top_picks_1d = top_picks_1d[
        [c for c in ["Symbol", "Sectors", "Last_Price", "VWAP", "Total_Qty", "Trades"] if c in top_picks_1d.columns]
    ]

    top_picks_1d.rename(
        columns={
            "Sectors": "Sector",
            "Last_Price": "Last Price",
            "Total_Qty": "Quantity",
        },
        inplace=True,
    )

    top_picks_1d = safe_round(top_picks_1d, ["Last Price", "VWAP"])

    # -----------------------------
    # Top Pick Stocks - Last 7 Days (Top 7)
    # Source: Top_Picks
    # -----------------------------
    top_picks_7d = (
        top_picks[top_picks["Window"] == "7D"]
        .sort_values("Total_Qty", ascending=False)
        .head(7)
        .copy()
    )

    top_picks_7d = top_picks_7d[
        [c for c in ["Symbol", "Sectors", "Last_Price", "VWAP", "Total_Qty", "Trades"] if c in top_picks_7d.columns]
    ]

    top_picks_7d.rename(
        columns={
            "Sectors": "Sector",
            "Last_Price": "Last Price",
            "Total_Qty": "Quantity",
        },
        inplace=True,
    )

    top_picks_7d = safe_round(top_picks_7d, ["Last Price", "VWAP"])

    # -----------------------------
    # Best Top 5 Swing Trading Stocks - 7 Day
    # Source: Symbol_Summary
    # -----------------------------
    swing_7d = (
        symbol_summary[symbol_summary["Window"] == "7D"]
        .sort_values("Score", ascending=False)
        .head(5)
        .copy()
    )

    swing_7d = swing_7d[
        [c for c in ["Symbol", "Sectors", "Last_Price", "Momentum", "Range_%", "Vol_Surge", "Score", "Recommendation"] if c in swing_7d.columns]
    ]

    swing_7d.rename(
        columns={
            "Sectors": "Sector",
            "Last_Price": "Last Price",
            "Range_%": "Range %",
            "Vol_Surge": "Volume Surge",
        },
        inplace=True,
    )

    swing_7d = safe_round(swing_7d, ["Last Price", "Range %", "Volume Surge", "Score"])
    swing_7d = format_percent_cols(swing_7d, ["Momentum"])

    # -----------------------------
    # Highest Movement Stocks - 7 Day
    # Source: Price_Movers
    # -----------------------------
    pm_7d = (
        price_movers[price_movers["Window"] == "7D"]
        .sort_values("Change_%", ascending=False)
        .head(7)
        .copy()
    )

    pm_7d = pm_7d[
        [c for c in ["Symbol", "Sectors", "Close_start", "Close_end", "Change_%"] if c in pm_7d.columns]
    ]

    pm_7d.rename(
        columns={
            "Sectors": "Sector",
            "Close_start": "Start Price",
            "Close_end": "Last Price",
            "Change_%": "Change %",
        },
        inplace=True,
    )

    pm_7d = safe_round(pm_7d, ["Start Price", "Last Price"])
    pm_7d = format_percent_cols(pm_7d, ["Change %"])

    # -----------------------------
    # Most Volatile Stocks - 7 Day
    # Source: Symbol_Summary
    # -----------------------------
    ss_7d = symbol_summary[symbol_summary["Window"] == "7D"].copy()

    volatile_7d = ss_7d.sort_values("Range_%", ascending=False).head(7).copy()

    volatile_7d = volatile_7d[
        [c for c in ["Symbol", "Sectors", "Last_Price", "Range_%", "Vol_Surge"] if c in volatile_7d.columns]
    ]

    volatile_7d.rename(
        columns={
            "Sectors": "Sector",
            "Last_Price": "Last Price",
            "Range_%": "Range %",
            "Vol_Surge": "Volume Surge",
        },
        inplace=True,
    )

    volatile_7d = safe_round(volatile_7d, ["Last Price", "Range %", "Volume Surge"])

    report_date = pd.Timestamp.now().strftime("%Y-%m-%d")

    html = f"""
    <html>
    <body style="font-family:Arial; font-size:13px;">
      <p><b>Subject:</b> NEPSE Smart Money & Trading Summary - {report_date}</p>
      <p>Hi,</p>
      <p>Please find below today’s NEPSE trading summary.</p>

      {format_html_table(sm_1d, "Best Smart Money Stocks - 1 Day (Top 5)")}
      {format_html_table(sm_7d, "Best Smart Money Stocks - 7 Day (Top 5)")}
      {format_html_table(sm_15d, "Best Smart Money Stocks - 15 Day (Top 5)")}

      {format_html_table(top_picks_1d, "Top Pick Stocks - 1 Day (Top 7)")}
      {format_html_table(top_picks_7d, "Top Pick Stocks - Last 7 Days (Top 7)")}

      {format_html_table(swing_7d, "Best Top 5 Swing Trading Stocks - 7 Day")}
      {format_html_table(pm_7d, "Highest Movement Stocks - 7 Day (Top 7)")}
      {format_html_table(volatile_7d, "Most Volatile Stocks - 7 Day (Top 7)")}

      <br>
      <p>Regards,<br>Automated Trading Report Bot</p>
    </body>
    </html>
    """
    return html, report_date


def send_email(subject, html_body):
    email_user = os.environ["EMAIL_USER"]
    email_pass = os.environ["EMAIL_PASS"]
    email_to = os.environ["EMAIL_TO"]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = email_user
    msg["To"] = email_to
    msg.attach(MIMEText(html_body, "html"))

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(email_user, email_pass)
        server.sendmail(email_user, email_to, msg.as_string())

    print("Email summary sent successfully.")


def main():
    report_path = load_latest_report()
    html_body, report_date = build_email_body(report_path)
    subject = f"NEPSE Smart Money & Trading Summary - {report_date}"
    send_email(subject, html_body)


if __name__ == "__main__":
    main()
