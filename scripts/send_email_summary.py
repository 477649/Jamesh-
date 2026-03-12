import os
import smtplib
import html
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
            series = pd.to_numeric(df[col], errors="coerce")
            df[col] = series.map(lambda x: f"{x:.2f}%" if pd.notna(x) else "")
    return df


def normalize_columns(df):
    df = df.copy()

    def norm(name):
        return (
            str(name)
            .strip()
            .replace("%", "pct")
            .replace("/", "_")
            .replace("-", "_")
            .replace(" ", "_")
            .lower()
        )

    df.columns = [norm(col) for col in df.columns]

    alias_map = {
        "window": "window",
        "symbol": "symbol",
        "company": "company",
        "sectors": "sector",
        "sector": "sector",
        "last_price": "last_price",
        "lastprice": "last_price",
        "vwap": "vwap",
        "momentum": "momentum",
        "range_pct": "range_pct",
        "range": "range_pct",
        "vol_surge": "vol_surge",
        "volume_surge": "vol_surge",
        "price_vs_vwap_pct": "price_vs_vwap_pct",
        "smartmoneyscore": "smart_money_score",
        "smart_money_score": "smart_money_score",
        "smartmoneysignal": "signal",
        "smart_money_signal": "signal",
        "signal": "signal",
        "total_qty": "quantity",
        "qty": "quantity",
        "trades": "trades",
        "score": "score",
        "recommendation": "recommendation",
        "close_start": "start_price",
        "close_end": "last_price",
        "change_pct": "change_pct",
    }

    rename_map = {}
    for col in df.columns:
        if col in alias_map:
            rename_map[col] = alias_map[col]

    df.rename(columns=rename_map, inplace=True)
    return df


def get_cell_style(column, value):
    text = "" if pd.isna(value) else str(value).strip().lower()

    if column in ["Signal", "Recommendation"]:
        if any(word in text for word in ["bull", "buy", "strong", "positive"]):
            return "background-color:#e8f5e9; color:#1b5e20; font-weight:bold;"
        if any(word in text for word in ["bear", "sell", "weak", "negative"]):
            return "background-color:#ffebee; color:#b71c1c; font-weight:bold;"

    if column in ["Momentum", "Change %", "Range %", "Price vs VWAP %"]:
        try:
            num = float(str(value).replace("%", "").replace(",", "").strip())
            if num > 0:
                return "color:#1b5e20; font-weight:bold;"
            if num < 0:
                return "color:#b71c1c; font-weight:bold;"
        except Exception:
            pass

    return ""


def format_html_table(df, title):
    title_html = html.escape(title)

    if df is None or df.empty:
        return f"<h3 style='margin-bottom:8px;color:#0b3d91;'>{title_html}</h3><p>No data available.</p>"

    parts = [
        f'<h3 style="margin-bottom:8px;color:#0b3d91;font-family:Arial;">{title_html}</h3>',
        '<table border="1" cellpadding="6" cellspacing="0" '
        'style="border-collapse:collapse; font-family:Arial; font-size:13px; margin-bottom:18px; width:100%;">',
        '<tr style="background-color:#d9eaf7;">',
    ]

    for col in df.columns:
        parts.append(f"<th style='padding:8px;text-align:center;'>{html.escape(str(col))}</th>")
    parts.append("</tr>")

    for _, row in df.iterrows():
        parts.append("<tr>")
        for col, val in row.items():
            cell = "" if pd.isna(val) else html.escape(str(val))
            style = get_cell_style(col, val)
            parts.append(f"<td style='padding:6px;text-align:center;{style}'>{cell}</td>")
        parts.append("</tr>")

    parts.append("</table>")
    return "".join(parts)


def build_summary_box(top_pick_df, sm1_df, movers_df, vol_df):
    top_pick = top_pick_df.iloc[0]["Symbol"] if not top_pick_df.empty else "-"
    smart_money = sm1_df.iloc[0]["Symbol"] if not sm1_df.empty else "-"
    highest_mover = movers_df.iloc[0]["Symbol"] if not movers_df.empty else "-"
    most_volatile = vol_df.iloc[0]["Symbol"] if not vol_df.empty else "-"

    return f"""
    <div style="
        border:1px solid #cfd8dc;
        background:#f8fbfd;
        padding:12px 14px;
        margin-bottom:18px;
        font-family:Arial;
        font-size:13px;
        border-radius:6px;">
      <h3 style="margin:0 0 10px 0; color:#0b3d91;">Market Snapshot</h3>
      <p style="margin:4px 0;"><b>Top Pick Today:</b> {html.escape(str(top_pick))}</p>
      <p style="margin:4px 0;"><b>Best Smart Money Stock:</b> {html.escape(str(smart_money))}</p>
      <p style="margin:4px 0;"><b>Highest 7D Mover:</b> {html.escape(str(highest_mover))}</p>
      <p style="margin:4px 0;"><b>Most Volatile 7D Stock:</b> {html.escape(str(most_volatile))}</p>
    </div>
    """


def filter_window(df, window_value):
    if "window" not in df.columns:
        return df.iloc[0:0].copy()
    return df[df["window"].astype(str).str.strip().str.upper() == window_value.upper()].copy()


def prepare_table(df, columns_map, round_cols=None, percent_cols=None, sort_by=None, ascending=False, limit=None):
    out = df.copy()

    if sort_by and sort_by in out.columns:
        out = out.sort_values(sort_by, ascending=ascending)

    if limit:
        out = out.head(limit).copy()

    result = pd.DataFrame()
    for source_col, display_col in columns_map:
        result[display_col] = out[source_col] if source_col in out.columns else ""

    if round_cols:
        result = safe_round(result, round_cols)

    if percent_cols:
        result = format_percent_cols(result, percent_cols)

    return result


def build_email_body(report_path):
    smart_money = normalize_columns(pd.read_excel(report_path, sheet_name="Smart_Money"))
    price_movers = normalize_columns(pd.read_excel(report_path, sheet_name="Price_Movers"))
    symbol_summary = normalize_columns(pd.read_excel(report_path, sheet_name="Symbol_Summary"))
    top_picks = normalize_columns(pd.read_excel(report_path, sheet_name="Top_Picks"))

    # Smart Money
    smart_money_map = [
        ("symbol", "Symbol"),
        ("sector", "Sector"),
        ("last_price", "Last Price"),
        ("vwap", "VWAP"),
        ("momentum", "Momentum"),
        ("range_pct", "Range %"),
        ("vol_surge", "Volume Surge"),
        ("price_vs_vwap_pct", "Price vs VWAP %"),
        ("smart_money_score", "Smart Money Score"),
        ("signal", "Signal"),
    ]

    sm1 = prepare_table(
        filter_window(smart_money, "1D"),
        smart_money_map,
        round_cols=["Last Price", "VWAP", "Range %", "Volume Surge", "Price vs VWAP %", "Smart Money Score"],
        percent_cols=["Momentum", "Range %", "Price vs VWAP %"],
        sort_by="smart_money_score",
        ascending=False,
        limit=5,
    )

    sm7 = prepare_table(
        filter_window(smart_money, "7D"),
        smart_money_map,
        round_cols=["Last Price", "VWAP", "Range %", "Volume Surge", "Price vs VWAP %", "Smart Money Score"],
        percent_cols=["Momentum", "Range %", "Price vs VWAP %"],
        sort_by="smart_money_score",
        ascending=False,
        limit=5,
    )

    sm15 = prepare_table(
        filter_window(smart_money, "15D"),
        smart_money_map,
        round_cols=["Last Price", "VWAP", "Range %", "Volume Surge", "Price vs VWAP %", "Smart Money Score"],
        percent_cols=["Momentum", "Range %", "Price vs VWAP %"],
        sort_by="smart_money_score",
        ascending=False,
        limit=5,
    )

    # Top Picks
    top_picks_map = [
        ("symbol", "Symbol"),
        ("sector", "Sector"),
        ("last_price", "Last Price"),
        ("vwap", "VWAP"),
        ("quantity", "Quantity"),
        ("trades", "Trades"),
    ]

    tp1 = prepare_table(
        filter_window(top_picks, "1D"),
        top_picks_map,
        round_cols=["Last Price", "VWAP"],
        sort_by="quantity",
        ascending=False,
        limit=5,
    )

    tp7 = prepare_table(
        filter_window(top_picks, "7D"),
        top_picks_map,
        round_cols=["Last Price", "VWAP"],
        sort_by="quantity",
        ascending=False,
        limit=5,
    )

    # Swing Trading
    swing_map = [
        ("symbol", "Symbol"),
        ("sector", "Sector"),
        ("last_price", "Last Price"),
        ("momentum", "Momentum"),
        ("range_pct", "Range %"),
        ("vol_surge", "Volume Surge"),
        ("score", "Score"),
        ("recommendation", "Recommendation"),
    ]

    swing = prepare_table(
        filter_window(symbol_summary, "7D"),
        swing_map,
        round_cols=["Last Price", "Range %", "Volume Surge", "Score"],
        percent_cols=["Momentum"],
        sort_by="score",
        ascending=False,
        limit=5,
    )

    # Highest Movement
    movers_map = [
        ("symbol", "Symbol"),
        ("sector", "Sector"),
        ("start_price", "Start Price"),
        ("last_price", "Last Price"),
        ("change_pct", "Change %"),
    ]

    movers = prepare_table(
        filter_window(price_movers, "7D"),
        movers_map,
        round_cols=["Start Price", "Last Price"],
        percent_cols=["Change %"],
        sort_by="change_pct",
        ascending=False,
        limit=7,
    )

    # Most Volatile
    vol_map = [
        ("symbol", "Symbol"),
        ("sector", "Sector"),
        ("last_price", "Last Price"),
        ("range_pct", "Range %"),
        ("vol_surge", "Volume Surge"),
    ]

    vol = prepare_table(
        filter_window(symbol_summary, "7D"),
        vol_map,
        round_cols=["Last Price", "Range %", "Volume Surge"],
        percent_cols=["Range %"],
        sort_by="range_pct",
        ascending=False,
        limit=7,
    )

    report_date = pd.Timestamp.now().strftime("%Y-%m-%d")

    header_html = f"""
    <div style="
        background:linear-gradient(90deg,#0b3d91,#1565c0);
        color:white;
        padding:16px 20px;
        border-radius:8px;
        margin-bottom:18px;">
        <h2 style="margin:0;font-family:Arial;">NEPSE Smart Money & Trading Summary</h2>
        <p style="margin:6px 0 0 0;font-family:Arial;font-size:13px;">
            Report Date: {report_date}
        </p>
    </div>
    """

    summary_box = build_summary_box(tp1, sm1, movers, vol)

    html_body = f"""
    <html>
    <body style="font-family:Arial;font-size:13px;background:#ffffff;color:#222;padding:10px;">

    {header_html}

    <p>Please find today's NEPSE trading summary.</p>

    {summary_box}

    {format_html_table(tp1, "Today Top Pick Stocks (Top 5)")}
    {format_html_table(tp7, "Top Pick Stocks - 7 Day")}
    {format_html_table(sm1, "Best Smart Money Stocks - 1 Day")}
    {format_html_table(sm7, "Best Smart Money Stocks - 7 Day")}
    {format_html_table(sm15, "Best Smart Money Stocks - 15 Day")}
    {format_html_table(swing, "Best Swing Trading Stocks - 7 Day")}
    {format_html_table(movers, "Highest Movement Stocks - 7 Day")}
    {format_html_table(vol, "Most Volatile Stocks - 7 Day")}

    <br>
    <p>Regards,<br><b>Trading Report Bot</b></p>

    </body>
    </html>
    """

    plain_text = f"""NEPSE Smart Money & Trading Summary
Report Date: {report_date}

Included Reports:
1. Today Top Pick Stocks (Top 5)
2. Top Pick Stocks - 7 Day
3. Best Smart Money Stocks - 1 Day
4. Best Smart Money Stocks - 7 Day
5. Best Smart Money Stocks - 15 Day
6. Best Swing Trading Stocks - 7 Day
7. Highest Movement Stocks - 7 Day
8. Most Volatile Stocks - 7 Day

Regards,
Trading Report Bot
"""

    return html_body, plain_text, report_date


def send_email(subject, html_body, plain_text):
    email_user = os.environ["EMAIL_USER"]
    email_pass = os.environ["EMAIL_PASS"]
    email_to_raw = os.environ["EMAIL_TO"]

    recipients = [e.strip() for e in email_to_raw.split(",") if e.strip()]
    if not recipients:
        raise ValueError("EMAIL_TO is empty.")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = email_user
    msg["To"] = ", ".join(recipients)

    msg.attach(MIMEText(plain_text, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(email_user, email_pass)
        server.sendmail(email_user, recipients, msg.as_string())

    print("Email summary sent successfully.")


def main():
    report_path = load_latest_report()
    html_body, plain_text, report_date = build_email_body(report_path)
    subject = f"NEPSE Smart Money & Trading Summary - {report_date}"
    send_email(subject, html_body, plain_text)


if __name__ == "__main__":
    main()
