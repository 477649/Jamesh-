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


def format_percent_cols(df, cols, scale_100_cols=None):
    df = df.copy()
    scale_100_cols = set(scale_100_cols or [])

    for col in cols:
        if col in df.columns:
            series = pd.to_numeric(df[col], errors="coerce")
            if col in scale_100_cols:
                series = series * 100
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
        "from": "from_date",
        "to": "to_date",
        "list": "list",
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
        "setupscore": "setup_score",
        "setup_score": "setup_score",
        "setup_tag": "setup_tag",
        "recommendation": "recommendation",
        "close_start": "start_price",
        "close_end": "last_price",
        "change_pct": "change_pct",
        "buy_pressure": "buy_pressure",
        "sell_pressure": "sell_pressure",
        "risk_flags": "risk_flags",
        "close_pos": "close_pos",
        "smartbrokerscore": "smart_broker_score",
        "institutionscore": "institution_score",
        "operatorscore": "operator_score",
        "amount_cr": "amount_cr",
        "avg_score": "avg_score",
        "avg_momentum": "avg_momentum",
        "avg_vol_surge": "avg_vol_surge",
        "buy_count": "buy_count",
        "hold_count": "hold_count",
        "sell_count": "sell_count",
        "median_score": "median_score",
        "avg_vol_surge": "avg_vol_surge",
        "top_sector": "top_sector",
        "broker": "broker",
        "brokername": "broker_name",
        "brokertype": "broker_type",
        "tag": "tag",
    }

    rename_map = {}
    for col in df.columns:
        if col in alias_map:
            rename_map[col] = alias_map[col]

    df.rename(columns=rename_map, inplace=True)
    return df


def get_report_date(report_path):
    try:
        overview = pd.read_excel(report_path, sheet_name="Market_Overview")
        overview = normalize_columns(overview)

        if "window" in overview.columns and "to_date" in overview.columns:
            one_day = overview[overview["window"].astype(str).str.upper() == "1D"]
            if not one_day.empty:
                val = one_day.iloc[0]["to_date"]
                return pd.to_datetime(val).strftime("%Y-%m-%d")
    except Exception:
        pass

    return pd.Timestamp.now().strftime("%Y-%m-%d")


def filter_window(df, window_value):
    if "window" not in df.columns:
        return df.iloc[0:0].copy()
    return df[df["window"].astype(str).str.strip().str.upper() == window_value.upper()].copy()


def filter_list(df, list_value):
    if "list" not in df.columns:
        return df.iloc[0:0].copy()
    return df[df["list"].astype(str).str.strip().str.upper() == list_value.upper()].copy()


def get_cell_style(column, value):
    text = "" if pd.isna(value) else str(value).strip().lower()

    if column in ["Signal", "Recommendation", "Setup Tag", "Tag"]:
        if any(word in text for word in ["bull", "buy", "strong", "positive", "accumulation", "early"]):
            return "background-color:#e8f5e9; color:#1b5e20; font-weight:bold;"
        if any(word in text for word in ["bear", "sell", "weak", "negative", "avoid", "caution", "distribution"]):
            return "background-color:#ffebee; color:#b71c1c; font-weight:bold;"
        if any(word in text for word in ["hold", "watch", "neutral"]):
            return "background-color:#fff8e1; color:#8d6e63; font-weight:bold;"

    if column in [
        "Momentum",
        "Change %",
        "Range %",
        "Price vs VWAP %",
        "Avg Momentum",
        "Avg Vol Surge",
    ]:
        try:
            num = float(str(value).replace("%", "").replace(",", "").strip())
            if num > 0:
                return "color:#1b5e20; font-weight:bold;"
            if num < 0:
                return "color:#b71c1c; font-weight:bold;"
        except Exception:
            pass

    if column in ["Operator Score"] and value not in ("", None):
        try:
            num = float(value)
            if num >= 75:
                return "background-color:#ffebee; color:#b71c1c; font-weight:bold;"
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


def prepare_table(
    df,
    columns_map,
    round_cols=None,
    percent_cols=None,
    scale_100_cols=None,
    sort_by=None,
    ascending=False,
    limit=None,
):
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
        result = format_percent_cols(result, percent_cols, scale_100_cols=scale_100_cols)

    return result


def build_summary_box(tp1_df, sm1_df, movers_df, vol_df, setup_df):
    top_pick = tp1_df.iloc[0]["Symbol"] if not tp1_df.empty else "-"
    smart_money = sm1_df.iloc[0]["Symbol"] if not sm1_df.empty else "-"
    highest_mover = movers_df.iloc[0]["Symbol"] if not movers_df.empty else "-"
    most_volatile = vol_df.iloc[0]["Symbol"] if not vol_df.empty else "-"
    best_setup = setup_df.iloc[0]["Symbol"] if not setup_df.empty else "-"

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
      <p style="margin:4px 0;"><b>Best Trade Setup:</b> {html.escape(str(best_setup))}</p>
      <p style="margin:4px 0;"><b>Best Smart Money Stock:</b> {html.escape(str(smart_money))}</p>
      <p style="margin:4px 0;"><b>Highest 7D Mover:</b> {html.escape(str(highest_mover))}</p>
      <p style="margin:4px 0;"><b>Most Volatile 7D Stock:</b> {html.escape(str(most_volatile))}</p>
    </div>
    """


def build_market_direction_box(market_overview, symbol_summary):
    box_style = """
        border:1px solid #cfd8dc;
        background:#f8fbfd;
        padding:12px 14px;
        margin-bottom:18px;
        font-family:Arial;
        font-size:13px;
        border-radius:6px;
    """

    ov7 = filter_window(market_overview, "7D")
    if not ov7.empty:
        row = ov7.iloc[0]
        buy_count = int(pd.to_numeric(row.get("buy_count", 0), errors="coerce") or 0)
        hold_count = int(pd.to_numeric(row.get("hold_count", 0), errors="coerce") or 0)
        sell_count = int(pd.to_numeric(row.get("sell_count", 0), errors="coerce") or 0)
        median_score = pd.to_numeric(row.get("median_score", None), errors="coerce")
        avg_vol_surge = pd.to_numeric(row.get("avg_vol_surge", None), errors="coerce")
    else:
        ss7 = filter_window(symbol_summary, "7D")
        buy_count = int((ss7.get("recommendation", pd.Series(dtype=str)) == "BUY").sum())
        hold_count = int((ss7.get("recommendation", pd.Series(dtype=str)) == "HOLD").sum())
        sell_count = int((ss7.get("recommendation", pd.Series(dtype=str)) == "SELL / AVOID").sum())
        median_score = pd.to_numeric(ss7.get("score", pd.Series(dtype=float)), errors="coerce").median()
        avg_vol_surge = pd.to_numeric(ss7.get("vol_surge", pd.Series(dtype=float)), errors="coerce").mean()

    if buy_count > sell_count and buy_count >= hold_count:
        bias = "Bullish"
        bias_color = "#1b5e20"
    elif sell_count > buy_count:
        bias = "Bearish"
        bias_color = "#b71c1c"
    else:
        bias = "Neutral"
        bias_color = "#8d6e63"

    median_score_txt = f"{median_score:.2f}" if pd.notna(median_score) else "-"
    avg_vol_surge_txt = f"{avg_vol_surge:.2f}" if pd.notna(avg_vol_surge) else "-"

    return f"""
    <div style="{box_style}">
      <h3 style="margin:0 0 10px 0; color:#0b3d91;">Market Direction (7D)</h3>
      <p style="margin:4px 0;"><b>Bullish Stocks:</b> {buy_count}</p>
      <p style="margin:4px 0;"><b>Neutral Stocks:</b> {hold_count}</p>
      <p style="margin:4px 0;"><b>Bearish Stocks:</b> {sell_count}</p>
      <p style="margin:4px 0;"><b>Median Score:</b> {median_score_txt}</p>
      <p style="margin:4px 0;"><b>Average Volume Surge:</b> {avg_vol_surge_txt}</p>
      <p style="margin:6px 0 0 0;"><b>Market Bias:</b> <span style="color:{bias_color};font-weight:bold;">{bias}</span></p>
    </div>
    """


def build_email_body(report_path):
    market_overview = normalize_columns(pd.read_excel(report_path, sheet_name="Market_Overview"))
    smart_money = normalize_columns(pd.read_excel(report_path, sheet_name="Smart_Money"))
    price_movers = normalize_columns(pd.read_excel(report_path, sheet_name="Price_Movers"))
    symbol_summary = normalize_columns(pd.read_excel(report_path, sheet_name="Symbol_Summary"))
    top_picks = normalize_columns(pd.read_excel(report_path, sheet_name="Top_Picks"))
    trade_setups = normalize_columns(pd.read_excel(report_path, sheet_name="Trade_Setups"))
    sector_summary = normalize_columns(pd.read_excel(report_path, sheet_name="Sector_Summary"))
    operator_radar = normalize_columns(pd.read_excel(report_path, sheet_name="Operator_Radar"))

    report_date = get_report_date(report_path)

    smart_money_map = [
        ("symbol", "Symbol"),
        ("sector", "Sector"),
        ("last_price", "Last Price"),
        ("vwap", "VWAP"),
        ("momentum", "Momentum"),
        ("range_pct", "Range %"),
        ("vol_surge", "Volume Surge"),
        ("smart_money_score", "Smart Money Score"),
        ("signal", "Signal"),
    ]

    sm1 = prepare_table(
        filter_window(smart_money, "1D"),
        smart_money_map,
        round_cols=["Last Price", "VWAP", "Range %", "Volume Surge", "Smart Money Score"],
        percent_cols=["Momentum", "Range %"],
        scale_100_cols=["Momentum"],
        sort_by="smart_money_score",
        ascending=False,
        limit=10,
    )

    sm7 = prepare_table(
        filter_window(smart_money, "7D"),
        smart_money_map,
        round_cols=["Last Price", "VWAP", "Range %", "Volume Surge", "Smart Money Score"],
        percent_cols=["Momentum", "Range %"],
        scale_100_cols=["Momentum"],
        sort_by="smart_money_score",
        ascending=False,
        limit=10,
    )

    sm15 = prepare_table(
        filter_window(smart_money, "15D"),
        smart_money_map,
        round_cols=["Last Price", "VWAP", "Range %", "Volume Surge", "Smart Money Score"],
        percent_cols=["Momentum", "Range %"],
        scale_100_cols=["Momentum"],
        sort_by="smart_money_score",
        ascending=False,
        limit=10,
    )

    top_picks_map = [
        ("symbol", "Symbol"),
        ("sector", "Sector"),
        ("last_price", "Last Price"),
        ("vwap", "VWAP"),
        ("score", "Score"),
        ("quantity", "Quantity"),
        ("trades", "Trades"),
        ("recommendation", "Recommendation"),
    ]

    tp1 = prepare_table(
        filter_list(filter_window(top_picks, "1D"), "TOP_BUY"),
        top_picks_map,
        round_cols=["Last Price", "VWAP", "Score"],
        sort_by="score",
        ascending=False,
        limit=10,
    )

    tp7 = prepare_table(
        filter_list(filter_window(top_picks, "7D"), "TOP_BUY"),
        top_picks_map,
        round_cols=["Last Price", "VWAP", "Score"],
        sort_by="score",
        ascending=False,
        limit=10,
    )

    swing_map = [
        ("symbol", "Symbol"),
        ("sector", "Sector"),
        ("last_price", "Last Price"),
        ("momentum", "Momentum"),
        ("range_pct", "Range %"),
        ("vol_surge", "Volume Surge"),
        ("score", "Score"),
        ("recommendation", "Recommendation"),
        ("risk_flags", "Risk Flags"),
    ]

    swing = prepare_table(
        filter_window(symbol_summary, "7D"),
        swing_map,
        round_cols=["Last Price", "Range %", "Volume Surge", "Score"],
        percent_cols=["Momentum"],
        scale_100_cols=["Momentum"],
        sort_by="score",
        ascending=False,
        limit=10,
    )

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
        limit=10,
    )

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
        limit=10,
    )

    setup_map = [
        ("symbol", "Symbol"),
        ("sector", "Sector"),
        ("setup_score", "Setup Score"),
        ("setup_tag", "Setup Tag"),
        ("score", "Score"),
        ("smart_money_score", "Smart Money Score"),
        ("vol_surge", "Volume Surge"),
        ("momentum", "Momentum"),
    ]

    setups = prepare_table(
        filter_window(trade_setups, "7D"),
        setup_map,
        round_cols=["Setup Score", "Score", "Smart Money Score", "Volume Surge"],
        percent_cols=["Momentum"],
        scale_100_cols=["Momentum"],
        sort_by="setup_score",
        ascending=False,
        limit=10,
    )

    sector_map = [
        ("sector", "Sector"),
        ("symbols", "Symbols"),
        ("amount_cr", "Amount Cr"),
        ("avg_score", "Avg Score"),
        ("avg_momentum", "Avg Momentum"),
        ("avg_vol_surge", "Avg Vol Surge"),
    ]

    sectors = prepare_table(
        filter_window(sector_summary, "7D"),
        sector_map,
        round_cols=["Amount Cr", "Avg Score", "Avg Vol Surge"],
        percent_cols=["Avg Momentum"],
        scale_100_cols=["Avg Momentum"],
        sort_by="avg_score",
        ascending=False,
        limit=10,
    )

    operator_map = [
        ("symbol", "Symbol"),
        ("broker", "Broker"),
        ("broker_name", "Broker Name"),
        ("broker_type", "Broker Type"),
        ("operator_score", "Operator Score"),
        ("concentration_pct", "Concentration %"),
        ("flip_ratio", "Flip Ratio"),
        ("tag", "Tag"),
    ]

    operator_warn = prepare_table(
        filter_window(operator_radar, "7D"),
        operator_map,
        round_cols=["Operator Score", "Concentration %", "Flip Ratio"],
        percent_cols=["Concentration %"],
        sort_by="operator_score",
        ascending=False,
        limit=10,
    )

    header_html = f"""
    <div style="
        background:linear-gradient(90deg,#0b3d91,#1565c0);
        color:white;
        padding:16px 20px;
        border-radius:8px;
        margin-bottom:18px;">
        <h2 style="margin:0;font-family:Arial;">NEPSE Smart Money & Trading Summary</h2>
        <p style="margin:6px 0 0 0;font-family:Arial;font-size:13px;">
            Report Date: {html.escape(report_date)}
        </p>
    </div>
    """

    summary_box = build_summary_box(tp1, sm1, movers, vol, setups)
    market_direction_box = build_market_direction_box(market_overview, symbol_summary)

    html_body = f"""
    <html>
    <body style="font-family:Arial;font-size:13px;background:#ffffff;color:#222;padding:10px;">

    {header_html}

    <p>Please find today's NEPSE trading summary generated from the latest Retail-Pro Excel report.</p>

    {summary_box}
    {market_direction_box}

    {format_html_table(tp1, "Today Top Buy Picks (1D)")}
    {format_html_table(tp7, "Top Buy Picks (7D)")}
    {format_html_table(sm1, "Best Smart Money Stocks (1D)")}
    {format_html_table(sm7, "Best Smart Money Stocks (7D)")}
    {format_html_table(sm15, "Best Smart Money Stocks (15D)")}
    {format_html_table(swing, "Best Swing Trading Stocks (7D)")}
    {format_html_table(setups, "Best Trade Setups (7D)")}
    {format_html_table(movers, "Highest Movement Stocks (7D)")}
    {format_html_table(vol, "Most Volatile Stocks (7D)")}
    {format_html_table(sectors, "Top Performing Sectors (7D)")}
    {format_html_table(operator_warn, "Operator Activity Warning (7D)")}

    <br>
    <p>Regards,<br><b>Trading Report Bot</b></p>

    </body>
    </html>
    """

    plain_text = f"""NEPSE Smart Money & Trading Summary
Report Date: {report_date}

Included Reports:
1. Market Snapshot
2. Market Direction (7D)
3. Today Top Buy Picks (1D)
4. Top Buy Picks (7D)
5. Best Smart Money Stocks (1D)
6. Best Smart Money Stocks (7D)
7. Best Smart Money Stocks (15D)
8. Best Swing Trading Stocks (7D)
9. Best Trade Setups (7D)
10. Highest Movement Stocks (7D)
11. Most Volatile Stocks (7D)
12. Top Performing Sectors (7D)
13. Operator Activity Warning (7D)

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
