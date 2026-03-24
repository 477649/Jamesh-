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


def get_table_theme(title):
    title_upper = title.upper()

    if "1D" in title_upper:
        return {
            "accent": "#1565c0",
            "border": "#90caf9",
            "header_bg": "linear-gradient(90deg,#1565c0,#1e88e5)",
            "section_bg": "linear-gradient(180deg,#f8fbff,#eef5ff)",
            "empty_bg": "linear-gradient(180deg,#f8fbff,#eef5ff)",
            "stripe": "#f7fbff",
        }
    if "7D" in title_upper:
        return {
            "accent": "#2e7d32",
            "border": "#a5d6a7",
            "header_bg": "linear-gradient(90deg,#2e7d32,#43a047)",
            "section_bg": "linear-gradient(180deg,#f7fff8,#eef9f0)",
            "empty_bg": "linear-gradient(180deg,#f7fff8,#eef9f0)",
            "stripe": "#f8fff8",
        }
    if "15D" in title_upper:
        return {
            "accent": "#6a1b9a",
            "border": "#d1c4e9",
            "header_bg": "linear-gradient(90deg,#6a1b9a,#8e24aa)",
            "section_bg": "linear-gradient(180deg,#fcf9ff,#f4ecfb)",
            "empty_bg": "linear-gradient(180deg,#fcf9ff,#f4ecfb)",
            "stripe": "#fcf8ff",
        }
    if "SECTOR" in title_upper:
        return {
            "accent": "#ef6c00",
            "border": "#ffcc80",
            "header_bg": "linear-gradient(90deg,#ef6c00,#fb8c00)",
            "section_bg": "linear-gradient(180deg,#fffaf5,#fff3e8)",
            "empty_bg": "linear-gradient(180deg,#fffaf5,#fff3e8)",
            "stripe": "#fffaf6",
        }
    if "OPERATOR" in title_upper or "WARNING" in title_upper:
        return {
            "accent": "#c62828",
            "border": "#ef9a9a",
            "header_bg": "linear-gradient(90deg,#c62828,#e53935)",
            "section_bg": "linear-gradient(180deg,#fff8f8,#ffefef)",
            "empty_bg": "linear-gradient(180deg,#fff8f8,#ffefef)",
            "stripe": "#fff9f9",
        }

    return {
        "accent": "#0b3d91",
        "border": "#cfd8dc",
        "header_bg": "linear-gradient(90deg,#0b3d91,#1565c0)",
        "section_bg": "linear-gradient(180deg,#f8fbfd,#eef4f8)",
        "empty_bg": "linear-gradient(180deg,#f8fbfd,#eef4f8)",
        "stripe": "#fafcfd",
    }


def format_html_table(df, title):
    title_html = html.escape(title)
    theme = get_table_theme(title)

    if df is None or df.empty:
        return f"""
        <div style="
            margin-bottom:20px;
            border:1px solid {theme['border']};
            border-left:6px solid {theme['accent']};
            border-radius:10px;
            background:{theme['empty_bg']};
            font-family:Arial;
            box-shadow:0 2px 8px rgba(0,0,0,0.05);
            overflow:hidden;">
            <div style="
                background:{theme['header_bg']};
                color:#ffffff;
                padding:10px 14px;
                font-size:15px;
                font-weight:bold;">
                {title_html}
            </div>
            <div style="padding:14px 16px;color:#546e7a;">
                No data available.
            </div>
        </div>
        """

    parts = [
        f"""
        <div style="
            margin-bottom:20px;
            border:1px solid {theme['border']};
            border-left:6px solid {theme['accent']};
            border-radius:10px;
            overflow:hidden;
            background:{theme['section_bg']};
            box-shadow:0 2px 8px rgba(0,0,0,0.05);">
            <div style="
                background:{theme['header_bg']};
                color:#ffffff;
                padding:10px 14px;
                font-family:Arial;
                font-size:15px;
                font-weight:bold;">
                {title_html}
            </div>
            <table border="0" cellpadding="0" cellspacing="0"
                style="border-collapse:collapse; font-family:Arial; font-size:13px; width:100%; background:#ffffff;">
                <tr style="background:#eaf2fb;">
        """
    ]

    for col in df.columns:
        parts.append(
            f"<th style='padding:10px 8px;text-align:center;color:#263238;"
            f"border-bottom:1px solid {theme['border']};border-right:1px solid #e6edf2;'>"
            f"{html.escape(str(col))}</th>"
        )
    parts.append("</tr>")

    for idx, (_, row) in enumerate(df.iterrows()):
        row_bg = "#ffffff" if idx % 2 == 0 else theme["stripe"]
        parts.append(f"<tr style='background:{row_bg};'>")
        for col, val in row.items():
            cell = "" if pd.isna(val) else html.escape(str(val))
            style = get_cell_style(col, val)
            parts.append(
                f"<td style='padding:8px 6px;text-align:center;border-bottom:1px solid #edf2f7;"
                f"border-right:1px solid #f1f4f7;{style}'>{cell}</td>"
            )
        parts.append("</tr>")

    parts.append("</table></div>")
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


def build_summary_box(tp_df, sm_df, movers_df, vol_df, setup_df, title="Market Snapshot", top_pick_label="Top Pick"):
    top_pick = tp_df.iloc[0]["Symbol"] if not tp_df.empty else "-"
    smart_money = sm_df.iloc[0]["Symbol"] if not sm_df.empty else "-"
    highest_mover = movers_df.iloc[0]["Symbol"] if not movers_df.empty else "-"
    most_volatile = vol_df.iloc[0]["Symbol"] if not vol_df.empty else "-"
    best_setup = setup_df.iloc[0]["Symbol"] if not setup_df.empty else "-"

    title_upper = title.upper()
    if "1D" in title_upper:
        accent = "#1565c0"
        border = "#90caf9"
        bg = "linear-gradient(180deg,#f4f9ff,#e3f2fd)"
    elif "7D" in title_upper:
        accent = "#2e7d32"
        border = "#a5d6a7"
        bg = "linear-gradient(180deg,#f4fff6,#e8f5e9)"
    elif "15D" in title_upper:
        accent = "#6a1b9a"
        border = "#d1c4e9"
        bg = "linear-gradient(180deg,#fbf7ff,#f3e5f5)"
    else:
        accent = "#0b3d91"
        border = "#cfd8dc"
        bg = "linear-gradient(180deg,#f8fbfd,#eef4f8)"

    return f"""
    <div style="
        border:1px solid {border};
        border-left:6px solid {accent};
        background:{bg};
        padding:14px 16px;
        margin-bottom:18px;
        font-family:Arial;
        font-size:13px;
        border-radius:10px;
        box-sizing:border-box;
        box-shadow:0 2px 8px rgba(0,0,0,0.05);">
      <h3 style="margin:0 0 12px 0; color:{accent}; font-size:16px;">{html.escape(title)}</h3>
      <p style="margin:6px 0;"><b>{html.escape(top_pick_label)}:</b> {html.escape(str(top_pick))}</p>
      <p style="margin:6px 0;"><b>Best Trade Setup:</b> {html.escape(str(best_setup))}</p>
      <p style="margin:6px 0;"><b>Best Smart Money Stock:</b> {html.escape(str(smart_money))}</p>
      <p style="margin:6px 0;"><b>Highest Mover:</b> {html.escape(str(highest_mover))}</p>
      <p style="margin:6px 0;"><b>Most Volatile Stock:</b> {html.escape(str(most_volatile))}</p>
    </div>
    """


def build_market_direction_box(market_overview, symbol_summary, window_label="7D"):
    ov = filter_window(market_overview, window_label)
    if not ov.empty:
        row = ov.iloc[0]
        buy_count = int(pd.to_numeric(row.get("buy_count", 0), errors="coerce") or 0)
        hold_count = int(pd.to_numeric(row.get("hold_count", 0), errors="coerce") or 0)
        sell_count = int(pd.to_numeric(row.get("sell_count", 0), errors="coerce") or 0)
        median_score = pd.to_numeric(row.get("median_score", None), errors="coerce")
        avg_vol_surge = pd.to_numeric(row.get("avg_vol_surge", None), errors="coerce")
    else:
        ss = filter_window(symbol_summary, window_label)
        rec = ss.get("recommendation", pd.Series(dtype=str)).astype(str).str.upper().str.strip()
        buy_count = int((rec == "BUY").sum())
        hold_count = int((rec == "HOLD").sum())
        sell_count = int((rec == "SELL / AVOID").sum())
        median_score = pd.to_numeric(ss.get("score", pd.Series(dtype=float)), errors="coerce").median()
        avg_vol_surge = pd.to_numeric(ss.get("vol_surge", pd.Series(dtype=float)), errors="coerce").mean()

    if buy_count > sell_count and buy_count >= hold_count:
        bias = "Bullish"
        bias_color = "#1b5e20"
        bias_bg = "#e8f5e9"
        accent = "#2e7d32"
        border = "#a5d6a7"
        bg = "linear-gradient(180deg,#f4fff6,#e8f5e9)"
    elif sell_count > buy_count:
        bias = "Bearish"
        bias_color = "#b71c1c"
        bias_bg = "#ffebee"
        accent = "#c62828"
        border = "#ef9a9a"
        bg = "linear-gradient(180deg,#fff7f7,#ffebee)"
    else:
        bias = "Neutral"
        bias_color = "#8d6e63"
        bias_bg = "#fff8e1"
        accent = "#f9a825"
        border = "#ffe082"
        bg = "linear-gradient(180deg,#fffdf4,#fff8e1)"

    median_score_txt = f"{median_score:.2f}" if pd.notna(median_score) else "-"
    avg_vol_surge_txt = f"{avg_vol_surge:.2f}" if pd.notna(avg_vol_surge) else "-"

    return f"""
    <div style="
        border:1px solid {border};
        border-left:6px solid {accent};
        background:{bg};
        padding:14px 16px;
        margin-bottom:18px;
        font-family:Arial;
        font-size:13px;
        border-radius:10px;
        box-sizing:border-box;
        box-shadow:0 2px 8px rgba(0,0,0,0.05);">
      <h3 style="margin:0 0 12px 0; color:{accent}; font-size:16px;">Market Direction ({html.escape(window_label)})</h3>
      <p style="margin:6px 0;"><b>Bullish Stocks:</b> {buy_count}</p>
      <p style="margin:6px 0;"><b>Neutral Stocks:</b> {hold_count}</p>
      <p style="margin:6px 0;"><b>Bearish Stocks:</b> {sell_count}</p>
      <p style="margin:6px 0;"><b>Median Score:</b> {median_score_txt}</p>
      <p style="margin:6px 0;"><b>Average Volume Surge:</b> {avg_vol_surge_txt}</p>
      <p style="margin:10px 0 0 0;">
        <b>Market Bias:</b>
        <span style="
            display:inline-block;
            padding:4px 10px;
            border-radius:999px;
            background:{bias_bg};
            color:{bias_color};
            font-weight:bold;
            border:1px solid {border};">{bias}</span>
      </p>
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
        limit=15,
    )

    sm7 = prepare_table(
        filter_window(smart_money, "7D"),
        smart_money_map,
        round_cols=["Last Price", "VWAP", "Range %", "Volume Surge", "Smart Money Score"],
        percent_cols=["Momentum", "Range %"],
        scale_100_cols=["Momentum"],
        sort_by="smart_money_score",
        ascending=False,
        limit=15,
    )

    sm15 = prepare_table(
        filter_window(smart_money, "15D"),
        smart_money_map,
        round_cols=["Last Price", "VWAP", "Range %", "Volume Surge", "Smart Money Score"],
        percent_cols=["Momentum", "Range %"],
        scale_100_cols=["Momentum"],
        sort_by="smart_money_score",
        ascending=False,
        limit=15,
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
        sort_by="quantity",
        ascending=False,
        limit=15,
    )

    tp7 = prepare_table(
        filter_list(filter_window(top_picks, "7D"), "TOP_BUY"),
        top_picks_map,
        round_cols=["Last Price", "VWAP", "Score"],
        sort_by="quantity",
        ascending=False,
        limit=15,
    )

    tp15 = prepare_table(
        filter_list(filter_window(top_picks, "15D"), "TOP_BUY"),
        top_picks_map,
        round_cols=["Last Price", "VWAP", "Score"],
        sort_by="quantity",
        ascending=False,
        limit=15,
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
        limit=15,
    )

    movers_map = [
        ("symbol", "Symbol"),
        ("sector", "Sector"),
        ("start_price", "Start Price"),
        ("last_price", "Last Price"),
        ("change_pct", "Change %"),
    ]

    movers1 = prepare_table(
        filter_window(price_movers, "1D"),
        movers_map,
        round_cols=["Start Price", "Last Price"],
        percent_cols=["Change %"],
        sort_by="change_pct",
        ascending=False,
        limit=15,
    )

    movers7 = prepare_table(
        filter_window(price_movers, "7D"),
        movers_map,
        round_cols=["Start Price", "Last Price"],
        percent_cols=["Change %"],
        sort_by="change_pct",
        ascending=False,
        limit=15,
    )

    movers15 = prepare_table(
        filter_window(price_movers, "15D"),
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

    vol1 = prepare_table(
        filter_window(symbol_summary, "1D"),
        vol_map,
        round_cols=["Last Price", "Range %", "Volume Surge"],
        percent_cols=["Range %"],
        sort_by="range_pct",
        ascending=False,
        limit=15,
    )

    vol7 = prepare_table(
        filter_window(symbol_summary, "7D"),
        vol_map,
        round_cols=["Last Price", "Range %", "Volume Surge"],
        percent_cols=["Range %"],
        sort_by="range_pct",
        ascending=False,
        limit=15,
    )

    vol15 = prepare_table(
        filter_window(symbol_summary, "15D"),
        vol_map,
        round_cols=["Last Price", "Range %", "Volume Surge"],
        percent_cols=["Range %"],
        sort_by="range_pct",
        ascending=False,
        limit=15,
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
        filter_window(operator_radar, "15D"),
        operator_map,
        round_cols=["Operator Score", "Concentration %", "Flip Ratio"],
        percent_cols=["Concentration %"],
        sort_by="operator_score",
        ascending=False,
        limit=20,
    )

    header_html = f"""
    <div style="
        background:linear-gradient(90deg,#0b3d91,#1565c0);
        color:white;
        padding:16px 20px;
        border-radius:10px;
        margin-bottom:18px;
        box-shadow:0 3px 10px rgba(21,101,192,0.25);">
        <h2 style="margin:0;font-family:Arial;">NEPSE Smart Money & Trading Summary</h2>
        <p style="margin:6px 0 0 0;font-family:Arial;font-size:13px;">
            Report Date: {html.escape(report_date)}
        </p>
    </div>
    """

    summary_box_1d = build_summary_box(
        tp1, sm1, movers1, vol1, setups,
        title="Market Snapshot (1D)",
        top_pick_label="Top Pick Today"
    )

    summary_box_7d = build_summary_box(
        tp7, sm7, movers7, vol7, setups,
        title="Market Snapshot (7D)",
        top_pick_label="Top Pick 7D"
    )

    summary_box_15d = build_summary_box(
        tp15, sm15, movers15, vol15, setups,
        title="Market Snapshot (15D)",
        top_pick_label="Top Pick 15D"
    )

    market_direction_box_1d = build_market_direction_box(market_overview, symbol_summary, "1D")
    market_direction_box_7d = build_market_direction_box(market_overview, symbol_summary, "7D")
    market_direction_box_15d = build_market_direction_box(market_overview, symbol_summary, "15D")

    html_body = f"""
    <html>
    <body style="font-family:Arial;font-size:13px;background:#f7f9fc;color:#222;padding:10px;">

    {header_html}

    <p style="font-size:14px;color:#37474f;">Please find today's NEPSE trading summary report.</p>

    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="margin-bottom:18px;">
      <tr>
        <td width="33.33%" valign="top" style="padding-right:8px;">
          {summary_box_1d}
          {market_direction_box_1d}
        </td>
        <td width="33.33%" valign="top" style="padding-left:4px;padding-right:4px;">
          {summary_box_7d}
          {market_direction_box_7d}
        </td>
        <td width="33.33%" valign="top" style="padding-left:8px;">
          {summary_box_15d}
          {market_direction_box_15d}
        </td>
      </tr>
    </table>

    {format_html_table(tp1, "Today Top Buy Picks (1D)")}
    {format_html_table(tp7, "Top Buy Picks (7D)")}
    {format_html_table(sm1, "Best Smart Money Stocks (1D)")}
    {format_html_table(sm7, "Best Smart Money Stocks (7D)")}
    {format_html_table(sm15, "Best Smart Money Stocks (15D)")}
    {format_html_table(swing, "Best Swing Trading Stocks (7D)")}
    {format_html_table(setups, "Best Trade Setups (7D)")}
    {format_html_table(movers7, "Highest Movement Stocks (7D)")}
    {format_html_table(vol7, "Most Volatile Stocks (7D)")}
    {format_html_table(sectors, "Top Performing Sectors (7D)")}
    {format_html_table(operator_warn, "Operator Activity Warning (15D)")}

    <br>
    <p>,<br><b></b></p>

    </body>
    </html>
    """

    plain_text = f"""NEPSE Smart Money & Trading Summary
Report Date: {report_date}

Included Reports:
1. Market Snapshot (1D)
2. Market Direction (1D)
3. Market Snapshot (7D)
4. Market Direction (7D)
5. Market Snapshot (15D)
6. Market Direction (15D)
7. Today Top Buy Picks (1D)
8. Top Buy Picks (7D)
9. Best Smart Money Stocks (1D)
10. Best Smart Money Stocks (7D)
11. Best Smart Money Stocks (15D)
12. Best Swing Trading Stocks (7D)
13. Best Trade Setups (7D)
14. Highest Movement Stocks (7D)
15. Most Volatile Stocks (7D)
16. Top Performing Sectors (7D)
17. Operator Activity Warning (15D)

Regards,

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
