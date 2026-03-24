# floor_ohlc_insight_report.py
# ------------------------------------------------------------
# 15D Floor Sheet + OHLC Insight Report
#
# Uses the SAME source structure from your generate_trading_report.py:
# - outputs/Floor Sheet/floorsheet_YYYY-MM-DD.csv
# - outputs/sharesansar/SharePrice_YYYY-MM-DD.csv
# - outputs/Sector/sector_master.csv
#
# Goal:
# - capture detailed insight from floor sheet + OHLC
# - classify Best Buy / Buy / Hold / Sell
# - generate entry / stop loss / target
# - show broker accumulation / distribution
#
# Sheets:
# - README
# - Detailed_Insights
# - Best_Buy
# - Buy_Setups
# - Hold_List
# - Sell_List
# - Broker_Accumulation
# - Broker_Distribution
# ------------------------------------------------------------

import re
from pathlib import Path
from datetime import datetime

import numpy as np
import pandas as pd

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter


# =========================
# CONFIG
# =========================
ROOT = Path(__file__).resolve().parents[1]

FLOOR_DIR = ROOT / "outputs" / "Floor Sheet"
PRICE_DIR = ROOT / "outputs" / "sharesansar"
SECTOR_PATH = ROOT / "outputs" / "Sector" / "sector_master.csv"

OUT_DIR = ROOT / "outputs" / "reports"
OUT_DIR.mkdir(parents=True, exist_ok=True)

FLOOR_RE = re.compile(r".*?(\d{4}-\d{2}-\d{2}).*\.csv$", re.IGNORECASE)
PRICE_RE = re.compile(r".*?(\d{4}-\d{2}-\d{2}).*\.csv$", re.IGNORECASE)

LOOKBACK_DAYS = 15
MIN_PRICE_HISTORY = 20
CRORE = 10_000_000

BEST_BUY_MIN_SCORE = 78
BUY_MIN_SCORE = 63
HOLD_MIN_SCORE = 45

TOP_BROKER_N = 8


# =========================
# HELPERS
# =========================
def safe_float(v):
    if v is None:
        return np.nan
    try:
        s = str(v).replace(",", "").strip()
        if s == "" or s.lower() == "nan":
            return np.nan
        return float(s)
    except Exception:
        return np.nan


def list_dates_from_folder(folder: Path, regex: re.Pattern):
    dates, files = [], []
    if not folder.exists():
        return dates, files

    for p in folder.glob("*.csv"):
        m = regex.match(p.name)
        if not m:
            continue
        try:
            d = pd.to_datetime(m.group(1)).date()
        except Exception:
            continue
        dates.append(d)
        files.append(p)

    pairs = sorted(zip(dates, files), key=lambda x: x[0])
    return [d for d, _ in pairs], [f for _, f in pairs]


def choose_window_dates(all_dates_sorted, n):
    if not all_dates_sorted:
        return []
    return all_dates_sorted[-min(n, len(all_dates_sorted)) :]


def zscore(s: pd.Series):
    s = pd.to_numeric(s, errors="coerce").replace([np.inf, -np.inf], np.nan).fillna(0.0)
    std = s.std(ddof=0)
    if std == 0 or np.isnan(std):
        return pd.Series(np.zeros(len(s)), index=s.index)
    return (s - s.mean()) / (std + 1e-9)


def rsi(close: pd.Series, n=14):
    delta = close.diff()
    gain = delta.clip(lower=0)
    loss = (-delta.clip(upper=0)).clip(lower=0)
    avg_gain = gain.rolling(n).mean()
    avg_loss = loss.rolling(n).mean()
    rs = avg_gain / (avg_loss + 1e-12)
    return 100 - (100 / (1 + rs))


def true_range(h, l, c):
    pc = c.shift(1)
    return pd.concat([(h - l), (h - pc).abs(), (l - pc).abs()], axis=1).max(axis=1)


# =========================
# LOADERS
# =========================
def load_sector_master(path: Path):
    if not path.exists():
        return pd.DataFrame(columns=["Symbol", "Company", "Sector"])

    df = pd.read_csv(path)
    df.columns = [c.strip() for c in df.columns]

    if "Sector" not in df.columns and "Sectors" in df.columns:
        df = df.rename(columns={"Sectors": "Sector"})

    for c in ["Symbol", "Company", "Sector"]:
        if c not in df.columns:
            df[c] = ""

    df["Symbol"] = df["Symbol"].astype(str).str.strip().str.upper()
    return df[["Symbol", "Company", "Sector"]].drop_duplicates(subset=["Symbol"])


def read_floorsheet_file(path: Path, trade_date):
    df = pd.read_csv(path)
    df.columns = [c.strip() for c in df.columns]

    rename_map = {
        "Transact No": "Transact No.",
        "Transact No.": "Transact No.",
        "Transaction No.": "Transact No.",
        "Qty": "Quantity",
        "QTY": "Quantity",
        "Amt": "Amount",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    needed = ["Transact No.", "Symbol", "Buyer", "Seller", "Quantity", "Rate", "Amount"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f"Floorsheet missing columns {missing} in {path.name}")

    df["TradeDate"] = pd.to_datetime(trade_date)
    df["Quantity"] = df["Quantity"].apply(safe_float).fillna(0).astype(float)
    df["Rate"] = df["Rate"].apply(safe_float).fillna(0).astype(float)
    df["Amount"] = df["Amount"].apply(safe_float).fillna(0).astype(float)

    for c in ["Buyer", "Seller"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").astype("Int64")

    df["Symbol"] = df["Symbol"].astype(str).str.strip().str.upper()
    return df[needed + ["TradeDate"]]


def read_price_file(path: Path, trade_date):
    df = pd.read_csv(path)
    df.columns = [c.strip() for c in df.columns]

    if "Close" not in df.columns and "LTP" in df.columns:
        df["Close"] = df["LTP"]
    if "Volume" not in df.columns and "Vol" in df.columns:
        df["Volume"] = df["Vol"]
    if "VWAP" not in df.columns:
        df["VWAP"] = np.nan

    core = ["Symbol", "Open", "High", "Low", "Close", "Volume", "VWAP"]
    missing = [c for c in core if c not in df.columns]
    if missing:
        raise ValueError(f"Share price missing columns {missing} in {path.name}")

    df["TradeDate"] = pd.to_datetime(trade_date)
    for c in ["Open", "High", "Low", "Close", "Volume", "VWAP"]:
        df[c] = df[c].apply(safe_float)

    df["Symbol"] = df["Symbol"].astype(str).str.strip().str.upper()
    return df[["TradeDate", "Symbol", "Open", "High", "Low", "Close", "Volume", "VWAP"]]


# =========================
# FLOOR METRICS
# =========================
def symbol_metrics_from_floorsheet(fs: pd.DataFrame):
    if fs.empty:
        return pd.DataFrame(columns=["Symbol", "Trades", "Floor_Qty", "Floor_Amount", "Floor_VWAP", "Floor_Amount_Cr"])

    x = fs.copy()
    x["_qxr"] = x["Quantity"] * x["Rate"]

    g = x.groupby("Symbol", as_index=False).agg(
        Trades=("Transact No.", "count"),
        Floor_Qty=("Quantity", "sum"),
        Floor_Amount=("Amount", "sum"),
        _qxr=("_qxr", "sum"),
    )
    g["Floor_VWAP"] = np.where(g["Floor_Qty"] > 0, g["_qxr"] / g["Floor_Qty"], np.nan)
    g["Floor_Amount_Cr"] = g["Floor_Amount"] / CRORE
    return g.drop(columns=["_qxr"], errors="ignore")


def broker_symbol_metrics(fs: pd.DataFrame):
    if fs.empty:
        return pd.DataFrame(columns=["TradeDate", "Symbol", "Broker", "Buy_Qty", "Sell_Qty", "Net_Qty", "Buy_Amount"])

    buy = fs[["TradeDate", "Symbol", "Buyer", "Quantity", "Rate"]].copy()
    buy = buy.rename(columns={"Buyer": "Broker"})
    buy["Buy_Qty"] = buy["Quantity"]
    buy["Sell_Qty"] = 0.0
    buy["Buy_Amount"] = buy["Quantity"] * buy["Rate"]

    sell = fs[["TradeDate", "Symbol", "Seller", "Quantity", "Rate"]].copy()
    sell = sell.rename(columns={"Seller": "Broker"})
    sell["Buy_Qty"] = 0.0
    sell["Sell_Qty"] = sell["Quantity"]
    sell["Buy_Amount"] = 0.0

    x = pd.concat([buy, sell], ignore_index=True)
    x["Broker"] = x["Broker"].astype("Int64")

    g = x.groupby(["TradeDate", "Symbol", "Broker"], as_index=False).agg(
        Buy_Qty=("Buy_Qty", "sum"),
        Sell_Qty=("Sell_Qty", "sum"),
        Buy_Amount=("Buy_Amount", "sum"),
    )
    g["Net_Qty"] = g["Buy_Qty"] - g["Sell_Qty"]
    g["Active_Buy_Day"] = (g["Net_Qty"] > 0).astype(int)
    return g


def compute_pressure(bs_window: pd.DataFrame, topn=5):
    if bs_window.empty:
        return pd.DataFrame(columns=["Symbol", "Buy_Pressure", "Sell_Pressure", "Net_Broker_Bias"])

    t = bs_window.copy()
    t["pos"] = t["Net_Qty"].clip(lower=0)
    t["neg_abs"] = (-t["Net_Qty"]).clip(lower=0)

    pos_total = t.groupby("Symbol", as_index=False)["pos"].sum().rename(columns={"pos": "pos_total"})
    neg_total = t.groupby("Symbol", as_index=False)["neg_abs"].sum().rename(columns={"neg_abs": "neg_total"})

    top_pos = (
        t.sort_values(["Symbol", "pos"], ascending=[True, False])
        .groupby("Symbol")
        .head(topn)
        .groupby("Symbol", as_index=False)["pos"]
        .sum()
        .rename(columns={"pos": "top_pos"})
    )
    top_neg = (
        t.sort_values(["Symbol", "neg_abs"], ascending=[True, False])
        .groupby("Symbol")
        .head(topn)
        .groupby("Symbol", as_index=False)["neg_abs"]
        .sum()
        .rename(columns={"neg_abs": "top_neg"})
    )

    out = (
        pos_total.merge(top_pos, on="Symbol", how="left")
        .merge(neg_total, on="Symbol", how="left")
        .merge(top_neg, on="Symbol", how="left")
    )

    out["Buy_Pressure"] = np.where(out["pos_total"] > 0, out["top_pos"].fillna(0) / out["pos_total"], np.nan)
    out["Sell_Pressure"] = np.where(out["neg_total"] > 0, out["top_neg"].fillna(0) / out["neg_total"], np.nan)
    out["Net_Broker_Bias"] = out["pos_total"].fillna(0) - out["neg_total"].fillna(0)

    return out[["Symbol", "Buy_Pressure", "Sell_Pressure", "Net_Broker_Bias"]]


def top_broker_accumulation_distribution(bs_window: pd.DataFrame, topn=TOP_BROKER_N):
    if bs_window.empty:
        empty = pd.DataFrame(columns=["Symbol", "Broker", "Net_Qty", "Buy_Qty", "Sell_Qty", "Active_Days", "Side"])
        return empty.copy(), empty.copy()

    acc = (
        bs_window.sort_values(["Symbol", "Net_Qty"], ascending=[True, False])
        .groupby("Symbol")
        .head(topn)
        .copy()
    )
    acc["Side"] = "ACCUMULATION"

    dist = (
        bs_window.sort_values(["Symbol", "Net_Qty"], ascending=[True, True])
        .groupby("Symbol")
        .head(topn)
        .copy()
    )
    dist["Side"] = "DISTRIBUTION"

    return acc, dist


# =========================
# PRICE FEATURES
# =========================
def add_price_features(g: pd.DataFrame):
    g = g.sort_values("TradeDate").copy()

    g["MA5"] = g["Close"].rolling(5).mean()
    g["MA10"] = g["Close"].rolling(10).mean()
    g["MA15"] = g["Close"].rolling(15).mean()
    g["MA20"] = g["Close"].rolling(20).mean()

    g["VMA5"] = g["Volume"].rolling(5).mean()
    g["VMA10"] = g["Volume"].rolling(10).mean()
    g["VMA15"] = g["Volume"].rolling(15).mean()

    g["HH5"] = g["High"].rolling(5).max()
    g["HH10"] = g["High"].rolling(10).max()
    g["HH15"] = g["High"].rolling(15).max()

    g["LL5"] = g["Low"].rolling(5).min()
    g["LL10"] = g["Low"].rolling(10).min()
    g["LL15"] = g["Low"].rolling(15).min()

    for n in [1, 3, 5, 10, 15]:
        g[f"RET{n}"] = g["Close"].pct_change(n) * 100

    rng = (g["High"] - g["Low"]).replace(0, np.nan)
    g["Close_Pos"] = (g["Close"] - g["Low"]) / (rng + 1e-12)
    g["Body_Pct"] = (g["Close"] - g["Open"]).abs() / (g["Open"] + 1e-12) * 100
    g["UpperWickPct"] = ((g["High"] - g[["Open", "Close"]].max(axis=1)) / (rng + 1e-12)).clip(0, 1)
    g["LowerWickPct"] = ((g[["Open", "Close"]].min(axis=1) - g["Low"]) / (rng + 1e-12)).clip(0, 1)
    g["RangePct"] = (rng / (g["Close"] + 1e-12)) * 100

    g["RSI14"] = rsi(g["Close"], 14)
    tr = true_range(g["High"], g["Low"], g["Close"])
    g["ATR14"] = tr.rolling(14).mean()
    g["ATR14_Pct"] = g["ATR14"] / (g["Close"] + 1e-12) * 100

    g["Price_vs_VWAP_Pct"] = np.where(
        g["VWAP"].notna() & (g["VWAP"] > 0),
        (g["Close"] / g["VWAP"] - 1.0) * 100,
        np.nan,
    )
    g["Vol_Surge"] = np.where(g["VMA10"] > 0, g["Volume"] / (g["VMA10"] + 1e-12), np.nan)

    return g


def build_price_snapshot(price_all: pd.DataFrame):
    rows = []

    for sym, g in price_all.groupby("Symbol"):
        g = add_price_features(g)
        if len(g) < MIN_PRICE_HISTORY:
            continue

        last = g.iloc[-1]

        rows.append({
            "Symbol": sym,
            "TradeDate": last["TradeDate"],
            "Open": last["Open"],
            "High": last["High"],
            "Low": last["Low"],
            "Close": last["Close"],
            "Volume": last["Volume"],
            "VWAP": last["VWAP"],
            "MA5": last["MA5"],
            "MA10": last["MA10"],
            "MA15": last["MA15"],
            "MA20": last["MA20"],
            "HH5": last["HH5"],
            "HH10": last["HH10"],
            "HH15": last["HH15"],
            "LL5": last["LL5"],
            "LL10": last["LL10"],
            "LL15": last["LL15"],
            "RET1_%": last["RET1"],
            "RET3_%": last["RET3"],
            "RET5_%": last["RET5"],
            "RET10_%": last["RET10"],
            "RET15_%": last["RET15"],
            "RSI14": last["RSI14"],
            "ATR14": last["ATR14"],
            "ATR14_Pct": last["ATR14_Pct"],
            "RangePct": last["RangePct"],
            "Body_Pct": last["Body_Pct"],
            "Close_Pos": last["Close_Pos"],
            "UpperWickPct": last["UpperWickPct"],
            "LowerWickPct": last["LowerWickPct"],
            "Price_vs_VWAP_Pct": last["Price_vs_VWAP_Pct"],
            "Vol_Surge": last["Vol_Surge"],
        })

    return pd.DataFrame(rows)


# =========================
# SCORING LOGIC
# =========================
def build_insight_table(price_snapshot, floor_symbol, pressure, bs_window, sector):
    if price_snapshot.empty:
        return pd.DataFrame()

    x = price_snapshot.copy()

    x = x.merge(floor_symbol, on="Symbol", how="left")
    x = x.merge(pressure, on="Symbol", how="left")
    x = x.merge(sector, on="Symbol", how="left")

    if bs_window.empty:
        broker_stats = pd.DataFrame(columns=["Symbol", "Accum_Brokers", "Dist_Brokers", "Active_Brokers"])
    else:
        broker_stats = bs_window.groupby("Symbol", as_index=False).agg(
            Accum_Brokers=("Net_Qty", lambda s: int((s > 0).sum())),
            Dist_Brokers=("Net_Qty", lambda s: int((s < 0).sum())),
            Active_Brokers=("Broker", "nunique"),
        )

    x = x.merge(broker_stats, on="Symbol", how="left")

    for c in [
        "Trades", "Floor_Qty", "Floor_Amount", "Floor_VWAP", "Floor_Amount_Cr",
        "Buy_Pressure", "Sell_Pressure", "Net_Broker_Bias",
        "Accum_Brokers", "Dist_Brokers", "Active_Brokers"
    ]:
        if c not in x.columns:
            x[c] = np.nan

    x["Broker_Balance"] = x["Accum_Brokers"].fillna(0) - x["Dist_Brokers"].fillna(0)

    # Sub-scores
    x["PriceScore"] = (
        (x["Close"] > x["MA5"]).astype(int) * 8
        + (x["MA5"] > x["MA10"]).astype(int) * 12
        + (x["MA10"] > x["MA15"]).astype(int) * 12
        + (x["Close"] >= 0.98 * x["HH10"]).astype(int) * 8
        + (x["Close"] >= 0.97 * x["HH15"]).astype(int) * 8
        + (x["RSI14"].between(52, 72, inclusive="both")).astype(int) * 8
        + (x["Close_Pos"] >= 0.65).astype(int) * 6
        + (x["UpperWickPct"] <= 0.30).astype(int) * 4
        + (x["LowerWickPct"] >= x["UpperWickPct"]).astype(int) * 4
    ).clip(0, 70)

    x["VolumeScore"] = (
        (x["Vol_Surge"] >= 1.20).astype(int) * 8
        + (x["Vol_Surge"] >= 1.50).astype(int) * 7
        + (x["Floor_Amount_Cr"] >= 1.0).astype(int) * 5
        + (x["Trades"] >= 80).astype(int) * 5
    ).clip(0, 25)

    x["FloorScore"] = (
        (x["Buy_Pressure"] > x["Sell_Pressure"]).astype(int) * 10
        + (x["Buy_Pressure"] >= 0.55).astype(int) * 10
        + (x["Broker_Balance"] > 0).astype(int) * 6
        + (x["Accum_Brokers"] >= x["Dist_Brokers"]).astype(int) * 6
        + (x["Close"] >= x["Floor_VWAP"]).astype(int) * 8
    ).clip(0, 40)

    x["RiskPenalty"] = (
        (x["Sell_Pressure"] >= 0.55).astype(int) * 14
        + (x["Close"] < x["VWAP"]).astype(int) * 10
        + (x["Close"] < x["MA10"]).astype(int) * 10
        + (x["RSI14"] < 45).astype(int) * 8
        + (x["UpperWickPct"] >= 0.45).astype(int) * 5
        + (x["Close_Pos"] <= 0.35).astype(int) * 5
        + (x["RET5_%"] <= -4).astype(int) * 8
    ).clip(0, 50)

    raw = x["PriceScore"] + x["VolumeScore"] + x["FloorScore"] - x["RiskPenalty"]
    mn, mx = float(raw.min()), float(raw.max())
    x["InsightScore"] = np.where(mx > mn, 100 * (raw - mn) / (mx - mn), 50.0)

    # Recommendation
    x["Recommendation"] = "SELL"
    x.loc[x["InsightScore"] >= HOLD_MIN_SCORE, "Recommendation"] = "HOLD"
    x.loc[x["InsightScore"] >= BUY_MIN_SCORE, "Recommendation"] = "BUY"
    x.loc[x["InsightScore"] >= BEST_BUY_MIN_SCORE, "Recommendation"] = "BEST BUY"

    # Trade plan
    x["Entry"] = x["Close"].round(2)

    stop_base_1 = x["LL5"]
    stop_base_2 = x["Close"] - 1.2 * x["ATR14"]
    x["StopLoss"] = np.where(
        stop_base_1.notna() & stop_base_2.notna(),
        np.minimum(stop_base_1, stop_base_2),
        np.where(stop_base_1.notna(), stop_base_1, stop_base_2),
    )

    x["RiskPerShare"] = (x["Entry"] - x["StopLoss"]).clip(lower=0)
    x["Target1"] = (x["Entry"] + 1.5 * x["RiskPerShare"]).round(2)
    x["Target2"] = (x["Entry"] + 2.5 * x["RiskPerShare"]).round(2)
    x["Target3"] = (x["Entry"] + 3.0 * x["RiskPerShare"]).round(2)

    # Hold / Sell action text
    x["ActionPlan"] = ""
    x.loc[x["Recommendation"] == "BEST BUY", "ActionPlan"] = "Fresh opportunity; entry near current price / slight dip; trail below stop."
    x.loc[x["Recommendation"] == "BUY", "ActionPlan"] = "Buy on confirmation; avoid chasing extended candle."
    x.loc[x["Recommendation"] == "HOLD", "ActionPlan"] = "Hold if above VWAP / MA10; wait for better follow-through."
    x.loc[x["Recommendation"] == "SELL", "ActionPlan"] = "Reduce / avoid; weak price-floor alignment."

    # Reasons
    reasons = []
    for _, r in x.iterrows():
        tags = []

        if pd.notna(r["Buy_Pressure"]) and pd.notna(r["Sell_Pressure"]):
            if r["Buy_Pressure"] > r["Sell_Pressure"]:
                tags.append("buyer pressure")
            elif r["Sell_Pressure"] > r["Buy_Pressure"]:
                tags.append("seller pressure")

        if pd.notna(r["Close"]) and pd.notna(r["VWAP"]):
            if r["Close"] > r["VWAP"]:
                tags.append("above VWAP")
            else:
                tags.append("below VWAP")

        if pd.notna(r["Vol_Surge"]) and r["Vol_Surge"] >= 1.5:
            tags.append("volume surge")

        if pd.notna(r["RSI14"]):
            if 52 <= r["RSI14"] <= 72:
                tags.append("healthy RSI")
            elif r["RSI14"] < 45:
                tags.append("weak RSI")

        if pd.notna(r["Close"]) and pd.notna(r["HH10"]) and r["Close"] >= 0.98 * r["HH10"]:
            tags.append("near 10D breakout")

        if pd.notna(r["Accum_Brokers"]) and pd.notna(r["Dist_Brokers"]):
            if r["Accum_Brokers"] > r["Dist_Brokers"]:
                tags.append("broker accumulation")
            elif r["Dist_Brokers"] > r["Accum_Brokers"]:
                tags.append("broker distribution")

        if pd.notna(r["UpperWickPct"]) and r["UpperWickPct"] >= 0.45:
            tags.append("upper-wick risk")

        reasons.append(", ".join(tags[:6]))

    x["InsightReason"] = reasons

    # Risk flags
    x["RiskFlags"] = ""
    x.loc[x["Sell_Pressure"].fillna(0) >= 0.55, "RiskFlags"] += "SellWall; "
    x.loc[x["Close_Pos"].fillna(0.5) <= 0.35, "RiskFlags"] += "WeakClose; "
    x.loc[x["UpperWickPct"].fillna(0) >= 0.45, "RiskFlags"] += "SupplyWick; "
    x.loc[x["Close"].fillna(0) < x["VWAP"].fillna(np.inf), "RiskFlags"] += "BelowVWAP; "
    x.loc[x["RET5_%"].fillna(0) <= -4, "RiskFlags"] += "ShortTermWeakness; "

    # Final order
    cols = [
        "TradeDate", "Symbol", "Company", "Sector",
        "Recommendation", "InsightScore",
        "Entry", "StopLoss", "Target1", "Target2", "Target3",
        "Close", "VWAP", "Floor_VWAP",
        "RET1_%", "RET3_%", "RET5_%", "RET10_%", "RET15_%",
        "MA5", "MA10", "MA15",
        "HH10", "HH15", "LL5", "LL10",
        "Volume", "Vol_Surge", "Floor_Qty", "Floor_Amount_Cr", "Trades",
        "Buy_Pressure", "Sell_Pressure", "Net_Broker_Bias",
        "Accum_Brokers", "Dist_Brokers", "Active_Brokers",
        "RSI14", "ATR14", "ATR14_Pct", "RangePct", "Body_Pct", "Close_Pos",
        "Price_vs_VWAP_Pct",
        "PriceScore", "VolumeScore", "FloorScore", "RiskPenalty",
        "InsightReason", "RiskFlags", "ActionPlan",
    ]
    cols = [c for c in cols if c in x.columns]

    x = x[cols].copy()
    x = x.sort_values(["InsightScore", "Buy_Pressure", "RET10_%"], ascending=[False, False, False])
    return x


# =========================
# EXCEL HELPERS
# =========================
def style_sheet(ws):
    ws.freeze_panes = "A2"
    thin = Side(style="thin", color="D0D7DE")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="DCE6F1")

    for cell in ws[1]:
        cell.font = Font(bold=True, color="000000")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.border = border
            if isinstance(c.value, str) and len(c.value) > 50:
                c.alignment = Alignment(wrap_text=True, vertical="top")


def add_table_sheet(ws, df, name):
    if df.empty:
        df = pd.DataFrame([["No data"]], columns=["Info"])

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    style_sheet(ws)

    nrows = ws.max_row
    ncols = ws.max_column
    ref = f"A1:{get_column_letter(ncols)}{nrows}"

    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)

    auto_fit_columns(ws)


def auto_fit_columns(ws, min_width=10, max_width=45):
    for col in ws.columns:
        max_len = 0
        letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is None:
                continue
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[letter].width = max(min_width, min(max_width, max_len + 2))


def color_scale(ws, col_name):
    header = [c.value for c in ws[1]]
    if col_name not in header or ws.max_row < 2:
        return
    idx = header.index(col_name) + 1
    col = get_column_letter(idx)
    rng = f"{col}2:{col}{ws.max_row}"
    ws.conditional_formatting.add(
        rng,
        ColorScaleRule(
            start_type="min", start_color="F8696B",
            mid_type="percentile", mid_value=50, mid_color="FFEB84",
            end_type="max", end_color="63BE7B"
        )
    )


def number_format(ws, fmt_map):
    header = [c.value for c in ws[1]]
    for col_name, fmt in fmt_map.items():
        if col_name not in header:
            continue
        idx = header.index(col_name) + 1
        for r in range(2, ws.max_row + 1):
            ws.cell(r, idx).number_format = fmt


# =========================
# MAIN
# =========================
def main():
    floor_dates, floor_files = list_dates_from_folder(FLOOR_DIR, FLOOR_RE)
    price_dates, price_files = list_dates_from_folder(PRICE_DIR, PRICE_RE)

    if not floor_dates:
        raise RuntimeError(f"No floorsheet csv files found in {FLOOR_DIR}")
    if not price_dates:
        raise RuntimeError(f"No share price csv files found in {PRICE_DIR}")

    floor_map = {d: f for d, f in zip(floor_dates, floor_files)}
    price_map = {d: f for d, f in zip(price_dates, price_files)}

    # use common dates only
    common_dates = sorted(set(floor_dates).intersection(set(price_dates)))
    if not common_dates:
        raise RuntimeError("No common dates found between floor sheet and share price files.")

    win_dates = choose_window_dates(common_dates, LOOKBACK_DAYS)
    latest_date = pd.to_datetime(win_dates[-1])

    print(f"✅ Using last {len(win_dates)} trading days")
    print(f"✅ Latest date: {win_dates[-1]}")

    # load sector
    sector = load_sector_master(SECTOR_PATH)

    # load floorsheet window
    fs_list = [read_floorsheet_file(floor_map[d], d) for d in win_dates if d in floor_map]
    fs = pd.concat(fs_list, ignore_index=True) if fs_list else pd.DataFrame()

    # load full enough price history for indicators
    # use more than 15 days so MA20/RSI14 are valid
    price_hist_dates = choose_window_dates(common_dates, max(LOOKBACK_DAYS + 10, 30))
    pr_list = [read_price_file(price_map[d], d) for d in price_hist_dates if d in price_map]
    pr = pd.concat(pr_list, ignore_index=True) if pr_list else pd.DataFrame()

    if fs.empty:
        raise RuntimeError("Floorsheet window data is empty.")
    if pr.empty:
        raise RuntimeError("Price window data is empty.")

    # Floor metrics
    floor_symbol = symbol_metrics_from_floorsheet(fs)

    bs_daily = broker_symbol_metrics(fs)
    if not bs_daily.empty:
        bs_window = bs_daily.groupby(["Symbol", "Broker"], as_index=False).agg(
            Buy_Qty=("Buy_Qty", "sum"),
            Sell_Qty=("Sell_Qty", "sum"),
            Net_Qty=("Net_Qty", "sum"),
            Buy_Amount=("Buy_Amount", "sum"),
            Active_Days=("Active_Buy_Day", "sum"),
        )
    else:
        bs_window = pd.DataFrame(columns=["Symbol", "Broker", "Buy_Qty", "Sell_Qty", "Net_Qty", "Buy_Amount", "Active_Days"])

    pressure = compute_pressure(bs_window, topn=5)

    # Price snapshot
    price_snapshot = build_price_snapshot(pr)

    # Insight table
    detailed = build_insight_table(price_snapshot, floor_symbol, pressure, bs_window, sector)

    if detailed.empty:
        raise RuntimeError("No symbol qualified after processing.")

    best_buy = detailed[detailed["Recommendation"] == "BEST BUY"].copy().head(40)
    buy_list = detailed[detailed["Recommendation"] == "BUY"].copy().head(80)
    hold_list = detailed[detailed["Recommendation"] == "HOLD"].copy().head(100)
    sell_list = detailed[detailed["Recommendation"] == "SELL"].copy().head(100)

    acc_brokers, dist_brokers = top_broker_accumulation_distribution(bs_window, topn=TOP_BROKER_N)
    acc_brokers = acc_brokers.merge(sector, on="Symbol", how="left")
    dist_brokers = dist_brokers.merge(sector, on="Symbol", how="left")

    acc_brokers = acc_brokers.sort_values(["Net_Qty", "Active_Days"], ascending=[False, False])
    dist_brokers = dist_brokers.sort_values(["Net_Qty", "Active_Days"], ascending=[True, False])

    readme = pd.DataFrame([
        ["Purpose", "Capture detailed short-term insight from floor sheet and OHLC data."],
        ["Lookback", f"{LOOKBACK_DAYS} trading days for floor insight; longer price history used for indicators."],
        ["Floor Source", "outputs/Floor Sheet/floorsheet_YYYY-MM-DD.csv"],
        ["Price Source", "outputs/sharesansar/SharePrice_YYYY-MM-DD.csv"],
        ["Required Floor Columns", "Transact No., Symbol, Buyer, Seller, Quantity, Rate, Amount"],
        ["Recommendations", "BEST BUY / BUY / HOLD / SELL"],
        ["Best Buy Idea", "Strong price structure + buyer pressure + broker accumulation + VWAP support + volume confirmation"],
        ["Targets", "Target1=1.5R, Target2=2.5R, Target3=3R from Entry vs StopLoss"],
        ["Stop Logic", "Lower of LL5 and Close - 1.2*ATR14"],
    ], columns=["Item", "Explanation"])

    wb = Workbook()
    wb.remove(wb.active)

    sheets = {
        "README": readme,
        "Detailed_Insights": detailed,
        "Best_Buy": best_buy,
        "Buy_Setups": buy_list,
        "Hold_List": hold_list,
        "Sell_List": sell_list,
        "Broker_Accumulation": acc_brokers,
        "Broker_Distribution": dist_brokers,
    }

    fmt_map = {
        "InsightScore": "0.00",
        "Entry": "#,##0.00",
        "StopLoss": "#,##0.00",
        "Target1": "#,##0.00",
        "Target2": "#,##0.00",
        "Target3": "#,##0.00",
        "Close": "#,##0.00",
        "VWAP": "#,##0.00",
        "Floor_VWAP": "#,##0.00",
        "RET1_%": "0.00",
        "RET3_%": "0.00",
        "RET5_%": "0.00",
        "RET10_%": "0.00",
        "RET15_%": "0.00",
        "Buy_Pressure": "0.00",
        "Sell_Pressure": "0.00",
        "RSI14": "0.00",
        "ATR14": "0.00",
        "ATR14_Pct": "0.00",
        "RangePct": "0.00",
        "Body_Pct": "0.00",
        "Close_Pos": "0.00",
        "Price_vs_VWAP_Pct": "0.00",
        "Vol_Surge": "0.00",
        "Floor_Amount_Cr": "0.000",
        "PriceScore": "0",
        "VolumeScore": "0",
        "FloorScore": "0",
        "RiskPenalty": "0",
        "Net_Qty": "#,##0",
        "Buy_Qty": "#,##0",
        "Sell_Qty": "#,##0",
        "Buy_Amount": "#,##0.00",
        "Active_Days": "0",
    }

    for i, (sheet_name, df) in enumerate(sheets.items(), start=1):
        ws = wb.create_sheet(sheet_name[:31])
        add_table_sheet(ws, df, f"T{i}")
        number_format(ws, fmt_map)

        for c in ["InsightScore", "Buy_Pressure", "Sell_Pressure", "Vol_Surge", "RET10_%", "PriceScore", "FloorScore"]:
            color_scale(ws, c)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = OUT_DIR / f"Floor_OHLC_Insight_Report_{ts}.xlsx"
    wb.save(out_path)

    print(f"✅ Report generated: {out_path}")


if __name__ == "__main__":
    main()
