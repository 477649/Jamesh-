# priceaction_signals.py
# ------------------------------------------------------------
# OHLCV + Floor Sheet Price Action + Early Smart Money Report
#
# Repo structure:
# - outputs/Floor Sheet/floorsheet_YYYY-MM-DD.csv
# - outputs/sharesansar/SharePrice_YYYY-MM-DD.csv
# - outputs/Sector/sector_master.csv
#
# Output:
# - outputs/PriceAction/nepse_signals_<latest_date>.xlsx
#
# Key fixes vs old version:
# - snapshot now only includes symbols whose latest price row is on latest_dt
# - latest-day floorsheet metrics are separated from rolling-window metrics
# - score logic is stable weighted scoring, not cross-sectional min-max scaling
# - Trades uses nunique on Transact No.
# - sentiment precedence is explicit: DISTRIBUTION > SHAKEOUT > ACCUMULATION > NEUTRAL
# - output lists are no longer truncated with head()
# - debug prints replaced with logging
# ------------------------------------------------------------

import logging
import re
from math import ceil
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo

# =========================
# PATHS / SETTINGS
# =========================
ROOT = Path(__file__).resolve().parent
PRICE_DIR = ROOT / "outputs" / "sharesansar"
FLOOR_DIR = ROOT / "outputs" / "Floor Sheet"
SECTOR_FILE = ROOT / "outputs" / "Sector" / "sector_master.csv"
OUT_DIR = ROOT / "outputs" / "PriceAction"
OUT_DIR.mkdir(parents=True, exist_ok=True)

PRICE_RE = re.compile(r"SharePrice_(\d{4}-\d{2}-\d{2})", re.IGNORECASE)
FLOOR_RE = re.compile(r".*?(\d{4}-\d{2}-\d{2}).*\.csv$", re.IGNORECASE)

LOOKBACK_SIGNAL = 15
PRICE_HISTORY_LOAD = 45
TOP_BROKER_N = 8
CRORE = 10_000_000

BEST_BUY_MIN = 78
BUY_MIN = 63
HOLD_MIN = 45
EARLY_SMART_MIN = 68
EARLY_WATCH_MIN = 52

SENTIMENT_TOP_N_EACH = 25
SENT_ACC_TH = 60
SENT_DIST_TH = 60
SENT_SHAKE_TH = 70

ACC_WIN = 4
ACC_MIN_HITS = 3
DIST_WIN = 4
DIST_MIN_HITS = 2
SHAKE_WIN = 3
SHAKE_MIN_HITS = 1

LOG_LEVEL = logging.INFO

logger = logging.getLogger("priceaction_signals")
logging.basicConfig(
    level=LOG_LEVEL,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

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


def extract_date_from_name(path: Path, regex: re.Pattern):
    m = regex.search(path.name)
    if not m:
        return None
    try:
        return pd.to_datetime(m.group(1)).date()
    except Exception:
        return None


def list_dated_files(folder: Path, regex: re.Pattern):
    if not folder.exists():
        return []
    out = []
    for p in sorted(folder.glob("*.csv")):
        d = extract_date_from_name(p, regex)
        if d is not None:
            out.append((d, p))
    return sorted(out, key=lambda x: x[0])


def choose_last_n_dates(dates_sorted, n):
    if not dates_sorted:
        return []
    return dates_sorted[-min(n, len(dates_sorted)) :]


def rsi(close, n=14):
    d = close.diff()
    gain = d.clip(lower=0)
    loss = (-d.clip(upper=0)).clip(lower=0)
    avg_gain = gain.rolling(n).mean()
    avg_loss = loss.rolling(n).mean()
    rs = avg_gain / (avg_loss + 1e-12)
    return 100 - (100 / (1 + rs))


def true_range(h, l, c):
    pc = c.shift(1)
    return pd.concat([(h - l), (h - pc).abs(), (l - pc).abs()], axis=1).max(axis=1)


def yesno(cond):
    return cond.fillna(False).astype(int)


def clip_score(s):
    return pd.Series(np.clip(s, 0, 100), index=s.index)


# =========================
# LOADERS
# =========================
def load_sector_master(path: Path):
    if not path.exists():
        return pd.DataFrame(columns=["Symbol", "Sector", "Company"])

    df = pd.read_csv(path)
    df.columns = [c.strip() for c in df.columns]

    if "Sector" not in df.columns and "Sectors" in df.columns:
        df = df.rename(columns={"Sectors": "Sector"})
    if "Company" not in df.columns and "Company Name" in df.columns:
        df = df.rename(columns={"Company Name": "Company"})

    for c in ["Symbol", "Sector", "Company"]:
        if c not in df.columns:
            df[c] = ""

    df["Symbol"] = df["Symbol"].astype(str).str.strip().str.upper()
    return df[["Symbol", "Sector", "Company"]].drop_duplicates(subset=["Symbol"], keep="last")


def read_price_file(path: Path, trade_date):
    df = pd.read_csv(path)
    df.columns = [c.strip() for c in df.columns]

    if "Close" not in df.columns and "LTP" in df.columns:
        df["Close"] = df["LTP"]

    if "Volume" not in df.columns:
        for c in ["Vol", "VOL", "Total Traded Quantity", "TotalQty", "Qty"]:
            if c in df.columns:
                df["Volume"] = df[c]
                break

    if "VWAP" not in df.columns:
        for c in ["vwap", "Vwap", "VWAP Price", "Daily VWAP", "AvgPrice", "Average Price"]:
            if c in df.columns:
                df["VWAP"] = df[c]
                break

    needed = ["Symbol", "Open", "High", "Low", "Close", "Volume", "VWAP"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f"SharePrice missing columns {missing} in {path.name}")

    for c in ["Open", "High", "Low", "Close", "Volume", "VWAP"]:
        df[c] = df[c].apply(safe_float)

    df["TradeDate"] = pd.to_datetime(trade_date)
    df["Symbol"] = df["Symbol"].astype(str).str.strip().str.upper()
    return df[["TradeDate", "Symbol", "Open", "High", "Low", "Close", "Volume", "VWAP"]]


def read_floorsheet_file(path: Path, trade_date):
    df = pd.read_csv(path)
    df.columns = [c.strip() for c in df.columns]

    ren = {
        "Transact No": "Transact No.",
        "Transact No.": "Transact No.",
        "Transaction No.": "Transact No.",
        "Qty": "Quantity",
        "QTY": "Quantity",
        "Amt": "Amount",
    }
    df = df.rename(columns={k: v for k, v in ren.items() if k in df.columns})

    needed = ["Transact No.", "Symbol", "Buyer", "Seller", "Quantity", "Rate", "Amount"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f"Floorsheet missing columns {missing} in {path.name}")

    df["TradeDate"] = pd.to_datetime(trade_date)
    for c in ["Quantity", "Rate", "Amount"]:
        df[c] = df[c].apply(safe_float).fillna(0.0)
    for c in ["Buyer", "Seller"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").astype("Int64")

    df["Symbol"] = df["Symbol"].astype(str).str.strip().str.upper()
    return df[needed + ["TradeDate"]]


# =========================
# FLOOR METRICS
# =========================
def symbol_metrics_from_floorsheet(fs: pd.DataFrame, latest_dt):
    """
    Returns BOTH latest-day metrics and rolling-window metrics.
    This avoids mixing a 15-day VWAP with a metric whose name suggests a single day.
    """
    base_cols = [
        "Symbol",
        "Trades_Latest",
        "Trades_Window",
        "Floor_Qty_Latest",
        "Floor_Qty_Window",
        "Floor_Amount_Latest",
        "Floor_Amount_Window",
        "Floor_VWAP_Latest",
        "Floor_VWAP_Window",
        "Floor_Amount_Cr_Latest",
        "Floor_Amount_Cr_Window",
    ]
    if fs.empty:
        return pd.DataFrame(columns=base_cols)

    x = fs.copy()
    x["_qxr"] = x["Quantity"] * x["Rate"]

    latest_only = x[x["TradeDate"].dt.date == latest_dt].copy()

    latest_g = latest_only.groupby("Symbol", as_index=False).agg(
        Trades_Latest=("Transact No.", "nunique"),
        Floor_Qty_Latest=("Quantity", "sum"),
        Floor_Amount_Latest=("Amount", "sum"),
        _qxr_latest=("_qxr", "sum"),
    )
    latest_g["Floor_VWAP_Latest"] = np.where(
        latest_g["Floor_Qty_Latest"] > 0,
        latest_g["_qxr_latest"] / latest_g["Floor_Qty_Latest"],
        np.nan,
    )
    latest_g["Floor_Amount_Cr_Latest"] = latest_g["Floor_Amount_Latest"] / CRORE
    latest_g = latest_g.drop(columns=["_qxr_latest"], errors="ignore")

    window_g = x.groupby("Symbol", as_index=False).agg(
        Trades_Window=("Transact No.", "nunique"),
        Floor_Qty_Window=("Quantity", "sum"),
        Floor_Amount_Window=("Amount", "sum"),
        _qxr_window=("_qxr", "sum"),
    )
    window_g["Floor_VWAP_Window"] = np.where(
        window_g["Floor_Qty_Window"] > 0,
        window_g["_qxr_window"] / window_g["Floor_Qty_Window"],
        np.nan,
    )
    window_g["Floor_Amount_Cr_Window"] = window_g["Floor_Amount_Window"] / CRORE
    window_g = window_g.drop(columns=["_qxr_window"], errors="ignore")

    out = window_g.merge(latest_g, on="Symbol", how="outer")

    # Backward-compatible convenience fields for signal logic:
    # Use latest-day VWAP for execution filter; use window fields for context.
    out["Trades"] = out["Trades_Latest"].fillna(0)
    out["Floor_Qty"] = out["Floor_Qty_Latest"].fillna(0)
    out["Floor_Amount"] = out["Floor_Amount_Latest"].fillna(0)
    out["Floor_VWAP"] = out["Floor_VWAP_Latest"]
    out["Floor_Amount_Cr"] = out["Floor_Amount_Cr_Latest"].fillna(0)

    return out


def broker_symbol_metrics(fs: pd.DataFrame):
    if fs.empty:
        return pd.DataFrame(columns=[
            "TradeDate", "Symbol", "Broker",
            "Buy_Qty", "Sell_Qty", "Net_Qty", "Buy_Amount", "Active_Buy_Day"
        ])

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
        .groupby("Symbol", as_index=False)["pos"].sum()
        .rename(columns={"pos": "top_pos"})
    )

    top_neg = (
        t.sort_values(["Symbol", "neg_abs"], ascending=[True, False])
        .groupby("Symbol")
        .head(topn)
        .groupby("Symbol", as_index=False)["neg_abs"].sum()
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
    empty = pd.DataFrame(columns=[
        "Symbol", "Broker", "Buy_Qty", "Sell_Qty", "Net_Qty",
        "Buy_Amount", "Active_Days", "Side"
    ])
    if bs_window.empty:
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


def build_broker_stats(bs_window: pd.DataFrame):
    if bs_window.empty:
        return pd.DataFrame(columns=[
            "Symbol", "Accum_Brokers", "Dist_Brokers", "Active_Brokers",
            "Broker_Balance", "TopBuyerNet", "TopSellerNetAbs",
            "TopBuyerActiveDays", "TopSellerActiveDays"
        ])

    pos = (
        bs_window.sort_values(["Symbol", "Net_Qty"], ascending=[True, False])
        .groupby("Symbol")
        .head(1)[["Symbol", "Net_Qty", "Active_Days"]]
        .rename(columns={"Net_Qty": "TopBuyerNet", "Active_Days": "TopBuyerActiveDays"})
    )

    neg = bs_window.copy()
    neg["Net_Qty_AbsNeg"] = (-neg["Net_Qty"]).clip(lower=0)
    neg = (
        neg.sort_values(["Symbol", "Net_Qty_AbsNeg"], ascending=[True, False])
        .groupby("Symbol")
        .head(1)[["Symbol", "Net_Qty_AbsNeg", "Active_Days"]]
        .rename(columns={"Net_Qty_AbsNeg": "TopSellerNetAbs", "Active_Days": "TopSellerActiveDays"})
    )

    base = bs_window.groupby("Symbol", as_index=False).agg(
        Accum_Brokers=("Net_Qty", lambda s: int((s > 0).sum())),
        Dist_Brokers=("Net_Qty", lambda s: int((s < 0).sum())),
        Active_Brokers=("Broker", "nunique"),
    )
    base["Broker_Balance"] = base["Accum_Brokers"] - base["Dist_Brokers"]

    return base.merge(pos, on="Symbol", how="left").merge(neg, on="Symbol", how="left")


# =========================
# PRICE FEATURES
# =========================
def ema(s, n):
    return s.ewm(span=n, adjust=False).mean()


def add_price_features(g):
    g = g.sort_values("TradeDate").copy()

    for n in [5, 7, 10, 15, 20, 30]:
        g[f"MA{n}"] = g["Close"].rolling(n).mean()
        g[f"VMA{n}"] = g["Volume"].rolling(n).mean()

    for n in [1, 2, 3, 5, 7, 10, 15, 20, 30]:
        g[f"RET{n}"] = g["Close"].pct_change(n) * 100

    for n in [5, 10, 15, 20]:
        g[f"HH{n}"] = g["High"].rolling(n).max()
        g[f"LL{n}"] = g["Low"].rolling(n).min()

    rng = (g["High"] - g["Low"]).replace(0, np.nan)
    g["UpperWickPct"] = ((g["High"] - g[["Open", "Close"]].max(axis=1)) / (rng + 1e-12)).clip(0, 1)
    g["LowerWickPct"] = ((g[["Open", "Close"]].min(axis=1) - g["Low"]) / (rng + 1e-12)).clip(0, 1)
    g["Close_Pos"] = (g["Close"] - g["Low"]) / (rng + 1e-12)
    g["Body_Pct"] = (g["Close"] - g["Open"]).abs() / (g["Open"] + 1e-12) * 100
    g["RangePct"] = (rng / (g["Close"] + 1e-12)) * 100
    g["RSI14"] = rsi(g["Close"], 14)

    tr = true_range(g["High"], g["Low"], g["Close"])
    g["ATR14"] = tr.rolling(14).mean()
    g["ATR14_Pct"] = (g["ATR14"] / (g["Close"] + 1e-12)) * 100

    g["EMA10"] = ema(g["Close"], 10)
    g["EMA20"] = ema(g["Close"], 20)
    g["EMA50"] = ema(g["Close"], 50)

    g["Vol_Surge"] = np.where(g["VMA10"] > 0, g["Volume"] / (g["VMA10"] + 1e-12), np.nan)
    g["Price_vs_VWAP_Pct"] = np.where(
        g["VWAP"].notna() & (g["VWAP"] > 0),
        (g["Close"] / g["VWAP"] - 1.0) * 100,
        np.nan,
    )

    g["HH10_PRIOR"] = g["HH10"].shift(1)
    g["HH15_PRIOR"] = g["HH15"].shift(1)
    return g


# =========================
# SENTIMENT DETECTOR
# =========================
def add_sentiment_detector(g):
    g = g.copy()

    rng = (g["High"] - g["Low"]).replace(0, np.nan)
    body = (g["Close"] - g["Open"]).abs()
    upper = (g["High"] - g[["Open", "Close"]].max(axis=1))
    lower = (g[["Open", "Close"]].min(axis=1) - g["Low"])

    g["BodyPct01"] = (body / (rng + 1e-12)).clip(0, 1)
    g["LowerWickPct01"] = (lower / (rng + 1e-12)).clip(0, 1)
    g["UpperWickPct01"] = (upper / (rng + 1e-12)).clip(0, 1)
    g["RangePct_D"] = (rng / (g["Close"] + 1e-12)) * 100
    g["RangePct_MA10"] = g["RangePct_D"].rolling(10).mean()
    g["VolAvg5"] = g["Volume"].rolling(5).mean()
    g["VolAvg20"] = g["Volume"].rolling(20).mean()
    g["VolRatio20"] = g["Volume"] / (g["VolAvg20"] + 1e-12)

    if "RET1" not in g.columns:
        g["RET1"] = g["Close"].pct_change(1) * 100

    g["HH20"] = g["High"].rolling(20).max()
    near_hh20 = (g["Close"] >= 0.97 * g["HH20"]).fillna(False)

    acc = np.zeros(len(g), dtype=float)
    acc += 20 * ((g["VolAvg5"] > g["VolAvg20"]).fillna(False)).astype(int)
    acc += 20 * ((g["RangePct_D"] <= (g["RangePct_MA10"] * 0.90)).fillna(False)).astype(int)
    acc += 20 * ((g["BodyPct01"] < 0.40).fillna(False)).astype(int)
    acc += 20 * ((g["Close"] > g["VWAP"]).fillna(False)).astype(int)
    acc += 20 * ((g["LowerWickPct01"] > g["UpperWickPct01"]).fillna(False)).astype(int)
    g["AccumulationScore"] = np.clip(acc, 0, 100).astype(int)

    dist = np.zeros(len(g), dtype=float)
    dist += 25 * ((g["VolRatio20"] >= 1.8).fillna(False)).astype(int)
    dist += 25 * ((g["RET1"] <= -0.50).fillna(False)).astype(int)  # fixed sign
    dist += 20 * ((g["UpperWickPct01"] >= 0.45).fillna(False)).astype(int)
    dist += 20 * ((g["Close"] < g["VWAP"]).fillna(False)).astype(int)
    dist += 10 * near_hh20.astype(int)
    g["DistributionScore"] = np.clip(dist, 0, 100).astype(int)

    prev_low = g["Low"].shift(1)
    shake = np.zeros(len(g), dtype=float)
    shake += 50 * ((g["Low"] < prev_low) & (g["Close"] > prev_low)).fillna(False).astype(int)
    shake += 30 * ((g["LowerWickPct01"] >= 0.50).fillna(False)).astype(int)
    shake += 20 * ((g["VolRatio20"] >= 1.5).fillna(False)).astype(int)
    g["ShakeoutScore"] = np.clip(shake, 0, 100).astype(int)

    acc_hits = (g["AccumulationScore"] >= SENT_ACC_TH).rolling(ACC_WIN).sum()
    dist_hits = (g["DistributionScore"] >= SENT_DIST_TH).rolling(DIST_WIN).sum()
    shake_hits = (g["ShakeoutScore"] >= SENT_SHAKE_TH).rolling(SHAKE_WIN).sum()

    g["AccHits4"] = acc_hits
    g["DistHits4"] = dist_hits
    g["ShakeHits3"] = shake_hits

    g["AccConfirmed"] = (acc_hits >= ACC_MIN_HITS).fillna(False)
    g["DistConfirmed"] = (dist_hits >= DIST_MIN_HITS).fillna(False)
    g["ShakeConfirmed"] = (shake_hits >= SHAKE_MIN_HITS).fillna(False)

    # Explicit precedence: Distribution > Shakeout > Accumulation > Neutral
    sig = np.array(["NEUTRAL"] * len(g), dtype=object)
    sig[g["AccConfirmed"]] = "ACCUMULATION"
    sig[g["ShakeConfirmed"]] = "SHAKEOUT"
    sig[g["DistConfirmed"]] = "DISTRIBUTION"
    g["SentimentSignal"] = sig

    reasons = []
    for i in range(len(g)):
        r = []
        if bool(g["DistConfirmed"].iloc[i]):
            r.append(f"Dist≥{SENT_DIST_TH} on {int(g['DistHits4'].iloc[i])}/{DIST_WIN}")
        if bool(g["ShakeConfirmed"].iloc[i]):
            r.append(f"Shake≥{SENT_SHAKE_TH} on {int(g['ShakeHits3'].iloc[i])}/{SHAKE_WIN}")
        if bool(g["AccConfirmed"].iloc[i]):
            r.append(f"Acc≥{SENT_ACC_TH} on {int(g['AccHits4'].iloc[i])}/{ACC_WIN}")
        reasons.append("; ".join(r) if r else "")
    g["SentimentReason"] = reasons

    return g


# =========================
# SNAPSHOT BUILDERS
# =========================
def build_symbol_snapshot(price_all: pd.DataFrame, latest_dt):
    """
    Only keeps symbols whose latest observed candle is exactly latest_dt.
    This prevents stale symbols from leaking into the report.
    """
    rows = []
    stale_rows = []

    for sym, g in price_all.groupby("Symbol"):
        g = add_price_features(g)
        g = add_sentiment_detector(g)
        if len(g) < 20:
            continue

        last = g.iloc[-1]
        last_date = pd.to_datetime(last["TradeDate"]).date()

        if last_date != latest_dt:
            stale_rows.append((sym, last_date))
            continue

        rows.append({
            "TradeDate": last["TradeDate"],
            "Symbol": sym,
            "DataFreshness": "FRESH",
            "LatestPriceDate": last["TradeDate"],
            "Open": last["Open"],
            "High": last["High"],
            "Low": last["Low"],
            "Close": last["Close"],
            "Volume": last["Volume"],
            "VWAP": last["VWAP"],
            "MA5": last["MA5"],
            "MA7": last["MA7"],
            "MA10": last["MA10"],
            "MA15": last["MA15"],
            "MA20": last["MA20"],
            "HH5": last["HH5"],
            "HH10": last["HH10"],
            "HH15": last["HH15"],
            "LL5": last["LL5"],
            "LL10": last["LL10"],
            "LL15": last["LL15"],
            "HH10_PRIOR": last["HH10_PRIOR"],
            "HH15_PRIOR": last["HH15_PRIOR"],
            "RET1_%": last["RET1"],
            "RET2_%": last["RET2"],
            "RET3_%": last["RET3"],
            "RET5_%": last["RET5"],
            "RET7_%": last["RET7"],
            "RET10_%": last["RET10"],
            "RET15_%": last["RET15"],
            "RET20_%": last["RET20"],
            "RSI14": last["RSI14"],
            "ATR14": last["ATR14"],
            "ATR14_Pct": last["ATR14_Pct"],
            "RangePct": last["RangePct"],
            "Body_Pct": last["Body_Pct"],
            "Close_Pos": last["Close_Pos"],
            "UpperWickPct": last["UpperWickPct"],
            "LowerWickPct": last["LowerWickPct"],
            "Vol_Surge": last["Vol_Surge"],
            "Price_vs_VWAP_Pct": last["Price_vs_VWAP_Pct"],
            "AccumulationScore": last["AccumulationScore"],
            "DistributionScore": last["DistributionScore"],
            "ShakeoutScore": last["ShakeoutScore"],
            "AccHits4": last["AccHits4"],
            "DistHits4": last["DistHits4"],
            "ShakeHits3": last["ShakeHits3"],
            "AccConfirmed": last["AccConfirmed"],
            "DistConfirmed": last["DistConfirmed"],
            "ShakeConfirmed": last["ShakeConfirmed"],
            "SentimentSignal": last["SentimentSignal"],
            "SentimentReason": last["SentimentReason"],
        })

    if stale_rows:
        logger.info("Excluded %s stale symbols not updated on %s", len(stale_rows), latest_dt)

    return pd.DataFrame(rows)


# =========================
# EARLY SMART MONEY
# =========================
def build_early_smart_money(snapshot, floor_symbol, pressure, broker_stats, sector):
    x = snapshot.copy()
    x = x.merge(floor_symbol, on="Symbol", how="left")
    x = x.merge(pressure, on="Symbol", how="left")
    x = x.merge(broker_stats, on="Symbol", how="left")
    x = x.merge(sector, on="Symbol", how="left")

    for c in [
        "Trades", "Trades_Latest", "Trades_Window",
        "Floor_Qty", "Floor_Qty_Latest", "Floor_Qty_Window",
        "Floor_Amount", "Floor_Amount_Latest", "Floor_Amount_Window",
        "Floor_VWAP", "Floor_VWAP_Latest", "Floor_VWAP_Window",
        "Floor_Amount_Cr", "Floor_Amount_Cr_Latest", "Floor_Amount_Cr_Window",
        "Buy_Pressure", "Sell_Pressure", "Net_Broker_Bias",
        "Accum_Brokers", "Dist_Brokers", "Active_Brokers", "Broker_Balance",
        "TopBuyerNet", "TopSellerNetAbs", "TopBuyerActiveDays", "TopSellerActiveDays"
    ]:
        if c not in x.columns:
            x[c] = np.nan

    x["Above_Floor_VWAP"] = np.where(
        x["Floor_VWAP_Latest"].notna() & x["Close"].notna(),
        x["Close"] > x["Floor_VWAP_Latest"],
        False,
    )

    raw = (
        yesno(x["Buy_Pressure"] > x["Sell_Pressure"]) * 16
        + yesno(x["Buy_Pressure"] >= 0.55) * 14
        + yesno(x["Broker_Balance"] > 0) * 10
        + yesno(x["Accum_Brokers"] >= x["Dist_Brokers"]) * 8
        + yesno(x["TopBuyerActiveDays"].fillna(0) >= 3) * 10
        + yesno(x["Floor_Amount_Cr_Window"].fillna(0) >= 1.0) * 8
        + yesno(x["Trades_Window"].fillna(0) >= 80) * 6
        + yesno(x["AccumulationScore"].fillna(0) >= 60) * 8
        + yesno(x["AccConfirmed"].fillna(False)) * 6
        + yesno(x["Close"] >= x["MA5"]) * 6
        + yesno(x["MA5"] >= x["MA10"]) * 6
        + yesno(x["RSI14"].between(48, 68, inclusive="both")) * 6
        + yesno(x["Vol_Surge"].fillna(0) >= 1.10) * 6
        + yesno(x["Close_Pos"].fillna(0) >= 0.55) * 4
        + yesno(x["Close"] >= 0.96 * x["HH10"]) * 6
        - yesno(x["Sell_Pressure"].fillna(0) >= 0.55) * 12
        - yesno(x["DistributionScore"].fillna(0) >= 70) * 12
        - yesno(x["DistConfirmed"].fillna(False)) * 12
        - yesno(x["UpperWickPct"].fillna(0) >= 0.45) * 6
        - yesno(x["RSI14"].fillna(100) < 42) * 6
    )

    x["EarlySmartScore"] = clip_score(raw)

    x["EarlySmartSignal"] = "NO"
    x.loc[x["EarlySmartScore"] >= EARLY_WATCH_MIN, "EarlySmartSignal"] = "WATCH"
    x.loc[x["EarlySmartScore"] >= EARLY_SMART_MIN, "EarlySmartSignal"] = "EARLY ACCUMULATION"

    reasons = []
    for _, r in x.iterrows():
        tags = []
        if r.get("Buy_Pressure", 0) > r.get("Sell_Pressure", 0):
            tags.append("buy pressure rising")
        if r.get("Broker_Balance", 0) > 0:
            tags.append("positive broker balance")
        if r.get("TopBuyerActiveDays", 0) >= 3:
            tags.append("persistent top buyer")
        if r.get("AccumulationScore", 0) >= 60:
            tags.append("accumulation score strong")
        if pd.notna(r.get("Close")) and pd.notna(r.get("HH10")) and r["Close"] >= 0.96 * r["HH10"]:
            tags.append("near breakout zone")
        if pd.notna(r.get("Close")) and pd.notna(r.get("Floor_VWAP_Latest")):
            tags.append("above latest floor VWAP" if r["Close"] > r["Floor_VWAP_Latest"] else "below latest floor VWAP")
        if r.get("DistConfirmed", False):
            tags.append("distribution risk")
        reasons.append(", ".join(tags[:7]))
    x["EarlyReason"] = reasons

    cols = [
        "TradeDate", "Symbol", "Sector", "Company", "DataFreshness",
        "EarlySmartSignal", "EarlySmartScore",
        "Close", "VWAP", "Floor_VWAP_Latest", "Floor_VWAP_Window", "Above_Floor_VWAP",
        "RET3_%", "RET5_%", "RET10_%",
        "MA5", "MA10", "HH10", "HH15",
        "Volume", "Vol_Surge",
        "Trades_Latest", "Trades_Window",
        "Floor_Qty_Latest", "Floor_Qty_Window",
        "Floor_Amount_Cr_Latest", "Floor_Amount_Cr_Window",
        "Buy_Pressure", "Sell_Pressure", "Net_Broker_Bias",
        "Accum_Brokers", "Dist_Brokers", "Active_Brokers", "Broker_Balance",
        "TopBuyerNet", "TopSellerNetAbs", "TopBuyerActiveDays", "TopSellerActiveDays",
        "RSI14", "Close_Pos", "UpperWickPct",
        "AccumulationScore", "DistributionScore", "ShakeoutScore",
        "AccHits4", "DistHits4", "ShakeHits3",
        "AccConfirmed", "DistConfirmed", "ShakeConfirmed",
        "SentimentSignal", "SentimentReason", "EarlyReason"
    ]
    cols = [c for c in cols if c in x.columns]
    x = x[cols].sort_values(["EarlySmartScore", "Buy_Pressure", "AccumulationScore"], ascending=[False, False, False])
    return x


# =========================
# WINDOW SCORING
# =========================
def score_window(x: pd.DataFrame, window_name: str):
    x = x.copy()

    if window_name == "5D":
        trend_score = (
            yesno(x["Close"] > x["MA5"]) * 12
            + yesno(x["MA5"] > x["MA10"]) * 16
            + yesno(x["Close"] >= 0.98 * x["HH5"]) * 12
            + yesno(x["RSI14"].between(50, 72, inclusive="both")) * 8
        )
    elif window_name == "10D":
        trend_score = (
            yesno(x["Close"] > x["MA5"]) * 8
            + yesno(x["MA5"] > x["MA10"]) * 14
            + yesno(x["MA10"] > x["MA15"]) * 12
            + yesno(x["Close"] >= 0.98 * x["HH10"]) * 10
            + yesno(x["RSI14"].between(52, 72, inclusive="both")) * 8
        )
    else:
        trend_score = (
            yesno(x["Close"] > x["MA10"]) * 8
            + yesno(x["MA10"] > x["MA15"]) * 14
            + yesno(x["MA15"] > x["MA20"]) * 12
            + yesno(x["Close"] >= 0.97 * x["HH15"]) * 10
            + yesno(x["RSI14"].between(52, 72, inclusive="both")) * 8
        )

    candle_score = (
        yesno(x["Close_Pos"] >= 0.65) * 8
        + yesno(x["UpperWickPct"] <= 0.30) * 6
        + yesno(x["LowerWickPct"] >= x["UpperWickPct"]) * 4
    )

    volume_score = (
        yesno(x["Vol_Surge"] >= 1.20) * 8
        + yesno(x["Vol_Surge"] >= 1.50) * 7
        + yesno(x["Trades_Window"] >= 80) * 5
        + yesno(x["Floor_Amount_Cr_Window"] >= 1.0) * 5
    )

    floor_score = (
        yesno(x["Buy_Pressure"] > x["Sell_Pressure"]) * 10
        + yesno(x["Buy_Pressure"] >= 0.55) * 10
        + yesno(x["Broker_Balance"] > 0) * 6
        + yesno(x["Accum_Brokers"] >= x["Dist_Brokers"]) * 6
        + yesno(x["Close"] >= x["VWAP"]) * 6
        + yesno(x["Close"] > x["Floor_VWAP_Latest"]) * 12
    )

    sentiment_bonus = (
        yesno(x["SentimentSignal"] == "ACCUMULATION") * 8
        + yesno(x["SentimentSignal"] == "SHAKEOUT") * 5
        - yesno(x["SentimentSignal"] == "DISTRIBUTION") * 10
    )

    risk_penalty = (
        yesno(x["Sell_Pressure"] >= 0.55) * 14
        + yesno(x["Close"] < x["VWAP"]) * 10
        + yesno(x["Close"] < x["Floor_VWAP_Latest"]) * 12
        + yesno(x["RSI14"] < 45) * 8
        + yesno(x["UpperWickPct"] >= 0.45) * 6
        + yesno(x["Close_Pos"] <= 0.35) * 5
        + yesno(x["DistributionScore"] >= 70) * 8
    )

    raw = trend_score + candle_score + volume_score + floor_score + sentiment_bonus - risk_penalty
    x[f"Score_{window_name}"] = clip_score(raw)

    signal = np.array(["SELL"] * len(x), dtype=object)
    signal[x[f"Score_{window_name}"] >= HOLD_MIN] = "HOLD"
    signal[x[f"Score_{window_name}"] >= BUY_MIN] = "BUY"
    signal[x[f"Score_{window_name}"] >= BEST_BUY_MIN] = "BEST BUY"

    x["Above_Floor_VWAP"] = np.where(
        x["Floor_VWAP_Latest"].notna() & x["Close"].notna(),
        x["Close"] > x["Floor_VWAP_Latest"],
        False,
    )
    signal[(signal == "BUY") & (~x["Above_Floor_VWAP"])] = "HOLD"
    signal[(signal == "BEST BUY") & (~x["Above_Floor_VWAP"])] = "HOLD"

    x[f"Signal_{window_name}"] = signal
    return x


def build_reason_row(r, window_name):
    parts = []

    if r.get("Buy_Pressure", np.nan) > r.get("Sell_Pressure", np.nan):
        parts.append("buyer pressure")
    elif r.get("Sell_Pressure", np.nan) > r.get("Buy_Pressure", np.nan):
        parts.append("seller pressure")

    if pd.notna(r.get("Close")) and pd.notna(r.get("VWAP")):
        parts.append("above VWAP" if r["Close"] >= r["VWAP"] else "below VWAP")

    if pd.notna(r.get("Floor_VWAP_Latest")) and pd.notna(r.get("Close")) and r["Close"] > r["Floor_VWAP_Latest"]:
        parts.append("above latest floor VWAP")

    if pd.notna(r.get("Vol_Surge")) and r["Vol_Surge"] >= 1.5:
        parts.append("volume surge")

    if pd.notna(r.get("RSI14")):
        if 52 <= r["RSI14"] <= 72:
            parts.append("healthy RSI")
        elif r["RSI14"] < 45:
            parts.append("weak RSI")

    if window_name == "5D" and pd.notna(r.get("HH5")) and r["Close"] >= 0.98 * r["HH5"]:
        parts.append("near 5D breakout")
    if window_name == "10D" and pd.notna(r.get("HH10")) and r["Close"] >= 0.98 * r["HH10"]:
        parts.append("near 10D breakout")
    if window_name == "15D" and pd.notna(r.get("HH15")) and r["Close"] >= 0.97 * r["HH15"]:
        parts.append("near 15D breakout")

    if r.get("Accum_Brokers", 0) > r.get("Dist_Brokers", 0):
        parts.append("broker accumulation")
    elif r.get("Dist_Brokers", 0) > r.get("Accum_Brokers", 0):
        parts.append("broker distribution")

    if r.get("SentimentSignal") == "ACCUMULATION":
        parts.append("sentiment accumulation")
    elif r.get("SentimentSignal") == "DISTRIBUTION":
        parts.append("sentiment distribution")

    if pd.notna(r.get("UpperWickPct")) and r["UpperWickPct"] >= 0.45:
        parts.append("upper-wick risk")

    return ", ".join(parts[:7])


def build_action_plan(signal):
    if signal == "BEST BUY":
        return "Fresh opportunity; entry near current price or slight dip; trail below stop."
    if signal == "BUY":
        return "Buy on confirmation; avoid chasing extended candle."
    if signal == "HOLD":
        return "Hold if above VWAP/short MA; wait for stronger follow-through."
    return "Reduce or avoid; price-floor alignment is weak."


def build_signal_sheet(snapshot, floor_symbol, pressure, broker_stats, sector, window_name):
    x = snapshot.copy()
    x = x.merge(floor_symbol, on="Symbol", how="left")
    x = x.merge(pressure, on="Symbol", how="left")
    x = x.merge(broker_stats, on="Symbol", how="left")
    x = x.merge(sector, on="Symbol", how="left")

    for c in [
        "Trades", "Trades_Latest", "Trades_Window",
        "Floor_Qty", "Floor_Qty_Latest", "Floor_Qty_Window",
        "Floor_Amount", "Floor_Amount_Latest", "Floor_Amount_Window",
        "Floor_VWAP", "Floor_VWAP_Latest", "Floor_VWAP_Window",
        "Floor_Amount_Cr", "Floor_Amount_Cr_Latest", "Floor_Amount_Cr_Window",
        "Buy_Pressure", "Sell_Pressure", "Net_Broker_Bias",
        "Accum_Brokers", "Dist_Brokers", "Active_Brokers", "Broker_Balance",
        "TopBuyerNet", "TopSellerNetAbs", "TopBuyerActiveDays", "TopSellerActiveDays"
    ]:
        if c not in x.columns:
            x[c] = np.nan

    x["Above_Floor_VWAP"] = np.where(
        x["Floor_VWAP_Latest"].notna() & x["Close"].notna(),
        x["Close"] > x["Floor_VWAP_Latest"],
        False,
    )

    x = score_window(x, window_name)
    score_col = f"Score_{window_name}"
    signal_col = f"Signal_{window_name}"

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

    x["RiskFlags"] = ""
    x.loc[x["Sell_Pressure"].fillna(0) >= 0.55, "RiskFlags"] += "SellWall; "
    x.loc[x["Close_Pos"].fillna(0.5) <= 0.35, "RiskFlags"] += "WeakClose; "
    x.loc[x["UpperWickPct"].fillna(0) >= 0.45, "RiskFlags"] += "SupplyWick; "
    x.loc[x["Close"].fillna(0) < x["VWAP"].fillna(np.inf), "RiskFlags"] += "BelowVWAP; "
    x.loc[x["Close"].fillna(0) <= x["Floor_VWAP_Latest"].fillna(np.inf), "RiskFlags"] += "BelowLatestFloorVWAP; "
    x.loc[x["DistributionScore"].fillna(0) >= 70, "RiskFlags"] += "Distribution; "

    x["InsightReason"] = x.apply(lambda r: build_reason_row(r, window_name), axis=1)
    x["ActionPlan"] = x[signal_col].apply(build_action_plan)

    out_cols = [
        "TradeDate", "Symbol", "Sector", "Company", "DataFreshness",
        signal_col, score_col, "Above_Floor_VWAP",
        "Entry", "StopLoss", "Target1", "Target2", "Target3",
        "Close", "VWAP", "Floor_VWAP_Latest", "Floor_VWAP_Window",
        "RET1_%", "RET3_%", "RET5_%", "RET7_%", "RET10_%", "RET15_%",
        "MA5", "MA7", "MA10", "MA15", "MA20",
        "HH5", "HH10", "HH15", "LL5", "LL10", "LL15",
        "Volume", "Vol_Surge",
        "Trades_Latest", "Trades_Window",
        "Floor_Qty_Latest", "Floor_Qty_Window",
        "Floor_Amount_Cr_Latest", "Floor_Amount_Cr_Window",
        "Buy_Pressure", "Sell_Pressure", "Net_Broker_Bias",
        "Accum_Brokers", "Dist_Brokers", "Active_Brokers", "Broker_Balance",
        "RSI14", "ATR14", "ATR14_Pct",
        "RangePct", "Body_Pct", "Close_Pos", "UpperWickPct", "LowerWickPct",
        "Price_vs_VWAP_Pct",
        "AccumulationScore", "DistributionScore", "ShakeoutScore",
        "AccHits4", "DistHits4", "ShakeHits3",
        "AccConfirmed", "DistConfirmed", "ShakeConfirmed",
        "SentimentSignal", "SentimentReason",
        "InsightReason", "RiskFlags", "ActionPlan"
    ]
    out_cols = [c for c in out_cols if c in x.columns]
    x = x[out_cols].sort_values([score_col, "Buy_Pressure", "RET10_%"], ascending=[False, False, False])
    return x


def build_sentiment_top(sheet_15d):
    if sheet_15d.empty:
        return pd.DataFrame(columns=[
            "Category", "TradeDate", "Symbol", "Sector", "Company", "SentimentSignal",
            "AccumulationScore", "DistributionScore", "ShakeoutScore",
            "AccHits4", "DistHits4", "ShakeHits3",
            "Close", "Volume", "Score_15D", "Signal_15D",
            "InsightReason", "SentimentReason"
        ])

    base_cols = [
        "TradeDate", "Symbol", "Sector", "Company", "SentimentSignal",
        "AccumulationScore", "DistributionScore", "ShakeoutScore",
        "AccHits4", "DistHits4", "ShakeHits3",
        "Close", "Volume", "Score_15D", "Signal_15D",
        "InsightReason", "SentimentReason"
    ]

    x = sheet_15d.copy()
    for c in ["AccConfirmed", "DistConfirmed", "ShakeConfirmed"]:
        if c not in x.columns:
            x[c] = False

    acc = x[x["AccConfirmed"].fillna(False)].copy()
    acc = acc.sort_values(["AccHits4", "Score_15D"], ascending=[False, False]).head(SENTIMENT_TOP_N_EACH)
    acc.insert(0, "Category", "ACCUMULATION")

    dist = x[x["DistConfirmed"].fillna(False)].copy()
    dist = dist.sort_values(["DistHits4", "Score_15D"], ascending=[False, False]).head(SENTIMENT_TOP_N_EACH)
    dist.insert(0, "Category", "DISTRIBUTION")

    shake = x[x["ShakeConfirmed"].fillna(False)].copy()
    shake = shake.sort_values(["ShakeHits3", "Score_15D"], ascending=[False, False]).head(SENTIMENT_TOP_N_EACH)
    shake.insert(0, "Category", "SHAKEOUT")

    out = pd.concat([acc, dist, shake], ignore_index=True)
    cols = ["Category"] + [c for c in base_cols if c in out.columns]
    return out[cols]


# =========================
# EXCEL HELPERS
# =========================
def style_sheet(ws):
    ws.freeze_panes = "A2"
    thin = Side(style="thin", color="D0D7DE")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="DCE6F1")

    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def auto_fit_columns(ws, min_w=10, max_w=45):
    for col_cells in ws.columns:
        letter = get_column_letter(col_cells[0].column)
        max_len = 0
        for cell in col_cells:
            if cell.value is None:
                continue
            text = str(cell.value)
            lines = text.splitlines() if "\n" in text else [text]
            longest = max(len(line) for line in lines) if lines else 0
            max_len = max(max_len, longest)
        ws.column_dimensions[letter].width = max(min_w, min(max_w, max_len + 2))


def auto_fit_rows(ws, base_height=18, max_height=90):
    col_widths = {}
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        width = ws.column_dimensions[letter].width
        col_widths[col_idx] = width if width else 10

    for row_idx in range(1, ws.max_row + 1):
        max_lines = 1
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value is None:
                continue
            text = str(cell.value)
            explicit_lines = text.splitlines() if "\n" in text else [text]
            width = max(col_widths.get(col_idx, 10), 1)
            wrapped_lines = 0
            for part in explicit_lines:
                wrapped_lines += max(1, ceil(len(part) / max(width - 1, 1)))
            max_lines = max(max_lines, wrapped_lines)
        ws.row_dimensions[row_idx].height = min(max_height, max(base_height, max_lines * base_height))


def setup_print(ws):
    ws.print_title_rows = "1:1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5
    ws.page_margins.header = 0.2
    ws.page_margins.footer = 0.2
    ws.oddFooter.center.text = "Page &[Page] of &[Pages]"
    ws.evenFooter.center.text = "Page &[Page] of &[Pages]"


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
    auto_fit_rows(ws)
    setup_print(ws)


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


def number_format(ws, mapping):
    header = [c.value for c in ws[1]]
    for col_name, fmt in mapping.items():
        if col_name not in header:
            continue
        idx = header.index(col_name) + 1
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=idx).number_format = fmt


# =========================
# MAIN
# =========================
def main():
    logger.info("SCRIPT = %s", Path(__file__).resolve())
    logger.info("ROOT = %s", ROOT)
    logger.info("FLOOR_DIR = %s", FLOOR_DIR)
    logger.info("PRICE_DIR = %s", PRICE_DIR)

    floor_pairs = list_dated_files(FLOOR_DIR, FLOOR_RE)
    price_pairs = list_dated_files(PRICE_DIR, PRICE_RE)

    if not floor_pairs:
        raise RuntimeError(f"No floorsheet csv files found in {FLOOR_DIR}")
    if not price_pairs:
        raise RuntimeError(f"No share price csv files found in {PRICE_DIR}")

    floor_dates = [d for d, _ in floor_pairs]
    price_dates = [d for d, _ in price_pairs]
    common_dates = sorted(set(floor_dates).intersection(set(price_dates)))

    if not common_dates:
        raise RuntimeError("No common dates found between floor sheet and share price files.")

    signal_dates = choose_last_n_dates(common_dates, LOOKBACK_SIGNAL)
    hist_dates = choose_last_n_dates(common_dates, max(PRICE_HISTORY_LOAD, LOOKBACK_SIGNAL + 10))
    latest_dt = pd.to_datetime(signal_dates[-1]).date()

    logger.info("Signal window: %s -> %s", signal_dates[0], signal_dates[-1])
    logger.info("History window: %s -> %s", hist_dates[0], hist_dates[-1])
    logger.info("Latest report date: %s", latest_dt)

    floor_map = {d: p for d, p in floor_pairs}
    price_map = {d: p for d, p in price_pairs}

    fs_list = [read_floorsheet_file(floor_map[d], d) for d in signal_dates if d in floor_map]
    pr_list = [read_price_file(price_map[d], d) for d in hist_dates if d in price_map]

    fs = pd.concat(fs_list, ignore_index=True) if fs_list else pd.DataFrame()
    pr = pd.concat(pr_list, ignore_index=True) if pr_list else pd.DataFrame()

    if fs.empty:
        raise RuntimeError("Floor sheet window data is empty after loading.")
    if pr.empty:
        raise RuntimeError("Price data is empty after loading.")

    sector = load_sector_master(SECTOR_FILE)
    floor_symbol = symbol_metrics_from_floorsheet(fs, latest_dt=latest_dt)

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
        bs_window = pd.DataFrame(columns=[
            "Symbol", "Broker", "Buy_Qty", "Sell_Qty",
            "Net_Qty", "Buy_Amount", "Active_Days"
        ])

    pressure = compute_pressure(bs_window, topn=5)
    broker_stats = build_broker_stats(bs_window)

    snapshot = build_symbol_snapshot(pr, latest_dt=latest_dt)
    if snapshot.empty:
        raise RuntimeError(f"No fresh symbol snapshots found for latest date {latest_dt}")

    early_smart = build_early_smart_money(snapshot, floor_symbol, pressure, broker_stats, sector)
    signals_5d = build_signal_sheet(snapshot, floor_symbol, pressure, broker_stats, sector, "5D")
    signals_10d = build_signal_sheet(snapshot, floor_symbol, pressure, broker_stats, sector, "10D")
    signals_15d = build_signal_sheet(snapshot, floor_symbol, pressure, broker_stats, sector, "15D")

    best_buy = signals_15d[
        (signals_15d["Signal_15D"] == "BEST BUY") &
        (signals_15d["Above_Floor_VWAP"] == True)
    ].copy()

    buy_setups = signals_15d[
        (signals_15d["Signal_15D"] == "BUY") &
        (signals_15d["Above_Floor_VWAP"] == True)
    ].copy()

    hold_list = signals_15d[signals_15d["Signal_15D"] == "HOLD"].copy()
    sell_list = signals_15d[signals_15d["Signal_15D"] == "SELL"].copy()

    acc_brokers, dist_brokers = top_broker_accumulation_distribution(bs_window, topn=TOP_BROKER_N)
    acc_brokers = acc_brokers.merge(sector, on="Symbol", how="left")
    dist_brokers = dist_brokers.merge(sector, on="Symbol", how="left")

    sentiment_top = build_sentiment_top(signals_15d)

    readme = pd.DataFrame([
        ["Purpose", "OHLCV + Floor Sheet combined short-term signal report."],
        ["Signal windows", "5D = early, 10D = confirmation, 15D = primary decision."],
        ["Early smart money", "EARLY ACCUMULATION/WATCH can appear before full BUY confirmation."],
        ["Buy filter", "BUY / BEST BUY only if Close > latest-day Floor_VWAP."],
        ["Floor metrics", "Latest-day Floor_VWAP is separated from rolling 15-day Floor_VWAP_Window."],
        ["Snapshot freshness", "Only symbols updated on latest common trading day are included."],
        ["Scoring", "Stable weighted scores (0-100 clipped), not daily min-max normalization."],
        ["Trades count", "Trades use unique Transact No. counts (nunique), not raw row counts."],
        ["Data needed", "At least 15 common trading days; extra price history loaded for RSI/ATR/MA calculations."],
        ["Floor source", str(FLOOR_DIR)],
        ["Price source", str(PRICE_DIR)],
        ["Sector source", str(SECTOR_FILE)],
        ["Trade plan", "Entry=current close, StopLoss=min(LL5, Close-1.2*ATR14), targets at 1.5R / 2.5R / 3R."],
        ["Auto-fit", "All sheets auto-fit column widths and estimated row heights."],
        ["Print setup", "Landscape, fit-to-width, repeated header row, page number footer."],
        ["Path debug", f"Script root resolved to: {ROOT}"],
    ], columns=["Item", "Explanation"])

    wb = Workbook()
    wb.remove(wb.active)

    sheets = {
        "README": readme,
        "Early_SmartMoney": early_smart,
        "Signals_5D": signals_5d,
        "Signals_10D": signals_10d,
        "Signals_15D": signals_15d,
        "Best_Buy": best_buy,
        "Buy_Setups": buy_setups,
        "Hold_List": hold_list,
        "Sell_List": sell_list,
        "Broker_Accumulation": acc_brokers,
        "Broker_Distribution": dist_brokers,
        "Sentiment_Top": sentiment_top,
    }

    fmt_map = {
        "EarlySmartScore": "0.00",
        "Score_5D": "0.00",
        "Score_10D": "0.00",
        "Score_15D": "0.00",
        "Entry": "#,##0.00",
        "StopLoss": "#,##0.00",
        "Target1": "#,##0.00",
        "Target2": "#,##0.00",
        "Target3": "#,##0.00",
        "Close": "#,##0.00",
        "VWAP": "#,##0.00",
        "Floor_VWAP": "#,##0.00",
        "Floor_VWAP_Latest": "#,##0.00",
        "Floor_VWAP_Window": "#,##0.00",
        "RET1_%": "0.00",
        "RET2_%": "0.00",
        "RET3_%": "0.00",
        "RET5_%": "0.00",
        "RET7_%": "0.00",
        "RET10_%": "0.00",
        "RET15_%": "0.00",
        "RET20_%": "0.00",
        "MA5": "#,##0.00",
        "MA7": "#,##0.00",
        "MA10": "#,##0.00",
        "MA15": "#,##0.00",
        "MA20": "#,##0.00",
        "HH5": "#,##0.00",
        "HH10": "#,##0.00",
        "HH15": "#,##0.00",
        "LL5": "#,##0.00",
        "LL10": "#,##0.00",
        "LL15": "#,##0.00",
        "Volume": "#,##0",
        "Floor_Qty": "#,##0",
        "Floor_Qty_Latest": "#,##0",
        "Floor_Qty_Window": "#,##0",
        "Trades": "0",
        "Trades_Latest": "0",
        "Trades_Window": "0",
        "Floor_Amount_Cr": "0.000",
        "Floor_Amount_Cr_Latest": "0.000",
        "Floor_Amount_Cr_Window": "0.000",
        "Buy_Pressure": "0.00",
        "Sell_Pressure": "0.00",
        "Net_Broker_Bias": "#,##0",
        "Accum_Brokers": "0",
        "Dist_Brokers": "0",
        "Active_Brokers": "0",
        "Broker_Balance": "0",
        "TopBuyerNet": "#,##0",
        "TopSellerNetAbs": "#,##0",
        "TopBuyerActiveDays": "0",
        "TopSellerActiveDays": "0",
        "RSI14": "0.00",
        "ATR14": "0.00",
        "ATR14_Pct": "0.00",
        "RangePct": "0.00",
        "Body_Pct": "0.00",
        "Close_Pos": "0.00",
        "UpperWickPct": "0.00",
        "LowerWickPct": "0.00",
        "Vol_Surge": "0.00",
        "Price_vs_VWAP_Pct": "0.00",
        "AccumulationScore": "0",
        "DistributionScore": "0",
        "ShakeoutScore": "0",
        "AccHits4": "0",
        "DistHits4": "0",
        "ShakeHits3": "0",
        "Buy_Qty": "#,##0",
        "Sell_Qty": "#,##0",
        "Net_Qty": "#,##0",
        "Buy_Amount": "#,##0.00",
        "Active_Days": "0",
    }

    for i, (sheet_name, df) in enumerate(sheets.items(), start=1):
        ws = wb.create_sheet(sheet_name[:31])
        add_table_sheet(ws, df, f"T{i}")
        number_format(ws, fmt_map)

        for c in [
            "EarlySmartScore",
            "Score_5D", "Score_10D", "Score_15D",
            "Buy_Pressure", "Sell_Pressure", "Vol_Surge", "RET10_%",
            "AccumulationScore", "DistributionScore",
            "AccHits4", "DistHits4", "ShakeHits3"
        ]:
            color_scale(ws, c)

    out_path = OUT_DIR / f"nepse_signals_{latest_dt}.xlsx"
    wb.save(out_path)
    logger.info("Excel created: %s", out_path)


if __name__ == "__main__":
    main()
