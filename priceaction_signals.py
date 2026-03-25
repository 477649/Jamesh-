# ==========================================
# FIXED 3-STATE MODEL (BUY / HOLD / SELL)
# WITH:
# - 15 trading-day floor analysis
# - fixed score calibration
# - backtest
# - excel output with table + autofit + wrap text
# ==========================================

import re
from math import ceil
from pathlib import Path

import numpy as np
import pandas as pd

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


# =========================================================
# SETTINGS / PATHS
# =========================================================
BUY_TH = 60
SELL_TH = 55

LOOKBACK_SIGNAL = 15
PRICE_HISTORY_LOAD = 60
TOP_BROKER_N = 8
CRORE = 10_000_000

ROOT = Path(__file__).resolve().parent

PRICE_DIR = ROOT / "outputs" / "sharesansar"
FLOOR_DIR = ROOT / "outputs" / "Floor Sheet"
SECTOR_FILE = ROOT / "outputs" / "Sector" / "sector_master.csv"
OUT_DIR = ROOT / "outputs" / "PriceAction"
OUT_DIR.mkdir(parents=True, exist_ok=True)

PRICE_RE = re.compile(r"SharePrice_(\d{4}-\d{2}-\d{2})", re.IGNORECASE)
FLOOR_RE = re.compile(r".*?(\d{4}-\d{2}-\d{2}).*\.csv$", re.IGNORECASE)


# =========================================================
# HELPERS
# =========================================================
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


def ema(s, n):
    return s.ewm(span=n, adjust=False).mean()


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


def clamp_score(s, lo=0, hi=100):
    return pd.Series(np.clip(pd.to_numeric(s, errors="coerce"), lo, hi), index=s.index)


# =========================================================
# LOADERS
# =========================================================
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

    return df[["Transact No.", "Symbol", "Buyer", "Seller", "Quantity", "Rate", "Amount", "TradeDate"]]


# =========================================================
# FLOOR METRICS
# =========================================================
def symbol_metrics_from_floorsheet_daily(fs: pd.DataFrame):
    if fs.empty:
        return pd.DataFrame(columns=[
            "TradeDate", "Symbol", "Trades", "Floor_Qty",
            "Floor_Amount", "Floor_VWAP", "Floor_Amount_Cr"
        ])

    x = fs.copy()
    x["_qxr"] = x["Quantity"] * x["Rate"]

    g = x.groupby(["TradeDate", "Symbol"], as_index=False).agg(
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
        return pd.DataFrame(columns=[
            "TradeDate", "Symbol", "Broker", "Buy_Qty", "Sell_Qty",
            "Net_Qty", "Buy_Amount", "Active_Buy_Day"
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

    out["Buy_Pressure"] = np.where(
        out["pos_total"] > 0,
        out["top_pos"].fillna(0) / out["pos_total"],
        np.nan
    )
    out["Sell_Pressure"] = np.where(
        out["neg_total"] > 0,
        out["top_neg"].fillna(0) / out["neg_total"],
        np.nan
    )
    out["Net_Broker_Bias"] = out["pos_total"].fillna(0) - out["neg_total"].fillna(0)

    return out[["Symbol", "Buy_Pressure", "Sell_Pressure", "Net_Broker_Bias"]]


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


# =========================================================
# PRICE FEATURES
# =========================================================
def add_price_features(g: pd.DataFrame):
    g = g.sort_values("TradeDate").copy()

    for n in [5, 10, 20]:
        g[f"MA{n}"] = g["Close"].rolling(n).mean()

    for n in [1, 3, 5, 10]:
        g[f"RET{n}"] = g["Close"].pct_change(n) * 100

    for n in [5, 10, 15]:
        g[f"HH{n}"] = g["High"].rolling(n).max()

    for n in [5, 10]:
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

    g["VMA10"] = g["Volume"].rolling(10).mean()
    g["Vol_Surge"] = np.where(g["VMA10"] > 0, g["Volume"] / (g["VMA10"] + 1e-12), np.nan)

    g["HH10_PRIOR"] = g["HH10"].shift(1)
    g["HH15_PRIOR"] = g["HH15"].shift(1)

    return g


def add_sentiment_detector(g: pd.DataFrame):
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
    dist += 25 * ((g["RET1"] <= 0.50).fillna(False)).astype(int)
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

    sig = np.array(["NEUTRAL"] * len(g), dtype=object)
    sig[g["AccumulationScore"] >= 60] = "ACCUMULATION"
    sig[g["ShakeoutScore"] >= 70] = "SHAKEOUT"
    sig[g["DistributionScore"] >= 60] = "DISTRIBUTION"
    g["SentimentSignal"] = sig

    reasons = []
    for i in range(len(g)):
        r = []
        if g["AccumulationScore"].iloc[i] >= 60:
            r.append("accumulation")
        if g["ShakeoutScore"].iloc[i] >= 70:
            r.append("shakeout")
        if g["DistributionScore"].iloc[i] >= 60:
            r.append("distribution")
        reasons.append(", ".join(r))
    g["SentimentReason"] = reasons

    return g


# =========================================================
# MODEL FEATURES
# =========================================================
def add_model_v2_features(g: pd.DataFrame):
    g = add_price_features(g)
    g = add_sentiment_detector(g)
    g = g.sort_values("TradeDate").copy()

    g["LL5_PRIOR"] = g["LL5"].shift(1)
    g["HH10_PRIOR"] = g["HH10"].shift(1)
    g["HH15_PRIOR"] = g["HH15"].shift(1)

    g["Stretch_MA10_Pct"] = np.where(g["MA10"] > 0, (g["Close"] / g["MA10"] - 1) * 100, np.nan)
    g["Stretch_MA20_Pct"] = np.where(g["MA20"] > 0, (g["Close"] / g["MA20"] - 1) * 100, np.nan)
    g["ATR_Stretch"] = np.where(g["ATR14"] > 0, (g["Close"] - g["MA10"]) / g["ATR14"], np.nan)

    g["GreenBar"] = (g["Close"] > g["Open"]).astype(int)
    g["RedBar"] = (g["Close"] < g["Open"]).astype(int)
    g["RedCount3"] = g["RedBar"].rolling(3).sum()
    g["RedCount4"] = g["RedBar"].rolling(4).sum()

    g["PrevHigh1"] = g["High"].shift(1)
    g["PrevHigh2"] = g["High"].shift(2)
    g["LowerHigh"] = ((g["High"] < g["PrevHigh1"]) & (g["PrevHigh1"] < g["PrevHigh2"])).astype(int)

    g["Near_HH10"] = (g["Close"] >= 0.98 * g["HH10"]).fillna(False)
    g["Near_HH15"] = (g["Close"] >= 0.98 * g["HH15"]).fillna(False)

    g["HighRejectBar"] = (
        (g["High"] >= g["HH10_PRIOR"]) &
        (g["Close"] < g["HH10_PRIOR"]) &
        (g["UpperWickPct"] >= 0.35)
    ).fillna(False)

    g["RejectHits3"] = g["HighRejectBar"].rolling(3).sum()

    g["FailedBreakout"] = (
        (
            (g["High"] > g["HH10_PRIOR"]) |
            (g["Close"] >= 0.99 * g["HH10_PRIOR"])
        ) &
        (g["Close"] < g["HH10_PRIOR"]) &
        (g["Close_Pos"] <= 0.45)
    ).fillna(False)

    g["Breakdown5"] = (
        (g["Close"] < g["LL5_PRIOR"]) &
        (g["Close"] < g["MA5"]) &
        (g["Close"] < g["VWAP"])
    ).fillna(False)

    g["Exhaustion"] = (
        (g["RET5"] >= 12) &
        (g["Stretch_MA10_Pct"] >= 8) &
        (
            (g["UpperWickPct"] >= 0.35) |
            (g["Close_Pos"] <= 0.45) |
            (g["RedCount3"] >= 2)
        )
    ).fillna(False)

    g["DeepRisk"] = (
        g["Exhaustion"] |
        (g["RejectHits3"] >= 2) |
        g["FailedBreakout"] |
        g["Breakdown5"] |
        (g["DistributionScore"] >= 70)
    )

    return g


# =========================================================
# 15-TRADING-DAY FLOOR FEATURES
# =========================================================
def add_floor_15d_features(floor_history: pd.DataFrame) -> pd.DataFrame:
    f = floor_history.copy()

    if f.empty:
        cols = [
            "Symbol", "TradeDate", "Trades", "Floor_Qty", "Floor_Amount_Cr", "Floor_VWAP",
            "FloorActive", "FloorDays15", "FloorQty15_Sum", "FloorAmt15_Sum",
            "FloorVWAP15_Avg", "FloorVWAP5_Avg", "FloorVWAP15_Max", "FloorVWAP15_Min",
            "RisingFloorVWAP", "FloorStrength15"
        ]
        return pd.DataFrame(columns=cols)

    f["TradeDate"] = pd.to_datetime(f["TradeDate"])
    f = f.sort_values(["Symbol", "TradeDate"]).copy()

    for c in ["Trades", "Floor_Qty", "Floor_Amount_Cr", "Floor_VWAP"]:
        if c not in f.columns:
            f[c] = np.nan

    def per_symbol(g: pd.DataFrame) -> pd.DataFrame:
        g = g.sort_values("TradeDate").copy()

        g["Floor_Qty"] = g["Floor_Qty"].fillna(0.0)
        g["Floor_Amount_Cr"] = g["Floor_Amount_Cr"].fillna(0.0)
        g["Trades"] = g["Trades"].fillna(0.0)

        g["FloorActive"] = (g["Floor_Qty"] > 0).astype(int)

        g["FloorDays15"] = g["FloorActive"].rolling(15, min_periods=15).sum()
        g["FloorQty15_Sum"] = g["Floor_Qty"].rolling(15, min_periods=15).sum()
        g["FloorAmt15_Sum"] = g["Floor_Amount_Cr"].rolling(15, min_periods=15).sum()
        g["FloorVWAP15_Avg"] = g["Floor_VWAP"].rolling(15, min_periods=15).mean()
        g["FloorVWAP15_Max"] = g["Floor_VWAP"].rolling(15, min_periods=15).max()
        g["FloorVWAP15_Min"] = g["Floor_VWAP"].rolling(15, min_periods=15).min()

        g["FloorVWAP5_Avg"] = g["Floor_VWAP"].rolling(5, min_periods=5).mean()
        g["RisingFloorVWAP"] = (
            (g["FloorVWAP5_Avg"] > g["FloorVWAP15_Avg"]) &
            g["FloorVWAP5_Avg"].notna() &
            g["FloorVWAP15_Avg"].notna()
        ).astype(int)

        g["FloorStrength15"] = np.where(
            g["FloorDays15"] > 0,
            g["FloorQty15_Sum"] / g["FloorDays15"],
            np.nan
        )

        return g

    f = f.groupby("Symbol", group_keys=False).apply(per_symbol)

    keep_cols = [
        "Symbol", "TradeDate", "Trades", "Floor_Qty", "Floor_Amount_Cr", "Floor_VWAP",
        "FloorActive", "FloorDays15", "FloorQty15_Sum", "FloorAmt15_Sum",
        "FloorVWAP15_Avg", "FloorVWAP5_Avg", "FloorVWAP15_Max", "FloorVWAP15_Min",
        "RisingFloorVWAP", "FloorStrength15"
    ]
    return f[keep_cols]


# =========================================================
# SNAPSHOT BUILDERS
# =========================================================
def build_symbol_snapshot_v2(price_all: pd.DataFrame):
    rows = []

    for sym, g in price_all.groupby("Symbol"):
        g = add_model_v2_features(g)
        if len(g) < 20:
            continue

        last = g.iloc[-1]

        rows.append({
            "TradeDate": last["TradeDate"],
            "Symbol": sym,
            "Open": last["Open"],
            "High": last["High"],
            "Low": last["Low"],
            "Close": last["Close"],
            "Volume": last["Volume"],
            "VWAP": last["VWAP"],

            "MA5": last["MA5"],
            "MA10": last["MA10"],
            "MA20": last["MA20"],
            "HH10": last["HH10"],
            "HH15": last["HH15"],
            "LL5": last["LL5"],
            "LL10": last["LL10"],

            "RET1_%": last["RET1"],
            "RET3_%": last["RET3"],
            "RET5_%": last["RET5"],
            "RET10_%": last["RET10"],

            "RSI14": last["RSI14"],
            "ATR14": last["ATR14"],
            "ATR14_Pct": last["ATR14_Pct"],
            "Vol_Surge": last["Vol_Surge"],

            "UpperWickPct": last["UpperWickPct"],
            "LowerWickPct": last["LowerWickPct"],
            "Close_Pos": last["Close_Pos"],
            "Body_Pct": last["Body_Pct"],
            "RangePct": last["RangePct"],

            "AccumulationScore": last["AccumulationScore"],
            "DistributionScore": last["DistributionScore"],
            "ShakeoutScore": last["ShakeoutScore"],
            "SentimentSignal": last["SentimentSignal"],
            "SentimentReason": last["SentimentReason"],

            "Stretch_MA10_Pct": last["Stretch_MA10_Pct"],
            "Stretch_MA20_Pct": last["Stretch_MA20_Pct"],
            "ATR_Stretch": last["ATR_Stretch"],
            "RedCount3": last["RedCount3"],
            "RedCount4": last["RedCount4"],
            "LowerHigh": last["LowerHigh"],
            "HighRejectBar": last["HighRejectBar"],
            "RejectHits3": last["RejectHits3"],
            "FailedBreakout": last["FailedBreakout"],
            "Breakdown5": last["Breakdown5"],
            "Exhaustion": last["Exhaustion"],
            "DeepRisk": last["DeepRisk"],
        })

    out = pd.DataFrame(rows)
    if not out.empty:
        out["TradeDate"] = pd.to_datetime(out["TradeDate"])
    return out


def build_snapshot_for_date(price_all: pd.DataFrame, asof_date):
    asof_date = pd.to_datetime(asof_date)
    sub = price_all[price_all["TradeDate"] <= asof_date].copy()
    if sub.empty:
        return pd.DataFrame()
    return build_symbol_snapshot_v2(sub)


# =========================================================
# FIXED SCORING
# =========================================================
def score_3state(x: pd.DataFrame):
    x = x.copy()

    default_zero_cols = [
        "Buy_Pressure", "Sell_Pressure", "Net_Broker_Bias",
        "Accum_Brokers", "Dist_Brokers", "Active_Brokers", "Broker_Balance",
        "FloorDays15", "FloorQty15_Sum", "FloorAmt15_Sum", "RisingFloorVWAP"
    ]
    for c in default_zero_cols:
        if c not in x.columns:
            x[c] = 0
        x[c] = x[c].fillna(0)

    for c in ["FloorVWAP15_Avg", "FloorVWAP5_Avg", "FloorStrength15", "Floor_VWAP", "VWAP"]:
        if c not in x.columns:
            x[c] = np.nan

    trend = (
        (x["Close"] > x["MA5"]).astype(int) * 12 +
        (x["MA5"] > x["MA10"]).astype(int) * 14 +
        (x["MA10"] > x["MA20"]).astype(int) * 12 +
        (x["Close"] >= 0.98 * x["HH10"]).astype(int) * 10 +
        (x["RSI14"].between(52, 70, inclusive="both")).astype(int) * 10
    )

    smart = (
        (x["Buy_Pressure"] > x["Sell_Pressure"]).astype(int) * 14 +
        (x["Buy_Pressure"] >= 0.55).astype(int) * 12 +
        (x["Broker_Balance"] > 0).astype(int) * 10 +
        (x["Accum_Brokers"] >= x["Dist_Brokers"]).astype(int) * 8 +
        (x["Close"] > x["VWAP"]).astype(int) * 10 +
        (x["AccumulationScore"] >= 60).astype(int) * 8 +
        (x["Close"] > x["FloorVWAP15_Avg"]).fillna(False).astype(int) * 14 +
        (x["FloorDays15"] >= 5).astype(int) * 8 +
        (x["RisingFloorVWAP"] == 1).astype(int) * 8 +
        (x["FloorQty15_Sum"] > 0).astype(int) * 6
    )

    quality = (
        (x["Close_Pos"] >= 0.60).astype(int) * 10 +
        (x["UpperWickPct"] <= 0.30).astype(int) * 8 +
        (x["Vol_Surge"] >= 1.20).astype(int) * 8 +
        (x["LowerWickPct"] >= x["UpperWickPct"]).astype(int) * 4
    )

    risk = (
        (x["Sell_Pressure"] >= 0.55).astype(int) * 16 +
        (x["DistributionScore"] >= 70).astype(int) * 14 +
        (x["Close"] < x["VWAP"]).astype(int) * 10 +
        (x["Close"] < x["FloorVWAP15_Avg"]).fillna(False).astype(int) * 14 +
        (x["UpperWickPct"] >= 0.45).astype(int) * 8 +
        (x["Close_Pos"] <= 0.35).astype(int) * 8 +
        (x["RedCount3"] >= 2).astype(int) * 8 +
        (x["LowerHigh"] == 1).astype(int) * 8 +
        (x["RejectHits3"] >= 2).astype(int) * 18 +
        (x["FailedBreakout"]).astype(int) * 18 +
        (x["Breakdown5"]).astype(int) * 22 +
        (x["Exhaustion"]).astype(int) * 16 +
        (x["FloorDays15"] <= 1).astype(int) * 6
    )

    x["TrendScore"] = trend
    x["SmartScore"] = smart
    x["QualityScore"] = quality
    x["RiskScore"] = risk

    buy_raw = x["TrendScore"] + x["SmartScore"] + x["QualityScore"] - 0.75 * x["RiskScore"]
    sell_raw = x["RiskScore"] + (x["DistributionScore"] >= 70).astype(int) * 8

    x["BuyScore"] = clamp_score(20 + 0.55 * buy_raw)
    x["SellScore"] = clamp_score(0.85 * sell_raw)

    x["Signal"] = "HOLD"

    hard_sell = (
        (x["SellScore"] >= SELL_TH) |
        (x["Breakdown5"]) |
        ((x["RejectHits3"] >= 2) & (x["Close"] < x["VWAP"])) |
        ((x["Exhaustion"]) & (x["RedCount3"] >= 2)) |
        ((x["Close"] < x["FloorVWAP15_Avg"]).fillna(False) & (x["FloorDays15"] >= 5))
    )

    buy_ok = (
        (x["BuyScore"] >= BUY_TH) &
        (x["SellScore"] < 45) &
        (x["Close"] > x["VWAP"]) &
        ((x["Close"] > x["FloorVWAP15_Avg"]).fillna(False))
    )

    x.loc[hard_sell, "Signal"] = "SELL"
    x.loc[(~hard_sell) & buy_ok, "Signal"] = "BUY"

    return x


# =========================================================
# REASONS / ACTION PLAN
# =========================================================
def build_reason_3state(r):
    parts = []

    if r["Signal"] == "BUY":
        if r.get("Buy_Pressure", 0) > r.get("Sell_Pressure", 0):
            parts.append("broker buying stronger")
        if pd.notna(r.get("VWAP")) and r.get("Close", 0) > r.get("VWAP", np.inf):
            parts.append("above VWAP")
        if pd.notna(r.get("FloorVWAP15_Avg")) and r.get("Close", 0) > r.get("FloorVWAP15_Avg", np.inf):
            parts.append("above 15-day floor VWAP")
        if r.get("FloorDays15", 0) >= 5:
            parts.append("repeated floor support")
        if r.get("RisingFloorVWAP", 0) == 1:
            parts.append("floor VWAP rising")
        if r.get("AccumulationScore", 0) >= 60:
            parts.append("accumulation visible")
        if r.get("Close", 0) >= 0.98 * r.get("HH10", np.inf):
            parts.append("near breakout zone")

    elif r["Signal"] == "SELL":
        if r.get("Breakdown5", False):
            parts.append("short-term breakdown")
        if r.get("FailedBreakout", False):
            parts.append("failed breakout")
        if r.get("RejectHits3", 0) >= 2:
            parts.append("repeated top rejection")
        if r.get("Exhaustion", False):
            parts.append("post-rally exhaustion")
        if r.get("DistributionScore", 0) >= 70:
            parts.append("distribution risk")
        if r.get("Sell_Pressure", 0) >= 0.55:
            parts.append("sell pressure high")
        if pd.notna(r.get("FloorVWAP15_Avg")) and r.get("Close", np.inf) < r.get("FloorVWAP15_Avg", -np.inf):
            parts.append("below 15-day floor support")

    else:
        if r.get("Close", 0) > r.get("VWAP", -np.inf):
            parts.append("holding above VWAP")
        if pd.notna(r.get("FloorVWAP15_Avg")) and r.get("Close", 0) > r.get("FloorVWAP15_Avg", -np.inf):
            parts.append("holding above 15-day floor VWAP")
        if r.get("FloorDays15", 0) >= 5:
            parts.append("floor activity present")
        if r.get("SellScore", 0) < SELL_TH:
            parts.append("no major sell trigger")
        if r.get("BuyScore", 0) < BUY_TH:
            parts.append("buy setup incomplete")

    return ", ".join(parts[:6])


def build_action_plan_3state(signal):
    if signal == "BUY":
        return "Buy on current strength or small dip; keep stop below LL5 / ATR stop."
    if signal == "SELL":
        return "Reduce or exit; avoid averaging until new base forms."
    return "Hold and wait; no fresh aggressive entry yet."


# =========================================================
# FINAL BUILD
# =========================================================
def build_signal_3state(snapshot, floor_history, pressure, broker_stats, sector):
    x = snapshot.copy()
    x["TradeDate"] = pd.to_datetime(x["TradeDate"])

    floor_15d = add_floor_15d_features(floor_history)

    x = x.merge(floor_15d, on=["Symbol", "TradeDate"], how="left")
    x = x.merge(pressure, on="Symbol", how="left")
    x = x.merge(broker_stats, on="Symbol", how="left")
    x = x.merge(sector, on="Symbol", how="left")

    for c in [
        "Trades", "Floor_Qty", "Floor_Amount_Cr", "Floor_VWAP",
        "FloorDays15", "FloorQty15_Sum", "FloorAmt15_Sum",
        "FloorVWAP15_Avg", "FloorVWAP5_Avg", "FloorVWAP15_Max", "FloorVWAP15_Min",
        "RisingFloorVWAP", "FloorStrength15",
        "Buy_Pressure", "Sell_Pressure", "Net_Broker_Bias",
        "Accum_Brokers", "Dist_Brokers", "Active_Brokers", "Broker_Balance"
    ]:
        if c not in x.columns:
            x[c] = np.nan

    x = score_3state(x)

    x["Entry"] = x["Close"].round(2)

    stop_base_1 = x["LL5"]
    stop_base_2 = x["Close"] - 1.2 * x["ATR14"]

    x["StopLoss"] = np.where(
        stop_base_1.notna() & stop_base_2.notna(),
        np.minimum(stop_base_1, stop_base_2),
        np.where(stop_base_1.notna(), stop_base_1, stop_base_2),
    ).round(2)

    x["RiskPerShare"] = (x["Entry"] - x["StopLoss"]).clip(lower=0)
    x["Target1"] = (x["Entry"] + 1.5 * x["RiskPerShare"]).round(2)
    x["Target2"] = (x["Entry"] + 2.5 * x["RiskPerShare"]).round(2)

    x["InsightReason"] = x.apply(build_reason_3state, axis=1)
    x["ActionPlan"] = x["Signal"].apply(build_action_plan_3state)

    out_cols = [
        "TradeDate", "Symbol", "Sector", "Company",
        "Signal", "BuyScore", "SellScore",
        "TrendScore", "SmartScore", "QualityScore", "RiskScore",
        "Entry", "StopLoss", "Target1", "Target2",
        "Close", "VWAP", "Floor_VWAP", "FloorVWAP15_Avg", "FloorVWAP5_Avg",
        "RET1_%", "RET3_%", "RET5_%", "RET10_%",
        "MA5", "MA10", "MA20", "HH10", "HH15", "LL5", "LL10",
        "Volume", "Vol_Surge",
        "Trades", "Floor_Qty", "Floor_Amount_Cr",
        "FloorDays15", "FloorQty15_Sum", "FloorAmt15_Sum",
        "FloorVWAP15_Max", "FloorVWAP15_Min", "RisingFloorVWAP", "FloorStrength15",
        "Buy_Pressure", "Sell_Pressure", "Net_Broker_Bias",
        "Accum_Brokers", "Dist_Brokers", "Active_Brokers", "Broker_Balance",
        "RSI14", "ATR14", "ATR14_Pct",
        "Close_Pos", "UpperWickPct", "LowerWickPct",
        "AccumulationScore", "DistributionScore", "ShakeoutScore",
        "Stretch_MA10_Pct", "Stretch_MA20_Pct", "ATR_Stretch",
        "RedCount3", "RedCount4", "LowerHigh",
        "HighRejectBar", "RejectHits3", "FailedBreakout", "Breakdown5", "Exhaustion", "DeepRisk",
        "SentimentSignal", "SentimentReason",
        "InsightReason", "ActionPlan"
    ]
    out_cols = [c for c in out_cols if c in x.columns]

    signal_order = {"BUY": 0, "HOLD": 1, "SELL": 2}
    x["_signal_order"] = x["Signal"].map(signal_order).fillna(9)

    out = x[out_cols + ["_signal_order"]].sort_values(
        ["_signal_order", "BuyScore", "SellScore"],
        ascending=[True, False, False]
    ).drop(columns=["_signal_order"])

    return out


# =========================================================
# DATA PIPELINE
# =========================================================
def load_pipeline_data():
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
    hist_dates = choose_last_n_dates(common_dates, max(PRICE_HISTORY_LOAD, LOOKBACK_SIGNAL + 20))

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

    floor_daily = symbol_metrics_from_floorsheet_daily(fs)
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
            "Symbol", "Broker", "Buy_Qty", "Sell_Qty", "Net_Qty", "Buy_Amount", "Active_Days"
        ])

    pressure = compute_pressure(bs_window, topn=5)
    broker_stats = build_broker_stats(bs_window)

    return pr, floor_daily, pressure, broker_stats, sector, signal_dates[-1]


def load_full_history_for_backtest():
    floor_pairs = list_dated_files(FLOOR_DIR, FLOOR_RE)
    price_pairs = list_dated_files(PRICE_DIR, PRICE_RE)

    if not floor_pairs or not price_pairs:
        raise RuntimeError("Backtest requires both floor and price history files.")

    floor_dates = [d for d, _ in floor_pairs]
    price_dates = [d for d, _ in price_pairs]
    common_dates = sorted(set(floor_dates).intersection(set(price_dates)))

    if len(common_dates) < 20:
        raise RuntimeError("Not enough common dates for backtest.")

    floor_map = {d: p for d, p in floor_pairs}
    price_map = {d: p for d, p in price_pairs}

    fs_list = [read_floorsheet_file(floor_map[d], d) for d in common_dates]
    pr_list = [read_price_file(price_map[d], d) for d in common_dates]

    fs = pd.concat(fs_list, ignore_index=True)
    pr = pd.concat(pr_list, ignore_index=True)
    sector = load_sector_master(SECTOR_FILE)

    return pr, fs, sector, common_dates


# =========================================================
# MAIN SIGNAL API
# =========================================================
def run_3state_model():
    price_all, floor_history, pressure, broker_stats, sector, latest_dt = load_pipeline_data()
    snapshot = build_symbol_snapshot_v2(price_all)

    final_signal = build_signal_3state(
        snapshot=snapshot,
        floor_history=floor_history,
        pressure=pressure,
        broker_stats=broker_stats,
        sector=sector
    )

    return final_signal, latest_dt


# =========================================================
# BACKTEST
# =========================================================
def build_signal_history():
    price_all, fs_all, sector, common_dates = load_full_history_for_backtest()

    floor_daily_all = symbol_metrics_from_floorsheet_daily(fs_all)
    bs_daily_all = broker_symbol_metrics(fs_all)

    rows = []

    for dt in common_dates[20:]:
        dt = pd.to_datetime(dt)

        pr_sub = price_all[price_all["TradeDate"] <= dt].copy()
        fs_sub = floor_daily_all[floor_daily_all["TradeDate"] <= dt].copy()
        bs_sub = bs_daily_all[bs_daily_all["TradeDate"] <= dt].copy()

        snapshot = build_snapshot_for_date(pr_sub, dt)
        if snapshot.empty:
            continue

        bs_window = pd.DataFrame(columns=[
            "Symbol", "Broker", "Buy_Qty", "Sell_Qty", "Net_Qty", "Buy_Amount", "Active_Days"
        ])
        if not bs_sub.empty:
            bs_window = bs_sub.groupby(["Symbol", "Broker"], as_index=False).agg(
                Buy_Qty=("Buy_Qty", "sum"),
                Sell_Qty=("Sell_Qty", "sum"),
                Net_Qty=("Net_Qty", "sum"),
                Buy_Amount=("Buy_Amount", "sum"),
                Active_Days=("Active_Buy_Day", "sum"),
            )

        pressure = compute_pressure(bs_window, topn=5)
        broker_stats = build_broker_stats(bs_window)

        sig = build_signal_3state(
            snapshot=snapshot,
            floor_history=fs_sub,
            pressure=pressure,
            broker_stats=broker_stats,
            sector=sector
        )
        rows.append(sig)

    if not rows:
        return pd.DataFrame()

    return pd.concat(rows, ignore_index=True)


def add_forward_returns(signal_history: pd.DataFrame, price_all: pd.DataFrame, horizons=(3, 5, 10, 15)):
    if signal_history.empty:
        return signal_history.copy()

    px = price_all[["TradeDate", "Symbol", "Close"]].copy()
    px = px.sort_values(["Symbol", "TradeDate"]).copy()

    for h in horizons:
        tmp = px.copy()
        tmp[f"Close_fwd_{h}"] = tmp.groupby("Symbol")["Close"].shift(-h)
        tmp[f"FwdRet_{h}D_%"] = (tmp[f"Close_fwd_{h}"] / tmp["Close"] - 1.0) * 100
        signal_history = signal_history.merge(
            tmp[["TradeDate", "Symbol", f"FwdRet_{h}D_%"]],
            on=["TradeDate", "Symbol"],
            how="left"
        )

    return signal_history


def summarize_backtest(signal_history_with_ret: pd.DataFrame, horizons=(3, 5, 10, 15)):
    if signal_history_with_ret.empty:
        return pd.DataFrame()

    frames = []
    for h in horizons:
        col = f"FwdRet_{h}D_%"
        if col not in signal_history_with_ret.columns:
            continue

        g = signal_history_with_ret.groupby("Signal", dropna=False)[col].agg(
            Count="count",
            AvgRet="mean",
            MedianRet="median",
            WinRate=lambda s: float((s > 0).mean()) if len(s.dropna()) else np.nan,
            LossRate=lambda s: float((s < 0).mean()) if len(s.dropna()) else np.nan,
        ).reset_index()

        g["Horizon"] = f"{h}D"
        frames.append(g)

    if not frames:
        return pd.DataFrame()

    out = pd.concat(frames, ignore_index=True)
    return out[["Horizon", "Signal", "Count", "AvgRet", "MedianRet", "WinRate", "LossRate"]]


def run_backtest():
    price_all, _, _, _, _, _ = load_pipeline_data()
    signal_history = build_signal_history()
    signal_history = add_forward_returns(signal_history, price_all, horizons=(3, 5, 10, 15))
    summary = summarize_backtest(signal_history, horizons=(3, 5, 10, 15))
    return signal_history, summary


# =========================================================
# EXCEL HELPERS
# =========================================================
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


def add_table_sheet(ws, df, table_name):
    if df.empty:
        df = pd.DataFrame([["No data"]], columns=["Info"])

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    style_sheet(ws)

    nrows = ws.max_row
    ncols = ws.max_column
    ref = f"A1:{get_column_letter(ncols)}{nrows}"

    tab = Table(displayName=table_name, ref=ref)
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


def apply_number_formats(ws):
    fmt_map = {
        "BuyScore": "0.00",
        "SellScore": "0.00",
        "TrendScore": "0.00",
        "SmartScore": "0.00",
        "QualityScore": "0.00",
        "RiskScore": "0.00",
        "Entry": "#,##0.00",
        "StopLoss": "#,##0.00",
        "Target1": "#,##0.00",
        "Target2": "#,##0.00",
        "Close": "#,##0.00",
        "VWAP": "#,##0.00",
        "Floor_VWAP": "#,##0.00",
        "FloorVWAP15_Avg": "#,##0.00",
        "FloorVWAP5_Avg": "#,##0.00",
        "FloorVWAP15_Max": "#,##0.00",
        "FloorVWAP15_Min": "#,##0.00",
        "RET1_%": "0.00",
        "RET3_%": "0.00",
        "RET5_%": "0.00",
        "RET10_%": "0.00",
        "MA5": "#,##0.00",
        "MA10": "#,##0.00",
        "MA20": "#,##0.00",
        "HH10": "#,##0.00",
        "HH15": "#,##0.00",
        "LL5": "#,##0.00",
        "LL10": "#,##0.00",
        "Volume": "#,##0",
        "Trades": "0",
        "Floor_Qty": "#,##0",
        "Floor_Amount_Cr": "0.000",
        "FloorDays15": "0",
        "FloorQty15_Sum": "#,##0",
        "FloorAmt15_Sum": "0.000",
        "RisingFloorVWAP": "0",
        "FloorStrength15": "#,##0.00",
        "Buy_Pressure": "0.00",
        "Sell_Pressure": "0.00",
        "Net_Broker_Bias": "#,##0",
        "Accum_Brokers": "0",
        "Dist_Brokers": "0",
        "Active_Brokers": "0",
        "Broker_Balance": "0",
        "RSI14": "0.00",
        "ATR14": "0.00",
        "ATR14_Pct": "0.00",
        "Close_Pos": "0.00",
        "UpperWickPct": "0.00",
        "LowerWickPct": "0.00",
        "AccumulationScore": "0",
        "DistributionScore": "0",
        "ShakeoutScore": "0",
        "Stretch_MA10_Pct": "0.00",
        "Stretch_MA20_Pct": "0.00",
        "ATR_Stretch": "0.00",
        "RedCount3": "0",
        "RedCount4": "0",
        "RejectHits3": "0",
        "FwdRet_3D_%": "0.00",
        "FwdRet_5D_%": "0.00",
        "FwdRet_10D_%": "0.00",
        "FwdRet_15D_%": "0.00",
        "AvgRet": "0.00",
        "MedianRet": "0.00",
        "WinRate": "0.00%",
        "LossRate": "0.00%",
    }

    header = [c.value for c in ws[1]]
    for col_name, fmt in fmt_map.items():
        if col_name not in header:
            continue
        idx = header.index(col_name) + 1
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=idx).number_format = fmt


def save_to_excel(current_signals, signal_history, backtest_summary, latest_dt):
    out_path = OUT_DIR / f"nepse_3state_signals_{latest_dt}.xlsx"

    wb = Workbook()
    wb.remove(wb.active)

    readme = pd.DataFrame([
        ["Purpose", "3-state BUY / HOLD / SELL model with 15-day floor analysis and backtest summary."],
        ["Floor path", str(FLOOR_DIR)],
        ["Price path", str(PRICE_DIR)],
        ["Sector file", str(SECTOR_FILE)],
        ["Thresholds", f"BUY_TH={BUY_TH}, SELL_TH={SELL_TH}"],
        ["Excel format", "Each sheet is a styled Excel table with wrap text and auto-fit row/column sizing."],
        ["Floor heading", "Transact No., Symbol, Buyer, Seller, Quantity, Rate, Amount"],
    ], columns=["Item", "Explanation"])

    sheets = {
        "README": readme,
        "Current_Signals": current_signals,
        "BUY_List": current_signals[current_signals["Signal"] == "BUY"].copy(),
        "HOLD_List": current_signals[current_signals["Signal"] == "HOLD"].copy(),
        "SELL_List": current_signals[current_signals["Signal"] == "SELL"].copy(),
        "Backtest_Summary": backtest_summary,
        "Signal_History": signal_history,
    }

    for i, (sheet_name, df) in enumerate(sheets.items(), start=1):
        ws = wb.create_sheet(sheet_name[:31])
        add_table_sheet(ws, df, f"T{i}")
        apply_number_formats(ws)

    wb.save(out_path)
    return out_path


# =========================================================
# SCRIPT RUN
# =========================================================
if __name__ == "__main__":
    print("ROOT =", ROOT)
    print("PRICE_DIR =", PRICE_DIR)
    print("FLOOR_DIR =", FLOOR_DIR)
    print("SECTOR_FILE =", SECTOR_FILE)

    current_signals, latest_dt = run_3state_model()
    signal_history, backtest_summary = run_backtest()

    excel_path = save_to_excel(
        current_signals=current_signals,
        signal_history=signal_history,
        backtest_summary=backtest_summary,
        latest_dt=latest_dt
    )

    print("\n=== CURRENT SIGNALS ===")
    print(current_signals.head(20).to_string(index=False))

    print("\n=== BACKTEST SUMMARY ===")
    print(backtest_summary.to_string(index=False))

    print(f"\n✅ Excel created: {excel_path}")
