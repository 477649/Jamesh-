# priceaction_signals.py
# ✅ Signals (7D/15D) + Circuits + Swing Scanner (3–10D)
# ✅ Adds TrendAge + DaysFromBreakout (freshness + ageing)
# ✅ Smart-Money Sentiment Detector: Accumulation / Distribution / Shakeout
# ✅ UPDATED: SentimentSignal is MULTI-DAY CONFIRMED (not single candle)
#     - Accumulation: AccScore>=60 on 3 of last 4 days
#     - Distribution: DistScore>=60 on 2 of last 4 days
#     - Shakeout: ShakeScore>=70 on 1 of last 3 days
# ✅ NEW: Sentiment_Top sheet uses confirmed logic
# ✅ Loads latest available SharePrice_YYYY-MM-DD.csv files (<= today)

import os, re, glob
import numpy as np
import pandas as pd
from datetime import date as _date

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter


# =========================
# PATHS / SETTINGS
# =========================
DATA_DIR = "outputs/sharesansar"                   # INPUT daily share price CSVs
SECTOR_FILE = "outputs/Sector/sector_master.csv"   # Symbol -> Sector/Company (label only)
OUT_DIR  = "outputs/PriceAction"

LATEST_FILES_TO_LOAD = 60
TOP_CIRCUIT_N = 25

# ---- Swing Scanner settings (NEPSE daily, ideal hold 3–10 trading days) ----
SWING_OUT_TOP_N = 120
SWING_MIN_HISTORY = 30                  # 30+ workable; 60–90 better
SWING_VALUE_TRADED_MIN = 2_500_000      # liquidity filter (tune)
SWING_VOL_SPIKE_BREAKOUT = 1.30         # breakout confirmation
SWING_VOL_SPIKE_COMPRESSION = 1.40      # compression expansion
SWING_RSI_MIN = 50
SWING_RSI_MAX = 70
SWING_MAX_HOLD_DAYS = 10
SWING_TIME_STOP_DAY = 6

# Breakout/trend-age settings
BREAKOUT_LOOKBACK = 20                  # 20 trading days for HH20
MAX_BREAKOUT_AGE_TRACK = 60             # only search last 60 trading days (enough)

# Sentiment Top sheet
SENTIMENT_TOP_N_EACH = 25               # top per category (acc/dist/shake)

# ✅ MULTI-DAY sentiment thresholds + windows (CONFIRMATION)
SENT_ACC_TH = 60
SENT_DIST_TH = 60
SENT_SHAKE_TH = 70

ACC_WIN = 4
ACC_MIN_HITS = 3     # 3 of last 4

DIST_WIN = 4
DIST_MIN_HITS = 2    # 2 of last 4

SHAKE_WIN = 3
SHAKE_MIN_HITS = 1   # 1 of last 3 (event)


# =========================
# LOAD DATA
# =========================
def _extract_date_from_filename(path):
    m = re.search(r"SharePrice_(\d{4}-\d{2}-\d{2})", os.path.basename(path))
    if not m:
        return None
    try:
        return pd.to_datetime(m.group(1)).date()
    except Exception:
        return None


def load_latest_files(folder, latest_n=60):
    """
    ✅ Loads latest available files up to *today* (not assuming yesterday exists).
    Filenames must look like SharePrice_YYYY-MM-DD.csv
    """
    files = sorted(glob.glob(os.path.join(folder, "SharePrice_*.csv")))
    if not files:
        raise FileNotFoundError(f"No SharePrice_*.csv files found in: {folder}")

    today = _date.today()

    dated = []
    for f in files:
        d = _extract_date_from_filename(f)
        if d is None:
            continue
        if d <= today:
            dated.append((d, f))

    if not dated:
        raise FileNotFoundError(f"No SharePrice_*.csv files <= today ({today}) in: {folder}")

    dated.sort(key=lambda x: x[0])
    dated = dated[-latest_n:]
    chosen_files = [f for _, f in dated]

    rows = []
    for f in chosen_files:
        file_dt = _extract_date_from_filename(f)
        date_val = pd.to_datetime(file_dt) if file_dt else pd.NaT

        df = pd.read_csv(f)
        df.columns = [c.strip() for c in df.columns]

        # normalize Close/LTP
        if "Close" not in df.columns and "LTP" in df.columns:
            df["Close"] = df["LTP"]

        # normalize Volume
        if "Volume" not in df.columns:
            for c in ["Vol", "VOL", "Total Traded Quantity", "TotalQty", "Qty"]:
                if c in df.columns:
                    df["Volume"] = df[c]
                    break

        # normalize VWAP
        if "VWAP" not in df.columns:
            for c in ["vwap", "Vwap", "VWAP Price", "Daily VWAP", "AvgPrice", "Average Price"]:
                if c in df.columns:
                    df["VWAP"] = df[c]
                    break

        df["Date"] = date_val

        required = ["Symbol", "Open", "High", "Low", "Close", "Volume", "VWAP"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"Missing columns {missing} in file: {f}")

        for c in ["Open", "High", "Low", "Close", "Volume", "VWAP"]:
            df[c] = df[c].astype(str).str.replace(",", "", regex=False)
            df[c] = pd.to_numeric(df[c], errors="coerce")

        rows.append(df[["Date", "Symbol", "Open", "High", "Low", "Close", "Volume", "VWAP"]])

    out = (
        pd.concat(rows, ignore_index=True)
        .sort_values(["Symbol", "Date"])
        .reset_index(drop=True)
    )

    out = out.dropna(subset=["Open", "High", "Low", "Close", "Volume", "VWAP"])
    return out


def load_sector_master(path):
    """
    Used ONLY for labeling Sector/Company in outputs.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"Sector master file not found: {path}")

    df = pd.read_csv(path)
    df.columns = [c.strip() for c in df.columns]

    if "Symbol" not in df.columns:
        raise ValueError("sector_master.csv must contain column: Symbol")

    sector_col = None
    for c in ["Sector", "Sectors", "sector", "sectors"]:
        if c in df.columns:
            sector_col = c
            break
    if sector_col is None:
        df["Sector"] = None
        sector_col = "Sector"

    company_col = None
    for c in ["Company", "Company Name", "company"]:
        if c in df.columns:
            company_col = c
            break

    out = df[["Symbol", sector_col]].copy().rename(columns={sector_col: "Sector"})
    out["Company"] = df[company_col] if company_col else None
    return out.drop_duplicates(subset=["Symbol"], keep="last")


# =========================
# MATH / INDICATORS
# =========================
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


def zscore(s, n):
    return (s - s.rolling(n).mean()) / (s.rolling(n).std() + 1e-12)


def slope_r2_last_n(series, n):
    y = series.tail(n).values
    if len(y) < n or np.isnan(y).any():
        return np.nan, np.nan
    x = np.arange(n)
    m, _ = np.polyfit(x, y, 1)
    r = np.corrcoef(x, y)[0, 1]
    return float(m), float(r * r)


def ema(s, n):
    return s.ewm(span=n, adjust=False).mean()


def add_features(g):
    g = g.copy()

    for n in [7, 10, 15, 30]:
        g[f"MA{n}"] = g["Close"].rolling(n).mean()
        g[f"VMA{n}"] = g["Volume"].rolling(n).mean()

    g["HH7"]  = g["High"].rolling(7).max()
    g["LL7"]  = g["Low"].rolling(7).min()
    g["HH15"] = g["High"].rolling(15).max()
    g["LL15"] = g["Low"].rolling(15).min()

    # Returns
    for n in [1, 2, 4, 7, 10, 15, 20, 30]:
        g[f"RET{n}"] = g["Close"].pct_change(n) * 100

    rng = (g["High"] - g["Low"]).replace(0, np.nan)
    g["UpperWickPct"] = ((g["High"] - g[["Open", "Close"]].max(axis=1)) / rng).clip(0, 1)

    g["RSI14"] = rsi(g["Close"], 14)

    tr = true_range(g["High"], g["Low"], g["Close"])
    g["ATR7"] = tr.rolling(7).mean()
    g["ATR15"] = tr.rolling(15).mean()
    g["ATR7_%"] = (g["ATR7"] / (g["Close"] + 1e-12)) * 100
    g["ATR_%"] = (g["ATR15"] / (g["Close"] + 1e-12)) * 100

    return g


# =========================
# SENTIMENT / MANIPULATION DETECTOR (OHLCV + VWAP only)
# ✅ UPDATED: SentimentSignal is MULTI-DAY confirmed (hit-count windows)
# =========================
def add_sentiment_detector(g):
    """
    Adds:
      - LowerWickPct, BodyPct01, RangePct_D
      - VolAvg5, VolAvg20, VolRatio20
      - AccumulationScore (0..100)
      - DistributionScore (0..100)
      - ShakeoutScore (0..100)
      ✅ Multi-day confirmation:
      - AccHits4, DistHits4, ShakeHits3
      - AccConfirmed, DistConfirmed, ShakeConfirmed
      - SentimentSignal (confirmed)
      - SentimentReason (confirmed)
    """
    g = g.copy()

    rng = (g["High"] - g["Low"]).replace(0, np.nan)

    body = (g["Close"] - g["Open"]).abs()
    upper = (g["High"] - g[["Open", "Close"]].max(axis=1))
    lower = (g[["Open", "Close"]].min(axis=1) - g["Low"])

    g["BodyPct01"] = (body / (rng + 1e-12)).clip(0, 1)
    g["LowerWickPct"] = (lower / (rng + 1e-12)).clip(0, 1)

    if "UpperWickPct" not in g.columns:
        g["UpperWickPct"] = (upper / (rng + 1e-12)).clip(0, 1)

    g["RangePct_D"] = (rng / (g["Close"] + 1e-12)) * 100
    g["RangePct_MA10"] = g["RangePct_D"].rolling(10).mean()

    g["VolAvg5"] = g["Volume"].rolling(5).mean()
    g["VolAvg20"] = g["Volume"].rolling(20).mean()
    g["VolRatio20"] = g["Volume"] / (g["VolAvg20"] + 1e-12)

    if "RET1" not in g.columns:
        g["RET1"] = g["Close"].pct_change(1) * 100

    g["HH20"] = g["High"].rolling(20).max()
    near_hh20 = (g["Close"] >= 0.97 * g["HH20"]).fillna(False)

    # ---- Accumulation (single-day score) ----
    acc = np.zeros(len(g), dtype=float)
    acc += 20 * ((g["VolAvg5"] > g["VolAvg20"]).fillna(False)).astype(int)
    acc += 20 * ((g["RangePct_D"] <= (g["RangePct_MA10"] * 0.90)).fillna(False)).astype(int)
    acc += 20 * ((g["BodyPct01"] < 0.40).fillna(False)).astype(int)
    acc += 20 * ((g["Close"] > g["VWAP"]).fillna(False)).astype(int)
    acc += 20 * ((g["LowerWickPct"] > g["UpperWickPct"]).fillna(False)).astype(int)
    g["AccumulationScore"] = np.clip(acc, 0, 100).astype(int)

    # ---- Distribution (single-day score) ----
    dist = np.zeros(len(g), dtype=float)
    dist += 25 * ((g["VolRatio20"] >= 1.8).fillna(False)).astype(int)
    dist += 25 * ((g["RET1"] <= 0.50).fillna(False)).astype(int)
    dist += 20 * ((g["UpperWickPct"] >= 0.45).fillna(False)).astype(int)
    dist += 20 * ((g["Close"] < g["VWAP"]).fillna(False)).astype(int)
    dist += 10 * near_hh20.astype(int)
    g["DistributionScore"] = np.clip(dist, 0, 100).astype(int)

    # ---- Shakeout (single-day score) ----
    prev_low = g["Low"].shift(1)
    shake = np.zeros(len(g), dtype=float)
    shake += 50 * ((g["Low"] < prev_low) & (g["Close"] > prev_low)).fillna(False).astype(int)
    shake += 30 * ((g["LowerWickPct"] >= 0.50).fillna(False)).astype(int)
    shake += 20 * ((g["VolRatio20"] >= 1.5).fillna(False)).astype(int)
    g["ShakeoutScore"] = np.clip(shake, 0, 100).astype(int)

    # -----------------------
    # ✅ MULTI-DAY CONFIRMATION
    # -----------------------
    acc_hits = (g["AccumulationScore"] >= SENT_ACC_TH).rolling(ACC_WIN).sum()
    dist_hits = (g["DistributionScore"] >= SENT_DIST_TH).rolling(DIST_WIN).sum()
    shake_hits = (g["ShakeoutScore"] >= SENT_SHAKE_TH).rolling(SHAKE_WIN).sum()

    g["AccHits4"] = acc_hits
    g["DistHits4"] = dist_hits
    g["ShakeHits3"] = shake_hits

    g["AccConfirmed"] = (acc_hits >= ACC_MIN_HITS).fillna(False)
    g["DistConfirmed"] = (dist_hits >= DIST_MIN_HITS).fillna(False)
    g["ShakeConfirmed"] = (shake_hits >= SHAKE_MIN_HITS).fillna(False)

    # Priority: Distribution > Shakeout > Accumulation
    sig = np.array(["NEUTRAL"] * len(g), dtype=object)
    sig[g["AccConfirmed"]] = "ACCUMULATION"
    sig[g["ShakeConfirmed"]] = "SHAKEOUT"
    sig[g["DistConfirmed"]] = "DISTRIBUTION"
    g["SentimentSignal"] = sig

    reasons = []
    for i in range(len(g)):
        r = []
        if bool(g["AccConfirmed"].iloc[i]):
            r.append(f"Acc≥{SENT_ACC_TH} on {int(g['AccHits4'].iloc[i])}/{ACC_WIN} days")
        if bool(g["ShakeConfirmed"].iloc[i]):
            r.append(f"Shake≥{SENT_SHAKE_TH} on {int(g['ShakeHits3'].iloc[i])}/{SHAKE_WIN} days")
        if bool(g["DistConfirmed"].iloc[i]):
            r.append(f"Dist≥{SENT_DIST_TH} on {int(g['DistHits4'].iloc[i])}/{DIST_WIN} days")
        reasons.append("; ".join(r) if r else "")
    g["SentimentReason"] = reasons

    return g


# =========================
# SWING FEATURES (3–10D)
# =========================
def add_swing_features(g):
    g = g.copy()

    g["EMA10"] = ema(g["Close"], 10)
    g["EMA20"] = ema(g["Close"], 20)
    g["EMA50"] = ema(g["Close"], 50)

    tr = true_range(g["High"], g["Low"], g["Close"])
    g["ATR14"] = tr.rolling(14).mean()

    g["VolAvg5"] = g["Volume"].rolling(5).mean()
    g["VolAvg20"] = g["Volume"].rolling(20).mean()
    g["VolSpike20"] = g["Volume"] / (g["VolAvg20"] + 1e-12)

    g["ValueTraded"] = g["Close"] * g["Volume"]

    rng = (g["High"] - g["Low"]).replace(0, np.nan)
    g["RangePct"] = (rng / (g["Close"] + 1e-12)) * 100
    g["ClosePos"] = (g["Close"] - g["Low"]) / (rng + 1e-12)   # 0..1
    g["BodyPct"] = (g["Close"] - g["Open"]).abs() / (rng + 1e-12) * 100

    g["AboveVWAP"] = g["Close"] > g["VWAP"]
    g["VWAP_DistPct"] = (g["Close"] - g["VWAP"]) / (g["VWAP"] + 1e-12) * 100
    g["VWAP_Slope5"] = g["VWAP"].diff(5)

    # HH20 prior (exclude today)
    g["HH20_PRIOR"] = g["High"].rolling(BREAKOUT_LOOKBACK).max().shift(1)

    # ATR contraction proxy
    g["ATR14_5ago"] = g["ATR14"].shift(5)

    return g


# =========================
# TREND AGE / DAYS FROM BREAKOUT
# =========================
def compute_breakout_age_cols(g):
    """
    Adds:
      - BreakoutFlag (True on breakout day)
      - DaysFromBreakout (0 on breakout day, 1 next day, ...)
      - TrendAge (same measure; used for Trend/Pullback ageing)
    Breakout definition aligns to your scanner:
      Close > HH20_PRIOR AND VolSpike20 >= threshold AND ClosePos >= 0.65
    """
    g = g.copy()

    closepos_ok = g["ClosePos"] >= 0.65
    breakout = (
        (g["Close"] > g["HH20_PRIOR"]) &
        (g["VolSpike20"] >= SWING_VOL_SPIKE_BREAKOUT) &
        closepos_ok
    ).fillna(False)

    g["BreakoutFlag"] = breakout

    ages = []
    last_idx = None
    for i in range(len(g)):
        if breakout.iloc[i]:
            last_idx = i
            ages.append(0)
        else:
            ages.append(np.nan if last_idx is None else (i - last_idx))

    g["DaysFromBreakout"] = ages
    g["TrendAge"] = ages
    return g


# =========================
# WINDOW-WISE BEHAVIOR LOGIC (7D/15D)
# =========================
def _vol_trend_ratio_last_n(volume_series, n):
    if len(volume_series) < n:
        return np.nan
    w = np.asarray(volume_series.tail(n).values, dtype=float)
    if np.isnan(w).any():
        return np.nan
    k = max(2, n // 3)
    first = np.mean(w[:k])
    last = np.mean(w[-k:])
    return float(last / (first + 1e-12))


def add_window_behavior(g, win):
    g = g.copy()
    W = str(win)

    g[f"RET{W}_Window"] = g["Close"].pct_change(win) * 100
    g[f"VolTrendRatio_{W}"] = g["Volume"].rolling(win).apply(lambda s: _vol_trend_ratio_last_n(s, win), raw=False)

    cond_up = (g[f"RET{W}_Window"] > 0)
    cond_down = (g[f"RET{W}_Window"] < 0)
    vol_falling = (g[f"VolTrendRatio_{W}"] < 0.90)
    vol_rising = (g[f"VolTrendRatio_{W}"] > 1.10)

    g[f"VolPriceFlag_{W}"] = "NEUTRAL"
    g.loc[cond_up & vol_rising, f"VolPriceFlag_{W}"] = "CONFIRMED"
    g.loc[cond_up & vol_falling, f"VolPriceFlag_{W}"] = "DIVERGENCE"
    g.loc[cond_down & vol_rising, f"VolPriceFlag_{W}"] = "SELL_PRESSURE"

    vma = g["Volume"].rolling(win).mean()
    dist_day = (g["UpperWickPct"] >= 0.55) & (g["Volume"] >= 1.5 * (vma + 1e-12)) & (g["RET1"] >= 0)
    dist_count = dist_day.rolling(win).sum()
    weak_progress = (g[f"RET{W}_Window"] < 5).fillna(False)
    g[f"Distribution_{W}"] = ((dist_count >= 2) & weak_progress).fillna(False)

    hi = g["High"].rolling(win).max()
    lo = g["Low"].rolling(win).min()
    range_pct = ((hi - lo) / (g["Close"] + 1e-12)) * 100
    vol_hot = ((g["Volume"] >= 1.5 * (vma + 1e-12)).rolling(win).sum() >= 2)
    wick_ok = (g["UpperWickPct"].rolling(win).mean() <= 0.40)
    tight = (range_pct <= 8.0)
    g[f"Absorption_{W}"] = (vol_hot & wick_ok & tight).fillna(False)

    pos3 = (g["RET1"] > 0).rolling(3).sum() >= 2
    vol_ok = g["Volume"] >= 0.8 * (vma + 1e-12)
    g[f"FollowThrough_{W}"] = (pos3 & vol_ok).fillna(False)

    hh = g["High"].rolling(win).max()
    break_level = hh.shift(1)
    g[f"RetestBuy_{W}"] = (
        (g["Close"] >= break_level) &
        (g["Close"] <= break_level * 1.02) &
        (g["Volume"] >= (vma + 1e-12)) &
        (g["UpperWickPct"] <= 0.35)
    ).fillna(False)

    return g


# =========================
# ADVANCED INSIGHTS
# =========================
def trend_health(slope, r2):
    if pd.isna(slope) or pd.isna(r2):
        return ""
    if slope <= 0:
        return "DOWN"
    if r2 >= 0.50:
        return "GOOD"
    if r2 >= 0.30:
        return "WEAK"
    return "NOISY"


def vol_expansion_flag(atr7_pct_series, atr15_pct_series):
    if len(atr7_pct_series) < 4 or len(atr15_pct_series) < 1:
        return False
    a7 = atr7_pct_series.iloc[-1]
    a15 = atr15_pct_series.iloc[-1]
    if pd.isna(a7) or pd.isna(a15):
        return False
    rising3 = (atr7_pct_series.diff().tail(3) > 0).all()
    return bool((a7 > a15) and rising3)


def false_breakout_metrics(close, hh, volume, vma, upperwickpct):
    if pd.isna(close) or pd.isna(hh) or pd.isna(volume) or pd.isna(vma) or pd.isna(upperwickpct):
        return (False, None)

    breakout_attempt = close >= hh
    reject = upperwickpct > 0.45
    weak_vol = volume < vma
    flag = bool(breakout_attempt and (reject or weak_vol))

    wick_comp = float(np.clip((upperwickpct - 0.35) / 0.30, 0, 1))
    vol_ratio = float(volume / (vma + 1e-12))
    vol_comp = float(np.clip((1.0 - vol_ratio) / 0.50, 0, 1))
    dist = float((close / (hh + 1e-12)) - 1.0)
    dist_comp = float(np.clip(dist / 0.03, 0, 1))

    score = int(round((0.45 * wick_comp + 0.35 * vol_comp + 0.20 * dist_comp) * 100, 0))
    return (flag, score)


# =========================
# SIGNAL HELPERS
# =========================
def stretch_flag(z):
    if pd.isna(z):
        return ""
    if z > 2.0:
        return "STRETCHED"
    if z < -2.0:
        return "OVERSOLD"
    return "HEALTHY"


def vol_regime_from_atr_pct(atr_pct):
    if pd.isna(atr_pct):
        return ""
    if atr_pct < 1.5:
        return "LOW"
    if atr_pct > 3.0:
        return "HIGH"
    return "NORMAL"


def position_hint_from_regime(reg):
    if reg == "LOW":
        return "BIG"
    if reg == "NORMAL":
        return "MEDIUM"
    if reg == "HIGH":
        return "SMALL"
    return ""


def confidence_advanced_window(early_score, trend_aligned, distribution, divergence, absorption, followthrough, stretched):
    es = (float(early_score) / 100.0) if early_score is not None else 0.0

    bonus = 0.00
    bonus += 0.03 if trend_aligned else 0.00
    bonus += 0.04 if absorption else 0.00
    bonus += 0.03 if followthrough else 0.00

    penalty = 0.00
    penalty += 0.08 if distribution else 0.00
    penalty += 0.05 if divergence else 0.00
    penalty += 0.04 if stretched else 0.00

    out = np.clip(es + bonus - penalty, 0.0, 1.0)
    return round(float(out), 2)


def rank_score_window(early_score, trend_aligned, breakout_quality,
                      followthrough, absorption, distribution, divergence, stretched):
    es = float(early_score) if early_score is not None else 0.0
    score = (
        es
        + 5.0 * int(trend_aligned)
        + 5.0 * int(breakout_quality)
        + 3.0 * int(followthrough)
        + 3.0 * int(absorption)
        - 4.0 * int(distribution)
        - 3.0 * int(divergence)
        - 2.0 * int(stretched)
    )
    return round(float(score), 1)


def early_score_7d(g):
    g = g.copy()
    cond_ma = (g["MA7"] > g["MA10"]).astype(int)
    near_hh7 = (g["Close"] >= 0.95 * g["HH7"]).astype(int)
    vol_ok = (g["Volume"] > g["VMA7"]).astype(int)
    wick_ok = (g["UpperWickPct"] < 0.35).astype(int)
    score = (0.30 * cond_ma + 0.30 * near_hh7 + 0.25 * vol_ok + 0.15 * wick_ok) * 100
    g["EarlyScore"] = score.round(0).clip(0, 100)
    return g


def early_score_15d(g):
    g = g.copy()
    cond_ma = (g["MA15"] > g["MA30"]).astype(int)
    near_hh15 = (g["Close"] >= 0.95 * g["HH15"]).astype(int)
    vol_ok = (g["Volume"] > g["VMA15"]).astype(int)
    wick_ok = (g["UpperWickPct"] < 0.35).astype(int)
    score = (0.35 * cond_ma + 0.30 * near_hh15 + 0.25 * vol_ok + 0.10 * wick_ok) * 100
    g["EarlyScore"] = score.round(0).clip(0, 100)
    return g


def signals_7d(g):
    g = early_score_7d(g).copy()
    g["Confidence"] = (g["EarlyScore"] / 100.0).round(2)

    buy = (g["EarlyScore"] >= 55) & (g["MA7"] > g["MA10"]) & (g["RSI14"] >= 50)
    sell = (g["MA7"] < g["MA10"]) & (g["RSI14"] < 45)

    g["Stock Signal"] = np.where(buy, "BUY", np.where(sell, "SELL", "HOLD"))

    g["BuyStrength"] = ""
    g.loc[(g["Stock Signal"] == "BUY") & (g["EarlyScore"] >= 75), "BuyStrength"] = "STRONG BUY"
    g.loc[(g["Stock Signal"] == "BUY") & (g["EarlyScore"] >= 65) & (g["EarlyScore"] < 75), "BuyStrength"] = "BUY"
    g.loc[(g["Stock Signal"] == "BUY") & (g["EarlyScore"] >= 55) & (g["EarlyScore"] < 65), "BuyStrength"] = "EARLY BUY"

    g["SellStrength"] = ""
    g.loc[(g["Stock Signal"] == "SELL") & (g["EarlyScore"] <= 25), "SellStrength"] = "STRONG SELL"
    g.loc[(g["Stock Signal"] == "SELL") & (g["EarlyScore"] > 25) & (g["EarlyScore"] <= 40), "SellStrength"] = "SELL"
    g.loc[(g["Stock Signal"] == "SELL") & (g["EarlyScore"] > 40), "SellStrength"] = "EARLY SELL"

    g["Bias"] = np.where(g["Stock Signal"] == "BUY", "UP",
                 np.where(g["Stock Signal"] == "SELL", "DOWN", "NEUTRAL"))
    return g


def signals_15d(g):
    g = early_score_15d(g).copy()
    g["Confidence"] = (g["EarlyScore"] / 100.0).round(2)

    buy = (g["EarlyScore"] >= 55) & (g["MA15"] > g["MA30"]) & (g["RSI14"] >= 50)
    sell = (g["MA15"] < g["MA30"]) & (g["RSI14"] < 45)

    g["Stock Signal"] = np.where(buy, "BUY", np.where(sell, "SELL", "HOLD"))

    g["BuyStrength"] = ""
    g.loc[(g["Stock Signal"] == "BUY") & (g["EarlyScore"] >= 75), "BuyStrength"] = "STRONG BUY"
    g.loc[(g["Stock Signal"] == "BUY") & (g["EarlyScore"] >= 65) & (g["EarlyScore"] < 75), "BuyStrength"] = "BUY"
    g.loc[(g["Stock Signal"] == "BUY") & (g["EarlyScore"] >= 55) & (g["EarlyScore"] < 65), "BuyStrength"] = "EARLY BUY"

    g["SellStrength"] = ""
    g.loc[(g["Stock Signal"] == "SELL") & (g["EarlyScore"] <= 25), "SellStrength"] = "STRONG SELL"
    g.loc[(g["Stock Signal"] == "SELL") & (g["EarlyScore"] > 25) & (g["EarlyScore"] <= 40), "SellStrength"] = "SELL"
    g.loc[(g["Stock Signal"] == "SELL") & (g["EarlyScore"] > 40), "SellStrength"] = "EARLY SELL"

    g["Bias"] = np.where(g["Stock Signal"] == "BUY", "UP",
                 np.where(g["Stock Signal"] == "SELL", "DOWN", "NEUTRAL"))
    return g


def build_reason(last, mode):
    parts = []
    if mode == "7D":
        if pd.notna(last.get("HH7")) and last["Close"] >= 0.97 * last["HH7"]:
            parts.append("near 7D breakout")
        if pd.notna(last.get("VMA7")) and last["Volume"] > last["VMA7"]:
            parts.append("volume > VMA7")
        if pd.notna(last.get("MA7")) and pd.notna(last.get("MA10")) and last["MA7"] > last["MA10"]:
            parts.append("MA7>MA10")
    else:
        if pd.notna(last.get("HH15")) and last["Close"] >= 0.97 * last["HH15"]:
            parts.append("near 15D breakout")
        if pd.notna(last.get("VMA15")) and last["Volume"] > last["VMA15"]:
            parts.append("volume > VMA15")
        if pd.notna(last.get("MA15")) and pd.notna(last.get("MA30")) and last["MA15"] > last["MA30"]:
            parts.append("MA15>MA30")

    if pd.notna(last.get("UpperWickPct")) and last["UpperWickPct"] < 0.30:
        parts.append("low upper wick")
    if last.get("StretchFlag") == "STRETCHED":
        parts.append("stretched")
    if last.get("VolPriceFlag_W") == "DIVERGENCE":
        parts.append("vol-price divergence")
    if bool(last.get("Distribution_W", False)):
        parts.append("distribution risk")
    if bool(last.get("Absorption_W", False)):
        parts.append("absorption")
    if bool(last.get("FollowThrough_W", False)):
        parts.append("follow-through")
    if bool(last.get("RetestBuy_W", False)):
        parts.append("breakout retest")

    if last.get("SentimentSignal") in ["ACCUMULATION", "DISTRIBUTION", "SHAKEOUT"]:
        parts.append(f"sentiment:{last.get('SentimentSignal').lower()}")

    return ", ".join(parts) if parts else "setup"


# =========================
# EXCEL HELPERS
# =========================
def autosize(ws, min_w=10, max_w=55):
    for col in ws.columns:
        mx = 0
        letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is None:
                continue
            mx = max(mx, len(str(cell.value)))
        ws.column_dimensions[letter].width = max(min_w, min(max_w, mx + 2))


def write_table(ws, df, name, header_color="1F4E79"):
    if df.empty:
        df = pd.DataFrame([["No data"]], columns=["Info"])

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    nrows, ncols = ws.max_row, ws.max_column
    ref = f"A1:{get_column_letter(ncols)}{nrows}"

    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tab)

    fill = PatternFill("solid", fgColor=header_color)
    font = Font(color="FFFFFF", bold=True)
    for c in ws[1]:
        c.fill = fill
        c.font = font
        c.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"
    autosize(ws)


def color_scale(ws, col_name):
    header = [c.value for c in ws[1]]
    if col_name not in header:
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
        if col_name in header:
            idx = header.index(col_name) + 1
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=idx).number_format = fmt


def apply_row_fill_by_value(ws, col_name, match_value, fill_hex):
    header = [c.value for c in ws[1]]
    if col_name not in header:
        return
    idx = header.index(col_name) + 1
    fill = PatternFill("solid", fgColor=fill_hex)
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=idx).value
        if v == match_value:
            for c in range(1, ws.max_column + 1):
                ws.cell(row=r, column=c).fill = fill


# =========================
# OUTPUT COLUMNS (Signals 7D/15D)
# =========================
FINAL_COLS = [
    "Date","Symbol","Sector","Company",
    "Stock Signal","BuyStrength","SellStrength","Bias",
    "EarlyScore","Confidence","Confidence_Advanced","RankScore",
    "Close","Volume",
    "RET1_%","RET2_%","RET4_%","RET7_%","RET10_%","RET15_%","RET20_%","RET30_%",
    "MA7","MA10","MA15","MA30",
    "HH7","HH15","LL7","LL15",
    "VMA7","VMA15",
    "UpperWickPct",
    "ATR7","ATR7_%","ATR15","ATR_%",
    "VolRegime","PositionSizeHint",
    "Close_Z20","Volume_Z20","StretchFlag",
    "Slope20","R2_20",
    "TrendHealth","VolExpansionFlag","FalseBreakoutFlag","FalseBreakoutScore",
    "VolPriceFlag_W","Distribution_W","Absorption_W","FollowThrough_W","RetestBuy_W",
    # Sentiment (scores + confirmation fields)
    "AccumulationScore","DistributionScore","ShakeoutScore",
    "AccHits4","DistHits4","ShakeHits3",
    "AccConfirmed","DistConfirmed","ShakeConfirmed",
    "SentimentSignal","SentimentReason",
    "Reason"
]


# =========================
# BUILD SIGNAL SHEETS (7D/15D)
# =========================
def build_sheet_df(data, sector_df, mode="7D"):
    rows = []

    if mode == "7D":
        win = 7
        z_win = 7
        slope_win = 7
    else:
        win = 15
        z_win = 15
        slope_win = 15

    W = str(win)

    for sym, g in data.groupby("Symbol"):
        if len(g) < 35:
            continue

        g = add_features(g)
        g = add_sentiment_detector(g)

        g["Close_Z20"] = zscore(g["Close"], z_win)
        g["Volume_Z20"] = zscore(g["Volume"], z_win)

        sl, r2 = slope_r2_last_n(g["Close"], slope_win)
        g["Slope20"] = np.nan
        g["R2_20"] = np.nan
        g.loc[g.index[-1], "Slope20"] = sl
        g.loc[g.index[-1], "R2_20"] = r2

        ve_flag = vol_expansion_flag(g["ATR7_%"], g["ATR_%"])

        if mode == "7D":
            g = signals_7d(g)
            reg = vol_regime_from_atr_pct(g.iloc[-1]["ATR7_%"])
            fb_flag, fb_score = false_breakout_metrics(
                g.iloc[-1]["Close"], g.iloc[-1]["HH7"], g.iloc[-1]["Volume"], g.iloc[-1]["VMA7"], g.iloc[-1]["UpperWickPct"]
            )
            trend_aligned_mode = bool(pd.notna(g.iloc[-1]["MA7"]) and pd.notna(g.iloc[-1]["MA10"]) and (g.iloc[-1]["MA7"] > g.iloc[-1]["MA10"]))
        else:
            g = signals_15d(g)
            reg = vol_regime_from_atr_pct(g.iloc[-1]["ATR_%"])
            fb_flag, fb_score = false_breakout_metrics(
                g.iloc[-1]["Close"], g.iloc[-1]["HH15"], g.iloc[-1]["Volume"], g.iloc[-1]["VMA15"], g.iloc[-1]["UpperWickPct"]
            )
            trend_aligned_mode = bool(pd.notna(g.iloc[-1]["MA15"]) and pd.notna(g.iloc[-1]["MA30"]) and (g.iloc[-1]["MA15"] > g.iloc[-1]["MA30"]))

        g = add_window_behavior(g, win)
        last = g.iloc[-1]

        th = trend_health(last.get("Slope20"), last.get("R2_20"))
        stretch = stretch_flag(last.get("Close_Z20"))
        stretched = (stretch == "STRETCHED")

        volpriceflag = last.get(f"VolPriceFlag_{W}", "NEUTRAL")
        distribution = bool(last.get(f"Distribution_{W}", False))
        absorption = bool(last.get(f"Absorption_{W}", False))
        followthrough = bool(last.get(f"FollowThrough_{W}", False))
        retestbuy = bool(last.get(f"RetestBuy_{W}", False))
        divergence = (volpriceflag == "DIVERGENCE")

        es = int(last["EarlyScore"]) if pd.notna(last["EarlyScore"]) else None
        conf_adv = confidence_advanced_window(
            es,
            trend_aligned_mode,
            distribution=distribution,
            divergence=divergence,
            absorption=absorption,
            followthrough=followthrough,
            stretched=stretched
        )

        reason = build_reason({
            **last.to_dict(),
            "StretchFlag": stretch,
            "VolPriceFlag_W": volpriceflag,
            "Distribution_W": distribution,
            "Absorption_W": absorption,
            "FollowThrough_W": followthrough,
            "RetestBuy_W": retestbuy,
            "SentimentSignal": last.get("SentimentSignal", "NEUTRAL"),
        }, mode=mode)

        rows.append({
            "Date": last["Date"].date() if pd.notna(last["Date"]) else None,
            "Symbol": sym,

            "Stock Signal": last["Stock Signal"],
            "BuyStrength": last["BuyStrength"],
            "SellStrength": last["SellStrength"],
            "Bias": last["Bias"],

            "EarlyScore": es,
            "Confidence": float(last["Confidence"]) if pd.notna(last["Confidence"]) else None,
            "Confidence_Advanced": conf_adv,
            "RankScore": None,

            "Close": float(last["Close"]),
            "Volume": float(last["Volume"]),

            "RET1_%": float(last["RET1"]) if pd.notna(last.get("RET1")) else None,
            "RET2_%": float(last["RET2"]) if pd.notna(last.get("RET2")) else None,
            "RET4_%": float(last["RET4"]) if pd.notna(last.get("RET4")) else None,

            "RET7_%": float(last["RET7"]) if pd.notna(last.get("RET7")) else None,
            "RET10_%": float(last["RET10"]) if pd.notna(last.get("RET10")) else None,
            "RET15_%": float(last["RET15"]) if pd.notna(last.get("RET15")) else None,
            "RET20_%": float(last["RET20"]) if pd.notna(last.get("RET20")) else None,
            "RET30_%": float(last["RET30"]) if pd.notna(last.get("RET30")) else None,

            "MA7": float(last["MA7"]) if pd.notna(last["MA7"]) else None,
            "MA10": float(last["MA10"]) if pd.notna(last["MA10"]) else None,
            "MA15": float(last["MA15"]) if pd.notna(last["MA15"]) else None,
            "MA30": float(last["MA30"]) if pd.notna(last["MA30"]) else None,

            "HH7": float(last["HH7"]) if pd.notna(last["HH7"]) else None,
            "HH15": float(last["HH15"]) if pd.notna(last["HH15"]) else None,
            "LL7": float(last["LL7"]) if pd.notna(last["LL7"]) else None,
            "LL15": float(last["LL15"]) if pd.notna(last["LL15"]) else None,

            "VMA7": float(last["VMA7"]) if pd.notna(last["VMA7"]) else None,
            "VMA15": float(last["VMA15"]) if pd.notna(last["VMA15"]) else None,

            "UpperWickPct": float(last["UpperWickPct"]) if pd.notna(last["UpperWickPct"]) else None,

            "ATR7": float(last["ATR7"]) if pd.notna(last["ATR7"]) else None,
            "ATR7_%": float(last["ATR7_%"]) if pd.notna(last["ATR7_%"]) else None,
            "ATR15": float(last["ATR15"]) if pd.notna(last["ATR15"]) else None,
            "ATR_%": float(last["ATR_%"]) if pd.notna(last["ATR_%"]) else None,

            "VolRegime": reg,
            "PositionSizeHint": position_hint_from_regime(reg),

            "Close_Z20": float(last["Close_Z20"]) if pd.notna(last["Close_Z20"]) else None,
            "Volume_Z20": float(last["Volume_Z20"]) if pd.notna(last["Volume_Z20"]) else None,
            "StretchFlag": stretch,

            "Slope20": float(last.get("Slope20")) if pd.notna(last.get("Slope20")) else None,
            "R2_20": float(last.get("R2_20")) if pd.notna(last.get("R2_20")) else None,

            "TrendHealth": th,
            "VolExpansionFlag": bool(ve_flag),
            "FalseBreakoutFlag": bool(fb_flag),
            "FalseBreakoutScore": fb_score,

            "VolPriceFlag_W": volpriceflag,
            "Distribution_W": distribution,
            "Absorption_W": absorption,
            "FollowThrough_W": followthrough,
            "RetestBuy_W": retestbuy,

            # sentiment scores + confirmation fields
            "AccumulationScore": int(last.get("AccumulationScore")) if pd.notna(last.get("AccumulationScore")) else None,
            "DistributionScore": int(last.get("DistributionScore")) if pd.notna(last.get("DistributionScore")) else None,
            "ShakeoutScore": int(last.get("ShakeoutScore")) if pd.notna(last.get("ShakeoutScore")) else None,

            "AccHits4": int(last.get("AccHits4")) if pd.notna(last.get("AccHits4")) else None,
            "DistHits4": int(last.get("DistHits4")) if pd.notna(last.get("DistHits4")) else None,
            "ShakeHits3": int(last.get("ShakeHits3")) if pd.notna(last.get("ShakeHits3")) else None,

            "AccConfirmed": bool(last.get("AccConfirmed", False)),
            "DistConfirmed": bool(last.get("DistConfirmed", False)),
            "ShakeConfirmed": bool(last.get("ShakeConfirmed", False)),

            "SentimentSignal": last.get("SentimentSignal", "NEUTRAL"),
            "SentimentReason": last.get("SentimentReason", ""),

            "Reason": reason,
        })

    df = pd.DataFrame(rows)
    df = df.merge(sector_df, on="Symbol", how="left")

    if df.empty:
        return df

    if mode == "7D":
        bq = ((df["Close"] >= df["HH7"]) & (df["Volume"] > df["VMA7"]) & (df["UpperWickPct"] < 0.30)).fillna(False)
        trend_aligned = (df["MA7"] > df["MA10"]).fillna(False)
    else:
        bq = ((df["Close"] >= df["HH15"]) & (df["Volume"] > df["VMA15"]) & (df["UpperWickPct"] < 0.30)).fillna(False)
        trend_aligned = (df["MA15"] > df["MA30"]).fillna(False)

    stretched = (df["StretchFlag"] == "STRETCHED").fillna(False)
    divergence = (df["VolPriceFlag_W"] == "DIVERGENCE").fillna(False)

    df["RankScore"] = [
        rank_score_window(es, ta, bqi, ft, ab, ds, dv, st)
        for es, ta, bqi, ft, ab, ds, dv, st in zip(
            df["EarlyScore"], trend_aligned, bq,
            df["FollowThrough_W"], df["Absorption_W"], df["Distribution_W"], divergence, stretched
        )
    ]

    df = df[[c for c in FINAL_COLS if c in df.columns]]
    df = df.sort_values(["RankScore", "EarlyScore", "Confidence_Advanced"], ascending=[False, False, False])
    return df


# =========================
# SWING SCANNER OUTPUT
# =========================
SWING_COLS = [
    "Date","Symbol","Sector","Company",
    "SwingCategory","SwingSignalType","SwingScore",
    "DaysFromBreakout","TrendAge",
    "AccumulationScore","DistributionScore","ShakeoutScore",
    "AccHits4","DistHits4","ShakeHits3",
    "AccConfirmed","DistConfirmed","ShakeConfirmed",
    "SentimentSignal",
    "Close","VWAP","VWAP_Dist%","Volume","ValueTraded","VolSpike20",
    "EMA10","EMA20","EMA50","RSI14","ATR14","Range%","ClosePos","Body%",
    "Entry","Stop","T1_1.5R","T2_2.5R","T3_3R",
    "ExpectedHoldDays","TimeStopDay","MaxHoldDays",
    "SwingReason","SwingRank"
]


def build_swing_scanner_df(data, sector_df):
    rows = []

    for sym, g in data.groupby("Symbol"):
        g = g.sort_values("Date").copy()
        if len(g) < SWING_MIN_HISTORY:
            continue

        g = add_features(g)
        g = add_sentiment_detector(g)
        g = add_swing_features(g)
        g = compute_breakout_age_cols(g)

        last = g.iloc[-1]

        if pd.isna(last["Close"]) or pd.isna(last["VWAP"]) or pd.isna(last["Volume"]):
            continue

        value_traded = float(last["ValueTraded"]) if pd.notna(last["ValueTraded"]) else 0.0
        liq_ok = value_traded >= SWING_VALUE_TRADED_MIN

        trend_ok = pd.notna(last["EMA20"]) and (last["Close"] > last["EMA20"])
        trend_up = pd.notna(last["EMA50"]) and pd.notna(last["EMA20"]) and (last["EMA20"] > last["EMA50"])
        vwap_ok = bool(last["AboveVWAP"]) and (pd.isna(last["VWAP_Slope5"]) or last["VWAP_Slope5"] >= 0)

        rsi_ok = pd.notna(last["RSI14"]) and (SWING_RSI_MIN <= last["RSI14"] <= SWING_RSI_MAX)

        vol_spike = float(last["VolSpike20"]) if pd.notna(last["VolSpike20"]) else np.nan
        vol_ok_breakout = pd.notna(vol_spike) and (vol_spike >= SWING_VOL_SPIKE_BREAKOUT)
        vol_ok_compress = pd.notna(vol_spike) and (vol_spike >= SWING_VOL_SPIKE_COMPRESSION)

        closepos_ok = pd.notna(last["ClosePos"]) and (last["ClosePos"] >= 0.65)
        body_ok = pd.notna(last["BodyPct"]) and (last["BodyPct"] >= 35)

        hh20_prior = float(last["HH20_PRIOR"]) if pd.notna(last["HH20_PRIOR"]) else np.nan
        breakout_today = pd.notna(hh20_prior) and (last["Close"] > hh20_prior) and vol_ok_breakout and closepos_ok

        pullback = False
        if pd.notna(last["ATR14"]) and last["ATR14"] > 0 and pd.notna(last["EMA20"]):
            pullback = trend_up and (abs(last["Close"] - last["EMA20"]) <= 1.0 * last["ATR14"])

        compression = False
        if pd.notna(last["ATR14"]) and pd.notna(last["ATR14_5ago"]):
            atr_contract = last["ATR14"] < last["ATR14_5ago"]
            compression = trend_ok and atr_contract and vol_ok_compress and (pd.notna(last["ClosePos"]) and last["ClosePos"] >= 0.60)

        too_extended = pd.notna(last["VWAP_DistPct"]) and (last["VWAP_DistPct"] > 6.0)
        too_volatile = pd.notna(last["RangePct"]) and (last["RangePct"] > 12.0)

        days_from_breakout = last["DaysFromBreakout"]
        trend_age = last["TrendAge"]

        score = 0
        reasons = []

        if trend_ok: score += 18; reasons.append("Close>EMA20")
        if trend_up: score += 8;  reasons.append("EMA20>EMA50")
        if vwap_ok:  score += 18; reasons.append("Close>VWAP & VWAP rising/flat")
        if rsi_ok:   score += 10; reasons.append("RSI in 50–70")
        if closepos_ok: score += 8; reasons.append("Strong close")
        if body_ok: score += 4; reasons.append("Good candle body")

        signal_type = ""
        exp_hold = ""

        if breakout_today:
            score += 22
            signal_type = "Breakout"
            exp_hold = "3–6"
            reasons.append("20D breakout + volume")
        elif pullback and trend_ok and vwap_ok:
            score += 18
            signal_type = "Pullback"
            exp_hold = "5–10"
            reasons.append("Pullback near EMA20 (ATR)")
        elif compression:
            score += 20
            signal_type = "Compression"
            exp_hold = "4–8"
            reasons.append("Compression→Expansion")
        else:
            if trend_ok and vwap_ok:
                score += 8
                signal_type = "Trend"
                exp_hold = "3–10"
                reasons.append("Trend continuation")

        if pd.notna(vol_spike) and vol_spike >= 1.2:
            score += 8; reasons.append("Volume>Avg20")
        elif pd.notna(last["VolAvg5"]) and pd.notna(last["VolAvg20"]) and (last["VolAvg5"] > last["VolAvg20"]):
            score += 6; reasons.append("VolAvg5>VolAvg20")

        if liq_ok:
            score += 10; reasons.append("Liquidity OK")
        else:
            score -= 18; reasons.append("Low liquidity")

        # ageing penalty
        if pd.notna(trend_age):
            if trend_age >= 12:
                score -= 10; reasons.append("Late-stage trend (age>=12)")
            elif trend_age >= 9:
                score -= 6; reasons.append("Maturing trend (age>=9)")

        # sentiment risk penalty (use confirmed distribution)
        dist_confirmed = bool(last.get("DistConfirmed", False))
        dist_score = int(last.get("DistributionScore")) if pd.notna(last.get("DistributionScore")) else 0
        if dist_confirmed:
            score -= 12; reasons.append("Distribution risk (confirmed)")
        elif dist_score >= 70:
            score -= 8; reasons.append("Distribution risk (1-day)")
        elif dist_score >= 55:
            score -= 4; reasons.append("Mild distribution risk")

        if too_extended:
            score -= 10; reasons.append("Extended vs VWAP")
        if too_volatile:
            score -= 8; reasons.append("High volatility")
        if pd.notna(last["RSI14"]) and last["RSI14"] > 75:
            score -= 6; reasons.append("RSI overbought")
        if pd.notna(last["RSI14"]) and last["RSI14"] < 45:
            score -= 4; reasons.append("Weak RSI")

        score = int(max(0, min(100, score)))

        category = "AVOID"
        if (score >= 70) and liq_ok and vwap_ok and trend_ok and (signal_type in ["Breakout","Pullback","Compression","Trend"]):
            category = "BUY"
        elif (score >= 50) and liq_ok:
            category = "WATCH"

        entry = float(last["Close"])
        stop = np.nan
        t1 = np.nan
        t2 = np.nan
        t3 = np.nan
        if pd.notna(last["ATR14"]) and last["ATR14"] > 0:
            atrv = float(last["ATR14"])
            stop = entry - 1.5 * atrv
            r = entry - stop
            t1 = entry + 1.5 * r
            t2 = entry + 2.5 * r
            t3 = entry + 3.0 * r

        rows.append({
            "Date": last["Date"].date() if pd.notna(last["Date"]) else None,
            "Symbol": sym,

            "SwingCategory": category,
            "SwingSignalType": signal_type,
            "SwingScore": score,

            "DaysFromBreakout": float(days_from_breakout) if pd.notna(days_from_breakout) else None,
            "TrendAge": float(trend_age) if pd.notna(trend_age) else None,

            "AccumulationScore": int(last.get("AccumulationScore")) if pd.notna(last.get("AccumulationScore")) else None,
            "DistributionScore": int(last.get("DistributionScore")) if pd.notna(last.get("DistributionScore")) else None,
            "ShakeoutScore": int(last.get("ShakeoutScore")) if pd.notna(last.get("ShakeoutScore")) else None,

            "AccHits4": int(last.get("AccHits4")) if pd.notna(last.get("AccHits4")) else None,
            "DistHits4": int(last.get("DistHits4")) if pd.notna(last.get("DistHits4")) else None,
            "ShakeHits3": int(last.get("ShakeHits3")) if pd.notna(last.get("ShakeHits3")) else None,

            "AccConfirmed": bool(last.get("AccConfirmed", False)),
            "DistConfirmed": bool(last.get("DistConfirmed", False)),
            "ShakeConfirmed": bool(last.get("ShakeConfirmed", False)),

            "SentimentSignal": last.get("SentimentSignal", "NEUTRAL"),

            "Close": float(last["Close"]),
            "VWAP": float(last["VWAP"]) if pd.notna(last["VWAP"]) else None,
            "VWAP_Dist%": float(last["VWAP_DistPct"]) if pd.notna(last["VWAP_DistPct"]) else None,

            "Volume": float(last["Volume"]),
            "ValueTraded": float(last["ValueTraded"]) if pd.notna(last["ValueTraded"]) else None,
            "VolSpike20": float(last["VolSpike20"]) if pd.notna(last["VolSpike20"]) else None,

            "EMA10": float(last["EMA10"]) if pd.notna(last["EMA10"]) else None,
            "EMA20": float(last["EMA20"]) if pd.notna(last["EMA20"]) else None,
            "EMA50": float(last["EMA50"]) if pd.notna(last["EMA50"]) else None,

            "RSI14": float(last["RSI14"]) if pd.notna(last["RSI14"]) else None,
            "ATR14": float(last["ATR14"]) if pd.notna(last["ATR14"]) else None,
            "Range%": float(last["RangePct"]) if pd.notna(last["RangePct"]) else None,
            "ClosePos": float(last["ClosePos"]) if pd.notna(last["ClosePos"]) else None,
            "Body%": float(last["BodyPct"]) if pd.notna(last["BodyPct"]) else None,

            "Entry": entry,
            "Stop": float(stop) if pd.notna(stop) else None,
            "T1_1.5R": float(t1) if pd.notna(t1) else None,
            "T2_2.5R": float(t2) if pd.notna(t2) else None,
            "T3_3R": float(t3) if pd.notna(t3) else None,

            "ExpectedHoldDays": exp_hold,
            "TimeStopDay": SWING_TIME_STOP_DAY,
            "MaxHoldDays": SWING_MAX_HOLD_DAYS,

            "SwingReason": "; ".join(reasons[:12]),
        })

    out = pd.DataFrame(rows)
    out = out.merge(sector_df, on="Symbol", how="left")

    if out.empty:
        return out, out, out, out

    out = out[[c for c in SWING_COLS if c in out.columns and c != "SwingRank"]]

    ranked = out.sort_values(["SwingScore","ValueTraded"], ascending=[False, False]).copy()
    ranked["SwingRank"] = np.arange(1, len(ranked) + 1)

    ranked = ranked.merge(sector_df, on="Symbol", how="left", suffixes=("", "_dup"))
    for c in list(ranked.columns):
        if c.endswith("_dup"):
            ranked.drop(columns=[c], inplace=True)

    ranked = ranked[[c for c in SWING_COLS if c in ranked.columns]]

    buy = ranked[ranked["SwingCategory"] == "BUY"].copy()
    watch = ranked[ranked["SwingCategory"] == "WATCH"].copy()
    avoid = ranked[ranked["SwingCategory"] == "AVOID"].copy()

    return buy, watch, avoid, ranked.head(SWING_OUT_TOP_N)


# =========================
# CIRCUIT WATCHLIST
# =========================
def build_circuit_watchlist(df):
    if df.empty:
        return df.copy(), df.copy()

    x = df.copy()

    x["UpperCircuitPrice"] = x["Close"] * 1.10
    x["LowerCircuitPrice"] = x["Close"] * 0.90

    ret1 = x["RET1_%"].fillna(0)
    ret2 = x["RET2_%"].fillna(0)
    ret4 = x["RET4_%"].fillna(0)
    vol = x["Volume"].fillna(0)
    vma7 = x["VMA7"].replace(0, np.nan).fillna(np.nan)
    vol_ratio7 = (vol / (vma7 + 1e-12)).replace([np.inf, -np.inf], 0).fillna(0)

    hh7 = x["HH7"].replace(0, np.nan)
    hh15 = x["HH15"].replace(0, np.nan)
    near_hh7 = ((x["Close"] / (hh7 + 1e-12)) >= 0.98).fillna(False)
    near_hh15 = ((x["Close"] / (hh15 + 1e-12)) >= 0.98).fillna(False)

    up_base = (
        18 * (ret1 >= 5).astype(int) +
        18 * (ret2 >= 7).astype(int) +
        12 * (ret4 >= 10).astype(int) +
        18 * (vol_ratio7 >= 1.5).astype(int) +
        12 * (near_hh7 | near_hh15).astype(int) +
        10 * (x["UpperWickPct"].fillna(1.0) <= 0.30).astype(int)
    ).clip(0, 100)

    penalty = (
        15 * x["Distribution_W"].fillna(False).astype(int) +
        10 * (x["StretchFlag"].fillna("") == "STRETCHED").astype(int) +
        10 * (x["VolPriceFlag_W"].fillna("") == "DIVERGENCE").astype(int) +
        8  * x["FalseBreakoutFlag"].fillna(False).astype(int) +
        8  * x.get("DistConfirmed", False).fillna(False).astype(int)
    )

    x["UpCircuitScore"] = (up_base - penalty).clip(0, 100).astype(int)

    down_score = (
        18 * (ret1 <= -5).astype(int) +
        18 * (ret2 <= -7).astype(int) +
        12 * (ret4 <= -10).astype(int) +
        16 * ((x["MA7"].fillna(0) < x["MA10"].fillna(0))).astype(int) +
        12 * (x["UpperWickPct"].fillna(0) >= 0.45).astype(int) +
        12 * (x["FalseBreakoutFlag"].fillna(False)).astype(int) +
        12 * (x["TrendHealth"].fillna("") == "DOWN").astype(int)
    ).clip(0, 100)

    x["DownCircuitScore"] = down_score.astype(int)

    x["CircuitPick"] = ""
    x.loc[(x["UpCircuitScore"] >= 60) & (x["UpCircuitScore"] > x["DownCircuitScore"]), "CircuitPick"] = "UP"
    x.loc[(x["DownCircuitScore"] >= 60) & (x["DownCircuitScore"] > x["UpCircuitScore"]), "CircuitPick"] = "DOWN"

    up_cols = [
        "Date","Symbol","Sector","Company","Close",
        "UpperCircuitPrice","UpCircuitScore",
        "RET1_%","RET2_%","RET4_%",
        "Volume","VMA7","UpperWickPct",
        "SentimentSignal","AccumulationScore","DistributionScore","ShakeoutScore",
        "AccHits4","DistHits4","ShakeHits3","DistConfirmed",
        "VolPriceFlag_W","Distribution_W","Absorption_W","FollowThrough_W","RetestBuy_W",
        "Stock Signal","BuyStrength","Bias","Reason",
        "CircuitPick"
    ]
    down_cols = [
        "Date","Symbol","Sector","Company","Close",
        "LowerCircuitPrice","DownCircuitScore",
        "RET1_%","RET2_%","RET4_%",
        "Volume","VMA7","UpperWickPct",
        "SentimentSignal","AccumulationScore","DistributionScore","ShakeoutScore",
        "AccHits4","DistHits4","ShakeHits3","DistConfirmed",
        "VolPriceFlag_W","Distribution_W","Absorption_W","FollowThrough_W","RetestBuy_W",
        "Stock Signal","SellStrength","Bias","Reason",
        "CircuitPick"
    ]

    up_df = x.sort_values(["UpCircuitScore","RankScore","EarlyScore"], ascending=[False, False, False])
    up_df = up_df[[c for c in up_cols if c in up_df.columns]].head(TOP_CIRCUIT_N)

    down_df = x.sort_values(["DownCircuitScore","RankScore","EarlyScore"], ascending=[False, False, False])
    down_df = down_df[[c for c in down_cols if c in down_df.columns]].head(TOP_CIRCUIT_N)

    return up_df, down_df


# =========================
# NEW: SENTIMENT TOP SHEET (confirmed logic)
# =========================
def build_sentiment_top(df7):
    """
    Fast scanner (confirmed):
      - Top ACCUMULATION: AccConfirmed True (rank by AccHits4)
      - Top DISTRIBUTION: DistConfirmed True (rank by DistHits4)
      - Top SHAKEOUT: ShakeConfirmed True (rank by ShakeHits3)
    """
    if df7 is None or df7.empty:
        return pd.DataFrame(columns=[
            "Category","Date","Symbol","Sector","Company",
            "SentimentSignal","AccumulationScore","DistributionScore","ShakeoutScore",
            "AccHits4","DistHits4","ShakeHits3",
            "Close","Volume","RankScore","EarlyScore","Reason","SentimentReason"
        ])

    base_cols = [
        "Date","Symbol","Sector","Company",
        "SentimentSignal","AccumulationScore","DistributionScore","ShakeoutScore",
        "AccHits4","DistHits4","ShakeHits3",
        "Close","Volume","RankScore","EarlyScore","Reason","SentimentReason"
    ]
    x = df7.copy()
    for c in base_cols + ["AccConfirmed","DistConfirmed","ShakeConfirmed"]:
        if c not in x.columns:
            x[c] = None

    acc = x[x["AccConfirmed"].fillna(False)].copy()
    acc = acc.sort_values(["AccHits4","RankScore","EarlyScore"], ascending=[False, False, False]).head(SENTIMENT_TOP_N_EACH)
    acc.insert(0, "Category", "ACCUMULATION")

    dist = x[x["DistConfirmed"].fillna(False)].copy()
    dist = dist.sort_values(["DistHits4","RankScore","EarlyScore"], ascending=[False, False, False]).head(SENTIMENT_TOP_N_EACH)
    dist.insert(0, "Category", "DISTRIBUTION")

    shak = x[x["ShakeConfirmed"].fillna(False)].copy()
    shak = shak.sort_values(["ShakeHits3","RankScore","EarlyScore"], ascending=[False, False, False]).head(SENTIMENT_TOP_N_EACH)
    shak.insert(0, "Category", "SHAKEOUT")

    out = pd.concat([acc, dist, shak], ignore_index=True)
    out = out[["Category"] + base_cols]
    return out


# =========================
# MAIN
# =========================
def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    data = load_latest_files(DATA_DIR, latest_n=LATEST_FILES_TO_LOAD)
    latest_dt = pd.to_datetime(data["Date"].max()).date()
    out_path = os.path.join(OUT_DIR, f"nepse_signals_{latest_dt}.xlsx")

    sector_df = load_sector_master(SECTOR_FILE)

    df7 = build_sheet_df(data, sector_df, mode="7D")
    df15 = build_sheet_df(data, sector_df, mode="15D")

    up_watch, down_watch = build_circuit_watchlist(df7)

    swing_buy, swing_watch, swing_avoid, swing_ranked = build_swing_scanner_df(data, sector_df)

    sentiment_top = build_sentiment_top(df7)

    wb = Workbook()

    fmt_map = {
        "RET1_%":"0.00","RET2_%":"0.00","RET4_%":"0.00",
        "RET7_%":"0.00","RET10_%":"0.00","RET15_%":"0.00","RET20_%":"0.00","RET30_%":"0.00",
        "Confidence":"0.00","Confidence_Advanced":"0.00",
        "UpperWickPct":"0.00",
        "ATR7":"0.00","ATR7_%":"0.00","ATR15":"0.00","ATR_%":"0.00",
        "Close_Z20":"0.00","Volume_Z20":"0.00",
        "Slope20":"0.00","R2_20":"0.00",
        "RankScore":"0.00",
        "FalseBreakoutScore":"0",
        "Close":"#,##0.00",
        "Volume":"#,##0",
        "MA7":"#,##0.00","MA10":"#,##0.00","MA15":"#,##0.00","MA30":"#,##0.00",
        "HH7":"#,##0.00","HH15":"#,##0.00","LL7":"#,##0.00","LL15":"#,##0.00",
        "VMA7":"#,##0","VMA15":"#,##0",
        "EarlyScore":"0",
        "UpperCircuitPrice":"#,##0.00",
        "LowerCircuitPrice":"#,##0.00",
        "UpCircuitScore":"0",
        "DownCircuitScore":"0",

        # sentiment
        "AccumulationScore":"0",
        "DistributionScore":"0",
        "ShakeoutScore":"0",
        "AccHits4":"0",
        "DistHits4":"0",
        "ShakeHits3":"0",

        # Swing formats
        "SwingScore":"0",
        "DaysFromBreakout":"0",
        "TrendAge":"0",
        "VWAP":"#,##0.00",
        "VWAP_Dist%":"0.00",
        "ValueTraded":"#,##0",
        "VolSpike20":"0.00",
        "EMA10":"#,##0.00","EMA20":"#,##0.00","EMA50":"#,##0.00",
        "RSI14":"0.00",
        "ATR14":"0.00",
        "Range%":"0.00",
        "ClosePos":"0.00",
        "Body%":"0.00",
        "Entry":"#,##0.00",
        "Stop":"#,##0.00",
        "T1_1.5R":"#,##0.00",
        "T2_2.5R":"#,##0.00",
        "T3_3R":"#,##0.00",
        "TimeStopDay":"0",
        "MaxHoldDays":"0",
        "SwingRank":"0",
    }

    # Sheet 1: Signals_7D
    wb.active.title = "Signals_7D"
    ws1 = wb["Signals_7D"]
    write_table(ws1, df7, "Signals7DTbl", header_color="1F4E79")
    number_format(ws1, fmt_map)
    for col in ["RankScore","Confidence_Advanced","EarlyScore","FalseBreakoutScore",
                "AccumulationScore","DistributionScore","ShakeoutScore",
                "AccHits4","DistHits4","ShakeHits3"]:
        color_scale(ws1, col)

    # Sheet 2: Signals_15D
    ws2 = wb.create_sheet("Signals_15D")
    write_table(ws2, df15, "Signals15DTbl", header_color="1F4E79")
    number_format(ws2, fmt_map)
    for col in ["RankScore","Confidence_Advanced","EarlyScore","FalseBreakoutScore",
                "AccumulationScore","DistributionScore","ShakeoutScore",
                "AccHits4","DistHits4","ShakeHits3"]:
        color_scale(ws2, col)

    # Sheet 3: Circuit_UP_Tomorrow
    ws3 = wb.create_sheet("Circuit_UP_Tomorrow")
    write_table(ws3, up_watch, "CircuitUpTbl", header_color="2E7D32")
    number_format(ws3, fmt_map)
    color_scale(ws3, "UpCircuitScore")
    apply_row_fill_by_value(ws3, "CircuitPick", "UP", "E8F5E9")

    # Sheet 4: Circuit_DOWN_Tomorrow
    ws4 = wb.create_sheet("Circuit_DOWN_Tomorrow")
    write_table(ws4, down_watch, "CircuitDownTbl", header_color="B71C1C")
    number_format(ws4, fmt_map)
    color_scale(ws4, "DownCircuitScore")
    apply_row_fill_by_value(ws4, "CircuitPick", "DOWN", "FFEBEE")

    # Sheet 5: Swing_BUY_3_10D
    ws5 = wb.create_sheet("Swing_BUY_3_10D")
    write_table(ws5, swing_buy, "SwingBuyTbl", header_color="2E7D32")
    number_format(ws5, fmt_map)
    for col in ["SwingScore","TrendAge","DaysFromBreakout","DistributionScore"]:
        color_scale(ws5, col)

    # Sheet 6: Swing_WATCH_3_10D
    ws6 = wb.create_sheet("Swing_WATCH_3_10D")
    write_table(ws6, swing_watch, "SwingWatchTbl", header_color="F9A825")
    number_format(ws6, fmt_map)
    for col in ["SwingScore","TrendAge","DistributionScore"]:
        color_scale(ws6, col)

    # Sheet 7: Swing_AVOID_3_10D
    ws7 = wb.create_sheet("Swing_AVOID_3_10D")
    write_table(ws7, swing_avoid, "SwingAvoidTbl", header_color="B71C1C")
    number_format(ws7, fmt_map)
    for col in ["SwingScore","DistributionScore"]:
        color_scale(ws7, col)

    # Sheet 8: Swing_RANKED_TOP
    ws8 = wb.create_sheet("Swing_RANKED_TOP")
    write_table(ws8, swing_ranked, "SwingRankedTbl", header_color="1565C0")
    number_format(ws8, fmt_map)
    for col in ["SwingScore","TrendAge","DaysFromBreakout","DistributionScore"]:
        color_scale(ws8, col)

    # Sheet 9: Sentiment_Top (confirmed)
    ws9 = wb.create_sheet("Sentiment_Top")
    write_table(ws9, sentiment_top, "SentimentTopTbl", header_color="6A1B9A")
    number_format(ws9, fmt_map)
    for col in ["AccHits4","DistHits4","ShakeHits3","RankScore","EarlyScore"]:
        color_scale(ws9, col)
    apply_row_fill_by_value(ws9, "Category", "ACCUMULATION", "E8F5E9")
    apply_row_fill_by_value(ws9, "Category", "DISTRIBUTION", "FFEBEE")
    apply_row_fill_by_value(ws9, "Category", "SHAKEOUT", "FFFDE7")

    wb.save(out_path)
    print(f"✅ Excel created (Signals + Circuits + Swing 3–10D + Sentiment_Top): {out_path}")


if __name__ == "__main__":
    main()
