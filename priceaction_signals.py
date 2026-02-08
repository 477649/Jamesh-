# priceaction_signals.py
# ✅ Updated:
# 1) Loads SharePrice_YYYY-MM-DD.csv by parsing dates (NOT string sort).
# 2) Uses latest available file date automatically (NOT system date).
# 3) VWAP handling: if VWAP missing in a file, fallback VWAP = Close (so script won't crash).
# 4) Adds SwingRank as LAST column in Swing outputs (rank by SwingScore, then ValueTraded).
# 5) Removed any reference to undefined add_trade_filter_rank() (your error).

import os, re, glob
import numpy as np
import pandas as pd

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter


# =========================
# PATHS / SETTINGS
# =========================
DATA_DIR = "outputs/sharesansar"                  # INPUT daily share price CSVs
SECTOR_FILE = "outputs/Sector/sector_master.csv"  # Symbol -> Sector/Company (used only for labeling)
OUT_DIR  = "outputs/PriceAction"
LATEST_FILES_TO_LOAD = 60

TOP_CIRCUIT_N = 25  # how many to show in circuit sheets

# ---- 3–10D Swing Scanner settings (NEPSE daily) ----
SWING_OUT_TOP_N = 80
SWING_MIN_HISTORY = 30                 # 30+ days workable; 60–90 best
SWING_VALUE_TRADED_MIN = 2_500_000     # liquidity filter (adjust based on your capital)
SWING_VOL_SPIKE_BREAKOUT = 1.30        # breakout confirmation
SWING_VOL_SPIKE_COMPRESSION = 1.40     # compression expansion
SWING_RSI_MIN = 50
SWING_RSI_MAX = 70
SWING_MAX_HOLD_DAYS = 10
SWING_TIME_STOP_DAY = 6


# =========================
# LOAD DATA (ROBUST BY DATE)
# =========================
def _dated_shareprice_files(folder: str):
    raw_files = glob.glob(os.path.join(folder, "SharePrice_*.csv"))
    if not raw_files:
        raise FileNotFoundError(f"No SharePrice_*.csv files found in: {folder}")

    dated = []
    for f in raw_files:
        m = re.search(r"SharePrice_(\d{4}-\d{2}-\d{2})", os.path.basename(f))
        if not m:
            continue
        dt = pd.to_datetime(m.group(1), errors="coerce")
        if pd.isna(dt):
            continue
        dated.append((dt, f))

    if not dated:
        raise ValueError(f"No valid dated SharePrice files found in: {folder}")

    dated.sort(key=lambda x: x[0])
    return dated


def load_latest_files(folder, latest_n=60):
    """
    ✅ Loads latest_n files by parsing date from filename and sorting by date.
    ✅ Uses those dates as the 'Date' column (not system date).
    ✅ VWAP fallback: if VWAP missing, use Close.
    """
    dated_files = _dated_shareprice_files(folder)
    dated_files = dated_files[-latest_n:]  # last N trading days available

    rows = []
    for date, f in dated_files:
        df = pd.read_csv(f)
        df.columns = [c.strip() for c in df.columns]

        # Normalize Close / Volume
        if "Close" not in df.columns and "LTP" in df.columns:
            df["Close"] = df["LTP"]

        if "Volume" not in df.columns:
            if "Vol" in df.columns:
                df["Volume"] = df["Vol"]
            elif "VOL" in df.columns:
                df["Volume"] = df["VOL"]

        # Normalize VWAP
        if "VWAP" not in df.columns:
            for c in ["vwap", "Vwap", "VWAP Price", "Daily VWAP"]:
                if c in df.columns:
                    df["VWAP"] = df[c]
                    break

        # ✅ VWAP fallback if still missing
        if "VWAP" not in df.columns:
            df["VWAP"] = df["Close"]

        df["Date"] = date

        required = ["Symbol", "Open", "High", "Low", "Close", "Volume", "VWAP"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"Missing columns {missing} in file: {f}")

        for c in ["Open", "High", "Low", "Close", "Volume", "VWAP"]:
            df[c] = df[c].astype(str).str.replace(",", "", regex=False).astype(float)

        rows.append(df[["Date", "Symbol", "Open", "High", "Low", "Close", "Volume", "VWAP"]])

    return (
        pd.concat(rows, ignore_index=True)
        .sort_values(["Symbol", "Date"])
        .reset_index(drop=True)
    )


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


def safe_div(a, b):
    # works for scalars/series
    return np.where((b == 0) | pd.isna(b), np.nan, a / b)


def add_features(g):
    g = g.copy()

    for n in [7, 10, 15, 30]:
        g[f"MA{n}"] = g["Close"].rolling(n).mean()
        g[f"VMA{n}"] = g["Volume"].rolling(n).mean()

    g["HH7"]  = g["High"].rolling(7).max()
    g["LL7"]  = g["Low"].rolling(7).min()
    g["HH15"] = g["High"].rolling(15).max()
    g["LL15"] = g["Low"].rolling(15).min()

    # Returns: these are computed from the Date-sorted series (NOT system date)
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
# 3–10D Swing Features (daily OHLCV + VWAP)
# =========================
def add_swing_features(g):
    g = g.copy()

    # EMAs for swing trend
    g["EMA10"] = ema(g["Close"], 10)
    g["EMA20"] = ema(g["Close"], 20)
    g["EMA50"] = ema(g["Close"], 50)

    # ATR14 for stops/targets
    tr = true_range(g["High"], g["Low"], g["Close"])
    g["ATR14"] = tr.rolling(14).mean()

    # Volume baselines
    g["VolAvg5"] = g["Volume"].rolling(5).mean()
    g["VolAvg20"] = g["Volume"].rolling(20).mean()
    g["VolSpike20"] = safe_div(g["Volume"], g["VolAvg20"])

    # Liquidity proxy
    g["ValueTraded"] = g["Close"] * g["Volume"]

    # Candle quality
    rng = (g["High"] - g["Low"]).replace(0, np.nan)
    g["RangePct"] = safe_div(rng, g["Close"]) * 100
    g["ClosePos"] = safe_div((g["Close"] - g["Low"]), rng)  # 0..1
    g["BodyPct"] = safe_div((g["Close"] - g["Open"]).abs(), rng) * 100

    # VWAP control (VWAP is daily; we only use last day VWAP + 5-day slope)
    g["AboveVWAP"] = g["Close"] > g["VWAP"]
    g["VWAP_DistPct"] = safe_div((g["Close"] - g["VWAP"]), g["VWAP"]) * 100
    g["VWAP_Slope5"] = g["VWAP"].diff(5)

    # 20D breakout level
    g["HH20"] = g["High"].rolling(20).max()

    # ATR contraction proxy for compression
    g["ATR14_5ago"] = g["ATR14"].shift(5)

    return g


# =========================
# WINDOW-WISE BEHAVIOR LOGIC (7D / 15D)
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
# OUTPUT COLUMNS (existing)
# =========================
FINAL_COLS = [
    "Date","Symbol","Sector","Company",
    "Stock Signal","BuyStrength","SellStrength","Bias",
    "EarlyScore","Confidence","Confidence_Advanced","RankScore",
    "Close","Volume",
    "RET1_%","RET2_%","RET4_%",
    "RET7_%","RET10_%","RET15_%","RET20_%","RET30_%",
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
    "Reason"
]


# =========================
# BUILD SHEET DF (existing 7D/15D logic)
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
        g = g.sort_values("Date").copy()
        if len(g) < 35:
            continue

        g = add_features(g)

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
            "RetestBuy_W": retestbuy
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
# 3–10D Swing Scanner (BUY / WATCH / AVOID / RANKED)
# =========================
SWING_COLS = [
    "Date","Symbol","Sector","Company",
    "SwingCategory","SwingSignalType","SwingScore",
    "Close","VWAP","VWAP_Dist%","Volume","ValueTraded","VolSpike20",
    "EMA10","EMA20","EMA50","RSI14","ATR14","Range%","ClosePos","Body%",
    "Entry","Stop","T1_1.5R","T2_2.5R","T3_3R",
    "ExpectedHoldDays","TimeStopDay","MaxHoldDays",
    "SwingReason",
    "SwingRank",  # ✅ last
]


def build_swing_scanner_df(data, sector_df):
    rows = []

    for sym, g in data.groupby("Symbol"):
        g = g.sort_values("Date").copy()
        if len(g) < SWING_MIN_HISTORY:
            continue

        g = add_features(g)         # RSI14 etc
        g = add_swing_features(g)   # EMA/VWAP/ATR/VolSpike etc

        last = g.iloc[-1]
        prior = g.iloc[:-1]

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

        # 20D high excluding today
        hh20_prior = prior["High"].tail(20).max() if len(prior) >= 20 else prior["High"].max()
        breakout = pd.notna(hh20_prior) and (last["Close"] > hh20_prior)

        pullback = False
        if pd.notna(last["ATR14"]) and last["ATR14"] > 0 and pd.notna(last["EMA20"]):
            pullback = trend_up and (abs(last["Close"] - last["EMA20"]) <= 1.0 * last["ATR14"])

        compression = False
        if pd.notna(last["ATR14"]) and pd.notna(last["ATR14_5ago"]):
            atr_contract = last["ATR14"] < last["ATR14_5ago"]
            compression = trend_ok and atr_contract and vol_ok_compress and (pd.notna(last["ClosePos"]) and last["ClosePos"] >= 0.60)

        too_extended = pd.notna(last["VWAP_DistPct"]) and (last["VWAP_DistPct"] > 6.0)
        too_volatile = pd.notna(last["RangePct"]) and (last["RangePct"] > 12.0)

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
        if breakout and vol_ok_breakout and closepos_ok and trend_ok and vwap_ok:
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

            "SwingReason": "; ".join(reasons[:10]),
        })

    out = pd.DataFrame(rows)
    if out.empty:
        out = out.assign(Sector=None, Company=None, SwingRank=None)
        return out, out, out, out

    out = out.merge(sector_df, on="Symbol", how="left")

    # ✅ Rank (1 = best) by SwingScore desc, then ValueTraded desc
    out = out.sort_values(["SwingScore", "ValueTraded"], ascending=[False, False])
    out["SwingRank"] = np.arange(1, len(out) + 1)

    # final order
    out = out[[c for c in SWING_COLS if c in out.columns]]

    buy = out[out["SwingCategory"] == "BUY"].copy()
    watch = out[out["SwingCategory"] == "WATCH"].copy()
    avoid = out[out["SwingCategory"] == "AVOID"].copy()

    ranked_top = out.head(SWING_OUT_TOP_N).copy()

    return buy, watch, avoid, ranked_top


# =========================
# CIRCUIT WATCHLIST (UNCHANGED)
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
        8  * x["FalseBreakoutFlag"].fillna(False).astype(int)
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
        "VolPriceFlag_W","Distribution_W","Absorption_W","FollowThrough_W","RetestBuy_W",
        "Stock Signal","BuyStrength","Bias","Reason",
        "CircuitPick"
    ]
    down_cols = [
        "Date","Symbol","Sector","Company","Close",
        "LowerCircuitPrice","DownCircuitScore",
        "RET1_%","RET2_%","RET4_%",
        "Volume","VMA7","UpperWickPct",
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
# MAIN
# =========================
def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    data = load_latest_files(DATA_DIR, latest_n=LATEST_FILES_TO_LOAD)

    # ✅ latest available trading date in loaded data (not system date)
    latest_dt = pd.to_datetime(data["Date"].max()).date()
    out_path = os.path.join(OUT_DIR, f"nepse_signals_{latest_dt}.xlsx")

    sector_df = load_sector_master(SECTOR_FILE)

    # existing signals
    df7 = build_sheet_df(data, sector_df, mode="7D")
    df15 = build_sheet_df(data, sector_df, mode="15D")

    # circuit sheets
    up_watch, down_watch = build_circuit_watchlist(df7 if not df7.empty else df15)

    # swing scanner
    swing_buy, swing_watch, swing_avoid, swing_ranked = build_swing_scanner_df(data, sector_df)

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

        # Swing
        "SwingScore":"0",
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
    color_scale(ws1, "RankScore")
    color_scale(ws1, "Confidence_Advanced")
    color_scale(ws1, "EarlyScore")
    color_scale(ws1, "FalseBreakoutScore")

    # Sheet 2: Signals_15D
    ws2 = wb.create_sheet("Signals_15D")
    write_table(ws2, df15, "Signals15DTbl", header_color="1F4E79")
    number_format(ws2, fmt_map)
    color_scale(ws2, "RankScore")
    color_scale(ws2, "Confidence_Advanced")
    color_scale(ws2, "EarlyScore")
    color_scale(ws2, "FalseBreakoutScore")

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

    # Swing sheets
    ws5 = wb.create_sheet("Swing_BUY_3_10D")
    write_table(ws5, swing_buy, "SwingBuyTbl", header_color="2E7D32")
    number_format(ws5, fmt_map)
    color_scale(ws5, "SwingScore")
    color_scale(ws5, "SwingRank")
    apply_row_fill_by_value(ws5, "SwingCategory", "BUY", "E8F5E9")

    ws6 = wb.create_sheet("Swing_WATCH_3_10D")
    write_table(ws6, swing_watch, "SwingWatchTbl", header_color="F9A825")
    number_format(ws6, fmt_map)
    color_scale(ws6, "SwingScore")
    color_scale(ws6, "SwingRank")

    ws7 = wb.create_sheet("Swing_AVOID_3_10D")
    write_table(ws7, swing_avoid, "SwingAvoidTbl", header_color="B71C1C")
    number_format(ws7, fmt_map)
    color_scale(ws7, "SwingScore")
    color_scale(ws7, "SwingRank")
    apply_row_fill_by_value(ws7, "SwingCategory", "AVOID", "FFEBEE")

    ws8 = wb.create_sheet("Swing_RANKED_TOP")
    write_table(ws8, swing_ranked, "SwingRankedTbl", header_color="1565C0")
    number_format(ws8, fmt_map)
    color_scale(ws8, "SwingScore")
    color_scale(ws8, "SwingRank")

    wb.save(out_path)
    print(f"✅ Excel created (Signals + Circuits + Swing 3–10D): {out_path}")


if __name__ == "__main__":
    main()
