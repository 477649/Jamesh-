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
SECTOR_FILE = "outputs/Sector/sector_master.csv"  # Sector master file (Symbol -> Sector/Company)
OUT_DIR  = "outputs/PriceAction"
OUT_PATH = os.path.join(OUT_DIR, "nepse_signals.xlsx")
LATEST_FILES_TO_LOAD = 60


# =========================
# LOAD DATA
# =========================
def load_latest_files(folder, latest_n=60):
    files = sorted(glob.glob(os.path.join(folder, "SharePrice_*.csv")))
    if not files:
        raise FileNotFoundError(f"No SharePrice_*.csv files found in: {folder}")

    files = files[-latest_n:]  # latest trading days (by filename sort)

    rows = []
    for f in files:
        m = re.search(r"SharePrice_(\d{4}-\d{2}-\d{2})", os.path.basename(f))
        date = pd.to_datetime(m.group(1)) if m else pd.NaT

        df = pd.read_csv(f)
        df.columns = [c.strip() for c in df.columns]

        if "Close" not in df.columns and "LTP" in df.columns:
            df["Close"] = df["LTP"]
        if "Volume" not in df.columns:
            if "Vol" in df.columns:
                df["Volume"] = df["Vol"]
            elif "VOL" in df.columns:
                df["Volume"] = df["VOL"]

        df["Date"] = date

        required = ["Symbol", "Open", "High", "Low", "Close", "Volume"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"Missing columns {missing} in file: {f}")

        for c in ["Open", "High", "Low", "Close", "Volume"]:
            df[c] = df[c].astype(str).str.replace(",", "", regex=False).astype(float)

        rows.append(df[["Date", "Symbol", "Open", "High", "Low", "Close", "Volume"]])

    return pd.concat(rows, ignore_index=True).sort_values(["Symbol", "Date"]).reset_index(drop=True)


def load_sector_master(path):
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
        raise ValueError("sector_master.csv must contain a Sector column (Sector/Sectors)")

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


def slope_r2_last_n(close, n):
    y = close.tail(n).values
    if len(y) < n or np.isnan(y).any():
        return np.nan, np.nan
    x = np.arange(n)
    m, _ = np.polyfit(x, y, 1)
    r = np.corrcoef(x, y)[0, 1]
    return float(m), float(r * r)


def add_features(g):
    g = g.copy()

    for n in [7, 10, 15, 30]:
        g[f"MA{n}"] = g["Close"].rolling(n).mean()
        g[f"VMA{n}"] = g["Volume"].rolling(n).mean()

    g["HH7"]  = g["High"].rolling(7).max()
    g["LL7"]  = g["Low"].rolling(7).min()
    g["HH15"] = g["High"].rolling(15).max()
    g["LL15"] = g["Low"].rolling(15).min()

    for n in [7, 10, 15, 20, 30]:
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
def sector_signal_from_ret(sector_ret10):
    if pd.isna(sector_ret10):
        return "HOLD"
    if sector_ret10 > 0:
        return "BUY"
    if sector_ret10 < 0:
        return "SELL"
    return "HOLD"


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


def confidence_advanced(early_score, trend_aligned):
    base = (early_score / 100.0) if early_score is not None else 0.0
    bonus = 1.0 if trend_aligned else 0.0
    return round(0.85 * base + 0.15 * bonus, 2)


def rank_score(early_score, trend_aligned, breakout_quality, sector_rel_pos):
    es = float(early_score) if early_score is not None else 0.0
    return round(es + 5.0 * int(trend_aligned) + 5.0 * int(breakout_quality) + 2.0 * int(sector_rel_pos), 1)


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


def build_reason_7d(last):
    parts = []
    if pd.notna(last.get("HH7")) and last["Close"] >= 0.97 * last["HH7"]:
        parts.append("near 7D breakout")
    if pd.notna(last.get("VMA7")) and last["Volume"] > last["VMA7"]:
        parts.append("volume > VMA7")
    if pd.notna(last.get("MA7")) and pd.notna(last.get("MA10")) and last["MA7"] > last["MA10"]:
        parts.append("MA7>MA10")
    if pd.notna(last.get("UpperWickPct")) and last["UpperWickPct"] < 0.30:
        parts.append("low upper wick")
    if pd.notna(last.get("Close_Z20")) and last["Close_Z20"] > 2:
        parts.append("stretched")
    return ", ".join(parts) if parts else "short-term setup"


def build_reason_15d(last):
    parts = []
    if pd.notna(last.get("HH15")) and last["Close"] >= 0.97 * last["HH15"]:
        parts.append("near 15D breakout")
    if pd.notna(last.get("VMA15")) and last["Volume"] > last["VMA15"]:
        parts.append("volume > VMA15")
    if pd.notna(last.get("MA15")) and pd.notna(last.get("MA30")) and last["MA15"] > last["MA30"]:
        parts.append("MA15>MA30")
    if pd.notna(last.get("UpperWickPct")) and last["UpperWickPct"] < 0.30:
        parts.append("low upper wick")
    return ", ".join(parts) if parts else "trend/accumulation setup"


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


def write_table(ws, df, name):
    if df.empty:
        df = pd.DataFrame([["No data"]], columns=["Info"])

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    nrows, ncols = ws.max_row, ws.max_column
    ref = f"A1:{get_column_letter(ncols)}{nrows}"

    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tab)

    fill = PatternFill("solid", fgColor="1F4E79")
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


# =========================
# OUTPUT COLUMNS
# =========================
FINAL_COLS = [
    "Date","Symbol","Sector","Company",
    "Stock Signal","Sector Signal","BuyStrength","SellStrength","Bias",
    "EarlyScore","Confidence","Confidence_Advanced","RankScore",
    "Close","Volume",
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
    "Reason","Sector_RET10"
]

BEST_COLS = ["Mode"] + FINAL_COLS  # for BUY_BEST / SELL_BEST


# =========================
# BUILD SHEET DF
# =========================
def build_sheet_df(data, sector_df, mode="7D"):
    """
    mode="7D": uses 7-day stats for Close_Z20/Volume_Z20/Slope20/R2_20/StretchFlag
    mode="15D": uses 20-day stats for Close_Z20/Volume_Z20/Slope20/R2_20/StretchFlag
    """
    rows = []

    for sym, g in data.groupby("Symbol"):
        if len(g) < 35:
            continue

        g = add_features(g)

        # window rule
        if mode == "7D":
            z_win = 7
            slope_win = 7
        else:
            z_win = 20
            slope_win = 20

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
        else:
            g = signals_15d(g)

        last = g.iloc[-1]

        if mode == "7D":
            reason = build_reason_7d(last)
            reg = vol_regime_from_atr_pct(last["ATR7_%"])
            fb_flag, fb_score = false_breakout_metrics(
                last["Close"], last["HH7"], last["Volume"], last["VMA7"], last["UpperWickPct"]
            )
        else:
            reason = build_reason_15d(last)
            reg = vol_regime_from_atr_pct(last["ATR_%"])
            fb_flag, fb_score = false_breakout_metrics(
                last["Close"], last["HH15"], last["Volume"], last["VMA15"], last["UpperWickPct"]
            )

        trend_aligned = bool(pd.notna(last["MA15"]) and pd.notna(last["MA30"]) and (last["MA15"] > last["MA30"]))
        es = int(last["EarlyScore"]) if pd.notna(last["EarlyScore"]) else None
        conf_adv = confidence_advanced(es, trend_aligned)
        th = trend_health(last.get("Slope20"), last.get("R2_20"))
        stretch = stretch_flag(last.get("Close_Z20"))

        rows.append({
            "Date": last["Date"].date() if pd.notna(last["Date"]) else None,
            "Symbol": sym,

            "Stock Signal": last["Stock Signal"],
            "Sector Signal": None,  # fill later
            "BuyStrength": last["BuyStrength"],
            "SellStrength": last["SellStrength"],
            "Bias": last["Bias"],

            "EarlyScore": es,
            "Confidence": float(last["Confidence"]) if pd.notna(last["Confidence"]) else None,
            "Confidence_Advanced": conf_adv,
            "RankScore": None,  # fill later

            "Close": float(last["Close"]),
            "Volume": float(last["Volume"]),

            "RET7_%": float(last["RET7"]) if pd.notna(last["RET7"]) else None,
            "RET10_%": float(last["RET10"]) if pd.notna(last["RET10"]) else None,
            "RET15_%": float(last["RET15"]) if pd.notna(last["RET15"]) else None,
            "RET20_%": float(last["RET20"]) if pd.notna(last["RET20"]) else None,
            "RET30_%": float(last["RET30"]) if pd.notna(last["RET30"]) else None,

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

            "Reason": reason,
        })

    df = pd.DataFrame(rows)
    df = df.merge(sector_df, on="Symbol", how="left")

    sector_mom = (
        df.groupby("Sector", dropna=True)["RET10_%"]
        .mean()
        .rename("Sector_RET10")
        .reset_index()
    )
    df = df.merge(sector_mom, on="Sector", how="left")
    df["Sector Signal"] = df["Sector_RET10"].apply(sector_signal_from_ret)

    df["__SectorRelPos"] = ((df["RET10_%"] - df["Sector_RET10"]) > 0).fillna(False)

    if mode == "7D":
        bq = ((df["Close"] >= df["HH7"]) & (df["Volume"] > df["VMA7"]) & (df["UpperWickPct"] < 0.30)).fillna(False)
    else:
        bq = ((df["Close"] >= df["HH15"]) & (df["Volume"] > df["VMA15"]) & (df["UpperWickPct"] < 0.30)).fillna(False)

    trend_aligned = (df["MA15"] > df["MA30"]).fillna(False)
    df["RankScore"] = [
        rank_score(es, ta, bqi, sr)
        for es, ta, bqi, sr in zip(df["EarlyScore"], trend_aligned, bq, df["__SectorRelPos"])
    ]
    df = df.drop(columns=["__SectorRelPos"], errors="ignore")

    df = df[[c for c in FINAL_COLS if c in df.columns]]
    df = df.sort_values(["RankScore", "EarlyScore", "Confidence_Advanced"], ascending=[False, False, False])
    return df


# =========================
# BEST FILTERS (NOT all BUY/SELL)
# =========================
def build_best_buy_sell(df, mode_label):
    """
    Returns (buy_best, sell_best) from one sheet (df) using only available columns.
    """
    d = df.copy()

    # --- BEST BUY: strict quality ---
    buy_best = d[
        (d["Stock Signal"] == "BUY") &
        (d["EarlyScore"] >= 70) &
        (d["Confidence_Advanced"] >= 0.75) &
        (d["RankScore"] >= 80) &
        (d["TrendHealth"] == "GOOD") &
        (d["FalseBreakoutFlag"] == False) &
        (d["StretchFlag"] != "STRETCHED") &
        (d["VolExpansionFlag"] == True)
    ].copy()

    # --- BEST SELL: risk/exit candidates (any trigger) ---
    # Ensure NaNs don't break comparisons
    slope = pd.to_numeric(d["Slope20"], errors="coerce")
    r2 = pd.to_numeric(d["R2_20"], errors="coerce")
    fb_score = pd.to_numeric(d["FalseBreakoutScore"], errors="coerce").fillna(0)

    sell_best = d[
        (d["Stock Signal"] == "SELL") |
        (d["TrendHealth"] == "DOWN") |
        ((d["FalseBreakoutFlag"] == True) & (fb_score >= 60)) |
        ((slope < 0) & (r2 < 0.30)) |
        ((d["StretchFlag"] == "STRETCHED") & (d["Confidence"] < 0.50))
    ].copy()

    buy_best.insert(0, "Mode", mode_label)
    sell_best.insert(0, "Mode", mode_label)

    buy_best = buy_best[[c for c in BEST_COLS if c in buy_best.columns]]
    sell_best = sell_best[[c for c in BEST_COLS if c in sell_best.columns]]

    return buy_best, sell_best


# =========================
# MAIN
# =========================
def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    data = load_latest_files(DATA_DIR, latest_n=LATEST_FILES_TO_LOAD)
    sector_df = load_sector_master(SECTOR_FILE)

    # Main sheets
    df7 = build_sheet_df(data, sector_df, mode="7D")
    df15 = build_sheet_df(data, sector_df, mode="15D")

    # Best filtered sheets (NOT all buy/sell)
    buy7_best, sell7_best = build_best_buy_sell(df7, "7D")
    buy15_best, sell15_best = build_best_buy_sell(df15, "15D")

    buy_best_all = pd.concat([buy7_best, buy15_best], ignore_index=True)
    sell_best_all = pd.concat([sell7_best, sell15_best], ignore_index=True)

    # Sort best buy by RankScore desc; sell by RankScore asc (worst first is optional)
    buy_best_all = buy_best_all.sort_values(["Mode", "RankScore"], ascending=[True, False])
    sell_best_all = sell_best_all.sort_values(["Mode", "RankScore"], ascending=[True, True])

    # ---------- WRITE EXCEL ----------
    wb = Workbook()

    fmt_map = {
        "RET7_%":"0.00","RET10_%":"0.00","RET15_%":"0.00","RET20_%":"0.00","RET30_%":"0.00",
        "Confidence":"0.00","Confidence_Advanced":"0.00",
        "UpperWickPct":"0.00",
        "ATR7":"0.00","ATR7_%":"0.00","ATR15":"0.00","ATR_%":"0.00",
        "Close_Z20":"0.00","Volume_Z20":"0.00",
        "Slope20":"0.00","R2_20":"0.00",
        "RankScore":"0.00",
        "Sector_RET10":"0.00",
        "FalseBreakoutScore":"0.00",
        "Close":"#,##0.00",
        "Volume":"#,##0",
        "MA7":"#,##0.00","MA10":"#,##0.00","MA15":"#,##0.00","MA30":"#,##0.00",
        "HH7":"#,##0.00","HH15":"#,##0.00","LL7":"#,##0.00","LL15":"#,##0.00",
        "VMA7":"#,##0","VMA15":"#,##0",
        "EarlyScore":"0",
    }

    # Sheet 1: Signals_7D
    wb.active.title = "Signals_7D"
    ws1 = wb["Signals_7D"]
    write_table(ws1, df7, "Signals7DTbl")
    number_format(ws1, fmt_map)
    color_scale(ws1, "RankScore")
    color_scale(ws1, "Confidence_Advanced")
    color_scale(ws1, "EarlyScore")
    color_scale(ws1, "FalseBreakoutScore")

    # Sheet 2: Signals_15D
    ws2 = wb.create_sheet("Signals_15D")
    write_table(ws2, df15, "Signals15DTbl")
    number_format(ws2, fmt_map)
    color_scale(ws2, "RankScore")
    color_scale(ws2, "Confidence_Advanced")
    color_scale(ws2, "EarlyScore")
    color_scale(ws2, "FalseBreakoutScore")

    # Sheet 3: BUY_ALL (BEST BUY ONLY)
    ws3 = wb.create_sheet("BUY_ALL")
    write_table(ws3, buy_best_all, "BuyBestTbl")
    number_format(ws3, fmt_map)
    color_scale(ws3, "RankScore")
    color_scale(ws3, "Confidence_Advanced")
    color_scale(ws3, "EarlyScore")
    color_scale(ws3, "FalseBreakoutScore")

    # Sheet 4: SELL_ALL (BEST SELL ONLY)
    ws4 = wb.create_sheet("SELL_ALL")
    write_table(ws4, sell_best_all, "SellBestTbl")
    number_format(ws4, fmt_map)
    color_scale(ws4, "RankScore")
    color_scale(ws4, "Confidence_Advanced")
    color_scale(ws4, "EarlyScore")
    color_scale(ws4, "FalseBreakoutScore")

    wb.save(OUT_PATH)
    print(f"✅ Excel created with 4 sheets (BEST filters): {OUT_PATH}")


if __name__ == "__main__":
    main()
