import re
from pathlib import Path
import numpy as np
import pandas as pd

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import PatternFill, Font, Alignment

# =========================================================
# CONFIG
# =========================================================
BASE_DIR = Path(".")

FLOOR_DIR = BASE_DIR / "outputs" / "Floor Sheet"
OHLC_DIR = BASE_DIR / "outputs" / "sharesansar"
SECTOR_FILE = BASE_DIR / "outputs" / "Sector" / "sector_master.csv"

FLOOR_PATTERN = "floorsheet_*.csv"
OHLC_PATTERN = "SharePrice_*.csv"

OUTPUT_DIR = BASE_DIR / "model_outputs"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

OUTPUT_FILE = OUTPUT_DIR / "nepse_signals_short_long.xlsx"

DATE_REGEX = re.compile(r"(\d{4}-\d{2}-\d{2})")

# Liquidity filters
MIN_VOLUME_FS = 1000
MIN_NUM_TRADES = 5

# Short-term thresholds
SHORT_STRONG_BUY = 11
SHORT_BUY = 8
SHORT_HOLD = 5

# Long-term thresholds
LONG_STRONG_BUY = 10
LONG_BUY = 7
LONG_HOLD = 4


# =========================================================
# HELPERS
# =========================================================
def extract_date_from_filename(path: Path) -> pd.Timestamp:
    m = DATE_REGEX.search(path.name)
    if not m:
        raise ValueError(f"Could not extract date from filename: {path.name}")
    return pd.to_datetime(m.group(1))


def safe_read_csv(path: Path) -> pd.DataFrame:
    try:
        return pd.read_csv(path)
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="latin1")


def clean_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce"
    )


def standardize_symbol(df: pd.DataFrame, col: str = "Symbol") -> pd.DataFrame:
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip().str.upper()
    return df


def safe_div(a, b):
    return np.where((b != 0) & pd.notna(b), a / b, np.nan)


def join_reasons(reasons):
    return " | ".join(reasons) if reasons else "No strong confirmation"


# =========================================================
# LOAD SHARE PRICE DATA
# ONLY TAKE: Symbol, Open, High, Low, Close
# =========================================================
def load_ohlc_data(ohlc_dir: Path) -> pd.DataFrame:
    files = sorted(ohlc_dir.glob(OHLC_PATTERN))
    if not files:
        raise FileNotFoundError(f"No share price files found in {ohlc_dir}")

    frames = []
    required_cols = {"Symbol", "Open", "High", "Low", "Close"}

    for f in files:
        df = safe_read_csv(f)
        missing = required_cols - set(df.columns)
        if missing:
            print(f"[WARN] Skipping {f.name}; missing columns: {missing}")
            continue

        df = df.copy()
        df["Date"] = extract_date_from_filename(f)

        for c in ["Open", "High", "Low", "Close"]:
            df[c] = clean_numeric(df[c])

        df = standardize_symbol(df, "Symbol")
        df = df[["Date", "Symbol", "Open", "High", "Low", "Close"]]
        df = df.dropna(subset=["Date", "Symbol", "Close"])

        frames.append(df)

    if not frames:
        raise ValueError("No valid share price data loaded.")

    out = pd.concat(frames, ignore_index=True)
    out = out.sort_values(["Symbol", "Date"]).drop_duplicates(["Date", "Symbol"], keep="last")
    return out


# =========================================================
# LOAD FLOOR SHEET DATA
# =========================================================
def load_floor_sheet_data(floor_dir: Path) -> pd.DataFrame:
    files = sorted(floor_dir.glob(FLOOR_PATTERN))
    if not files:
        raise FileNotFoundError(f"No floor sheet files found in {floor_dir}")

    frames = []
    required_cols = {"Transact No.", "Symbol", "Buyer", "Seller", "Quantity", "Rate", "Amount"}

    for f in files:
        df = safe_read_csv(f)
        missing = required_cols - set(df.columns)
        if missing:
            print(f"[WARN] Skipping {f.name}; missing columns: {missing}")
            continue

        df = df.copy()
        df["Date"] = extract_date_from_filename(f)
        df = standardize_symbol(df, "Symbol")
        df["Buyer"] = df["Buyer"].astype(str).str.strip()
        df["Seller"] = df["Seller"].astype(str).str.strip()

        for c in ["Quantity", "Rate", "Amount"]:
            df[c] = clean_numeric(df[c])

        df = df.dropna(subset=["Date", "Symbol", "Quantity", "Amount"])
        frames.append(df)

    if not frames:
        raise ValueError("No valid floor sheet data loaded.")

    out = pd.concat(frames, ignore_index=True)
    out = out.sort_values(["Date", "Symbol"])
    return out


# =========================================================
# LOAD SECTOR MASTER
# =========================================================
def load_sector_master(sector_file: Path) -> pd.DataFrame:
    if not sector_file.exists():
        return pd.DataFrame(columns=["Symbol", "Sector", "Company"])

    sector = safe_read_csv(sector_file)
    required = {"Symbol", "Sector", "Company"}
    if not required.issubset(sector.columns):
        return pd.DataFrame(columns=["Symbol", "Sector", "Company"])

    sector = standardize_symbol(sector, "Symbol")
    sector = sector[["Symbol", "Sector", "Company"]].drop_duplicates("Symbol")
    return sector


# =========================================================
# FLOOR SHEET AGGREGATION
# VOLUME IS TAKEN FROM FLOOR SHEET:
# volume_fs = sum(Quantity)
# VWAP / WAP IS TAKEN FROM FLOOR SHEET:
# wap_fs = sum(Amount) / sum(Quantity)
# =========================================================
def compute_broker_concentration(group: pd.DataFrame) -> pd.Series:
    total_qty = group["Quantity"].sum()
    if total_qty <= 0:
        return pd.Series({
            "top_buyer_ratio": 0.0,
            "top_seller_ratio": 0.0,
            "top3_buyer_ratio": 0.0,
            "top3_seller_ratio": 0.0,
            "unique_buyers": 0,
            "unique_sellers": 0,
        })

    buyer_qty = group.groupby("Buyer")["Quantity"].sum().sort_values(ascending=False)
    seller_qty = group.groupby("Seller")["Quantity"].sum().sort_values(ascending=False)

    return pd.Series({
        "top_buyer_ratio": buyer_qty.iloc[0] / total_qty if len(buyer_qty) else 0.0,
        "top_seller_ratio": seller_qty.iloc[0] / total_qty if len(seller_qty) else 0.0,
        "top3_buyer_ratio": buyer_qty.head(3).sum() / total_qty,
        "top3_seller_ratio": seller_qty.head(3).sum() / total_qty,
        "unique_buyers": int(buyer_qty.shape[0]),
        "unique_sellers": int(seller_qty.shape[0]),
    })


def aggregate_floor_sheet_daily(floor: pd.DataFrame) -> pd.DataFrame:
    agg = (
        floor.groupby(["Date", "Symbol"], as_index=False)
        .agg(
            volume_fs=("Quantity", "sum"),
            turnover_fs=("Amount", "sum"),
            num_trades=("Transact No.", "count"),
            avg_rate_fs=("Rate", "mean"),
            max_trade_qty=("Quantity", "max"),
        )
    )

    # DAY-WISE VWAP / WAP FROM FLOOR SHEET
    agg["wap_fs"] = safe_div(agg["turnover_fs"], agg["volume_fs"])

    agg["avg_trade_size"] = safe_div(agg["volume_fs"], agg["num_trades"])

    broker = floor.groupby(["Date", "Symbol"]).apply(compute_broker_concentration).reset_index()
    out = agg.merge(broker, on=["Date", "Symbol"], how="left")
    out["imbalance"] = out["top_buyer_ratio"] - out["top_seller_ratio"]

    return out


# =========================================================
# MERGE DATA
# =========================================================
def build_master_dataset(ohlc: pd.DataFrame, floor_daily: pd.DataFrame, sector: pd.DataFrame) -> pd.DataFrame:
    df = ohlc.merge(floor_daily, on=["Date", "Symbol"], how="left")
    df = df.merge(sector, on="Symbol", how="left")
    return df.sort_values(["Symbol", "Date"]).reset_index(drop=True)


# =========================================================
# FEATURE ENGINEERING
# =========================================================
def add_features(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    g = df.groupby("Symbol", group_keys=False)

    # Moving averages from Close
    for w in [5, 10, 20, 35, 50]:
        df[f"ma{w}"] = g["Close"].transform(lambda s, w=w: s.rolling(w, min_periods=w).mean())

    # Returns from Close
    for w in [3, 5, 10, 20]:
        df[f"ret_{w}"] = g["Close"].transform(lambda s, w=w: s.pct_change(w))

    # Rolling averages for FLOOR SHEET VOLUME and other floor sheet data
    for w in [10, 20]:
        df[f"vol_fs_avg{w}"] = g["volume_fs"].transform(lambda s, w=w: s.rolling(w, min_periods=w).mean())
        df[f"turnover_fs_avg{w}"] = g["turnover_fs"].transform(lambda s, w=w: s.rolling(w, min_periods=w).mean())
        df[f"num_trades_avg{w}"] = g["num_trades"].transform(lambda s, w=w: s.rolling(w, min_periods=w).mean())
        df[f"imbalance_avg{w}"] = g["imbalance"].transform(lambda s, w=w: s.rolling(w, min_periods=w).mean())

    # Ratios using FLOOR SHEET VOLUME
    df["vol_ratio_10"] = safe_div(df["volume_fs"], df["vol_fs_avg10"])
    df["vol_ratio_20"] = safe_div(df["volume_fs"], df["vol_fs_avg20"])
    df["turnover_ratio_10"] = safe_div(df["turnover_fs"], df["turnover_fs_avg10"])
    df["turnover_ratio_20"] = safe_div(df["turnover_fs"], df["turnover_fs_avg20"])
    df["num_trades_ratio_10"] = safe_div(df["num_trades"], df["num_trades_avg10"])
    df["num_trades_ratio_20"] = safe_div(df["num_trades"], df["num_trades_avg20"])

    # WAP / VWAP strength from FLOOR SHEET WAP
    df["wap_diff"] = df["Close"] - df["wap_fs"]
    df["wap_strength"] = safe_div(df["wap_diff"], df["wap_fs"])

    # Breakouts
    df["high_5"] = g["High"].transform(lambda s: s.rolling(5, min_periods=5).max())
    df["high_20"] = g["High"].transform(lambda s: s.rolling(20, min_periods=20).max())
    df["prev_high_5"] = g["high_5"].shift(1)
    df["prev_high_20"] = g["high_20"].shift(1)
    df["breakout_5d"] = (df["Close"] > df["prev_high_5"]).astype(int)
    df["breakout_20d"] = (df["Close"] > df["prev_high_20"]).astype(int)

    # Multi-day bullish confirmation
    df["bull_day_short"] = (
        (df["Close"] > df["ma10"]) &
        (df["Close"] > df["wap_fs"]) &
        (df["ret_3"] > 0) &
        (df["imbalance"] > 0)
    ).astype(int)

    df["bull_day_long"] = (
        (df["Close"] > df["ma20"]) &
        (df["Close"] > df["wap_fs"]) &
        (df["ret_10"] > 0) &
        (df["imbalance"] > 0)
    ).astype(int)

    df["confirm_3d_short"] = g["bull_day_short"].transform(lambda s: s.rolling(3, min_periods=3).sum())
    df["confirm_5d_long"] = g["bull_day_long"].transform(lambda s: s.rolling(5, min_periods=5).sum())

    # Persistence features
    df["wap_strength_avg3"] = g["wap_strength"].transform(lambda s: s.rolling(3, min_periods=3).mean())
    df["wap_strength_avg5"] = g["wap_strength"].transform(lambda s: s.rolling(5, min_periods=5).mean())
    df["top_buy_persist_3"] = g["top_buyer_ratio"].transform(lambda s: s.rolling(3, min_periods=3).mean())
    df["top_sell_persist_3"] = g["top_seller_ratio"].transform(lambda s: s.rolling(3, min_periods=3).mean())

    return df


# =========================================================
# SHORT-TERM MODEL
# =========================================================
def short_term_score_and_reason(row):
    score = 0
    reasons = []
    negative_reasons = []

    # Liquidity from FLOOR SHEET VOLUME
    if pd.notna(row["volume_fs"]) and row["volume_fs"] >= MIN_VOLUME_FS:
        score += 1
        reasons.append("Good liquidity from floor-sheet volume")
    else:
        negative_reasons.append("Low floor-sheet volume")

    if pd.notna(row["num_trades"]) and row["num_trades"] >= MIN_NUM_TRADES:
        score += 1
        reasons.append("Enough trade activity")
    else:
        negative_reasons.append("Low number of trades")

    # Trend / momentum
    if pd.notna(row["ma5"]) and row["Close"] > row["ma5"]:
        score += 1
        reasons.append("Close is above MA5")
    else:
        negative_reasons.append("Close is below MA5")

    if pd.notna(row["ma10"]) and row["Close"] > row["ma10"]:
        score += 2
        reasons.append("Close is above MA10")
    else:
        negative_reasons.append("Close is below MA10")

    if pd.notna(row["ret_3"]) and row["ret_3"] > 0:
        score += 1
        reasons.append("3-day momentum is positive")
    else:
        negative_reasons.append("3-day momentum is weak")

    if pd.notna(row["ret_5"]) and row["ret_5"] > 0:
        score += 1
        reasons.append("5-day momentum is positive")
    else:
        negative_reasons.append("5-day momentum is weak")

    # Participation from FLOOR SHEET VOLUME
    if pd.notna(row["vol_ratio_10"]) and row["vol_ratio_10"] > 1.15:
        score += 1
        reasons.append("Floor-sheet volume is above 10-day average")
    else:
        negative_reasons.append("Floor-sheet volume is not strong vs 10-day average")

    if pd.notna(row["turnover_ratio_10"]) and row["turnover_ratio_10"] > 1.10:
        score += 1
        reasons.append("Turnover is above 10-day average")
    else:
        negative_reasons.append("Turnover is not strong vs 10-day average")

    if pd.notna(row["num_trades_ratio_10"]) and row["num_trades_ratio_10"] > 1.05:
        score += 1
        reasons.append("Trade count is increasing")
    else:
        negative_reasons.append("Trade count is not expanding")

    # WAP and broker behavior from FLOOR SHEET
    if pd.notna(row["wap_fs"]) and row["Close"] > row["wap_fs"]:
        score += 2
        reasons.append("Close is above floor-sheet WAP, buyers controlled the session")
    else:
        negative_reasons.append("Close is below WAP, sellers had better control")

    if pd.notna(row["imbalance"]) and row["imbalance"] > 0:
        score += 2
        reasons.append("Buyer broker concentration is stronger than seller broker concentration")
    else:
        negative_reasons.append("Broker imbalance is not positive")

    if pd.notna(row["top_buyer_ratio"]) and row["top_buyer_ratio"] > 0.25:
        score += 1
        reasons.append("Top buyer broker concentration is high")
    else:
        negative_reasons.append("Top buyer broker concentration is not strong")

    # Multi-day confirmation
    if pd.notna(row["confirm_3d_short"]) and row["confirm_3d_short"] >= 2:
        score += 2
        reasons.append("At least 2 of the last 3 sessions confirmed bullish conditions")
    else:
        negative_reasons.append("Recent sessions do not confirm strong short-term trend")

    if pd.notna(row["wap_strength_avg3"]) and row["wap_strength_avg3"] > 0:
        score += 1
        reasons.append("3-day WAP strength is positive")
    else:
        negative_reasons.append("3-day WAP strength is weak")

    if int(row.get("breakout_5d", 0)) == 1:
        score += 1
        reasons.append("Price broke above recent 5-day high")
    else:
        negative_reasons.append("No 5-day breakout")

    # Distribution override
    if (
        pd.notna(row["wap_fs"]) and row["wap_fs"] > row["Close"]
        and pd.notna(row["top_sell_persist_3"]) and row["top_sell_persist_3"] > 0.28
    ):
        signal = "SELL"
        insight = "SELL because price is below WAP and seller-broker pressure has persisted for 3 days"
        why = join_reasons(["Price below WAP", "Persistent seller concentration", "Distribution warning"])
        return score, signal, insight, why

    if score >= SHORT_STRONG_BUY:
        signal = "STRONG BUY"
        insight = "Strong short-term bullish setup with trend, momentum, floor-sheet volume, WAP, and broker confirmation aligned"
        why = join_reasons(reasons)
    elif score >= SHORT_BUY:
        signal = "BUY"
        insight = "Short-term bullish setup with multiple confirmations, but not the strongest possible"
        why = join_reasons(reasons)
    elif score >= SHORT_HOLD:
        signal = "HOLD"
        insight = "Mixed short-term setup. Some bullish signs exist, but confirmation is incomplete"
        why = join_reasons(reasons + negative_reasons[:3])
    else:
        signal = "SELL"
        insight = "Weak short-term structure. Trend, floor-sheet volume, participation, or broker control is not supportive"
        why = join_reasons(negative_reasons[:6])

    return score, signal, insight, why


# =========================================================
# LONG-TERM MODEL
# =========================================================
def long_term_score_and_reason(row):
    score = 0
    reasons = []
    negative_reasons = []

    # Liquidity from FLOOR SHEET VOLUME
    if pd.notna(row["volume_fs"]) and row["volume_fs"] >= MIN_VOLUME_FS:
        score += 1
        reasons.append("Good liquidity from floor-sheet volume")
    else:
        negative_reasons.append("Low floor-sheet volume")

    if pd.notna(row["num_trades"]) and row["num_trades"] >= MIN_NUM_TRADES:
        score += 1
        reasons.append("Enough trade activity")
    else:
        negative_reasons.append("Low number of trades")

    # Trend
    if pd.notna(row["ma20"]) and row["Close"] > row["ma20"]:
        score += 2
        reasons.append("Close is above MA20")
    else:
        negative_reasons.append("Close is below MA20")

    if pd.notna(row["ma35"]) and row["Close"] > row["ma35"]:
        score += 2
        reasons.append("Close is above MA35")
    else:
        negative_reasons.append("Close is below MA35")

    if pd.notna(row["ma50"]) and row["Close"] > row["ma50"]:
        score += 1
        reasons.append("Close is above MA50")
    else:
        negative_reasons.append("Close is below MA50")

    # Momentum
    if pd.notna(row["ret_10"]) and row["ret_10"] > 0:
        score += 2
        reasons.append("10-day momentum is positive")
    else:
        negative_reasons.append("10-day momentum is weak")

    if pd.notna(row["ret_20"]) and row["ret_20"] > 0:
        score += 1
        reasons.append("20-day momentum is positive")
    else:
        negative_reasons.append("20-day momentum is weak")

    # Participation from FLOOR SHEET VOLUME
    if pd.notna(row["vol_ratio_20"]) and row["vol_ratio_20"] > 1.10:
        score += 1
        reasons.append("Floor-sheet volume is above 20-day average")
    else:
        negative_reasons.append("Floor-sheet volume is not strong vs 20-day average")

    if pd.notna(row["turnover_ratio_20"]) and row["turnover_ratio_20"] > 1.05:
        score += 1
        reasons.append("Turnover is above 20-day average")
    else:
        negative_reasons.append("Turnover is not strong vs 20-day average")

    # WAP / broker persistence
    if pd.notna(row["wap_fs"]) and row["Close"] > row["wap_fs"]:
        score += 1
        reasons.append("Close is above floor-sheet WAP")
    else:
        negative_reasons.append("Close is below WAP")

    if pd.notna(row["imbalance_avg20"]) and row["imbalance_avg20"] > 0:
        score += 2
        reasons.append("20-day average broker imbalance is positive")
    else:
        negative_reasons.append("20-day broker imbalance is not positive")

    if pd.notna(row["top3_buyer_ratio"]) and pd.notna(row["top3_seller_ratio"]) and row["top3_buyer_ratio"] > row["top3_seller_ratio"]:
        score += 1
        reasons.append("Top 3 buyer brokers are stronger than top 3 seller brokers")
    else:
        negative_reasons.append("Top 3 buyer brokers are not dominating")

    # Multi-day confirmation
    if pd.notna(row["confirm_5d_long"]) and row["confirm_5d_long"] >= 3:
        score += 2
        reasons.append("At least 3 of the last 5 sessions confirmed bullish conditions")
    else:
        negative_reasons.append("Recent sessions do not confirm strong long-term trend")

    if pd.notna(row["wap_strength_avg5"]) and row["wap_strength_avg5"] > 0:
        score += 1
        reasons.append("5-day WAP strength is positive")
    else:
        negative_reasons.append("5-day WAP strength is weak")

    if int(row.get("breakout_20d", 0)) == 1:
        score += 1
        reasons.append("Price broke above recent 20-day high")
    else:
        negative_reasons.append("No 20-day breakout")

    # Distribution override
    if (
        pd.notna(row["wap_fs"]) and row["wap_fs"] > row["Close"]
        and pd.notna(row["top3_seller_ratio"]) and row["top3_seller_ratio"] > 0.55
    ):
        signal = "SELL"
        insight = "SELL because price is below WAP and top seller brokers dominate the structure"
        why = join_reasons(["Price below WAP", "Top 3 seller brokers dominate", "Distribution warning"])
        return score, signal, insight, why

    if score >= LONG_STRONG_BUY:
        signal = "STRONG BUY"
        insight = "Strong longer-term bullish structure with trend, momentum, floor-sheet volume, WAP, and broker persistence aligned"
        why = join_reasons(reasons)
    elif score >= LONG_BUY:
        signal = "BUY"
        insight = "Longer-term bullish structure is positive, though not at the strongest level"
        why = join_reasons(reasons)
    elif score >= LONG_HOLD:
        signal = "HOLD"
        insight = "Mixed longer-term structure. Trend is not fully broken, but confirmation is incomplete"
        why = join_reasons(reasons + negative_reasons[:3])
    else:
        signal = "SELL"
        insight = "Weak longer-term structure. Trend, floor-sheet volume, broker persistence, or momentum is not supportive"
        why = join_reasons(negative_reasons[:6])

    return score, signal, insight, why


# =========================================================
# APPLY MODELS
# =========================================================
def apply_models(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    short_results = out.apply(short_term_score_and_reason, axis=1)
    out["short_score"] = [x[0] for x in short_results]
    out["short_signal"] = [x[1] for x in short_results]
    out["short_insight"] = [x[2] for x in short_results]
    out["short_why"] = [x[3] for x in short_results]

    long_results = out.apply(long_term_score_and_reason, axis=1)
    out["long_score"] = [x[0] for x in long_results]
    out["long_signal"] = [x[1] for x in long_results]
    out["long_insight"] = [x[2] for x in long_results]
    out["long_why"] = [x[3] for x in long_results]

    return out


# =========================================================
# FINAL SHEET BUILDERS
# =========================================================
def latest_rows(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df.sort_values(["Symbol", "Date"])
        .groupby("Symbol", as_index=False)
        .tail(1)
        .reset_index(drop=True)
    )


def build_short_sheet(df_latest: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "Date", "Symbol", "Sector", "Company", "Open", "High", "Low", "Close",
        "volume_fs", "wap_fs", "num_trades",
        "ma5", "ma10", "ret_3", "ret_5",
        "vol_ratio_10", "turnover_ratio_10",
        "imbalance", "top_buyer_ratio", "top_seller_ratio",
        "confirm_3d_short", "wap_strength_avg3", "breakout_5d",
        "short_score", "short_signal", "short_insight", "short_why"
    ]
    out = df_latest[[c for c in cols if c in df_latest.columns]].copy()
    return out.sort_values(["short_score", "Symbol"], ascending=[False, True])


def build_long_sheet(df_latest: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "Date", "Symbol", "Sector", "Company", "Open", "High", "Low", "Close",
        "volume_fs", "wap_fs", "num_trades",
        "ma20", "ma35", "ma50", "ret_10", "ret_20",
        "vol_ratio_20", "turnover_ratio_20",
        "imbalance_avg20", "top3_buyer_ratio", "top3_seller_ratio",
        "confirm_5d_long", "wap_strength_avg5", "breakout_20d",
        "long_score", "long_signal", "long_insight", "long_why"
    ]
    out = df_latest[[c for c in cols if c in df_latest.columns]].copy()
    return out.sort_values(["long_score", "Symbol"], ascending=[False, True])


# =========================================================
# EXCEL HELPERS
# =========================================================
def autofit_worksheet(ws):
    for column_cells in ws.columns:
        max_length = 0
        col_letter = column_cells[0].column_letter

        for cell in column_cells:
            try:
                value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(value))
            except Exception:
                pass

        ws.column_dimensions[col_letter].width = min(max(max_length + 2, 10), 55)


def style_header(ws):
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    align = Alignment(horizontal="center", vertical="center")

    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = align


def add_excel_table(ws, df: pd.DataFrame, table_name: str):
    from openpyxl.utils import get_column_letter

    if df.empty:
        return

    last_row = len(df) + 1
    last_col = len(df.columns)
    ref = f"A1:{get_column_letter(last_col)}{last_row}"

    tab = Table(displayName=table_name, ref=ref)
    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )
    tab.tableStyleInfo = style
    ws.add_table(tab)


def write_df_to_sheet(ws, df: pd.DataFrame):
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)


def write_excel(short_df: pd.DataFrame, long_df: pd.DataFrame, output_file: Path):
    wb = Workbook()

    ws_short = wb.active
    ws_short.title = "Short_Term"

    ws_long = wb.create_sheet("Long_Term")

    write_df_to_sheet(ws_short, short_df)
    write_df_to_sheet(ws_long, long_df)

    style_header(ws_short)
    style_header(ws_long)

    add_excel_table(ws_short, short_df, "ShortTermTable")
    add_excel_table(ws_long, long_df, "LongTermTable")

    ws_short.freeze_panes = "A2"
    ws_long.freeze_panes = "A2"

    autofit_worksheet(ws_short)
    autofit_worksheet(ws_long)

    wb.save(output_file)
    print(f"Saved workbook: {output_file}")


# =========================================================
# MAIN
# =========================================================
def main():
    print("Loading share price data...")
    ohlc = load_ohlc_data(OHLC_DIR)

    print("Loading floor sheet data...")
    floor = load_floor_sheet_data(FLOOR_DIR)

    print("Loading sector master...")
    sector = load_sector_master(SECTOR_FILE)

    print("Aggregating floor sheet...")
    floor_daily = aggregate_floor_sheet_daily(floor)

    print("Building master dataset...")
    df = build_master_dataset(ohlc, floor_daily, sector)

    print("Adding features...")
    df = add_features(df)

    print("Applying short-term and long-term models...")
    df = apply_models(df)

    # Keep rows with at least enough history for short-term features
    valid_rows = df.groupby("Symbol").cumcount() + 1
    df_model = df.loc[valid_rows >= 20].copy()

    latest = latest_rows(df_model)

    short_sheet = build_short_sheet(latest)
    long_sheet = build_long_sheet(latest)

    write_excel(short_sheet, long_sheet, OUTPUT_FILE)

    print("\nTop Short-Term Signals:")
    print(short_sheet.head(10).to_string(index=False))

    print("\nTop Long-Term Signals:")
    print(long_sheet.head(10).to_string(index=False))


if __name__ == "__main__":
    main()
