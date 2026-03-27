import re
from pathlib import Path
from datetime import datetime
import numpy as np
import pandas as pd

# =========================================================
# CONFIG
# =========================================================
BASE_DIR = Path(".")

FLOOR_DIR = BASE_DIR / "outputs" / "Floor Sheet"
OHLC_DIR = BASE_DIR / "outputs" / "sharesansar"
SECTOR_FILE = BASE_DIR / "outputs" / "Sector" / "sector_master.csv"

FLOOR_PATTERN = "floorsheet_*.csv"
OHLC_PATTERN = "SharePrice_*.csv"

# OUTPUT LOCATION YOU REQUESTED
OUTPUT_DIR = BASE_DIR / "outputs" / "PriceAction"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# fixed file name
OUTPUT_FILE = OUTPUT_DIR / "nepse_signals_report.html"

# if you want daily file instead, use this:
# OUTPUT_FILE = OUTPUT_DIR / f"nepse_signals_report_{datetime.now().strftime('%Y-%m-%d')}.html"

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


def format_num(x, decimals=2):
    if pd.isna(x):
        return ""
    return f"{x:.{decimals}f}"


def format_pct(x, decimals=2):
    if pd.isna(x):
        return ""
    return f"{x * 100:.{decimals}f}%"


def signal_badge(signal: str) -> str:
    signal = str(signal).strip().upper()
    if signal == "STRONG BUY":
        return '<span class="badge strong-buy">STRONG BUY</span>'
    if signal == "BUY":
        return '<span class="badge buy">BUY</span>'
    if signal == "HOLD":
        return '<span class="badge hold">HOLD</span>'
    if signal == "SELL":
        return '<span class="badge sell">SELL</span>'
    return f'<span class="badge neutral">{signal}</span>'


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
# VOLUME FROM FLOOR SHEET
# WAP / VWAP FROM FLOOR SHEET
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

    # day-wise WAP from floor sheet
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

    # Rolling averages for floor-sheet data
    for w in [10, 20]:
        df[f"vol_fs_avg{w}"] = g["volume_fs"].transform(lambda s, w=w: s.rolling(w, min_periods=w).mean())
        df[f"turnover_fs_avg{w}"] = g["turnover_fs"].transform(lambda s, w=w: s.rolling(w, min_periods=w).mean())
        df[f"num_trades_avg{w}"] = g["num_trades"].transform(lambda s, w=w: s.rolling(w, min_periods=w).mean())
        df[f"imbalance_avg{w}"] = g["imbalance"].transform(lambda s, w=w: s.rolling(w, min_periods=w).mean())

    # Ratios using floor-sheet volume
    df["vol_ratio_10"] = safe_div(df["volume_fs"], df["vol_fs_avg10"])
    df["vol_ratio_20"] = safe_div(df["volume_fs"], df["vol_fs_avg20"])
    df["turnover_ratio_10"] = safe_div(df["turnover_fs"], df["turnover_fs_avg10"])
    df["turnover_ratio_20"] = safe_div(df["turnover_fs"], df["turnover_fs_avg20"])
    df["num_trades_ratio_10"] = safe_div(df["num_trades"], df["num_trades_avg10"])
    df["num_trades_ratio_20"] = safe_div(df["num_trades"], df["num_trades_avg20"])

    # WAP strength from floor-sheet WAP
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

    if (
        pd.notna(row["top3_buyer_ratio"])
        and pd.notna(row["top3_seller_ratio"])
        and row["top3_buyer_ratio"] > row["top3_seller_ratio"]
    ):
        score += 1
        reasons.append("Top 3 buyer brokers are stronger than top 3 seller brokers")
    else:
        negative_reasons.append("Top 3 buyer brokers are not dominating")

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
# HTML REPORT
# =========================================================
def prepare_short_html_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["Date"] = pd.to_datetime(out["Date"]).dt.strftime("%Y-%m-%d")
    out["Open"] = out["Open"].map(format_num)
    out["High"] = out["High"].map(format_num)
    out["Low"] = out["Low"].map(format_num)
    out["Close"] = out["Close"].map(format_num)
    out["volume_fs"] = out["volume_fs"].map(lambda x: format_num(x, 0))
    out["wap_fs"] = out["wap_fs"].map(format_num)
    out["ma5"] = out["ma5"].map(format_num)
    out["ma10"] = out["ma10"].map(format_num)
    out["ret_3"] = out["ret_3"].map(format_pct)
    out["ret_5"] = out["ret_5"].map(format_pct)
    out["vol_ratio_10"] = out["vol_ratio_10"].map(format_num)
    out["turnover_ratio_10"] = out["turnover_ratio_10"].map(format_num)
    out["imbalance"] = out["imbalance"].map(format_num, na_action=None)
    out["top_buyer_ratio"] = out["top_buyer_ratio"].map(format_pct)
    out["top_seller_ratio"] = out["top_seller_ratio"].map(format_pct)
    out["wap_strength_avg3"] = out["wap_strength_avg3"].map(format_pct)
    out["short_signal"] = out["short_signal"].apply(signal_badge)
    return out


def prepare_long_html_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["Date"] = pd.to_datetime(out["Date"]).dt.strftime("%Y-%m-%d")
    out["Open"] = out["Open"].map(format_num)
    out["High"] = out["High"].map(format_num)
    out["Low"] = out["Low"].map(format_num)
    out["Close"] = out["Close"].map(format_num)
    out["volume_fs"] = out["volume_fs"].map(lambda x: format_num(x, 0))
    out["wap_fs"] = out["wap_fs"].map(format_num)
    out["ma20"] = out["ma20"].map(format_num)
    out["ma35"] = out["ma35"].map(format_num)
    out["ma50"] = out["ma50"].map(format_num)
    out["ret_10"] = out["ret_10"].map(format_pct)
    out["ret_20"] = out["ret_20"].map(format_pct)
    out["vol_ratio_20"] = out["vol_ratio_20"].map(format_num)
    out["turnover_ratio_20"] = out["turnover_ratio_20"].map(format_num)
    out["imbalance_avg20"] = out["imbalance_avg20"].map(format_num, na_action=None)
    out["top3_buyer_ratio"] = out["top3_buyer_ratio"].map(format_pct)
    out["top3_seller_ratio"] = out["top3_seller_ratio"].map(format_pct)
    out["wap_strength_avg5"] = out["wap_strength_avg5"].map(format_pct)
    out["long_signal"] = out["long_signal"].apply(signal_badge)
    return out


def dataframe_to_html_table(df: pd.DataFrame, table_id: str) -> str:
    return df.to_html(index=False, escape=False, table_id=table_id, classes="report-table")


def build_html_report(short_df: pd.DataFrame, long_df: pd.DataFrame) -> str:
    short_html_df = prepare_short_html_df(short_df)
    long_html_df = prepare_long_html_df(long_df)

    report_date = datetime.now().strftime("%Y-%m-%d %I:%M %p")

    short_counts = short_df["short_signal"].value_counts().to_dict() if not short_df.empty else {}
    long_counts = long_df["long_signal"].value_counts().to_dict() if not long_df.empty else {}

    short_summary = (
        f"STRONG BUY: {short_counts.get('STRONG BUY', 0)} | "
        f"BUY: {short_counts.get('BUY', 0)} | "
        f"HOLD: {short_counts.get('HOLD', 0)} | "
        f"SELL: {short_counts.get('SELL', 0)}"
    )

    long_summary = (
        f"STRONG BUY: {long_counts.get('STRONG BUY', 0)} | "
        f"BUY: {long_counts.get('BUY', 0)} | "
        f"HOLD: {long_counts.get('HOLD', 0)} | "
        f"SELL: {long_counts.get('SELL', 0)}"
    )

    short_table = dataframe_to_html_table(short_html_df, "short_term_table")
    long_table = dataframe_to_html_table(long_html_df, "long_term_table")

    html = f"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NEPSE Price Action Report</title>
<style>
    body {{
        font-family: Arial, sans-serif;
        background: #0f172a;
        color: #e2e8f0;
        margin: 0;
        padding: 24px;
    }}

    .container {{
        max-width: 1800px;
        margin: 0 auto;
    }}

    .header {{
        background: linear-gradient(135deg, #1e293b, #0f172a);
        border: 1px solid #334155;
        border-radius: 14px;
        padding: 20px 24px;
        margin-bottom: 24px;
    }}

    .title {{
        margin: 0;
        font-size: 28px;
        font-weight: 700;
        color: #f8fafc;
    }}

    .subtitle {{
        margin-top: 8px;
        color: #94a3b8;
        font-size: 14px;
    }}

    .summary-grid {{
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 16px;
        margin-bottom: 24px;
    }}

    .card {{
        background: #111827;
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 18px;
    }}

    .card h3 {{
        margin: 0 0 10px 0;
        color: #f8fafc;
        font-size: 18px;
    }}

    .card p {{
        margin: 0;
        color: #cbd5e1;
        line-height: 1.6;
    }}

    .section {{
        margin-bottom: 28px;
    }}

    .section h2 {{
        margin: 0 0 12px 0;
        font-size: 22px;
        color: #f8fafc;
    }}

    .table-wrap {{
        background: #111827;
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 14px;
        overflow-x: auto;
    }}

    table.report-table {{
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
        min-width: 1400px;
    }}

    table.report-table thead th {{
        background: #1f4e78;
        color: white;
        padding: 10px 8px;
        border: 1px solid #334155;
        text-align: left;
        position: sticky;
        top: 0;
    }}

    table.report-table tbody td {{
        padding: 8px;
        border: 1px solid #334155;
        color: #e5e7eb;
        vertical-align: top;
    }}

    table.report-table tbody tr:nth-child(even) {{
        background: #0b1220;
    }}

    table.report-table tbody tr:nth-child(odd) {{
        background: #111827;
    }}

    .badge {{
        display: inline-block;
        padding: 4px 10px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 700;
        letter-spacing: 0.2px;
        white-space: nowrap;
    }}

    .strong-buy {{
        background: #14532d;
        color: #dcfce7;
    }}

    .buy {{
        background: #166534;
        color: #dcfce7;
    }}

    .hold {{
        background: #92400e;
        color: #fef3c7;
    }}

    .sell {{
        background: #991b1b;
        color: #fee2e2;
    }}

    .neutral {{
        background: #475569;
        color: #f8fafc;
    }}

    .footer {{
        margin-top: 24px;
        color: #94a3b8;
        font-size: 12px;
        text-align: center;
    }}

    @media (max-width: 900px) {{
        .summary-grid {{
            grid-template-columns: 1fr;
        }}
    }}
</style>
</head>
<body>
<div class="container">

    <div class="header">
        <h1 class="title">NEPSE Price Action Report</h1>
        <div class="subtitle">Generated on {report_date}</div>
    </div>

    <div class="summary-grid">
        <div class="card">
            <h3>Short-Term Summary</h3>
            <p>{short_summary}</p>
        </div>
        <div class="card">
            <h3>Long-Term Summary</h3>
            <p>{long_summary}</p>
        </div>
    </div>

    <div class="section">
        <h2>Short-Term Signals</h2>
        <div class="table-wrap">
            {short_table}
        </div>
    </div>

    <div class="section">
        <h2>Long-Term Signals</h2>
        <div class="table-wrap">
            {long_table}
        </div>
    </div>

    <div class="footer">
        Auto-generated from share price OHLC and floor-sheet-based volume, WAP, and broker activity.
    </div>

</div>
</body>
</html>
"""
    return html


def save_html_report(html_content: str, output_file: Path):
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(html_content)


# =========================================================
# MAIN
# =========================================================
def main():
    ohlc = load_ohlc_data(OHLC_DIR)
    floor = load_floor_sheet_data(FLOOR_DIR)
    sector = load_sector_master(SECTOR_FILE)

    floor_daily = aggregate_floor_sheet_daily(floor)
    df = build_master_dataset(ohlc, floor_daily, sector)
    df = add_features(df)
    df = apply_models(df)

    # keep rows with at least enough history for short-term features
    valid_rows = df.groupby("Symbol").cumcount() + 1
    df_model = df.loc[valid_rows >= 20].copy()

    latest = latest_rows(df_model)

    short_sheet = build_short_sheet(latest)
    long_sheet = build_long_sheet(latest)

    html_report = build_html_report(short_sheet, long_sheet)
    save_html_report(html_report, OUTPUT_FILE)


if __name__ == "__main__":
    main()
