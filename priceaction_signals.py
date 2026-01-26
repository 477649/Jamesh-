import os, re, glob
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter

# =========================
# PATHS (your requirement)
# =========================
DATA_DIR = "outputs/sharesansar"                  # INPUT daily share price CSVs
SECTOR_FILE = "outputs/Sector/sector_master.csv"  # Sector master file (Symbol -> Sector/Company)
OUT_DIR  = "outputs/PriceAction"                  # OUTPUT folder
OUT_PATH = os.path.join(OUT_DIR, "nepse_signals.xlsx")

# Load latest N trading-day files (not calendar filtering)
LATEST_FILES_TO_LOAD = 60

# Use 15-day context (you asked: no 20-day for key logic)
CONTEXT_DAYS = 15
# =========================


# ---------- LOAD DATA ----------
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

        # Map common columns
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

        # Numeric cleanup
        for c in ["Open", "High", "Low", "Close", "Volume"]:
            df[c] = df[c].astype(str).str.replace(",", "", regex=False).astype(float)

        rows.append(df[["Date", "Symbol", "Open", "High", "Low", "Close", "Volume"]])

    data = pd.concat(rows, ignore_index=True).sort_values(["Symbol", "Date"]).reset_index(drop=True)
    return data


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


# ---------- INDICATORS ----------
def rsi(close, n=14):
    d = close.diff()
    gain = d.clip(lower=0)
    loss = (-d).clip(lower=0)
    avg_gain = gain.rolling(n).mean()
    avg_loss = loss.rolling(n).mean()
    rs = avg_gain / (avg_loss + 1e-12)
    return 100 - (100 / (1 + rs))


def true_range(h, l, c):
    pc = c.shift(1)
    return pd.concat([(h - l), (h - pc).abs(), (l - pc).abs()], axis=1).max(axis=1)


def add_features(g):
    g = g.copy()

    # windows you use
    for n in [7, 10, 15, 20]:
        g[f"MA{n}"] = g["Close"].rolling(n).mean()
        g[f"VMA{n}"] = g["Volume"].rolling(n).mean()

    # breakout context (15D)
    g["HH15"] = g["High"].rolling(15).max()
    g["LL15"] = g["Low"].rolling(15).min()

    # candle / price-action
    rng = (g["High"] - g["Low"]).replace(0, np.nan)
    g["ClosePos"] = ((g["Close"] - g["Low"]) / rng).clip(0, 1)
    g["UpperWickPct"] = ((g["High"] - g[["Open", "Close"]].max(axis=1)) / rng).clip(0, 1)

    # momentum
    g["RSI14"] = rsi(g["Close"], 14)
    g["RET7"]  = g["Close"].pct_change(7) * 100
    g["RET10"] = g["Close"].pct_change(10) * 100
    g["RET15"] = g["Close"].pct_change(15) * 100
    g["RET20"] = g["Close"].pct_change(20) * 100

    # volatility / compression: compare TR7 vs TR15
    g["TR"] = true_range(g["High"], g["Low"], g["Close"])
    g["TR7"] = g["TR"].rolling(7).mean()
    g["TR15"] = g["TR"].rolling(CONTEXT_DAYS).mean()
    g["Compression"] = g["TR7"] / (g["TR15"] + 1e-12)  # <1 = compressed

    # distribution helpers
    g["RedDay"] = g["Close"] <= g["Open"]

    # ✅ UPDATED: VolSpike uses 15-day volume average (not 20)
    g["VolSpike"] = g["Volume"] > 1.5 * g["VMA15"]

    g["FailNewHigh"] = g["Close"] < g["HH15"]

    red_vol = (g["Volume"] * g["RedDay"].astype(int)).rolling(10).sum()
    tot_vol = g["Volume"].rolling(10).sum()
    g["RedVolShare10"] = (red_vol / (tot_vol + 1e-12)).clip(0, 1)

    return g


# ---------- ACCUMULATION + DISTRIBUTION + EarlyScore ----------
def early_score_bias(g):
    g = g.copy()

    # accumulation components
    hl = (g["Low"] > g["Low"].shift(1)).rolling(10).mean()
    vtrend = (g["VMA7"] / (g["VMA15"] + 1e-12)).clip(0, 2) / 2
    strong_close = g["ClosePos"].rolling(7).mean()
    comp_score = (1 - (g["Compression"].clip(0, 2) / 2)).clip(0, 1)

    near = (g["Close"] / (g["HH15"] + 1e-12)).clip(0, 1.5)
    near_score = ((near - 0.90) / 0.10).clip(0, 1)

    score01 = (0.28 * hl + 0.22 * vtrend + 0.18 * strong_close + 0.20 * comp_score + 0.12 * near_score).clip(0, 1)
    g["EarlyScore"] = (score01 * 100).round(0)

    # accumulation flag = 3/4
    acc1_higher_lows = (hl >= 0.6)
    acc2_vol_trend   = (g["VMA7"] > g["VMA15"])
    acc3_strongclose = (g["ClosePos"].rolling(7).mean() > 0.60)
    acc4_compress    = (g["Compression"] < 1.0)

    acc_count = acc1_higher_lows.astype(int) + acc2_vol_trend.astype(int) + acc3_strongclose.astype(int) + acc4_compress.astype(int)
    g["AccumulationCount"] = acc_count
    g["AccumulationFlag"] = acc_count >= 3

    # distribution rules (2+)
    dist1_highvol_weak = g["VolSpike"] & (g["Close"] <= g["Open"])
    dist2_reject = (g["UpperWickPct"].rolling(5).mean() > 0.45) & (g["ClosePos"].rolling(5).mean() < 0.45)
    dist3_fail_high = (g["FailNewHigh"].rolling(5).mean() > 0.70)
    dist4_red_vol_dom = (g["RedVolShare10"] > 0.55)

    dist_count = dist1_highvol_weak.astype(int) + dist2_reject.astype(int) + dist3_fail_high.astype(int) + dist4_red_vol_dom.astype(int)
    g["DistributionCount"] = dist_count
    g["DistributionFlag"] = dist_count >= 2

    up = g["AccumulationFlag"] & (g["EarlyScore"] >= 70) & (g["RSI14"] > 50)
    down = g["DistributionFlag"] | ((g["EarlyScore"] <= 35) & (g["RSI14"] < 45))
    g["Bias"] = np.where(up, "UP", np.where(down, "DOWN", "NEUTRAL"))
    g["Confidence"] = np.round((g["EarlyScore"] / 100).clip(0, 1), 2)

    # reasons flags
    g["HigherLows"] = acc1_higher_lows
    g["VolTrendUp"] = acc2_vol_trend
    g["StrongClose"] = acc3_strongclose
    g["CompressionGood"] = acc4_compress
    g["NearBreakout"] = g["Close"] >= 0.97 * g["HH15"]
    g["MAStack"] = (g["MA7"] > g["MA10"]) & (g["MA10"] > g["MA15"]) & (g["MA15"] > g["MA20"])

    return g


# ---------- STOCK SIGNAL (EarlyScore >= 50, RSI >= 50) ----------
def stock_signals(g):
    g = g.copy()

    buy = (
        (g["AccumulationFlag"]) &
        (g["EarlyScore"] >= 50) &
        (g["MA7"] > g["MA10"]) &
        (g["RSI14"] >= 50) &
        (g["DistributionFlag"] == False)
    )

    sell = (
        (g["DistributionFlag"]) |
        (g["Close"] < g["MA10"]) |
        (g["RSI14"] < 45)
    )

    g["Stock Signal"] = np.where(buy, "BUY", np.where(sell, "SELL", "HOLD"))

    g["BuyStrength"] = ""
    g.loc[(g["Stock Signal"] == "BUY") & (g["EarlyScore"] >= 75), "BuyStrength"] = "STRONG BUY"
    g.loc[(g["Stock Signal"] == "BUY") & (g["EarlyScore"] >= 60) & (g["EarlyScore"] < 75), "BuyStrength"] = "BUY"
    g.loc[(g["Stock Signal"] == "BUY") & (g["EarlyScore"] >= 50) & (g["EarlyScore"] < 60), "BuyStrength"] = "EARLY BUY"

    return g


def build_reason(row):
    parts = []
    if row.get("AccumulationFlag", False): parts.append(f"accumulation({int(row.get('AccumulationCount',0))}/4)")
    if row.get("DistributionFlag", False): parts.append(f"distribution({int(row.get('DistributionCount',0))}/4)")
    if row.get("HigherLows", False): parts.append("higher lows")
    if row.get("VolTrendUp", False): parts.append("volume trend up")
    if row.get("StrongClose", False): parts.append("strong closes")
    if row.get("CompressionGood", False): parts.append("range compression")
    if row.get("NearBreakout", False): parts.append("near 15D breakout")
    if row.get("MAStack", False): parts.append("MA stack 7>10>15>20")
    return ", ".join(parts) if parts else "mixed / no clear edge"


def sector_signal_from_ret(sector_ret10):
    if pd.isna(sector_ret10):
        return "HOLD"
    if sector_ret10 > 0:
        return "BUY"
    if sector_ret10 < 0:
        return "SELL"
    return "HOLD"


# ---------- EXCEL HELPERS ----------
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


def number_format(ws, mapping):
    header = [c.value for c in ws[1]]
    for col_name, fmt in mapping.items():
        if col_name in header:
            idx = header.index(col_name) + 1
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=idx).number_format = fmt


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


# ---------- MAIN ----------
def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    data = load_latest_files(DATA_DIR, latest_n=LATEST_FILES_TO_LOAD)
    sector_df = load_sector_master(SECTOR_FILE)

    latest_rows = []
    for sym, g in data.groupby("Symbol"):
        if len(g) < 25:
            continue
        g = stock_signals(early_score_bias(add_features(g)))
        last = g.iloc[-1]

        latest_rows.append({
            "Date": last["Date"].date() if pd.notna(last["Date"]) else None,
            "Symbol": sym,
            "Stock Signal": last["Stock Signal"],
            "Sector Signal": None,  # filled later
            "BuyStrength": last["BuyStrength"],
            "Bias": last["Bias"],
            "EarlyScore": int(last["EarlyScore"]) if not pd.isna(last["EarlyScore"]) else None,
            "Confidence": float(last["Confidence"]) if not pd.isna(last["Confidence"]) else None,
            "Close": float(last["Close"]),
            "Volume": float(last["Volume"]),
            "RET7_%": float(last["RET7"]) if not pd.isna(last["RET7"]) else None,
            "RET10_%": float(last["RET10"]) if not pd.isna(last["RET10"]) else None,
            "RET15_%": float(last["RET15"]) if not pd.isna(last["RET15"]) else None,
            "RET20_%": float(last["RET20"]) if not pd.isna(last["RET20"]) else None,
            "Reason": build_reason(last),
        })

    signals_df = pd.DataFrame(latest_rows)
    signals_df = signals_df.merge(sector_df, on="Symbol", how="left")

    sector_mom = (
        signals_df.groupby("Sector", dropna=True)["RET10_%"]
        .mean()
        .rename("Sector_RET10")
        .reset_index()
    )
    signals_df = signals_df.merge(sector_mom, on="Sector", how="left")
    signals_df["Sector Signal"] = signals_df["Sector_RET10"].apply(sector_signal_from_ret)

    final_cols = [
        "Date","Symbol","Stock Signal","Sector Signal",
        "BuyStrength","Bias","EarlyScore","Confidence","Close","Volume",
        "RET7_%","RET10_%","RET15_%","RET20_%","Reason",
        "Sector","Company","Sector_RET10"
    ]
    final_cols = [c for c in final_cols if c in signals_df.columns]
    signals_df = signals_df[final_cols]

    order = {"BUY":0,"HOLD":1,"SELL":2}
    signals_df["__r"] = signals_df["Stock Signal"].map(order).fillna(9)
    signals_df = signals_df.sort_values(["__r","EarlyScore","Confidence"], ascending=[True, False, False]).drop(columns="__r")

    wb = Workbook()
    wb.active.title = "Signals"
    ws = wb["Signals"]
    write_table(ws, signals_df, "SignalsTbl")

    number_format(ws, {
        "Close":"#,##0.00",
        "Volume":"#,##0",
        "Confidence":"0.00",
        "EarlyScore":"0",
        "RET7_%":"0.00",
        "RET10_%":"0.00",
        "RET15_%":"0.00",
        "RET20_%":"0.00",
        "Sector_RET10":"0.00",
    })
    color_scale(ws, "EarlyScore")
    color_scale(ws, "Confidence")

    wb.save(OUT_PATH)
    print(f"✅ Excel created: {OUT_PATH}")


if __name__ == "__main__":
    main()
