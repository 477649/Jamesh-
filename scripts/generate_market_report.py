import os
import re
import numpy as np
import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


# =========================
# CONFIG
# =========================
# ✅ GitHub folder: outputs/Floor Sheet/floorsheet_YYYY-MM-DD.csv
INPUT_FOLDER = os.path.join("outputs", "Floor Sheet")

REPORT_FOLDER = os.path.join("outputs", "reports")
REPORT_NAME_PREFIX = "Market_Overview_Report"

# ✅ WINDOWS = number of AVAILABLE trading days (dates) from latest going backwards
WINDOWS = {"1D": 1, "7D": 7, "15D": 15, "1M": 30}

CRORE = 10_000_000

PRESSURE_HIGH_TOP3_SHARE = 0.35
PRESSURE_MED_TOP3_SHARE = 0.20

MOM_STRONG_PCT = 5.0
MOM_WEAK_PCT = 1.0

# price vs vwap tolerance => NEAR if within +/-0.25%
NEAR_TOL = 0.0025


# =========================
# EXCEL STYLING
# =========================
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# =========================
# Helpers
# =========================
def _safe_float(x):
    try:
        if pd.isna(x):
            return np.nan
        s = str(x).strip().replace(",", "")
        return float(s) if s else np.nan
    except Exception:
        return np.nan


def _parse_date_from_filename(fname: str):
    """
    Accepts:
      floorsheet_2026-01-10.csv
      floorsheet_2026-01-10.xlsx
      floorsheet_20260110.xlsx
      floorsheet_20260110.csv
    """
    base = os.path.basename(fname)

    m = re.search(r"(\d{4}-\d{2}-\d{2})", base)
    if m:
        return pd.to_datetime(m.group(1), errors="coerce")

    m = re.search(r"(\d{8})", base)
    if m:
        return pd.to_datetime(m.group(1), format="%Y%m%d", errors="coerce")

    return pd.NaT


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Supports BOTH floorsheet header styles:

    Style A:
      Transact No. | Symbol | Buyer Broker | Seller Broker | Quantity | Rate | Amount

    Style B (your GitHub output):
      Transact No. | Symbol | Buyer | Seller | Quantity | Rate | Amount
    """
    col_map = {c: re.sub(r"\s+", " ", str(c).strip().lower()) for c in df.columns}
    df = df.rename(columns=col_map)

    buyer_col = "buyer broker" if "buyer broker" in df.columns else ("buyer" if "buyer" in df.columns else None)
    seller_col = "seller broker" if "seller broker" in df.columns else ("seller" if "seller" in df.columns else None)

    missing = []
    if "symbol" not in df.columns:
        missing.append("symbol")
    if buyer_col is None:
        missing.append("buyer/buyer broker")
    if seller_col is None:
        missing.append("seller/seller broker")
    if "quantity" not in df.columns:
        missing.append("quantity")
    if "rate" not in df.columns:
        missing.append("rate")

    if missing:
        raise ValueError(
            "Missing required columns: " + ", ".join(missing) +
            "\nColumns found: " + ", ".join(df.columns.astype(str).tolist())
        )

    out = pd.DataFrame()
    out["Symbol"] = df["symbol"].astype(str).str.upper().str.strip()
    out["Buyer"] = df[buyer_col].astype(str).str.strip()
    out["Seller"] = df[seller_col].astype(str).str.strip()
    out["Qty"] = df["quantity"].apply(_safe_float).fillna(0).astype(float)
    out["Rate"] = df["rate"].apply(_safe_float).astype(float)

    if "amount" in df.columns:
        out["Amount"] = df["amount"].apply(_safe_float).astype(float)
    else:
        out["Amount"] = out["Qty"] * out["Rate"]

    out["Date"] = pd.NaT
    return out


def load_all_floorsheets(folder: str) -> pd.DataFrame:
    if not os.path.isdir(folder):
        raise ValueError(f"Input folder not found: {folder}")

    files = []
    for fn in sorted(os.listdir(folder)):
        low = fn.lower()
        if "floorsheet" not in low:
            continue
        if not (low.endswith(".csv") or low.endswith(".xlsx") or low.endswith(".xls")):
            continue
        files.append(os.path.join(folder, fn))

    if not files:
        raise ValueError(f"No floorsheet files found in: {folder}")

    all_rows = []
    for fpath in files:
        fn = os.path.basename(fpath)

        if fn.lower().endswith(".csv"):
            df = pd.read_csv(fpath, dtype=str, encoding_errors="ignore")
        else:
            df = pd.read_excel(fpath, dtype=str)

        df2 = _normalize_columns(df)

        d = _parse_date_from_filename(fn)
        if pd.isna(d):
            print(f"[WARN] Skipping {fn} (date not in filename).")
            continue

        df2["Date"] = pd.to_datetime(d).date()
        all_rows.append(df2)

    if not all_rows:
        raise ValueError("All files were skipped. Ensure filenames include date like floorsheet_YYYY-MM-DD.csv")

    big = pd.concat(all_rows, ignore_index=True)
    big = big.dropna(subset=["Symbol"])
    big = big[big["Qty"] > 0]
    big = big.dropna(subset=["Rate"])
    return big


def build_daily_symbol_prices(trades: pd.DataFrame) -> pd.DataFrame:
    t = trades.copy()
    t["Date"] = pd.to_datetime(t["Date"])
    t["_row"] = np.arange(len(t))

    daily = (
        t.sort_values(["Date", "Symbol", "_row"])
        .groupby(["Date", "Symbol"], as_index=False)
        .agg(
            Total_Qty=("Qty", "sum"),
            Total_Amt=("Amount", "sum"),
            Last_Price=("Rate", "last"),
        )
    )
    daily["VWAP"] = daily["Total_Amt"] / daily["Total_Qty"].replace(0, np.nan)
    daily["Date"] = pd.to_datetime(daily["Date"]).dt.date
    return daily


# ✅ ALWAYS pick last N AVAILABLE trading dates (from latest backwards)
def _rolling_dates(all_dates, window_days):
    """
    all_dates: sorted list of unique dates (python date)
    returns: last N dates from the latest available date
    """
    if not all_dates:
        return []
    n = int(window_days) if window_days else 0
    if n <= 0:
        return []
    return all_dates[-n:] if len(all_dates) >= n else all_dates


def compute_price_gain(daily_prices: pd.DataFrame, window_days: int) -> pd.DataFrame:
    dp = daily_prices.copy()
    dp["Date"] = pd.to_datetime(dp["Date"])
    all_dates = sorted(dp["Date"].dt.date.unique())
    if not all_dates:
        return pd.DataFrame(columns=["Symbol", "Today Price", f"Price_Gain_%_{window_days}D"])

    win_dates = _rolling_dates(all_dates, window_days)

    sub = dp[dp["Date"].dt.date.isin(win_dates)].copy()
    if sub.empty:
        return pd.DataFrame(columns=["Symbol", "Today Price", f"Price_Gain_%_{window_days}D"])

    first = (
        sub.sort_values(["Symbol", "Date"])
        .groupby("Symbol", as_index=False)
        .first()[["Symbol", "Last_Price"]]
        .rename(columns={"Last_Price": "Start_Price"})
    )
    last = (
        sub.sort_values(["Symbol", "Date"])
        .groupby("Symbol", as_index=False)
        .last()[["Symbol", "Last_Price"]]
        .rename(columns={"Last_Price": "Today Price"})
    )

    out = first.merge(last, on="Symbol", how="inner")
    out[f"Price_Gain_%_{window_days}D"] = (out["Today Price"] / out["Start_Price"] - 1.0) * 100.0
    return out


def compute_broker_net(trades: pd.DataFrame, window_days: int) -> pd.DataFrame:
    t = trades.copy()
    t["Date"] = pd.to_datetime(t["Date"]).dt.date
    all_dates = sorted(t["Date"].unique())
    if not all_dates:
        return pd.DataFrame(columns=["Symbol", "Broker", "Buy_Qty", "Buy_Amt", "Sell_Qty", "Sell_Amt", "Net_Qty", "Net_Amount"])

    latest = all_dates[-1]
    win_dates = _rolling_dates(all_dates, window_days)
    sub = t[t["Date"].isin(win_dates)].copy()

    buy = (
        sub.groupby(["Symbol", "Buyer"], as_index=False)
        .agg(Buy_Qty=("Qty", "sum"), Buy_Amt=("Amount", "sum"))
        .rename(columns={"Buyer": "Broker"})
    )
    sell = (
        sub.groupby(["Symbol", "Seller"], as_index=False)
        .agg(Sell_Qty=("Qty", "sum"), Sell_Amt=("Amount", "sum"))
        .rename(columns={"Seller": "Broker"})
    )

    m = buy.merge(sell, on=["Symbol", "Broker"], how="outer").fillna(0)
    m["Net_Qty"] = m["Buy_Qty"] - m["Sell_Qty"]
    m["Net_Amount"] = m["Buy_Amt"] - m["Sell_Amt"]
    m["Average Cost"] = np.where(m["Net_Qty"] > 0, m["Net_Amount"] / m["Net_Qty"].replace(0, np.nan), np.nan)

    latest_prices = (
        sub[sub["Date"] == latest]
        .sort_values(["Symbol"])
        .groupby("Symbol", as_index=False)
        .agg(Last_Price=("Rate", "last"))
    )
    m = m.merge(latest_prices, on="Symbol", how="left")
    return m


def pressure_label(top3_share: float) -> str:
    if top3_share >= PRESSURE_HIGH_TOP3_SHARE:
        return "HIGH"
    if top3_share >= PRESSURE_MED_TOP3_SHARE:
        return "MED"
    return "LOW"


def momentum_label(pct_gain: float) -> str:
    if pct_gain >= MOM_STRONG_PCT:
        return "STRONG"
    if pct_gain < MOM_WEAK_PCT:
        return "WEAK"
    return "MODERATE"


def trader_view(price_trend: str, buy_pressure: str, sell_pressure: str, momentum: str) -> str:
    if sell_pressure == "HIGH":
        return "🔴 AVOID"
    if price_trend == "UP" and buy_pressure in ("HIGH", "MED") and momentum in ("STRONG", "MODERATE"):
        return "🟢 BUY"
    return "🟡 WATCH"


def price_vs_vwap_label(last_price: float, vwap: float) -> str:
    if vwap is None or pd.isna(vwap) or vwap == 0 or last_price is None or pd.isna(last_price):
        return "NEAR"
    diff = (last_price - vwap) / vwap
    if abs(diff) <= NEAR_TOL:
        return "NEAR"
    return "ABOVE" if diff > 0 else "BELOW"


def build_market_overview_detailed(trades: pd.DataFrame, daily_prices: pd.DataFrame, window_days: int) -> pd.DataFrame:
    t = trades.copy()
    t["Date"] = pd.to_datetime(t["Date"]).dt.date
    all_dates = sorted(t["Date"].unique())
    if not all_dates:
        return pd.DataFrame()

    latest = all_dates[-1]
    win_dates = _rolling_dates(all_dates, window_days)
    sub = t[t["Date"].isin(win_dates)].copy()

    if sub.empty:
        return pd.DataFrame()

    sym_tr = sub.groupby("Symbol", as_index=False).agg(
        Total_Trades=("Symbol", "size"),
        Total_Qty=("Qty", "sum"),
        Total_Amt=("Amount", "sum"),
    )
    sym_tr["VWAP"] = sym_tr["Total_Amt"] / sym_tr["Total_Qty"].replace(0, np.nan)

    sub_sorted = sub.copy()
    sub_sorted["_row"] = np.arange(len(sub_sorted))
    last_px = (
        sub_sorted.sort_values(["Date", "Symbol", "_row"])
        .groupby("Symbol", as_index=False)
        .last()[["Symbol", "Rate"]]
        .rename(columns={"Rate": "Last_Price"})
    )

    gain = compute_price_gain(daily_prices, window_days)
    pct_col = f"Price_Gain_%_{window_days}D"

    bn = compute_broker_net(trades, window_days)

    top_buyer = (
        bn.sort_values(["Symbol", "Net_Amount"], ascending=[True, False])
        .groupby("Symbol", as_index=False)
        .first()[["Symbol", "Net_Amount", "Broker"]]
        .rename(columns={"Net_Amount": "Top_Net_Broker_Amt", "Broker": "Top_Net_Broker"})
    )
    top_seller = (
        bn.sort_values(["Symbol", "Net_Amount"], ascending=[True, True])
        .groupby("Symbol", as_index=False)
        .first()[["Symbol", "Net_Amount", "Broker"]]
        .rename(columns={"Net_Amount": "Top_Net_Seller_Amt", "Broker": "Top_Net_Seller_Broker"})
    )

    pressure_rows = []
    for sym, g in bn.groupby("Symbol"):
        total_qty = sym_tr.loc[sym_tr["Symbol"] == sym, "Total_Qty"]
        total_qty = float(total_qty.iloc[0]) if len(total_qty) else 0.0
        if total_qty <= 0:
            continue

        buyers = g[g["Net_Qty"] > 0].copy().sort_values("Net_Qty", ascending=False)
        sellers = g[g["Net_Qty"] < 0].copy()
        sellers["AbsNet"] = sellers["Net_Qty"].abs()
        sellers = sellers.sort_values("AbsNet", ascending=False)

        top3_buy_share = buyers.head(3)["Net_Qty"].sum() / total_qty if not buyers.empty else 0.0
        top3_sell_share = sellers.head(3)["AbsNet"].sum() / total_qty if not sellers.empty else 0.0
        pressure_rows.append({"Symbol": sym, "Top3_Buy_Share": top3_buy_share, "Top3_Sell_Share": top3_sell_share})

    pr = pd.DataFrame(pressure_rows)
    if pr.empty:
        pr = pd.DataFrame({"Symbol": sym_tr["Symbol"], "Top3_Buy_Share": 0.0, "Top3_Sell_Share": 0.0})

    out = (
        sym_tr.merge(last_px, on="Symbol", how="left")
        .merge(gain[["Symbol", "Today Price", pct_col]], on="Symbol", how="left")
        .merge(top_buyer, on="Symbol", how="left")
        .merge(top_seller, on="Symbol", how="left")
        .merge(pr, on="Symbol", how="left")
        .fillna(0)
    )

    out["Buy_Trades"] = out["Total_Trades"]
    out["Sell_Trades"] = out["Total_Trades"]

    out["Total_Amount (Cr)"] = out["Total_Amt"] / CRORE
    out["Top_Net_Broker_Amt (Cr)"] = out["Top_Net_Broker_Amt"] / CRORE
    out["Top_Net_Seller_Amt (Cr)"] = out["Top_Net_Seller_Amt"] / CRORE

    out["Price_vs_VWAP"] = out.apply(lambda r: price_vs_vwap_label(r["Last_Price"], r["VWAP"]), axis=1)
    out["Price Trend"] = np.where(out["Last_Price"] >= out["VWAP"].fillna(out["Last_Price"]), "UP", "DOWN")

    out["Buy_Pressure"] = out["Top3_Buy_Share"].apply(pressure_label)
    out["Sell_Pressure"] = out["Top3_Sell_Share"].apply(pressure_label)
    out["Momentum"] = out[pct_col].apply(lambda x: momentum_label(float(x) if pd.notna(x) else 0.0))
    out["Trader_View"] = out.apply(
        lambda r: trader_view(r["Price Trend"], r["Buy_Pressure"], r["Sell_Pressure"], r["Momentum"]),
        axis=1
    )

    sell_rank = {"HIGH": 0, "MED": 1, "LOW": 2}
    out["_sell_rank"] = out["Sell_Pressure"].map(sell_rank).fillna(9)
    out = out.sort_values(["_sell_rank", "Total_Qty"], ascending=[True, False]).reset_index(drop=True)
    out.insert(0, "Rank", np.arange(1, len(out) + 1))

    label = f"{window_days}D" if window_days != 30 else "1M"

    final = pd.DataFrame({
        "Rank": out["Rank"],
        "Symbol": out["Symbol"],
        f"Buy_Trades_{label}": out["Buy_Trades"].astype(int),
        f"Sell_Trades_{label}": out["Sell_Trades"].astype(int),
        "Total_Trades": out["Total_Trades"].astype(int),
        "Total_Qty": out["Total_Qty"].round(0).astype(int),
        "Total_Amount (Cr)": out["Total_Amount (Cr)"].round(3),
        f"VWAP_{label}": out["VWAP"].round(2),
        "Last_Price": pd.to_numeric(out["Last_Price"], errors="coerce").round(2),
        "Price_vs_VWAP": out["Price_vs_VWAP"],
        "Top_Net_Broker_Amt (Cr)": out["Top_Net_Broker_Amt (Cr)"].round(3),
        "Top_Net_Broker": out["Top_Net_Broker"],
        "Top_Net_Seller_Amt (Cr)": out["Top_Net_Seller_Amt (Cr)"].round(3),
        "Top_Net_Seller_Broker": out["Top_Net_Seller_Broker"],
        "Price Trend": out["Price Trend"],
        "Buy_Pressure": out["Buy_Pressure"],
        "Sell_Pressure": out["Sell_Pressure"],
        "Momentum": out["Momentum"],
        "Trader_View": out["Trader_View"],
    })

    return final


# =========================
# Price Movement + Broker Holding sheets
# =========================
def build_price_movement(daily_prices: pd.DataFrame) -> pd.DataFrame:
    g7 = compute_price_gain(daily_prices, 7)
    g15 = compute_price_gain(daily_prices, 15)
    g1m = compute_price_gain(daily_prices, 30)

    g7 = g7.rename(columns={"Today Price": "Today Price", "Price_Gain_%_7D": "Gain_7D"})
    g15 = g15.rename(columns={"Price_Gain_%_15D": "Gain_15D"})
    g1m = g1m.rename(columns={"Price_Gain_%_30D": "Gain_1M"})

    df = g7[["Symbol", "Today Price", "Gain_7D"]].merge(
        g15[["Symbol", "Gain_15D"]], on="Symbol", how="left"
    ).merge(
        g1m[["Symbol", "Gain_1M"]], on="Symbol", how="left"
    ).fillna(0)

    def fmt(today_price, pct):
        try:
            tp = float(today_price)
        except Exception:
            return ""
        try:
            p = float(pct)
        except Exception:
            p = 0.0
        return f"{tp:.0f}({p:.0f}%)"

    out = pd.DataFrame()
    out["Rank"] = np.arange(1, len(df) + 1)
    out["Symbol"] = df["Symbol"]
    out["Today Price"] = pd.to_numeric(df["Today Price"], errors="coerce").round(2)
    out["Price_Gain_%_7D"] = df.apply(lambda r: fmt(r["Today Price"], r["Gain_7D"]), axis=1)
    out["Price_Gain_%_15D"] = df.apply(lambda r: fmt(r["Today Price"], r["Gain_15D"]), axis=1)
    out["Price_Gain_%_1M"] = df.apply(lambda r: fmt(r["Today Price"], r["Gain_1M"]), axis=1)
    return out


def build_broker_holding(trades: pd.DataFrame) -> pd.DataFrame:
    bn_1d = compute_broker_net(trades, 1)
    bn_7d = compute_broker_net(trades, 7)
    bn_15d = compute_broker_net(trades, 15)
    bn_1m = compute_broker_net(trades, 30)

    base = bn_1d.copy()
    base["Net_Amt_Cr"] = base["Net_Amount"] / CRORE
    base["Abs_Net_Amt_Cr"] = base["Net_Amt_Cr"].abs()
    base = base.sort_values("Abs_Net_Amt_Cr", ascending=False).head(500)

    k = ["Symbol", "Broker"]
    out = base.merge(
        bn_7d[k + ["Net_Qty"]].rename(columns={"Net_Qty": "Net_Qty_7D"}), on=k, how="left"
    ).merge(
        bn_15d[k + ["Net_Qty"]].rename(columns={"Net_Qty": "Net_Qty_15D"}), on=k, how="left"
    ).merge(
        bn_1m[k + ["Net_Qty"]].rename(columns={"Net_Qty": "Net_Qty_1M"}), on=k, how="left"
    ).fillna(0)

    final = pd.DataFrame({
        "Symbol": out["Symbol"],
        "Broker": out["Broker"],
        "Buy_Qty": out["Buy_Qty"].round(0).astype(int),
        "Sell_Qty": out["Sell_Qty"].round(0).astype(int),
        "Net_Qty_1D": out["Net_Qty"].round(0).astype(int),
        "Net_Amount_1D (Cr)": (out["Net_Amount"] / CRORE).round(2),
        "Average Cost": pd.to_numeric(out["Average Cost"], errors="coerce").round(2),
        "Last_Price": pd.to_numeric(out["Last_Price"], errors="coerce").round(2),
        "Net_Qty_7D": out["Net_Qty_7D"].round(0).astype(int),
        "Net_Qty_15D": out["Net_Qty_15D"].round(0).astype(int),
        "Net_Qty_1M": out["Net_Qty_1M"].round(0).astype(int),
    })
    return final


# =========================
# Excel writing utilities
# =========================
def add_excel_table(ws, start_row, start_col, end_row, end_col, table_name):
    ref = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"
    tab = Table(displayName=table_name, ref=ref)
    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    tab.tableStyleInfo = style
    ws.add_table(tab)


def style_block(ws, start_row, start_col, end_row, end_col):
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_align = Alignment(horizontal="center", vertical="center")

    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(r, c)
            cell.border = BORDER
            if r == start_row:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_align
            else:
                cell.alignment = cell_align

    for c in range(start_col, end_col + 1):
        max_len = 10
        for r in range(start_row, end_row + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(c)].width = min(max_len + 2, 45)


def write_df_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame, table_name: str):
    ws = wb.create_sheet(sheet_name)
    ws["A1"] = sheet_name
    ws["A1"].font = Font(bold=True, size=14)

    start_row = 3
    start_col = 1

    if df is None or df.empty:
        ws["A3"] = "No data"
        return

    rows = list(dataframe_to_rows(df, index=False, header=True))
    for i, row in enumerate(rows, start=start_row):
        for j, val in enumerate(row, start=start_col):
            ws.cell(i, j, value=val)

    end_row = start_row + len(rows) - 1
    end_col = start_col + df.shape[1] - 1
    style_block(ws, start_row, start_col, end_row, end_col)
    add_excel_table(ws, start_row, start_col, end_row, end_col, table_name)


def write_market_overview_all_in_one_sheet(wb: Workbook, latest_date, df_1d, df_7d, df_15d, df_1m):
    ws = wb.create_sheet("Market_Overview_All")

    ws["A1"] = f"Market Overview — 1D / 7D / 15D / 1M (Trade Date: {latest_date})"
    ws["A1"].font = Font(bold=True, size=14)

    title_row = 3
    header_row = 4

    start_col = 1
    gap_cols = 1

    blocks = [
        ("Market_Overview_1D", df_1d, "TBL_OV_1D"),
        ("Market_Overview_7D", df_7d, "TBL_OV_7D"),
        ("Market_Overview_15D", df_15d, "TBL_OV_15D"),
        ("Market_Overview_1M", df_1m, "TBL_OV_1M"),
    ]

    for title, df, tname in blocks:
        ws.cell(row=title_row, column=start_col, value=title).font = Font(bold=True, size=12)

        if df is None or df.empty:
            ws.cell(row=header_row, column=start_col, value="No data").font = Font(bold=True)
            start_col += 2 + gap_cols
            continue

        rows = list(dataframe_to_rows(df, index=False, header=True))

        for i, row in enumerate(rows, start=header_row):
            for j, val in enumerate(row, start=start_col):
                ws.cell(i, j, value=val)

        end_row = header_row + len(rows) - 1
        end_col = start_col + df.shape[1] - 1

        style_block(ws, header_row, start_col, end_row, end_col)
        add_excel_table(ws, header_row, start_col, end_row, end_col, tname)

        start_col = end_col + 1 + gap_cols


# =========================
# MAIN
# =========================
def main():
    print(f"[1/5] Loading floorsheets from {INPUT_FOLDER} ...")
    trades = load_all_floorsheets(INPUT_FOLDER)

    print("[2/5] Building daily prices ...")
    daily_prices = build_daily_symbol_prices(trades)

    all_dates = sorted(pd.to_datetime(daily_prices["Date"]).dt.date.unique())
    latest_date = all_dates[-1] if all_dates else None
    print(f"[INFO] Latest trade date: {latest_date}")

    print("[3/5] Building Market Overview windows (latest -> back by available dates) ...")
    df_1d = build_market_overview_detailed(trades, daily_prices, WINDOWS["1D"]).head(200)
    df_7d = build_market_overview_detailed(trades, daily_prices, WINDOWS["7D"]).head(200)
    df_15d = build_market_overview_detailed(trades, daily_prices, WINDOWS["15D"]).head(200)
    df_1m = build_market_overview_detailed(trades, daily_prices, WINDOWS["1M"]).head(200)

    print("[4/5] Building Price Movement + Broker Holding ...")
    df_pm = build_price_movement(daily_prices).head(500)
    df_bh = build_broker_holding(trades).head(500)

    print("[5/5] Writing Excel report ...")
    os.makedirs(REPORT_FOLDER, exist_ok=True)
    out_path = os.path.join(REPORT_FOLDER, f"{REPORT_NAME_PREFIX}_{latest_date}.xlsx")

    wb = Workbook()
    wb.remove(wb.active)

    write_market_overview_all_in_one_sheet(wb, latest_date, df_1d, df_7d, df_15d, df_1m)
    write_df_sheet(wb, "Price Movement", df_pm, "TBL_PRICE_MOVE")
    write_df_sheet(wb, "Broker Holding", df_bh, "TBL_BROKER_HOLD")

    wb.save(out_path)
    print(f"[DONE] Saved: {out_path}")


if __name__ == "__main__":
    main()
