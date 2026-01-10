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
# CONFIG (Repo relative)
# =========================
INPUT_FOLDER = "outputs"                         # daily floorsheets are here
REPORT_FOLDER = os.path.join("outputs", "reports")
REPORT_NAME_PREFIX = "Market_Overview_Report"

WINDOWS = {"1D": 1, "7D": 7, "15D": 15, "1M": 30}

# Pressure thresholds
PRESSURE_HIGH_TOP3_SHARE = 0.35
PRESSURE_MED_TOP3_SHARE = 0.20

# Momentum thresholds
MOM_STRONG_PCT = 5.0
MOM_WEAK_PCT = 1.0


# =========================
# Styling
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
    except:
        return np.nan


def _parse_date_from_filename(fname: str):
    """
    Accepts:
      floorsheet_2026-01-10.xlsx
      floorsheet_20260110.xlsx
    """
    base = os.path.basename(fname)

    m = re.search(r"(\d{4}-\d{2}-\d{2})", base)
    if m:
        return pd.to_datetime(m.group(1), errors="coerce")

    m = re.search(r"(\d{8})", base)
    if m:
        return pd.to_datetime(m.group(1), format="%Y%m%d", errors="coerce")

    return pd.NaT


def _normalize_columns_exact(df: pd.DataFrame) -> pd.DataFrame:
    """
    Supports BOTH header styles:

    Style A:
      Transact No. | Symbol | Buyer Broker | Seller Broker | Quantity | Rate | Amount

    Style B (your GitHub output):
      Transact No. | Symbol | Buyer | Seller | Quantity | Rate | Amount
    """
    # Normalize column names
    col_map = {c: re.sub(r"\s+", " ", str(c).strip().lower()) for c in df.columns}
    df = df.rename(columns=col_map)

    # detect buyer/seller columns
    buyer_col = None
    seller_col = None

    if "buyer broker" in df.columns:
        buyer_col = "buyer broker"
    elif "buyer" in df.columns:
        buyer_col = "buyer"

    if "seller broker" in df.columns:
        seller_col = "seller broker"
    elif "seller" in df.columns:
        seller_col = "seller"

    required_missing = []
    if "symbol" not in df.columns:
        required_missing.append("symbol")
    if buyer_col is None:
        required_missing.append("buyer broker OR buyer")
    if seller_col is None:
        required_missing.append("seller broker OR seller")
    if "quantity" not in df.columns:
        required_missing.append("quantity")
    if "rate" not in df.columns:
        required_missing.append("rate")

    if required_missing:
        raise ValueError(
            "Missing required columns in floorsheet file: "
            + ", ".join(required_missing)
            + "\nColumns found: " + ", ".join(df.columns.astype(str).tolist())
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
        if not (low.endswith(".xlsx") or low.endswith(".xls") or low.endswith(".csv")):
            continue
        if "floorsheet" not in low:
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

        df2 = _normalize_columns_exact(df)

        d = _parse_date_from_filename(fn)
        if pd.isna(d):
            print(f"[WARN] Skipping {fn} (could not parse date from filename).")
            continue

        df2["Date"] = pd.to_datetime(d).date()
        all_rows.append(df2)

    if not all_rows:
        raise ValueError("All files were skipped. Check filenames include date like floorsheet_YYYY-MM-DD.xlsx")

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


def _rolling_dates(all_dates, end_date, window_days):
    dates = [d for d in all_dates if d <= end_date]
    return dates[-window_days:] if dates else []


def compute_price_gain(daily_prices: pd.DataFrame, window_days: int) -> pd.DataFrame:
    dp = daily_prices.copy()
    dp["Date"] = pd.to_datetime(dp["Date"])
    all_dates = sorted(dp["Date"].dt.date.unique())
    latest_date = max(all_dates)
    win_dates = _rolling_dates(all_dates, latest_date, window_days)

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
    latest = max(all_dates)
    win_dates = _rolling_dates(all_dates, latest, window_days)
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


def build_market_overview(trades: pd.DataFrame, daily_prices: pd.DataFrame, window_days: int) -> pd.DataFrame:
    gain = compute_price_gain(daily_prices, window_days)
    if gain.empty:
        return pd.DataFrame()

    pct_col = f"Price_Gain_%_{window_days}D"

    t = trades.copy()
    t["Date"] = pd.to_datetime(t["Date"]).dt.date
    all_dates = sorted(t["Date"].unique())
    latest = max(all_dates)
    win_dates = _rolling_dates(all_dates, latest, window_days)
    sub = t[t["Date"].isin(win_dates)].copy()

    qty = sub.groupby("Symbol", as_index=False).agg(Window_Qty=("Qty", "sum"), Window_Amt=("Amount", "sum"))
    bn = compute_broker_net(trades, window_days)

    pressure_rows = []
    for sym, g in bn.groupby("Symbol"):
        total_qty = qty.loc[qty["Symbol"] == sym, "Window_Qty"]
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
        pr = pd.DataFrame({"Symbol": gain["Symbol"], "Top3_Buy_Share": 0.0, "Top3_Sell_Share": 0.0})

    out = gain.merge(qty, on="Symbol", how="left").merge(pr, on="Symbol", how="left").fillna(0)

    vwap = sub.groupby("Symbol", as_index=False).agg(Window_Amt=("Amount", "sum"), Window_Qty=("Qty", "sum"))
    vwap["Window_VWAP"] = vwap["Window_Amt"] / vwap["Window_Qty"].replace(0, np.nan)
    out = out.merge(vwap[["Symbol", "Window_VWAP"]], on="Symbol", how="left")

    out["Price Trend"] = np.where(out["Today Price"] >= out["Window_VWAP"].fillna(out["Today Price"]), "UP", "DOWN")
    out["Buy_Pressure"] = out["Top3_Buy_Share"].apply(pressure_label)
    out["Sell_Pressure"] = out["Top3_Sell_Share"].apply(pressure_label)
    out["Momentum"] = out[pct_col].apply(lambda x: momentum_label(float(x) if pd.notna(x) else 0.0))

    out["Trader_View"] = out.apply(
        lambda r: trader_view(r["Price Trend"], r["Buy_Pressure"], r["Sell_Pressure"], r["Momentum"]),
        axis=1
    )

    sell_rank = {"HIGH": 0, "MED": 1, "LOW": 2}
    out["_sell_rank"] = out["Sell_Pressure"].map(sell_rank).fillna(9)
    out = out.sort_values(["_sell_rank", "Window_Qty"], ascending=[True, False]).reset_index(drop=True)
    out.insert(0, "Rank", np.arange(1, len(out) + 1))

    out = out.rename(columns={
        "Window_Qty": f"Total_Qty_{window_days}D",
        "Window_Amt": f"Total_Amt_{window_days}D",
        "Top3_Buy_Share": "Buy_Pressure_Share(Top3)",
        "Top3_Sell_Share": "Sell_Pressure_Share(Top3)",
        pct_col: f"Price_Gain_%_{window_days}D",
    })

    keep = [
        "Rank", "Symbol", "Today Price", f"Price_Gain_%_{window_days}D",
        f"Total_Qty_{window_days}D", f"Total_Amt_{window_days}D",
        "Price Trend", "Buy_Pressure", "Sell_Pressure", "Momentum", "Trader_View",
        "Buy_Pressure_Share(Top3)", "Sell_Pressure_Share(Top3)"
    ]
    return out[keep]


# =========================
# Excel Table Helpers
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
        ws.column_dimensions[get_column_letter(c)].width = min(max_len + 2, 40)


def write_market_overview_all_in_one_sheet(wb, latest_date, df_1d, df_7d, df_15d, df_1m):
    ws = wb.create_sheet("Market_Overview_All")

    ws["A1"] = f"Market Overview Summary (All Windows) — {latest_date}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "1D | 7D | 15D | 1M tables in one sheet (each is Excel Table)"
    ws["A2"].font = Font(italic=True, size=10)

    section_title_row = 4
    table_header_row = 5
    gap = 1
    start_col = 1

    tables = [
        ("1D", df_1d, "TBL_OV_1D"),
        ("7D", df_7d, "TBL_OV_7D"),
        ("15D", df_15d, "TBL_OV_15D"),
        ("1M", df_1m, "TBL_OV_1M"),
    ]

    for label, df, tname in tables:
        ws.cell(row=section_title_row, column=start_col, value=f"Market Overview – {label}").font = Font(bold=True, size=12)

        if df is None or df.empty:
            ws.cell(row=table_header_row, column=start_col, value="No data").font = Font(bold=True)
            start_col += 1 + gap
            continue

        rows = list(dataframe_to_rows(df, index=False, header=True))
        for i, row in enumerate(rows, start=table_header_row):
            for j, val in enumerate(row, start=start_col):
                ws.cell(i, j, value=val)

        end_row = table_header_row + len(rows) - 1
        end_col = start_col + df.shape[1] - 1

        style_block(ws, table_header_row, start_col, end_row, end_col)
        add_excel_table(ws, table_header_row, start_col, end_row, end_col, tname)

        start_col = end_col + 1 + gap


# =========================
# MAIN
# =========================
def main():
    print("[1/4] Loading floorsheets from outputs/ ...")
    trades = load_all_floorsheets(INPUT_FOLDER)

    print("[2/4] Building daily prices ...")
    daily_prices = build_daily_symbol_prices(trades)

    latest_date = max(pd.to_datetime(daily_prices["Date"]).dt.date.unique())
    print(f"[INFO] Latest trade date: {latest_date}")

    print("[3/4] Computing market overview tables ...")
    df_1d = build_market_overview(trades, daily_prices, WINDOWS["1D"]).head(50)
    df_7d = build_market_overview(trades, daily_prices, WINDOWS["7D"]).head(50)
    df_15 = build_market_overview(trades, daily_prices, WINDOWS["15D"]).head(50)
    df_1m = build_market_overview(trades, daily_prices, WINDOWS["1M"]).head(50)

    print("[4/4] Writing Excel report ...")
    os.makedirs(REPORT_FOLDER, exist_ok=True)
    out_path = os.path.join(REPORT_FOLDER, f"{REPORT_NAME_PREFIX}_{latest_date}.xlsx")

    wb = Workbook()
    wb.remove(wb.active)

    write_market_overview_all_in_one_sheet(wb, latest_date, df_1d, df_7d, df_15, df_1m)

    wb.save(out_path)
    print(f"[DONE] Saved: {out_path}")


if __name__ == "__main__":
    main()
