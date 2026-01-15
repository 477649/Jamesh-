# scripts/generate_trading_report.py
import os
import re
import json
from pathlib import Path
from datetime import datetime
import numpy as np
import pandas as pd

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


# =========================
# CONFIG (edit only if needed)
# =========================
ROOT = Path(__file__).resolve().parents[1]

FLOOR_DIR = ROOT / "outputs" / "Floor Sheet"
PRICE_DIR = ROOT / "outputs" / "sharesansar"
SECTOR_PATH = ROOT / "outputs" / "Sector" / "sector_master.csv"

REPORT_DIR = ROOT / "outputs" / "reports"
REPORT_DIR.mkdir(parents=True, exist_ok=True)

# rolling windows in "available trading days"
WINDOWS = {
    "1D": 1,
    "7D": 7,
    "15D": 15,
    "1M": 30,   # 30 available trading days
}

# filename patterns
FLOOR_RE = re.compile(r".*?(\d{4}-\d{2}-\d{2}).*\.csv$", re.IGNORECASE)
PRICE_RE = re.compile(r".*?(\d{4}-\d{2}-\d{2}).*\.csv$", re.IGNORECASE)

CRORE = 10_000_000


# =========================
# HELPERS
# =========================
def safe_float(s):
    if s is None:
        return np.nan
    try:
        return float(str(s).replace(",", "").strip())
    except Exception:
        return np.nan

def list_dates_from_folder(folder: Path, regex: re.Pattern):
    dates = []
    files = []
    if not folder.exists():
        return dates, files
    for p in folder.glob("*.csv"):
        m = regex.match(p.name)
        if m:
            d = m.group(1)
            try:
                dt = pd.to_datetime(d).date()
                dates.append(dt)
                files.append(p)
            except Exception:
                pass
    # sort by date
    dates_files = sorted(zip(dates, files), key=lambda x: x[0])
    return [d for d, _ in dates_files], [f for _, f in dates_files]

def choose_window_dates(all_dates_sorted, n):
    # last n available trading dates
    if not all_dates_sorted:
        return []
    return all_dates_sorted[-min(n, len(all_dates_sorted)):]

def load_sector_master(path: Path):
    if not path.exists():
        # allow empty master
        return pd.DataFrame(columns=["Symbol", "Company", "Sectors"])
    df = pd.read_csv(path)
    # normalize col names
    df.columns = [c.strip() for c in df.columns]
    # expected: Symbol, Company, Sectors
    if "Sector" in df.columns and "Sectors" not in df.columns:
        df = df.rename(columns={"Sector": "Sectors"})
    return df[["Symbol", "Company", "Sectors"]].drop_duplicates()

def read_floorsheet_file(path: Path, trade_date):
    df = pd.read_csv(path)
    df.columns = [c.strip() for c in df.columns]
    # expected: Transact No., Symbol, Buyer, Seller, Quantity, Rate, Amount
    # normalize variations
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
    df["Quantity"] = df["Quantity"].apply(safe_float).fillna(0).astype(float)
    df["Rate"] = df["Rate"].apply(safe_float).fillna(0).astype(float)
    df["Amount"] = df["Amount"].apply(safe_float).fillna(0).astype(float)
    # brokers as int where possible
    for col in ["Buyer", "Seller"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")
    df["Symbol"] = df["Symbol"].astype(str).str.strip().str.upper()
    return df[needed + ["TradeDate"]]

def read_price_file(path: Path, trade_date):
    df = pd.read_csv(path)
    df.columns = [c.strip() for c in df.columns]
    # expected at least these:
    needed = ["Symbol", "Open", "High", "Low", "Close", "LTP", "VWAP", "Vol", "Turnover"]
    # some files use Volume
    if "Volume" in df.columns and "Vol" not in df.columns:
        df = df.rename(columns={"Volume": "Vol"})
    missing = [c for c in needed if c not in df.columns]
    if missing:
        # be tolerant: require only core OHLC + LTP
        core = ["Symbol", "Open", "High", "Low", "Close", "LTP"]
        core_missing = [c for c in core if c not in df.columns]
        if core_missing:
            raise ValueError(f"SharePrice missing columns {core_missing} in {path.name}")
        # if VWAP/Vol/Turnover not present, create
        for c in ["VWAP", "Vol", "Turnover"]:
            if c not in df.columns:
                df[c] = np.nan

    df["TradeDate"] = pd.to_datetime(trade_date)
    for c in ["Open", "High", "Low", "Close", "LTP", "VWAP", "Vol", "Turnover"]:
        if c in df.columns:
            df[c] = df[c].apply(safe_float)
    df["Symbol"] = df["Symbol"].astype(str).str.strip().str.upper()
    keep_cols = ["TradeDate", "Symbol", "Open", "High", "Low", "Close", "LTP", "VWAP", "Vol", "Turnover"]
    keep_cols = [c for c in keep_cols if c in df.columns]
    return df[keep_cols]


# =========================
# METRICS + SCORING
# =========================
def symbol_metrics_from_floorsheet(fs: pd.DataFrame):
    # Symbol-level trades + VWAP from floorsheet
    g = fs.groupby("Symbol", as_index=False).agg(
        Trades=("Transact No.", "count"),
        Total_Qty=("Quantity", "sum"),
        Total_Amount=("Amount", "sum"),
        # numerator for VWAP
        _pv=("Rate", lambda s: 0.0),
    )
    # compute VWAP safely: sum(qty*rate)/sum(qty)
    fs["_qxr"] = fs["Quantity"] * fs["Rate"]
    vw = fs.groupby("Symbol", as_index=False).agg(
        _qxr=("__qxr" if "__qxr" in fs.columns else "_qxr", "sum"),
        _qty=("Quantity", "sum"),
    )
    vw["VWAP"] = np.where(vw["_qty"] > 0, vw["_qxr"] / vw["_qty"], np.nan)
    vw = vw[["Symbol", "VWAP"]]
    out = g.merge(vw, on="Symbol", how="left")
    out["Total_Amount_Cr"] = out["Total_Amount"] / CRORE
    return out.drop(columns=["_pv"], errors="ignore")

def broker_symbol_metrics(fs: pd.DataFrame):
    # Convert each trade into buyer-side (+qty) and seller-side (-qty)
    buy = fs[["Symbol", "Buyer", "Quantity", "Rate", "Amount"]].copy()
    buy = buy.rename(columns={"Buyer": "Broker"})
    buy["Buy_Qty"] = buy["Quantity"]
    buy["Sell_Qty"] = 0.0
    buy["_buy_cost"] = buy["Quantity"] * buy["Rate"]

    sell = fs[["Symbol", "Seller", "Quantity", "Rate", "Amount"]].copy()
    sell = sell.rename(columns={"Seller": "Broker"})
    sell["Buy_Qty"] = 0.0
    sell["Sell_Qty"] = sell["Quantity"]
    sell["_buy_cost"] = 0.0

    x = pd.concat([buy, sell], ignore_index=True)
    x["Broker"] = x["Broker"].astype("Int64")
    g = x.groupby(["Symbol", "Broker"], as_index=False).agg(
        Buy_Qty=("Buy_Qty", "sum"),
        Sell_Qty=("Sell_Qty", "sum"),
        _buy_cost=("_buy_cost", "sum"),
    )
    g["Net_Qty"] = g["Buy_Qty"] - g["Sell_Qty"]
    g["Avg_Buy_Cost"] = np.where(g["Buy_Qty"] > 0, g["_buy_cost"] / g["Buy_Qty"], np.nan)
    return g.drop(columns=["_buy_cost"])

def top_net_brokers(bsym: pd.DataFrame, topn=5):
    # For each symbol, get top net buyers and sellers (by Net_Qty)
    buyers = bsym.sort_values(["Symbol", "Net_Qty"], ascending=[True, False]).groupby("Symbol").head(topn)
    sellers = bsym.sort_values(["Symbol", "Net_Qty"], ascending=[True, True]).groupby("Symbol").head(topn)
    buyers = buyers.copy()
    sellers = sellers.copy()
    buyers["Side"] = "Top_Net_Buyers"
    sellers["Side"] = "Top_Net_Sellers"
    return pd.concat([buyers, sellers], ignore_index=True)

def compute_pressure(bsym: pd.DataFrame):
    # Pressure based on top-3 net qty dominance share
    # BuyPressure: (sum of top3 positive Net_Qty) / (sum of positive Net_Qty)
    # SellPressure: (sum of top3 negative abs(Net_Qty)) / (sum of negative abs(Net_Qty))
    if bsym.empty:
        return pd.DataFrame(columns=["Symbol", "Buy_Pressure", "Sell_Pressure"])

    t = bsym.copy()
    t["pos"] = t["Net_Qty"].clip(lower=0)
    t["neg_abs"] = (-t["Net_Qty"]).clip(lower=0)

    pos_total = t.groupby("Symbol", as_index=False)["pos"].sum().rename(columns={"pos": "pos_total"})
    neg_total = t.groupby("Symbol", as_index=False)["neg_abs"].sum().rename(columns={"neg_abs": "neg_total"})

    top3_pos = t.sort_values(["Symbol", "pos"], ascending=[True, False]).groupby("Symbol").head(3)
    top3_pos = top3_pos.groupby("Symbol", as_index=False)["pos"].sum().rename(columns={"pos": "top3_pos"})

    top3_neg = t.sort_values(["Symbol", "neg_abs"], ascending=[True, False]).groupby("Symbol").head(3)
    top3_neg = top3_neg.groupby("Symbol", as_index=False)["neg_abs"].sum().rename(columns={"neg_abs": "top3_neg"})

    out = pos_total.merge(top3_pos, on="Symbol", how="left").merge(neg_total, on="Symbol", how="left").merge(top3_neg, on="Symbol", how="left")
    out["Buy_Pressure"] = np.where(out["pos_total"] > 0, out["top3_pos"].fillna(0) / out["pos_total"], np.nan)
    out["Sell_Pressure"] = np.where(out["neg_total"] > 0, out["top3_neg"].fillna(0) / out["neg_total"], np.nan)
    return out[["Symbol", "Buy_Pressure", "Sell_Pressure"]]

def momentum_from_prices(price_window: pd.DataFrame, latest_date: pd.Timestamp):
    # simple momentum: Close(latest) / Close(first_in_window) - 1
    # requires multiple dates; works best for 7D/15D/1M
    if price_window.empty:
        return pd.DataFrame(columns=["Symbol", "Momentum"])

    p = price_window.sort_values(["Symbol", "TradeDate"])
    first = p.groupby("Symbol", as_index=False).first()[["Symbol", "Close"]].rename(columns={"Close": "Close_first"})
    last = p[p["TradeDate"] == latest_date][["Symbol", "Close"]].rename(columns={"Close": "Close_last"})
    out = first.merge(last, on="Symbol", how="inner")
    out["Momentum"] = np.where(out["Close_first"] > 0, (out["Close_last"] / out["Close_first"]) - 1.0, np.nan)
    return out[["Symbol", "Momentum"]]

def classify_signal(df):
    """
    Final recommendation:
    - BUY: Score >= 70
    - HOLD: 50-69
    - SELL/AVOID: <50
    """
    bins = [-1e9, 50, 70, 1e9]
    labels = ["SELL / AVOID", "HOLD", "BUY"]
    return pd.cut(df["Score"], bins=bins, labels=labels, right=False)

def build_score(df):
    """
    Score is built from:
    + Price vs VWAP (premium/discount)
    + Momentum (window)
    + Net flow strength (Net_Qty scaled)
    + Pressure quality (not too concentrated on sell side)
    + Liquidity filter (Turnover/Amount)
    """
    x = df.copy()

    # Normalize helper
    def z(s):
        s = s.replace([np.inf, -np.inf], np.nan)
        return (s - s.mean()) / (s.std(ddof=0) + 1e-9)

    # Features (safe)
    x["Price_vs_VWAP_pct"] = np.where(x["VWAP"].notna() & (x["VWAP"] > 0),
                                     (x["Last_Price"] / x["VWAP"]) - 1.0, np.nan)

    # Net strength proxy (use Total_Qty as activity proxy; Net flow derived from broker table later if present)
    x["Activity"] = np.log1p(x["Total_Qty"].fillna(0))

    # Pressure signals: prefer higher buy pressure and lower sell pressure
    x["BuyP"] = x["Buy_Pressure"].fillna(0)
    x["SellP"] = x["Sell_Pressure"].fillna(0)

    # Liquidity: Amount (Cr)
    x["Liq"] = np.log1p(x["Total_Amount_Cr"].fillna(0))

    # Build score components (z-scored then weighted)
    comp_pvv = z(x["Price_vs_VWAP_pct"].fillna(0)) * 25
    comp_mom = z(x["Momentum"].fillna(0)) * 25
    comp_act = z(x["Activity"].fillna(0)) * 15
    comp_liq = z(x["Liq"].fillna(0)) * 15
    comp_pressure = (z(x["BuyP"].fillna(0)) - z(x["SellP"].fillna(0))) * 20

    raw = comp_pvv + comp_mom + comp_act + comp_liq + comp_pressure

    # Scale to 0-100
    raw_min, raw_max = raw.min(), raw.max()
    x["Score"] = np.where(raw_max > raw_min, 100 * (raw - raw_min) / (raw_max - raw_min), 50.0)

    # Add signal label
    x["Recommendation"] = classify_signal(x)

    # Reason tags (interpretable)
    reasons = []
    for _, r in x.iterrows():
        tags = []
        if r.get("Price_vs_VWAP_pct", 0) > 0.01: tags.append("Above VWAP")
        if r.get("Price_vs_VWAP_pct", 0) < -0.01: tags.append("Below VWAP")
        if r.get("Momentum", 0) > 0.03: tags.append("Strong Momentum")
        if r.get("Momentum", 0) < -0.03: tags.append("Weak Momentum")
        if r.get("Buy_Pressure", 0) >= 0.40: tags.append("Buyer Dominance")
        if r.get("Sell_Pressure", 0) >= 0.40: tags.append("Seller Dominance")
        if r.get("Total_Amount_Cr", 0) >= 5: tags.append("High Liquidity")
        reasons.append(", ".join(tags[:4]))
    x["Reason"] = reasons
    return x


# =========================
# EXCEL WRITER (format)
# =========================
def style_sheet(ws):
    ws.freeze_panes = "A2"
    thin = Side(style="thin", color="D0D7DE")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="EEF2FF")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.border = border
            if isinstance(c.value, str) and len(c.value) > 60:
                c.alignment = Alignment(wrap_text=True, vertical="top")

def add_table(ws, df, name="Table1"):
    # write df then convert to an Excel Table
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    style_sheet(ws)
    nrows = ws.max_row
    ncols = ws.max_column
    if nrows >= 2 and ncols >= 1:
        from openpyxl.utils import get_column_letter
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

def write_excel_report(path, sheets: dict):
    wb = Workbook()
    # remove default
    wb.remove(wb.active)

    for i, (sname, df) in enumerate(sheets.items(), start=1):
        ws = wb.create_sheet(title=sname[:31])
        df = df.copy()
        # avoid nan in excel readability
        df = df.replace([np.inf, -np.inf], np.nan)
        add_table(ws, df, name=f"T{i}")

        # auto width
        for col in ws.columns:
            max_len = 10
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value is None:
                    continue
                max_len = max(max_len, len(str(cell.value))[:200].__len__())
            ws.column_dimensions[col_letter].width = min(45, max_len + 2)

    wb.save(path)


# =========================
# MAIN PIPELINE
# =========================
def main():
    # 1) discover available dates from floorsheet & price
    floor_dates, floor_files = list_dates_from_folder(FLOOR_DIR, FLOOR_RE)
    price_dates, price_files = list_dates_from_folder(PRICE_DIR, PRICE_RE)

    if not floor_dates:
        raise RuntimeError(f"No floorsheet csv files found in {FLOOR_DIR}")

    # Use floorsheet dates as primary trading calendar (prices may be missing some)
    trading_dates = floor_dates
    latest_date = pd.to_datetime(trading_dates[-1])

    sector = load_sector_master(SECTOR_PATH)

    # quick map date->file
    floor_map = {d: f for d, f in zip(floor_dates, floor_files)}
    price_map = {d: f for d, f in zip(price_dates, price_files)}

    # 2) for each window, load + compute
    symbol_summary_all = []
    broker_by_symbol_all = []
    broker_summary_all = []
    market_overview_rows = []
    top_picks_rows = []
    price_movers_rows = []
    sector_summary_rows = []

    # Load all price history needed once (optional, but efficient)
    # We'll load per window to keep memory low.

    for wname, n in WINDOWS.items():
        w_dates = choose_window_dates(trading_dates, n)
        w_dates_ts = [pd.to_datetime(d) for d in w_dates]
        w_latest = pd.to_datetime(w_dates[-1])

        # ----- Floorsheet load -----
        fs_list = []
        for d in w_dates:
            fp = floor_map.get(d)
            if fp is None:
                continue
            fs_list.append(read_floorsheet_file(fp, d))
        fs = pd.concat(fs_list, ignore_index=True) if fs_list else pd.DataFrame()

        # ----- Price load -----
        pr_list = []
        for d in w_dates:
            pp = price_map.get(d)
            if pp is None:
                continue
            pr_list.append(read_price_file(pp, d))
        pr = pd.concat(pr_list, ignore_index=True) if pr_list else pd.DataFrame()

        # Latest prices (for Last_Price)
        last_price = pr[pr["TradeDate"] == w_latest][["Symbol", "Close", "LTP", "VWAP", "Turnover", "Vol"]] \
                        .rename(columns={"Close": "Close_latest", "LTP": "Last_Price", "VWAP": "VWAP_price"}) \
                        if not pr.empty else pd.DataFrame(columns=["Symbol","Close_latest","Last_Price","VWAP_price","Turnover","Vol"])

        # Floorsheet symbol metrics (VWAP from trades)
        sym = symbol_metrics_from_floorsheet(fs) if not fs.empty else pd.DataFrame(columns=["Symbol","Trades","Total_Qty","Total_Amount","VWAP","Total_Amount_Cr"])

        # Broker-symbol and pressure
        bsym = broker_symbol_metrics(fs) if not fs.empty else pd.DataFrame(columns=["Symbol","Broker","Buy_Qty","Sell_Qty","Net_Qty","Avg_Buy_Cost"])
        pressure = compute_pressure(bsym)

        # Momentum from price
        mom = momentum_from_prices(pr, w_latest)

        # Merge symbol summary
        sym = sym.merge(last_price[["Symbol","Last_Price","Turnover","Vol"]], on="Symbol", how="left")
        sym = sym.merge(pressure, on="Symbol", how="left")
        sym = sym.merge(mom, on="Symbol", how="left")

        # Add sector/company
        sym = sym.merge(sector, on="Symbol", how="left")

        # Fill last price fallback
        if "Last_Price" in sym.columns:
            sym["Last_Price"] = sym["Last_Price"].fillna(sym.get("Close_latest", np.nan))

        # Scoring + recommendation for this window
        scored = build_score(sym)

        # keep & rename with window suffix for final wide table
        keep_cols = ["Symbol","Company","Sectors","Trades","Total_Qty","Total_Amount_Cr","VWAP","Last_Price",
                     "Price_vs_VWAP_pct","Momentum","Buy_Pressure","Sell_Pressure","Score","Recommendation","Reason"]
        keep_cols = [c for c in keep_cols if c in scored.columns]
        scored = scored[keep_cols].copy()

        # Convert % fields to readable
        if "Price_vs_VWAP_pct" in scored.columns:
            scored["Price_vs_VWAP_pct"] = (scored["Price_vs_VWAP_pct"] * 100).round(2)
        if "Momentum" in scored.columns:
            scored["Momentum"] = (scored["Momentum"] * 100).round(2)
        if "Buy_Pressure" in scored.columns:
            scored["Buy_Pressure"] = scored["Buy_Pressure"].round(3)
        if "Sell_Pressure" in scored.columns:
            scored["Sell_Pressure"] = scored["Sell_Pressure"].round(3)

        # Save symbol summary long format
        scored.insert(0, "Window", wname)
        symbol_summary_all.append(scored)

        # ----- Top Picks for this window -----
        top_buy = scored.sort_values("Score", ascending=False).head(20).copy()
        top_buy["List"] = "TOP_BUY"
        top_sell = scored.sort_values("Score", ascending=True).head(20).copy()
        top_sell["List"] = "TOP_SELL"
        top_hold = scored[(scored["Recommendation"] == "HOLD")].sort_values("Score", ascending=False).head(20).copy()
        top_hold["List"] = "TOP_HOLD"

        top_picks_rows.append(pd.concat([top_buy, top_hold, top_sell], ignore_index=True))

        # ----- Broker Summary (overall) -----
        if not bsym.empty:
            bsum = bsym.groupby("Broker", as_index=False).agg(
                Buy_Qty=("Buy_Qty", "sum"),
                Sell_Qty=("Sell_Qty", "sum"),
                Net_Qty=("Net_Qty", "sum"),
            )
            bsum.insert(0, "Window", wname)
            broker_summary_all.append(bsum)

            # Broker-by-symbol top lists
            tb = top_net_brokers(bsym, topn=5)
            tb.insert(0, "Window", wname)
            broker_by_symbol_all.append(tb)

        # ----- Market Overview -----
        mrow = {
            "Window": wname,
            "From": str(w_dates[0]),
            "To": str(w_dates[-1]),
            "Symbols_Traded": int(sym["Symbol"].nunique()) if not sym.empty else 0,
            "Total_Amount_Cr": float(sym["Total_Amount_Cr"].sum()) if "Total_Amount_Cr" in sym.columns else 0.0,
            "Total_Qty": float(sym["Total_Qty"].sum()) if "Total_Qty" in sym.columns else 0.0,
            "Avg_Buy_Pressure": float(sym["Buy_Pressure"].mean()) if "Buy_Pressure" in sym.columns else np.nan,
            "Avg_Sell_Pressure": float(sym["Sell_Pressure"].mean()) if "Sell_Pressure" in sym.columns else np.nan,
            "Advancers(Score>=70)": int((scored["Recommendation"] == "BUY").sum()) if "Recommendation" in scored.columns else 0,
            "Decliners(Score<50)": int((scored["Recommendation"] == "SELL / AVOID").sum()) if "Recommendation" in scored.columns else 0,
        }
        market_overview_rows.append(mrow)

        # ----- Price Movers (based on prices only) -----
        if not pr.empty and pr["TradeDate"].nunique() >= 2:
            p = pr.sort_values(["Symbol","TradeDate"])
            first = p.groupby("Symbol", as_index=False).first()[["Symbol","Close"]].rename(columns={"Close":"Close_start"})
            last = p[p["TradeDate"] == w_latest][["Symbol","Close"]].rename(columns={"Close":"Close_end"})
            mv = first.merge(last, on="Symbol", how="inner")
            mv["Change_%"] = np.where(mv["Close_start"]>0, (mv["Close_end"]/mv["Close_start"]-1)*100, np.nan)
            mv = mv.merge(sector, on="Symbol", how="left")
            mv["Window"] = wname
            price_movers_rows.append(mv.sort_values("Change_%", ascending=False).head(50))

        # ----- Sector Summary (net score / liquidity) -----
        if not sym.empty and "Sectors" in sym.columns:
            sec = sym.groupby("Sectors", as_index=False).agg(
                Symbols=("Symbol", "nunique"),
                Amount_Cr=("Total_Amount_Cr", "sum"),
                Avg_Score=("Score", "mean") if "Score" in sym.columns else ("Total_Amount_Cr","mean"),
                Avg_Momentum=("Momentum", "mean") if "Momentum" in sym.columns else ("Total_Amount_Cr","mean"),
            )
            sec["Window"] = wname
            sector_summary_rows.append(sec.sort_values("Amount_Cr", ascending=False))

    # Combine long outputs
    symbol_summary = pd.concat(symbol_summary_all, ignore_index=True) if symbol_summary_all else pd.DataFrame()
    broker_summary = pd.concat(broker_summary_all, ignore_index=True) if broker_summary_all else pd.DataFrame()
    broker_by_symbol = pd.concat(broker_by_symbol_all, ignore_index=True) if broker_by_symbol_all else pd.DataFrame()
    market_overview = pd.DataFrame(market_overview_rows)
    top_picks = pd.concat(top_picks_rows, ignore_index=True) if top_picks_rows else pd.DataFrame()
    price_movers = pd.concat(price_movers_rows, ignore_index=True) if price_movers_rows else pd.DataFrame()
    sector_summary = pd.concat(sector_summary_rows, ignore_index=True) if sector_summary_rows else pd.DataFrame()

    # README sheet content
    readme = pd.DataFrame([
        ["Advanced Trading Insight Report", ""],
        ["Windows", "1D=1 day, 7D=7 trading days, 15D=15 trading days, 1M=30 trading days (based on available files)"],
        ["VWAP", "Computed from floorsheet: sum(Qty*Rate)/sum(Qty)"],
        ["Buy Pressure", "Top-3 net buyers dominance share among positive net flows (0-1)"],
        ["Sell Pressure", "Top-3 net sellers dominance share among negative net flows (0-1)"],
        ["Score", "0-100 scale from price vs VWAP, momentum, activity, liquidity, pressure"],
        ["Recommendation", "BUY (>=70), HOLD (50-69), SELL/AVOID (<50)"],
        ["Notes", "Use TOP_PICKS sheet for decision. Always verify with chart/news before trade."],
    ], columns=["Item", "Explanation"])

    # Output file
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = REPORT_DIR / f"Advanced_Trading_Insight_Report_{ts}.xlsx"

    sheets = {
        "README": readme,
        "Market_Overview": market_overview.sort_values("Window"),
        "Top_Picks": top_picks,
        "Symbol_Summary": symbol_summary,
        "Broker_Summary": broker_summary,
        "Broker_by_Symbol": broker_by_symbol,
        "Sector_Summary": sector_summary,
        "Price_Movers": price_movers,
    }

    write_excel_report(out_path, sheets)
    print(f"✅ Report generated: {out_path}")

    # Optional: write a small json pointer for dashboards
    latest_json = REPORT_DIR / "latest_report.json"
    latest_json.write_text(json.dumps({"latest_report": out_path.name}, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
