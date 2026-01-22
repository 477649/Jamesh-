# scripts/generate_advanced_market_dashboard.py
# -----------------------------------------------------------------------------
# Advanced NEPSE Retail-Pro Dashboard (Excel) with BUY / HOLD / EXIT engine
# Based on repo data:
#   outputs/floorsheet_YYYY-MM-DD.csv
#   outputs/sharesansar/SharePrice_YYYY-MM-DD.csv
#   outputs/Sector/sector_master.csv (recommended)
#   outputs/Brokers/broker_master.csv (optional)
#
# Creates an Excel report with multi-window analysis, built for RETAIL decision-making.
#
# Sheets (core):
#   README
#   INPUTS_SUMMARY
#   MARKET_OVERVIEW
#   RISK_GRID
#   SYMBOL_SCORECARD   (multi-window scoring + BUY/HOLD/EXIT)
#   TRADE_SETUPS       (top actionable per window)
#   SECTOR_SUMMARY     (if sector_master exists)
#   BROKER_SUMMARY     (net flow by broker per window, if floorsheet exists)
#   BROKER_BY_SYMBOL   (top brokers per symbol per window)
#   SMART_MONEY        (symbol-level accumulation/distribution features)
#   TRAP_WARNINGS      (retail protection)
#   PRICE_MOVERS       (gainers/losers + volume confirmation)
#
# Extra (window-specific top sheets):
#   SM_TOP_<W>, SB_TOP_<W>, INST_TOP_<W>, OPR_TOP_<W>  (if broker_master has types)
#
# IMPORTANT:
# - Does NOT create "Market_Breadth" sheet.
# - Uses trading-day windows (by file count) and works with MANY days of data.
# -----------------------------------------------------------------------------

from __future__ import annotations

import re
from pathlib import Path
from datetime import datetime
import numpy as np
import pandas as pd

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.utils import get_column_letter


# -----------------------------
# CONFIG
# -----------------------------
ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "outputs"
PRICE_DIR = OUT_DIR / "sharesansar"
REPORT_DIR = OUT_DIR / "reports"

# Some repos store floorsheets inside a subfolder like outputs/"Floor Sheet"
FLOOR_DIR_CANDIDATES = [
    OUT_DIR,
    OUT_DIR / "Floor Sheet",
    OUT_DIR / "FloorSheet",
    OUT_DIR / "floorsheet",
]

SECTOR_MASTER = OUT_DIR / "Sector" / "sector_master.csv"
BROKER_MASTER = OUT_DIR / "Brokers" / "broker_master.csv"

FLOOR_PATTERN = re.compile(r"floorsheet_(\d{4}-\d{2}-\d{2})\.csv$", re.I)
PRICE_PATTERN = re.compile(r"SharePrice_(\d{4}-\d{2}-\d{2})\.csv$", re.I)

# Trading-day windows (by number of days/files)
WINDOWS = [
    ("1D", 1),
    ("7D", 7),
    ("15D", 15),
    ("1M", 30),
    ("3M", 90),
    ("6M", 180),
    ("1Y", 250),
]

# Top-N logic (your earlier preference)
def topn_for_window(wname: str) -> int:
    if wname == "1D":
        return 4
    if wname == "7D":
        return 5
    return 10


# -----------------------------
# Helpers: parsing, safe numeric
# -----------------------------
def parse_date(s: str) -> datetime:
    return datetime.strptime(s, "%Y-%m-%d")

def to_num(s):
    if pd.isna(s):
        return np.nan
    if isinstance(s, (int, float, np.number)):
        return float(s)
    s = str(s).strip()
    if s == "":
        return np.nan
    s = s.replace(",", "")
    s = re.sub(r"\s+", "", s)
    try:
        return float(s)
    except Exception:
        return pd.to_numeric(s, errors="coerce")

def safe_div(a, b):
    b = np.asarray(b, dtype="float64")
    a = np.asarray(a, dtype="float64")
    return np.where(b == 0, np.nan, a / b)

def clamp(x, lo, hi):
    return np.minimum(np.maximum(x, lo), hi)


# -----------------------------
# Locate input files
# -----------------------------
def list_floor_files() -> list[tuple[datetime, Path]]:
    files: list[tuple[datetime, Path]] = []
    # Search in multiple possible folders (handles outputs/Floor Sheet structure)
    for base in FLOOR_DIR_CANDIDATES:
        if not base.exists():
            continue
        for p in base.glob('floorsheet_*.csv'):
            m = FLOOR_PATTERN.search(p.name)
            if m:
                files.append((parse_date(m.group(1)), p))
    if not files:
        return []
    files = sorted(files, key=lambda x: x[0])
    # De-duplicate by date (keep last path if duplicates exist)
    dedup = {}
    for d, p in files:
        dedup[d] = p
    return sorted([(d, p) for d, p in dedup.items()], key=lambda x: x[0])

def list_price_files() -> list[tuple[datetime, Path]]:
    files = []
    if not PRICE_DIR.exists():
        return []
    for p in PRICE_DIR.glob("SharePrice_*.csv"):
        m = PRICE_PATTERN.search(p.name)
        if m:
            files.append((parse_date(m.group(1)), p))
    return sorted(files, key=lambda x: x[0])

def latest_common_dates(floor_list, price_list) -> list[datetime]:
    floor_dates = {d for d, _ in floor_list}
    price_dates = {d for d, _ in price_list}
    common = sorted(floor_dates.intersection(price_dates))
    return common


# -----------------------------
# Load masters
# -----------------------------
def load_sector_master() -> pd.DataFrame:
    if not SECTOR_MASTER.exists():
        return pd.DataFrame(columns=["Symbol", "Sector"])
    df = pd.read_csv(SECTOR_MASTER)
    df.columns = [c.strip() for c in df.columns]

    if "Symbol" not in df.columns:
        for c in df.columns:
            if c.lower() == "symbol":
                df = df.rename(columns={c: "Symbol"})
                break

    if "Sector" not in df.columns:
        for c in df.columns:
            if c.lower() in ("sector", "sector_name", "sectorname"):
                df = df.rename(columns={c: "Sector"})
                break

    if "Symbol" in df.columns:
        df["Symbol"] = df["Symbol"].astype(str).str.strip().str.upper()
    if "Sector" in df.columns:
        df["Sector"] = df["Sector"].astype(str).str.strip()

    return df[["Symbol", "Sector"]].drop_duplicates()

def load_broker_master() -> pd.DataFrame:
    if not BROKER_MASTER.exists():
        return pd.DataFrame(columns=["Broker", "BrokerName", "BrokerType"])

    df = pd.read_csv(BROKER_MASTER)
    df.columns = [c.strip() for c in df.columns]

    if "Broker" not in df.columns:
        for c in df.columns:
            if c.lower() in ("broker", "broker_id", "brokerid", "code"):
                df = df.rename(columns={c: "Broker"})
                break
    if "BrokerName" not in df.columns:
        for c in df.columns:
            if c.lower() in ("brokername", "name"):
                df = df.rename(columns={c: "BrokerName"})
                break
    if "BrokerType" not in df.columns:
        for c in df.columns:
            if c.lower() in ("brokertype", "type", "category"):
                df = df.rename(columns={c: "BrokerType"})
                break

    if "Broker" in df.columns:
        df["Broker"] = pd.to_numeric(df["Broker"], errors="coerce").astype("Int64")

    for col in ["BrokerName", "BrokerType"]:
        if col in df.columns:
            df[col] = df[col].astype(str).fillna("").str.strip()

    if "BrokerType" in df.columns:
        df["BrokerType"] = df["BrokerType"].replace({"nan": ""}).fillna("")

    return df[["Broker", "BrokerName", "BrokerType"]].drop_duplicates()


# -----------------------------
# Load daily floorsheet + price
# -----------------------------
def load_floorsheet(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    df.columns = [c.strip() for c in df.columns]

    rename_map = {}
    for c in df.columns:
        cl = c.lower().strip()
        if cl in ("transact no.", "transactno", "transaction", "transaction no", "transact no"):
            rename_map[c] = "TransactNo"
        elif cl == "symbol":
            rename_map[c] = "Symbol"
        elif cl == "buyer":
            rename_map[c] = "Buyer"
        elif cl == "seller":
            rename_map[c] = "Seller"
        elif cl in ("quantity", "qty"):
            rename_map[c] = "Quantity"
        elif cl in ("rate", "price"):
            rename_map[c] = "Rate"
        elif cl in ("amount", "amt"):
            rename_map[c] = "Amount"

    df = df.rename(columns=rename_map)

    needed = ["Symbol", "Buyer", "Seller", "Quantity", "Rate", "Amount"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f"Floorsheet missing columns: {missing} in {path.name}")

    df["Symbol"] = df["Symbol"].astype(str).str.strip().str.upper()
    df["Buyer"] = pd.to_numeric(df["Buyer"], errors="coerce").astype("Int64")
    df["Seller"] = pd.to_numeric(df["Seller"], errors="coerce").astype("Int64")
    for c in ["Quantity", "Rate", "Amount"]:
        df[c] = df[c].map(to_num).astype(float)

    df = df.dropna(subset=["Symbol", "Buyer", "Seller", "Quantity", "Amount"])
    df = df[df["Quantity"] > 0]
    return df

def load_shareprice(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    df.columns = [c.strip() for c in df.columns]

    rename = {}
    for c in df.columns:
        cl = c.lower().strip()
        if cl == "symbol":
            rename[c] = "Symbol"
        elif cl == "open":
            rename[c] = "Open"
        elif cl == "high":
            rename[c] = "High"
        elif cl == "low":
            rename[c] = "Low"
        elif cl == "close":
            rename[c] = "Close"
        elif cl == "ltp":
            rename[c] = "LTP"
        elif cl == "vwap":
            rename[c] = "VWAP"
        elif cl in ("vol", "volume"):
            rename[c] = "Vol"
        elif cl in ("prev. close", "prev close", "prev_close"):
            rename[c] = "PrevClose"
        elif cl in ("turnover", "turn over"):
            rename[c] = "Turnover"
        elif cl in ("conf.", "conf"):
            rename[c] = "Conf"

    df = df.rename(columns=rename)

    # Minimal required set (more robust than forcing Turnover/VWAP)
    required = ["Symbol", "Close", "Vol", "PrevClose", "High", "Low"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"SharePrice missing columns: {missing} in {path.name}")

    df["Symbol"] = df["Symbol"].astype(str).str.strip().str.upper()

    # Convert numerics
    for c in ["Open", "High", "Low", "Close", "LTP", "VWAP", "Vol", "PrevClose", "Turnover", "Conf"]:
        if c in df.columns:
            df[c] = df[c].map(to_num).astype(float)

    # Fill optional columns if absent
    if "VWAP" not in df.columns:
        df["VWAP"] = df["Close"]
    else:
        df["VWAP"] = df["VWAP"].fillna(df["Close"])

    if "Turnover" not in df.columns:
        # Approx turnover if not provided
        df["Turnover"] = df["Close"] * df["Vol"]
    else:
        df["Turnover"] = df["Turnover"].fillna(df["Close"] * df["Vol"])

    df = df.dropna(subset=["Symbol", "Close", "Vol", "High", "Low", "PrevClose"])
    return df


# -----------------------------
# Floorsheet aggregation
# -----------------------------
def broker_symbol_net(floor_df: pd.DataFrame) -> pd.DataFrame:
    buy = floor_df.groupby(["Symbol", "Buyer"], as_index=False).agg(
        Buy_Qty=("Quantity", "sum"),
        Buy_Amt=("Amount", "sum"),
        Trades_Buy=("Quantity", "count"),
    ).rename(columns={"Buyer": "Broker"})

    sell = floor_df.groupby(["Symbol", "Seller"], as_index=False).agg(
        Sell_Qty=("Quantity", "sum"),
        Sell_Amt=("Amount", "sum"),
        Trades_Sell=("Quantity", "count"),
    ).rename(columns={"Seller": "Broker"})

    out = buy.merge(sell, on=["Symbol", "Broker"], how="outer").fillna(0.0)
    out["Net_Qty"] = out["Buy_Qty"] - out["Sell_Qty"]
    out["Net_Amt"] = out["Buy_Amt"] - out["Sell_Amt"]
    out["Trades"] = out["Trades_Buy"] + out["Trades_Sell"]
    return out

def symbol_flow_features(bs_net: pd.DataFrame) -> pd.DataFrame:
    if bs_net.empty:
        return pd.DataFrame(columns=[
            "Symbol","Buy_Qty","Sell_Qty","Net_Pos_Qty","Net_Neg_Qty",
            "Net_Pos_Ratio","Top3_Pos_Ratio","Broker_Concentration","Active_Brokers"
        ])

    sym = bs_net.groupby("Symbol", as_index=False).agg(
        Buy_Qty=("Buy_Qty", "sum"),
        Sell_Qty=("Sell_Qty", "sum"),
        Active_Brokers=("Broker", "nunique"),
    )

    tmp = bs_net.copy()
    tmp["Net_Pos_Qty"] = tmp["Net_Qty"].clip(lower=0)
    tmp["Net_Neg_Qty"] = (-tmp["Net_Qty"]).clip(lower=0)

    pos = tmp.groupby("Symbol", as_index=False).agg(
        Net_Pos_Qty=("Net_Pos_Qty", "sum"),
        Net_Neg_Qty=("Net_Neg_Qty", "sum"),
    )

    sym = sym.merge(pos, on="Symbol", how="left").fillna(0.0)
    sym["Total_Abs_Net"] = sym["Net_Pos_Qty"] + sym["Net_Neg_Qty"]
    sym["Net_Pos_Ratio"] = np.where(sym["Total_Abs_Net"] > 0,
                                    sym["Net_Pos_Qty"] / sym["Total_Abs_Net"], 0.0)

    top3 = (
        tmp[tmp["Net_Qty"] > 0]
        .sort_values(["Symbol","Net_Qty"], ascending=[True, False])
        .groupby("Symbol")
        .head(3)
        .groupby("Symbol", as_index=False)
        .agg(Top3_Pos_Qty=("Net_Qty", "sum"))
    )
    sym = sym.merge(top3, on="Symbol", how="left").fillna(0.0)
    sym["Top3_Pos_Ratio"] = np.where(sym["Net_Pos_Qty"] > 0,
                                     sym["Top3_Pos_Qty"] / sym["Net_Pos_Qty"], 0.0)

    tmp["Total_Qty_Act"] = tmp["Buy_Qty"] + tmp["Sell_Qty"]
    top3_act = (
        tmp.sort_values(["Symbol","Total_Qty_Act"], ascending=[True, False])
        .groupby("Symbol").head(3)
        .groupby("Symbol", as_index=False).agg(Top3_Act=("Total_Qty_Act","sum"))
    )
    tot_act = tmp.groupby("Symbol", as_index=False).agg(Total_Act=("Total_Qty_Act","sum"))
    conc = tot_act.merge(top3_act, on="Symbol", how="left").fillna(0.0)
    conc["Broker_Concentration"] = np.where(conc["Total_Act"] > 0, conc["Top3_Act"]/conc["Total_Act"], 0.0)

    sym = sym.merge(conc[["Symbol","Broker_Concentration"]], on="Symbol", how="left").fillna(0.0)
    return sym.drop(columns=["Total_Abs_Net"])


# -----------------------------
# Price panel + window stats
# -----------------------------
def build_price_panel(common_dates: list[datetime], price_map: dict) -> pd.DataFrame:
    frames = []
    for d in common_dates:
        p = price_map[d]
        dfp = load_shareprice(p).copy()
        dfp["Date"] = d
        dfp["Range"] = (dfp["High"] - dfp["Low"]).replace(0, np.nan)
        dfp["Ret1D"] = np.where(dfp["PrevClose"] > 0, (dfp["Close"] - dfp["PrevClose"]) / dfp["PrevClose"], np.nan)
        frames.append(dfp[["Date","Symbol","Close","VWAP","Vol","Turnover","High","Low","PrevClose","Ret1D","Range"]])
    panel = pd.concat(frames, ignore_index=True)
    panel = panel.sort_values(["Symbol","Date"]).reset_index(drop=True)
    return panel

def window_price_features(panel: pd.DataFrame, latest_date: datetime, wlen: int) -> pd.DataFrame:
    sub = panel[panel["Date"] <= latest_date].copy()
    sub = sub.sort_values(["Symbol","Date"])

    last_w = sub.groupby("Symbol").tail(wlen).copy()
    latest = last_w.sort_values(["Symbol","Date"]).groupby("Symbol").tail(1).copy()

    denom = (latest["High"] - latest["Low"]).replace(0, np.nan)
    latest["CSI"] = (latest["Close"] - latest["Low"]) / denom
    latest["Close_gt_VWAP"] = (latest["Close"] > latest["VWAP"]).astype(int)

    base = (
        last_w.sort_values(["Symbol","Date"])
        .groupby("Symbol")
        .head(1)[["Symbol","Close"]]
        .rename(columns={"Close":"BaseClose"})
    )
    latest = latest.merge(base, on="Symbol", how="left")
    latest["Ret_W"] = (latest["Close"] - latest["BaseClose"]) / latest["BaseClose"].replace(0, np.nan)

    agg = last_w.groupby("Symbol", as_index=False).agg(
        AvgVol=("Vol","mean"),
        AvgTurnover=("Turnover","mean"),
        SumTurnover=("Turnover","sum"),
        SumVol=("Vol","sum"),
        AvgRange=("Range","mean"),
    )
    out = latest.merge(agg, on="Symbol", how="left")
    out["Vol_Surge"] = out["Vol"] / out["AvgVol"].replace(0, np.nan)
    out["Close_vs_VWAP_pct"] = (out["Close"] - out["VWAP"]) / out["VWAP"].replace(0, np.nan)

    return out[[
        "Symbol","Close","VWAP","Vol","Turnover","Ret1D","Ret_W","Vol_Surge","AvgVol","AvgTurnover","SumTurnover","SumVol",
        "Close_gt_VWAP","CSI","High","Low","AvgRange","Close_vs_VWAP_pct"
    ]]


# -----------------------------
# Scoring engine (Retail-Pro)
# -----------------------------
def score_symbols(price_feat: pd.DataFrame, flow_feat: pd.DataFrame, sector_map: pd.DataFrame, wname: str) -> pd.DataFrame:
    df = price_feat.merge(flow_feat, on="Symbol", how="left").fillna({
        "Buy_Qty":0,"Sell_Qty":0,"Net_Pos_Qty":0,"Net_Neg_Qty":0,"Net_Pos_Ratio":0,
        "Top3_Pos_Ratio":0,"Broker_Concentration":0,"Active_Brokers":0
    })

    if not sector_map.empty:
        df = df.merge(sector_map, on="Symbol", how="left")
    else:
        df["Sector"] = ""

    # 1) Momentum (0..25)
    r = df["Ret_W"].fillna(0.0).values
    r_scaled = (clamp(r, -0.10, 0.15) + 0.10) / 0.25
    df["Score_Momentum"] = (25 * r_scaled).round(2)

    # 2) Volume confirmation (0..20)
    vs = df["Vol_Surge"].replace([np.inf,-np.inf], np.nan).fillna(1.0).values
    vs_scaled = clamp((vs - 0.5) / 1.5, 0, 1)
    df["Score_Volume"] = (20 * vs_scaled).round(2)

    # 3) Flow / accumulation (0..25) with concentration penalty
    npr = df["Net_Pos_Ratio"].fillna(0.0).values
    conc = df["Broker_Concentration"].fillna(0.0).values
    conc_pen = clamp((conc - 0.60) / 0.30, 0, 1)
    df["Score_Flow"] = (25 * npr * (1 - 0.6 * conc_pen)).round(2)

    # 4) VWAP + Close strength (0..15)
    cgv = df["Close_gt_VWAP"].fillna(0).values
    csi = df["CSI"].replace([np.inf,-np.inf], np.nan).fillna(0.5).values
    csi_scaled = clamp((csi - 0.3) / 0.5, 0, 1)
    df["Score_VWAP_Close"] = (15 * (0.55 * cgv + 0.45 * csi_scaled)).round(2)

    # 5) Liquidity (0..10)
    lt = np.log1p(df["AvgTurnover"].fillna(0.0).values)
    mx = np.nanmax(lt) if len(lt) else 0
    lt_scaled = (lt / mx) if mx and mx > 0 else np.zeros_like(lt)
    df["Score_Liquidity"] = (10 * lt_scaled).round(2)

    # ---- Risk penalties
    trapA = ((df["Close_gt_VWAP"] == 0) & (df["Vol_Surge"].fillna(1.0) >= 1.6)).astype(int)
    trapB = ((df["Ret_W"].fillna(0.0) > 0.04) & (df["Vol_Surge"].fillna(1.0) < 0.9)).astype(int)
    trapC = ((df["Broker_Concentration"].fillna(0.0) > 0.75) & (df["Ret_W"].fillna(0.0) > 0.03)).astype(int)

    churn = df["SumVol"].fillna(0.0) / df["AvgRange"].replace(0, np.nan)
    churn_thr = np.nanpercentile(churn.replace([np.inf,-np.inf], np.nan).fillna(0.0), 85) if len(churn) else 0
    trapD = (churn.replace([np.inf,-np.inf], np.nan).fillna(0.0) > churn_thr).astype(int)

    df["Penalty_Risk"] = -(15*trapA + 20*trapB + 15*trapC + 10*trapD)

    df["Score"] = (
        df["Score_Momentum"]
        + df["Score_Volume"]
        + df["Score_Flow"]
        + df["Score_VWAP_Close"]
        + df["Score_Liquidity"]
        + df["Penalty_Risk"]
    ).round(2)

    def action_from_score(x):
        if x >= 75:
            return "BUY"
        if x >= 60:
            return "HOLD/ACCUMULATE"
        if x >= 45:
            return "HOLD/WAIT"
        if x >= 30:
            return "EXIT ON BOUNCE"
        return "AVOID/EXIT"

    df["Action"] = df["Score"].map(action_from_score)
    df.insert(0, "Window", wname)

    notes = []
    for _, row in df.iterrows():
        n = []
        n.append("Close>VWAP" if row.get("Close_gt_VWAP", 0) == 1 else "Close<VWAP")
        if row.get("Vol_Surge", 1.0) >= 1.5:
            n.append("VolSurge")
        if row.get("Net_Pos_Ratio", 0.0) >= 0.60:
            n.append("Accumulation")
        if row.get("Broker_Concentration", 0.0) >= 0.75:
            n.append("HighConcentration")
        if row.get("Penalty_Risk", 0) < 0:
            n.append("RiskPenalty")
        notes.append(", ".join(n))
    df["Notes"] = notes

    cols = [
        "Window","Symbol","Sector","Action","Score",
        "Score_Momentum","Score_Volume","Score_Flow","Score_VWAP_Close","Score_Liquidity","Penalty_Risk",
        "Close","VWAP","Close_vs_VWAP_pct","Ret1D","Ret_W","Vol","AvgVol","Vol_Surge","Turnover","AvgTurnover","SumTurnover",
        "Net_Pos_Ratio","Broker_Concentration","Active_Brokers","CSI","Notes"
    ]
    for c in cols:
        if c not in df.columns:
            df[c] = np.nan
    return df[cols].sort_values(["Window","Score"], ascending=[True, False]).reset_index(drop=True)


# -----------------------------
# Excel writer helpers
# -----------------------------
def style_sheet(ws, freeze_row=1, freeze_col=0):
    ws.freeze_panes = ws.cell(row=freeze_row+1, column=freeze_col+1)
    ws.sheet_view.showGridLines = False

def autofit(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is None:
                continue
            v = str(cell.value)
            max_len = max(max_len, len(v))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 45)

def add_table(ws, df: pd.DataFrame, name: str):
    if df.empty:
        return
    end_row = ws.max_row
    end_col = ws.max_column
    ref = f"A1:{get_column_letter(end_col)}{end_row}"
    tab = Table(displayName=name, ref=ref)
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    tab.tableStyleInfo = style
    ws.add_table(tab)

def write_df(ws, df: pd.DataFrame, table_name: str | None = None):
    if df is None:
        df = pd.DataFrame()

    if df.empty:
        ws["A1"] = "No data available."
        style_sheet(ws, freeze_row=0, freeze_col=0)
        return

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = border
            if cell.row > 1:
                cell.alignment = Alignment(vertical="center", wrap_text=False)

    style_sheet(ws, freeze_row=1, freeze_col=0)
    autofit(ws)

    if table_name:
        tname = re.sub(r"[^A-Za-z0-9_]", "_", table_name)[:250]
        add_table(ws, df, tname)

def add_score_formatting(ws, score_col_letter: str, action_col_letter: str):
    ws.conditional_formatting.add(
        f"{score_col_letter}2:{score_col_letter}{ws.max_row}",
        ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                       mid_type="num", mid_value=50, mid_color="FFEB84",
                       end_type="num", end_value=100, end_color="63BE7B")
    )
    ws.conditional_formatting.add(
        f"{action_col_letter}2:{action_col_letter}{ws.max_row}",
        FormulaRule(formula=[f'${action_col_letter}2="BUY"'],
                    fill=PatternFill("solid", fgColor="C6EFCE"))
    )
    ws.conditional_formatting.add(
        f"{action_col_letter}2:{action_col_letter}{ws.max_row}",
        FormulaRule(formula=[f'OR(${action_col_letter}2="HOLD/ACCUMULATE",${action_col_letter}2="HOLD/WAIT")'],
                    fill=PatternFill("solid", fgColor="FFEB9C"))
    )
    ws.conditional_formatting.add(
        f"{action_col_letter}2:{action_col_letter}{ws.max_row}",
        FormulaRule(formula=[f'OR(${action_col_letter}2="EXIT ON BOUNCE",${action_col_letter}2="AVOID/EXIT")'],
                    fill=PatternFill("solid", fgColor="FFC7CE"))
    )

def col_letter_by_header(ws, header_name: str) -> str | None:
    for cell in ws[1]:
        if str(cell.value).strip() == header_name:
            return get_column_letter(cell.column)
    return None


# -----------------------------
# Trade setups builder
# -----------------------------
def build_trade_setups(scorecard_all: pd.DataFrame) -> pd.DataFrame:
    if scorecard_all.empty:
        return pd.DataFrame()

    out_rows = []
    for wname, _ in WINDOWS:
        sub = scorecard_all[scorecard_all["Window"] == wname].copy()
        if sub.empty:
            continue

        # Retail-friendly gating (avoid obvious traps)
        # - Prefer: Close>VWAP, Vol_Surge >= 1.1, Penalty not too severe
        good = sub[
            (sub["Action"].isin(["BUY", "HOLD/ACCUMULATE"]))
            & (sub["Close_gt_VWAP"].fillna(0) == 1)
            & (sub["Vol_Surge"].fillna(1.0) >= 1.1)
            & (sub["Penalty_Risk"].fillna(0) >= -20)
        ].copy()

        # If too few, relax a bit (still not allowing extreme risk)
        if len(good) < 12:
            good = sub[
                (sub["Action"].isin(["BUY", "HOLD/ACCUMULATE", "HOLD/WAIT"]))
                & (sub["Vol_Surge"].fillna(1.0) >= 0.9)
                & (sub["Penalty_Risk"].fillna(0) >= -30)
            ].copy()

        # pick top setups
        topN = 15 if wname in ("1D", "7D", "15D") else 20
        good = good.sort_values("Score", ascending=False).head(topN)

        for _, r in good.iterrows():
            plan = []
            if r.get("Action") == "BUY":
                plan.append("Entry: Break above day-high / pullback to VWAP")
                plan.append("Stop: Below VWAP or last swing low")
                plan.append("Add: If Vol keeps > avg")
            elif r.get("Action") == "HOLD/ACCUMULATE":
                plan.append("Entry: Small batches near VWAP / support")
                plan.append("Stop: Close below VWAP with volume")
            else:
                plan.append("Wait for confirmation (Close>VWAP + volume)")

            out_rows.append({
                "Window": wname,
                "Symbol": r.get("Symbol"),
                "Sector": r.get("Sector", ""),
                "Score": r.get("Score"),
                "Action": r.get("Action"),
                "Ret_W": r.get("Ret_W"),
                "Vol_Surge": r.get("Vol_Surge"),
                "Net_Pos_Ratio": r.get("Net_Pos_Ratio"),
                "Broker_Concentration": r.get("Broker_Concentration"),
                "RiskPenalty": r.get("Penalty_Risk"),
                "QuickPlan": " | ".join(plan),
                "Notes": r.get("Notes", "")
            })

    df = pd.DataFrame(out_rows)
    if df.empty:
        return df

    df = df.sort_values(["Window", "Score"], ascending=[True, False]).reset_index(drop=True)
    return df


# -----------------------------
# Build all outputs
# -----------------------------
def main():
    REPORT_DIR.mkdir(parents=True, exist_ok=True)

    floor_list = list_floor_files()
    price_list = list_price_files()

    if not floor_list:
        raise SystemExit("❌ No floorsheet_YYYY-MM-DD.csv found. Checked: outputs/, outputs/Floor Sheet/, outputs/FloorSheet/")
    if not price_list:
        raise SystemExit("❌ No SharePrice_YYYY-MM-DD.csv found in outputs/sharesansar/")

    floor_map = {d: p for d, p in floor_list}
    price_map = {d: p for d, p in price_list}

    common_dates = latest_common_dates(floor_list, price_list)
    if len(common_dates) < 2:
        raise SystemExit("❌ Not enough common dates between floorsheet and shareprice files.")

    latest_date = common_dates[-1]
    latest_date_str = latest_date.strftime("%Y-%m-%d")

    sector_map = load_sector_master()
    broker_master = load_broker_master()

    panel = build_price_panel(common_dates, price_map)

    max_w = max(w for _, w in WINDOWS)
    use_dates = common_dates[-max_w:] if len(common_dates) > max_w else common_dates[:]

    floors_by_date = {}
    for d in use_dates:
        floors_by_date[d] = load_floorsheet(floor_map[d])

    scorecards = []
    market_overview_rows = []
    risk_rows = []
    broker_summary_all = []
    broker_by_symbol_all = []
    smart_money_all = []
    trap_all = []
    price_movers_all = []
    window_top_sheets: list[tuple[str, pd.DataFrame]] = []

    # ---------- window loop ----------
    for wname, wlen in WINDOWS:
        w_dates = common_dates[-wlen:] if len(common_dates) >= wlen else common_dates[:]
        w_floor = pd.concat([floors_by_date[d] for d in w_dates if d in floors_by_date], ignore_index=True)

        bs_net = broker_symbol_net(w_floor)
        flow_feat = symbol_flow_features(bs_net)
        price_feat = window_price_features(panel, latest_date, wlen)

        sc = score_symbols(price_feat, flow_feat, sector_map, wname)
        scorecards.append(sc)

        # MARKET OVERVIEW
        m = sc.copy()
        total_turnover = float(m["SumTurnover"].fillna(0).sum())
        pct_above_vwap = float((m["Close"].fillna(0) > m["VWAP"].fillna(np.inf)).mean() * 100)
        pct_acc = float((m["Net_Pos_Ratio"].fillna(0) >= 0.60).mean() * 100)
        adv = int((m["Ret1D"].fillna(0) > 0).sum())
        dec = int((m["Ret1D"].fillna(0) < 0).sum())

        med_score = float(np.nanmedian(m["Score"].values)) if len(m) else 0.0
        if pct_above_vwap >= 55 and pct_acc >= 35:
            regime = "ACCUMULATION / HEALTHY"
            risk = "LOW-MODERATE"
        elif pct_above_vwap < 45 and pct_acc < 25:
            regime = "DISTRIBUTION / WEAK"
            risk = "HIGH"
        else:
            regime = "MIXED / SELECTIVE"
            risk = "MODERATE"

        market_overview_rows.append({
            "Window": wname,
            "TradingDaysUsed": len(w_dates),
            "TotalTurnover": total_turnover,
            "%CloseAboveVWAP": round(pct_above_vwap, 2),
            "%AccumulationSymbols(Net_Pos_Ratio>=0.60)": round(pct_acc, 2),
            "Advancers": adv,
            "Decliners": dec,
            "MedianScore": round(med_score, 2),
            "MarketRegime": regime,
            "RiskLevel": risk
        })

        risk_rows.append({
            "Window": wname,
            "Regime": regime,
            "RiskLevel": risk,
            "RetailRule": "If WEAK: avoid fresh BUY; trade only best setups or protect capital."
        })

        # SMART MONEY
        sm = sc[["Window","Symbol","Sector","Score","Action","Net_Pos_Ratio","Broker_Concentration","Vol_Surge","Ret_W","Close_gt_VWAP","CSI","Notes"]].copy()
        smart_money_all.append(sm)

        # TRAPS
        traps = sc[sc["Penalty_Risk"] < 0].copy()
        traps["RiskLevel"] = np.where(traps["Penalty_Risk"] <= -30, "HIGH",
                               np.where(traps["Penalty_Risk"] <= -15, "MEDIUM", "LOW"))
        traps = traps[["Window","Symbol","Sector","Score","Action","Penalty_Risk","RiskLevel","Notes"]].sort_values(["Window","Penalty_Risk"])
        trap_all.append(traps)

        # PRICE MOVERS
        pm = sc[["Window","Symbol","Sector","Ret1D","Ret_W","Vol_Surge","Close","VWAP","Score","Action"]].copy()
        pm["Direction"] = np.where(pm["Ret1D"].fillna(0) > 0, "GAINER",
                            np.where(pm["Ret1D"].fillna(0) < 0, "LOSER", "FLAT"))
        pm = pm.sort_values(["Window","Ret1D"], ascending=[True, False])
        price_movers_all.append(pm)

        # BROKER outputs
        if not bs_net.empty:
            bsum = bs_net.groupby("Broker", as_index=False).agg(
                Buy_Qty=("Buy_Qty","sum"),
                Sell_Qty=("Sell_Qty","sum"),
                Net_Qty=("Net_Qty","sum"),
                Buy_Amt=("Buy_Amt","sum"),
                Sell_Amt=("Sell_Amt","sum"),
                Net_Amt=("Net_Amt","sum"),
                Symbols=("Symbol","nunique"),
            )
            bsum.insert(0, "Window", wname)
            if not broker_master.empty:
                bsum = bsum.merge(broker_master, on="Broker", how="left")
                bsum["BrokerName"] = bsum["BrokerName"].fillna("")
                bsum["BrokerType"] = bsum["BrokerType"].fillna("")
            else:
                bsum["BrokerName"] = ""
                bsum["BrokerType"] = ""

            bsum = bsum.sort_values(["Window","Net_Qty"], ascending=[True, False])
            broker_summary_all.append(bsum)

            # Broker-by-symbol Top lists
            tN = topn_for_window(wname)

            bb = bs_net.sort_values(["Symbol","Net_Qty"], ascending=[True, False]).copy()
            top_buy = bb.groupby("Symbol").head(tN).copy()
            top_buy["Side"] = "TopNetBuyers"

            bb2 = bs_net.sort_values(["Symbol","Net_Qty"], ascending=[True, True]).copy()
            top_sell = bb2.groupby("Symbol").head(tN).copy()
            top_sell["Side"] = "TopNetSellers"

            bbs = pd.concat([top_buy, top_sell], ignore_index=True)
            bbs.insert(0, "Window", wname)

            if not broker_master.empty:
                bbs = bbs.merge(broker_master, on="Broker", how="left")
                bbs["BrokerName"] = bbs["BrokerName"].fillna("")
                bbs["BrokerType"] = bbs["BrokerType"].fillna("")
            else:
                bbs["BrokerName"] = ""
                bbs["BrokerType"] = ""

            broker_by_symbol_all.append(bbs)

            # Window-specific tops
            sm_top = sc.sort_values("Score", ascending=False).head(20)
            window_top_sheets.append((f"SM_TOP_{wname}", sm_top[["Symbol","Sector","Score","Action","Ret_W","Vol_Surge","Net_Pos_Ratio","Broker_Concentration","Notes"]]))

            sb_top = bsum.sort_values("Net_Qty", ascending=False).head(15)
            window_top_sheets.append((f"SB_TOP_{wname}", sb_top[["Broker","BrokerName","BrokerType","Net_Qty","Buy_Qty","Sell_Qty","Symbols"]]))

            # INST / OPR tops if broker type exists
            has_types = ("BrokerType" in bsum.columns) and (bsum["BrokerType"].astype(str).str.strip().str.len().sum() > 0)
            if has_types:
                inst = bsum[bsum["BrokerType"].str.upper().str.contains("INST|INSTIT", na=False)].copy()
                opr = bsum[bsum["BrokerType"].str.upper().str.contains("OPR|OPERAT", na=False)].copy()

                if not inst.empty:
                    inst_top = inst.sort_values("Net_Qty", ascending=False).head(15)
                    window_top_sheets.append((f"INST_TOP_{wname}", inst_top[["Broker","BrokerName","BrokerType","Net_Qty","Buy_Qty","Sell_Qty","Symbols"]]))

                if not opr.empty:
                    opr_top = opr.sort_values("Net_Qty", ascending=False).head(15)
                    window_top_sheets.append((f"OPR_TOP_{wname}", opr_top[["Broker","BrokerName","BrokerType","Net_Qty","Buy_Qty","Sell_Qty","Symbols"]]))

    # ---------- combine ----------
    scorecard_all = pd.concat(scorecards, ignore_index=True) if scorecards else pd.DataFrame()
    market_overview = pd.DataFrame(market_overview_rows)
    risk_grid = pd.DataFrame(risk_rows)

    smart_money_df = pd.concat(smart_money_all, ignore_index=True) if smart_money_all else pd.DataFrame()
    traps_df = pd.concat(trap_all, ignore_index=True) if trap_all else pd.DataFrame()
    price_movers_df = pd.concat(price_movers_all, ignore_index=True) if price_movers_all else pd.DataFrame()

    broker_summary_df = pd.concat(broker_summary_all, ignore_index=True) if broker_summary_all else pd.DataFrame()
    broker_by_symbol_df = pd.concat(broker_by_symbol_all, ignore_index=True) if broker_by_symbol_all else pd.DataFrame()

    trade_setups_df = build_trade_setups(scorecard_all)

    # Sector summary
    sector_summary_df = pd.DataFrame()
    if not scorecard_all.empty and "Sector" in scorecard_all.columns and scorecard_all["Sector"].astype(str).str.strip().ne("").any():
        sector_summary_df = scorecard_all.groupby(["Window","Sector"], as_index=False).agg(
            Symbols=("Symbol","nunique"),
            AvgScore=("Score","mean"),
            BuyCount=("Action", lambda s: int((s=="BUY").sum())),
            HoldAccCount=("Action", lambda s: int((s=="HOLD/ACCUMULATE").sum())),
            AvgRetW=("Ret_W","mean"),
            AvgVolSurge=("Vol_Surge","mean"),
            AvgAcc=("Net_Pos_Ratio","mean"),
        ).sort_values(["Window","AvgScore"], ascending=[True, False]).reset_index(drop=True)

    # Inputs summary
    inputs_summary = pd.DataFrame([{
        "LatestDate": latest_date_str,
        "CommonTradingDays": len(common_dates),
        "FloorsheetFiles": len(floor_list),
        "SharePriceFiles": len(price_list),
        "SectorMasterExists": bool(SECTOR_MASTER.exists()),
        "BrokerMasterExists": bool(BROKER_MASTER.exists()),
        "Windows": ", ".join([w for w,_ in WINDOWS]),
        "TopNRule": "1D=4, 7D=5, 15D+=10"
    }])

    # README
    readme = pd.DataFrame([
        {"Section":"How to use", "Details":"Open SYMBOL_SCORECARD → filter by Window and Action. Start from TRADE_SETUPS for ready ideas."},
        {"Section":"Interpret Action", "Details":"BUY: best momentum+flow+volume. HOLD/ACCUMULATE: build position carefully. EXIT/AVOID: protect capital."},
        {"Section":"Retail risk control", "Details":"Always use stops. Avoid HighConcentration + RiskPenalty combos unless you are a short-term trader."},
        {"Section":"Top sheets", "Details":"SM_TOP_* = top symbols. SB_TOP_* = top brokers. INST/OPR sheets appear if broker_master has types."},
        {"Section":"Note", "Details":"This report intentionally does NOT create Market_Breadth sheet."},
    ])

    # ---------- Write workbook ----------
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    def add_sheet(name: str) -> any:
        # Excel sheet name max 31 chars
        safe = name[:31]
        return wb.create_sheet(safe)

    ws = add_sheet("README")
    write_df(ws, readme, "README")

    ws = add_sheet("INPUTS_SUMMARY")
    write_df(ws, inputs_summary, "INPUTS_SUMMARY")

    ws = add_sheet("MARKET_OVERVIEW")
    write_df(ws, market_overview, "MARKET_OVERVIEW")

    ws = add_sheet("RISK_GRID")
    write_df(ws, risk_grid, "RISK_GRID")

    ws = add_sheet("SYMBOL_SCORECARD")
    write_df(ws, scorecard_all, "SYMBOL_SCORECARD")

    # Apply score formatting on SYMBOL_SCORECARD
    score_col = col_letter_by_header(ws, "Score")
    action_col = col_letter_by_header(ws, "Action")
    if score_col and action_col and ws.max_row >= 2:
        add_score_formatting(ws, score_col, action_col)

    ws = add_sheet("TRADE_SETUPS")
    write_df(ws, trade_setups_df, "TRADE_SETUPS")

    if not sector_summary_df.empty:
        ws = add_sheet("SECTOR_SUMMARY")
        write_df(ws, sector_summary_df, "SECTOR_SUMMARY")

    ws = add_sheet("SMART_MONEY")
    write_df(ws, smart_money_df, "SMART_MONEY")

    ws = add_sheet("TRAP_WARNINGS")
    write_df(ws, traps_df, "TRAP_WARNINGS")

    ws = add_sheet("PRICE_MOVERS")
    write_df(ws, price_movers_df, "PRICE_MOVERS")

    if not broker_summary_df.empty:
        ws = add_sheet("BROKER_SUMMARY")
        write_df(ws, broker_summary_df, "BROKER_SUMMARY")

    if not broker_by_symbol_df.empty:
        ws = add_sheet("BROKER_BY_SYMBOL")
        write_df(ws, broker_by_symbol_df, "BROKER_BY_SYMBOL")

    # Add window top sheets
    # (Keep total sheets reasonable—still fine for your use, but Excel can slow if too many.)
    for sname, sdf in window_top_sheets:
        ws = add_sheet(sname)
        write_df(ws, sdf, sname)

    report_path = REPORT_DIR / f"Advanced_NEPSERetailPro_Dashboard_{latest_date_str}.xlsx"
    wb.save(report_path)

    print(f"✅ Report saved: {report_path}")


if __name__ == "__main__":
    main()
