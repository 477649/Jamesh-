import requests
import pandas as pd
from datetime import datetime
from pathlib import Path

# ================= CONFIG =================
BASE_URL = "https://www.nepalstock.com/api/nots/security/today-price"
OUTPUT_DIR = Path("outputs/share_price")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Today (NEPSE trading date)
TRADE_DATE = datetime.now().strftime("%Y-%m-%d")
# =========================================

headers = {
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json",
    "Origin": "https://www.nepalstock.com",
    "Referer": "https://www.nepalstock.com/today-price",
    "User-Agent": "Mozilla/5.0"
}

payload = {
    "size": 500,
    "page": 0,
    "date": TRADE_DATE
}

print(f"📥 Fetching NEPSE Today Price for {TRADE_DATE}")

resp = requests.post(BASE_URL, json=payload, headers=headers, timeout=30)
resp.raise_for_status()

data = resp.json().get("content", [])

if not data:
    raise ValueError("No data returned – market closed or API changed")

df = pd.DataFrame(data)

# ---------- Column cleanup ----------
rename_map = {
    "symbol": "Symbol",
    "closePrice": "Close",
    "openPrice": "Open",
    "highPrice": "High",
    "lowPrice": "Low",
    "totalTradedQuantity": "Total_Qty",
    "totalTrades": "Trades",
    "totalTradedValue": "Total_Value",
    "lastTradedPrice": "LTP",
    "previousClose": "Prev_Close",
    "averageTradedPrice": "Avg_Price",
    "fiftyTwoWeekHigh": "Week52_High",
    "fiftyTwoWeekLow": "Week52_Low",
    "marketCapitalization": "Market_Cap"
}

df = df.rename(columns=rename_map)

# ---------- Save ----------
out_file = OUTPUT_DIR / f"{TRADE_DATE}_today_price.csv"
df.to_csv(out_file, index=False)

print(f"✅ Saved: {out_file}")
