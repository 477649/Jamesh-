import json
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]   # repo root
OUTPUTS = ROOT / "outputs"
DATA = ROOT / "docs" / "data"
DATA.mkdir(parents=True, exist_ok=True)

INDEX_PATH = DATA / "index.json"

# -------------------------------------------------
# 1) Find latest XLSX only
# -------------------------------------------------
xlsx_files = sorted(OUTPUTS.glob("floorsheet_*.xlsx"))
if not xlsx_files:
    raise SystemExit("No floorsheet_*.xlsx found in outputs/")

latest_xlsx = xlsx_files[-1]
latest_date = latest_xlsx.stem.replace("floorsheet_", "").strip()
csv_name = f"floorsheet_{latest_date}.csv"
csv_path = DATA / csv_name

# -------------------------------------------------
# 2) Convert XLSX -> CSV (always overwrite latest)
# -------------------------------------------------
df = pd.read_excel(latest_xlsx)

rename_map = {
    "Transact No": "Transact No.",
    "Transact": "Transact No.",
    "Quantity ": "Quantity",
    "Rate ": "Rate",
    "Amount ": "Amount",
}
df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, inplace=True)

required = {"Symbol", "Buyer", "Seller", "Quantity", "Rate", "Amount"}
missing = required - set(df.columns)
if missing:
    raise SystemExit(f"Missing required columns in {latest_xlsx.name}: {sorted(missing)}")

df.to_csv(csv_path, index=False)

# -------------------------------------------------
# 3) Write index.json (LATEST DAY ONLY)
# -------------------------------------------------
manifest = {
    "files": [
        {
            "date": latest_date,
            "file": f"data/{csv_name}"
        }
    ],
    "latest": latest_date
}

INDEX_PATH.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

# -------------------------------------------------
# 4) Cleanup: remove all other CSVs
# -------------------------------------------------
for csv in DATA.glob("floorsheet_*.csv"):
    if csv.name != csv_name:
        csv.unlink(missing_ok=True)

print("✅ Published latest day only")
print(f"   Date   : {latest_date}")
print(f"   CSV    : {csv_name}")
print(f"   Index  : {INDEX_PATH}")
