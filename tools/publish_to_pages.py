import json
from pathlib import Path
import pandas as pd
import shutil

ROOT = Path(__file__).resolve().parents[1]
OUTPUTS = ROOT / "outputs"
DOCS = ROOT / "docs"
DATA = DOCS / "data"
DATA.mkdir(parents=True, exist_ok=True)

# ✅ Your actual folder (from screenshot)
FLOOR_SHEET_DIR = OUTPUTS / "Floor Sheet"

# ✅ Read all floorsheet CSVs
files = sorted(FLOOR_SHEET_DIR.glob("floorsheet_*.csv"))

if not files:
    raise SystemExit(f"No floorsheet CSV files found in: {FLOOR_SHEET_DIR}")

manifest = {"files": [], "latest": None}

latest_csv_path = None

for f in files:
    date = f.stem.replace("floorsheet_", "")      # floorsheet_YYYY-MM-DD
    csv_name = f"floorsheet_{date}.csv"
    csv_path = DATA / csv_name

    # ✅ Copy directly (no need to read+write again)
    shutil.copyfile(f, csv_path)

    manifest["files"].append({
        "date": date,
        "file": f"data/{csv_name}"
    })

    latest_csv_path = csv_path

# ✅ latest date
manifest["latest"] = manifest["files"][-1]["date"]

# ✅ Always create a fixed latest file
if latest_csv_path:
    latest_fixed = DATA / "floorsheet_latest.csv"
    shutil.copyfile(latest_csv_path, latest_fixed)

(DATA / "index.json").write_text(
    json.dumps(manifest, indent=2),
    encoding="utf-8"
)

print("Published", len(files), "days")
print("Latest day:", manifest["latest"])
print("Updated dashboard file:", "docs/data/floorsheet_latest.csv")
