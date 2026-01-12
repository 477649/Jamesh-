import json
from pathlib import Path
import pandas as pd
import shutil

ROOT = Path(__file__).resolve().parents[1]
OUTPUTS = ROOT / "outputs"
DOCS = ROOT / "docs"
DATA = DOCS / "data"

DATA.mkdir(parents=True, exist_ok=True)

files = sorted(OUTPUTS.glob("floorsheet_*.xlsx"))
if not files:
    raise SystemExit("No floorsheet files found")

manifest = {"files": [], "latest": None}

latest_csv_path = None  # ✅ will store latest day's csv path

for f in files:
    date = f.stem.replace("floorsheet_", "")
    csv_name = f"floorsheet_{date}.csv"
    csv_path = DATA / csv_name

    df = pd.read_excel(f)
    df.to_csv(csv_path, index=False)

    manifest["files"].append({
        "date": date,
        "file": f"data/{csv_name}"
    })

    latest_csv_path = csv_path  # ✅ keep updating; last one becomes latest

# ✅ latest date (already your logic)
manifest["latest"] = manifest["files"][-1]["date"]

# ✅ NEW: create "today only" file for dashboard
# This will always represent the latest trading day
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
