import json
from pathlib import Path
import shutil
import re

ROOT = Path(__file__).resolve().parents[1]
OUTPUTS = ROOT / "outputs"
DOCS = ROOT / "docs"
DATA = DOCS / "data"
DATA.mkdir(parents=True, exist_ok=True)

# ✅ Your actual folder
FLOOR_SHEET_DIR = OUTPUTS / "Floor Sheet"

# ✅ Find all floorsheet CSVs
files = sorted(FLOOR_SHEET_DIR.glob("floorsheet_*.csv"))

if not files:
    raise SystemExit(f"No floorsheet CSV files found in: {FLOOR_SHEET_DIR}")

# ✅ Pick latest file (by filename sort: floorsheet_YYYY-MM-DD.csv)
latest_src = files[-1]
date = latest_src.stem.replace("floorsheet_", "")  # YYYY-MM-DD

# Optional: validate date format
if not re.match(r"^\d{4}-\d{2}-\d{2}$", date):
    raise SystemExit(f"Latest filename date is not YYYY-MM-DD: {latest_src.name}")

# ✅ (Optional but recommended) Remove old published daily CSVs in docs/data
# Keeps only floorsheet_latest.csv and index.json
for old in DATA.glob("floorsheet_*.csv"):
    if old.name != "floorsheet_latest.csv":
        old.unlink(missing_ok=True)

# ✅ Copy latest into fixed file
latest_fixed = DATA / "floorsheet_latest.csv"
shutil.copyfile(latest_src, latest_fixed)

# ✅ Create index.json with ONLY latest
manifest = {
    "files": [
        {"date": date, "file": "data/floorsheet_latest.csv"}
    ],
    "latest": date
}

(DATA / "index.json").write_text(
    json.dumps(manifest, indent=2),
    encoding="utf-8"
)

print("✅ Published ONLY latest day")
print("Latest day:", date)
print("Updated:", latest_fixed)
print("Updated:", DATA / "index.json")
