import json
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUTPUTS = ROOT / "outputs"      # where your daily XLSX files are
DOCS = ROOT / "docs"
DATA = DOCS / "data"

DATA.mkdir(parents=True, exist_ok=True)

# Read all daily XLSX floorsheets
xlsx_files = sorted(OUTPUTS.glob("floorsheet_*.xlsx"))
if not xlsx_files:
    raise SystemExit("No floorsheet_*.xlsx found in outputs/")

manifest = {"files": [], "latest": None}

for f in xlsx_files:
    # floorsheet_YYYY-MM-DD.xlsx -> YYYY-MM-DD
    date = f.stem.replace("floorsheet_", "")
    csv_name = f"floorsheet_{date}.csv"
    csv_path = DATA / csv_name

    # Convert XLSX → CSV (dashboard can read CSV)
    df = pd.read_excel(f)
    df.to_csv(csv_path, index=False)

    manifest["files"].append({
        "date": date,
        "file": f"data/{csv_name}"
    })

manifest["latest"] = manifest["files"][-1]["date"]

# Save index.json (dashboard uses this for multi-day)
(DATA / "index.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")

print(f"Published {len(xlsx_files)} day(s). Latest: {manifest['latest']}")
