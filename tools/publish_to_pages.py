import json
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUTPUTS = ROOT / "outputs"
DOCS = ROOT / "docs"
DATA = DOCS / "data"

DATA.mkdir(parents=True, exist_ok=True)

# Read all daily XLSX floorsheets
files = sorted(OUTPUTS.glob("floorsheet_*.xlsx"))
if not files:
    raise SystemExit("No floorsheet XLSX files found.")

manifest = {"files": [], "latest": None}

for f in files:
    date = f.stem.replace("floorsheet_", "")
    csv_name = f"floorsheet_{date}.csv"
    csv_path = DATA / csv_name

    # Convert XLSX → CSV
    df = pd.read_excel(f)
    df.to_csv(csv_path, index=False)

    manifest["files"].append({
        "date": date,
        "file": f"data/{csv_name}"
    })

manifest["latest"] = manifest["files"][-1]["date"]

# Write index.json for dashboard
(DATA / "index.json").write_text(
    json.dumps(manifest, indent=2),
    encoding="utf-8"
)

print(f"Published {len(files)} days. Latest = {manifest['latest']}")
