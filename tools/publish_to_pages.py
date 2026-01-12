# scripts/publish_pages_data.py
import json
from pathlib import Path
import pandas as pd

# ----------------- SETTINGS -----------------
KEEP_LAST_DAYS = None
# Examples:
# KEEP_LAST_DAYS = 400  # keep last 400 days only
# KEEP_LAST_DAYS = 120  # keep last 120 days
# KEEP_LAST_DAYS = None # keep all available

# If True: reconvert CSV even if already exists (slow)
FORCE_REBUILD_ALL = False
# --------------------------------------------

ROOT = Path(__file__).resolve().parents[1]   # repo root
OUTPUTS = ROOT / "outputs"
DATA = ROOT / "docs" / "data"
DATA.mkdir(parents=True, exist_ok=True)

INDEX_PATH = DATA / "index.json"

# Load existing index.json (so we don't rebuild everything every time)
existing = {"files": [], "latest": None}
if INDEX_PATH.exists():
    try:
        existing = json.loads(INDEX_PATH.read_text(encoding="utf-8") or "{}")
        if "files" not in existing:
            existing["files"] = []
    except Exception:
        existing = {"files": [], "latest": None}

existing_map = {x.get("date"): x.get("file") for x in existing.get("files", []) if x.get("date")}

# Find day-wise XLSX files
xlsx_files = sorted(OUTPUTS.glob("floorsheet_*.xlsx"))
if not xlsx_files:
    raise SystemExit("No floorsheet_*.xlsx found in outputs/")

def extract_date_from_name(path: Path) -> str:
    # floorsheet_YYYY-MM-DD.xlsx -> YYYY-MM-DD
    return path.stem.replace("floorsheet_", "").strip()

def needs_rebuild(xlsx_path: Path, csv_path: Path) -> bool:
    if FORCE_REBUILD_ALL:
        return True
    if not csv_path.exists():
        return True
    # If xlsx is newer than csv, rebuild
    return xlsx_path.stat().st_mtime > csv_path.stat().st_mtime

# Convert XLSX -> CSV (only if needed)
published = []
rebuilt_count = 0

for f in xlsx_files:
    date = extract_date_from_name(f)
    csv_name = f"floorsheet_{date}.csv"
    csv_path = DATA / csv_name

    if needs_rebuild(f, csv_path):
        df = pd.read_excel(f)

        # Optional: normalize columns (keeps your dashboard happy)
        # If your excel already has these exact names, no change happens.
        rename_map = {
            "Transact No": "Transact No.",
            "Transact": "Transact No.",
            "Quantity ": "Quantity",
            "Rate ": "Rate",
            "Amount ": "Amount",
        }
        df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, inplace=True)

        # Ensure required columns exist (fail early if your XLSX structure changed)
        required = {"Symbol", "Buyer", "Seller", "Quantity", "Rate", "Amount"}
        missing = required - set(df.columns)
        if missing:
            raise SystemExit(f"Missing required columns in {f.name}: {sorted(missing)}")

        df.to_csv(csv_path, index=False)
        rebuilt_count += 1

    published.append({"date": date, "file": f"data/{csv_name}"})

# Sort by date (string sort works for YYYY-MM-DD)
published.sort(key=lambda x: x["date"])

# Keep only last N days (optional)
if KEEP_LAST_DAYS is not None and KEEP_LAST_DAYS > 0:
    published = published[-KEEP_LAST_DAYS:]

manifest = {
    "files": published,
    "latest": published[-1]["date"] if published else None
}

INDEX_PATH.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

# Optional cleanup: delete CSVs not in manifest (when KEEP_LAST_DAYS is used)
keep_files = {Path(item["file"]).name for item in published}  # floorsheet_YYYY-MM-DD.csv
for csv in DATA.glob("floorsheet_*.csv"):
    if csv.name not in keep_files:
        csv.unlink(missing_ok=True)

print(f"Published {len(published)} day(s). Latest: {manifest['latest']}. Rebuilt CSVs: {rebuilt_count}")
print(f"Wrote: {INDEX_PATH}")
