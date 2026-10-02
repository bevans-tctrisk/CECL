"""Check ER monthly_loan_data snapshots and Raw_Uploads contents."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from sqlalchemy import create_engine, text
from cecl_credentials import get_database_url
import yaml

CU = "Emergency Responders CU"
SHORT = "emergency_responders_cu"

eng = create_engine(get_database_url())
with eng.begin() as conn:
    rows = conn.execute(text(
        "SELECT snapshot_date, COUNT(*) FROM monthly_loan_data "
        "WHERE credit_union = :cu GROUP BY snapshot_date "
        "ORDER BY snapshot_date DESC LIMIT 10"
    ), {"cu": CU}).fetchall()
    print("=== monthly_loan_data snapshots for ER ===")
    for r in rows:
        print(r)

    # Also try ILIKE in case the cu name shifts
    rows2 = conn.execute(text(
        "SELECT DISTINCT credit_union FROM monthly_loan_data "
        "WHERE credit_union ILIKE '%emergency%'"
    )).fetchall()
    print("\n=== distinct ER-ish names ===")
    for r in rows2:
        print(r)

# Raw_Uploads contents
ws_yaml = Path(r"Z:\Shared\TCT Files\CECL - CM Files")
upload_dir = ws_yaml / "Raw_Uploads" / SHORT
print(f"\n=== Files in {upload_dir} ===")
if upload_dir.exists():
    for p in sorted(upload_dir.iterdir()):
        print(f"  {p.name}  ({p.stat().st_size} bytes)")
else:
    print("  (directory does not exist)")

# Look at YAML config
cfg_path = ws_yaml / "client_configs" / f"{SHORT}.yaml"
print(f"\n=== {cfg_path} ===")
if cfg_path.exists():
    cfg = yaml.safe_load(cfg_path.read_text())
    print(f"  report_period: {cfg.get('report_period')!r}")
    print(f"  file_pattern: {cfg.get('file_pattern')!r}")
    print(f"  charter_number: {cfg.get('charter_number')!r}")
    print(f"  credit_union: {cfg.get('credit_union')!r}")
    for i, ex in enumerate(cfg.get("loan_data_extracts") or []):
        print(f"  extract[{i}]: name={ex.get('name')!r} file_pattern={ex.get('file_pattern')!r}")
else:
    print("  (missing)")
