"""Re-import all WSSC 2024-2025 loan data so Option 3 (derive current score from
the member's most-recent loan) populates credit migration. Stages the 96 loan
extracts (4 x 24 months) into a temp folder, points the config there with
archiving OFF, runs process_client (wholesale DELETE per snapshot then insert),
restores the config, and removes the known bogus 2024-01 'Exclude' totals row."""
import os, re, shutil, datetime, importlib
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
WS = os.environ["CECL_WORKSPACE_ROOT"]
import yaml
from sqlalchemy import create_engine, text
from cecl_credentials import get_database_url

SRC = r"Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients\WSSC FCU 16268\Portfolio Management\CECL-Migration IDLR"
CU = "WSSC FCU"
STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
TEMP = os.path.join(os.environ.get("TEMP", r"C:\Temp"), f"wssc_reimport_all_{STAMP}")
os.makedirs(TEMP, exist_ok=True)

cpath = os.path.join(WS, "client_configs", "wssc_fcu.yaml")
cfg = yaml.safe_load(open(cpath, encoding="utf-8"))
pats = [re.compile(e["file_pattern"]) for e in cfg["loan_data_extracts"]]

n = 0
for yr in ("2024", "2025"):
    for dp, dn, fns in os.walk(os.path.join(SRC, yr)):
        for f in fns:
            if any(p.match(f) for p in pats):
                shutil.copy2(os.path.join(dp, f), os.path.join(TEMP, f))
                n += 1
print(f"Staged {n} loan files -> {TEMP}")

shutil.copy2(cpath, cpath + f".bak.phase_reimport_all_{STAMP}")
orig_lff = "loan_file_folder" in cfg
orig_arch = "archive_imported_files" in cfg
cfg["loan_file_folder"] = TEMP
cfg["archive_imported_files"] = False
with open(cpath, "w", encoding="utf-8") as fh:
    yaml.safe_dump(cfg, fh, sort_keys=False, default_flow_style=False, allow_unicode=True)

try:
    import_data = importlib.import_module("import_data")
    import_data.process_client("wssc_fcu")
finally:
    c2 = yaml.safe_load(open(cpath, encoding="utf-8"))
    if not orig_lff:
        c2.pop("loan_file_folder", None)
    if not orig_arch:
        c2.pop("archive_imported_files", None)
    with open(cpath, "w", encoding="utf-8") as fh:
        yaml.safe_dump(c2, fh, sort_keys=False, default_flow_style=False, allow_unicode=True)
    shutil.rmtree(TEMP, ignore_errors=True)
    print("Restored config, removed temp folder.")

# Remove the bogus 2024-01 Exclude totals row (member_number IS NULL).
eng = create_engine(get_database_url())
with eng.begin() as conn:
    res = conn.execute(text(
        "DELETE FROM monthly_loan_data WHERE credit_union=:c "
        "AND snapshot_date='2024-01-31' AND loan_pool IN ('Exclude','Ignore') "
        "AND member_number IS NULL"), {"c": CU})
    print("Deleted junk Exclude rows:", res.rowcount)
