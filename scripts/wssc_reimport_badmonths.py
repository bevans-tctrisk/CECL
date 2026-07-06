"""Re-import WSSC 2024-01, 2024-02, 2025-05 from source (fixes: adds CUMA
mortgages for Jan/Feb 2024, imports the full May-2025 '(1)' extract set, and the
wholesale per-snapshot wipe clears the bogus $15.6M 'Exclude' totals row).

Stages the 12 source files (4 extracts x 3 months) in a temp folder, points the
config's loan_file_folder at it with archiving OFF, runs process_client, then
restores the config."""
import os, shutil, datetime, importlib
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
WS = os.environ["CECL_WORKSPACE_ROOT"]
import yaml

SRC = r"Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients\WSSC FCU 16268\Portfolio Management\CECL-Migration IDLR"
FILES = [
    r"2024\Jan 2024\2024.01.31 Loan ARIES.xlsx",
    r"2024\Jan 2024\2024.01 Credit Card ARIES.xls",
    r"2024\Jan 2024\2024.01 ODP Accounts.xlsx",
    r"2024\Jan 2024\2024.01 CUMA CECL Report V2.xlsx",
    r"2024\Feb 2024\2024.02.29 Loan ARIES.xlsx",
    r"2024\Feb 2024\2024.02 Credit Card ARIES.xls",
    r"2024\Feb 2024\2024.02 ODP Accounts.xlsx",
    r"2024\Feb 2024\2024.02 CUMA CECL Report V2.xlsx",
    r"2025\May 2025\2025.05.31 Loan ARIES (1).xlsx",
    r"2025\May 2025\2025.05 Credit Card ARIES (1).xls",
    r"2025\May 2025\2025.05.31 ODP Accounts (1).xlsx",
    r"2025\May 2025\2025.05 CUMA CECL Report (1).xlsx",
]
STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
TEMP = os.path.join(os.environ.get("TEMP", r"C:\Temp"), f"wssc_reimport_{STAMP}")
os.makedirs(TEMP, exist_ok=True)

print("Staging 12 files ->", TEMP)
for rel in FILES:
    src = os.path.join(SRC, rel)
    if not os.path.isfile(src):
        raise SystemExit("MISSING source file: " + src)
    shutil.copy2(src, os.path.join(TEMP, os.path.basename(src)))

cpath = os.path.join(WS, "client_configs", "wssc_fcu.yaml")
shutil.copy2(cpath, cpath + f".bak.phase_reimport_{STAMP}")
cfg = yaml.safe_load(open(cpath, encoding="utf-8"))
orig_has_lff = "loan_file_folder" in cfg
orig_has_arch = "archive_imported_files" in cfg
cfg["loan_file_folder"] = TEMP
cfg["archive_imported_files"] = False
with open(cpath, "w", encoding="utf-8") as fh:
    yaml.safe_dump(cfg, fh, sort_keys=False, default_flow_style=False, allow_unicode=True)

try:
    import_data = importlib.import_module("import_data")
    import_data.process_client("wssc_fcu")
finally:
    # Restore config: drop the temp overrides.
    cfg2 = yaml.safe_load(open(cpath, encoding="utf-8"))
    if not orig_has_lff:
        cfg2.pop("loan_file_folder", None)
    if not orig_has_arch:
        cfg2.pop("archive_imported_files", None)
    with open(cpath, "w", encoding="utf-8") as fh:
        yaml.safe_dump(cfg2, fh, sort_keys=False, default_flow_style=False, allow_unicode=True)
    shutil.rmtree(TEMP, ignore_errors=True)
    print("Restored config, removed temp folder.")
