"""Make the ODP Accounts pattern's day-of-month optional (Jan/Feb 2024 ODP files
are named 'YYYY.MM ODP Accounts.xlsx' with no day, while others have a day).
Updates draft + YAML, then verifies all target files match."""
import os, shutil, datetime, re
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
WS = os.environ["CECL_WORKSPACE_ROOT"]
from cecl_ui.services import config_service, wizard_drafts

SHORT = "wssc_fcu"
STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

def fix_odp(pat):
    if not pat or "ODP" not in pat:
        return pat
    # \.\d{2}\s+ODP  ->  (?:\.\d{2})?\s+ODP
    return pat.replace(r"\.\d{2}\s+ODP", r"(?:\.\d{2})?\s+ODP")

state = wizard_drafts.load_draft(WS, SHORT, model="migration")
for e in (state.get("sample_uploads") or {}).get("loan_data_files") or []:
    if "ODP" in (e.get("name") or "") or "ODP" in (e.get("file_pattern") or ""):
        e["file_pattern"] = fix_odp(e.get("file_pattern"))

dpath = wizard_drafts._file_for(WS, SHORT, "migration")
shutil.copy2(dpath, str(dpath) + f".bak.phase_odpdayopt_{STAMP}")
wizard_drafts.save_draft(WS, state, active_step=state.get("_active_step") or "sample", model="migration")
cpath = config_service.configs_dir(WS) / f"{SHORT}.yaml"
shutil.copy2(cpath, str(cpath) + f".bak.phase_odpdayopt_{STAMP}")
cfg = config_service.build_yaml_from_wizard(state)
config_service.save_client_config(WS, SHORT, cfg, overwrite=True)

pats = {e["label"]: e["file_pattern"] for e in cfg["loan_data_extracts"]}
print("ODP pattern now:", pats["ODP Accounts"])
targets = {
    "Loan ARIES": ["2024.01.31 Loan ARIES.xlsx", "2024.02.29 Loan ARIES.xlsx", "2025.05.31 Loan ARIES (1).xlsx"],
    "Credit Card AIRES": ["2024.01 Credit Card ARIES.xls", "2024.02 Credit Card ARIES.xls", "2025.05 Credit Card ARIES (1).xls"],
    "ODP Accounts": ["2024.01 ODP Accounts.xlsx", "2024.02 ODP Accounts.xlsx", "2025.05.31 ODP Accounts (1).xlsx"],
    "CUMA CECL Report": ["2024.01 CUMA CECL Report V2.xlsx", "2024.02 CUMA CECL Report V2.xlsx", "2025.05 CUMA CECL Report (1).xlsx"],
}
print("\nAll 12 target files:")
allok = True
for label, files in targets.items():
    rx = re.compile(pats[label])
    for f in files:
        ok = bool(rx.match(f)); allok &= ok
        print(f"  [{'OK ' if ok else 'FAIL'}] {label}: {f}")
print("\nALL MATCH:", allok)
