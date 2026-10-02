"""Fix WSSC loan-extract file patterns so they tolerate the ' V2' version suffix
(CUMA Jan/Feb 2024) and the ' (1)' copy suffix (May 2025). Updates the draft
(top file_pattern + each loan_data_files extract) and regenerates the YAML.
Idempotent; backs up draft + YAML first."""
import os, shutil, datetime, re
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
WS = os.environ["CECL_WORKSPACE_ROOT"]
from cecl_ui.services import config_service, wizard_drafts

SHORT = "wssc_fcu"
STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
SUFFIX = r"(?:\s+V\d+)?(?:\s*\(\d+\))?"  # optional ' V2' and/or ' (1)'

def loosen(pat):
    """Insert the optional version/copy suffix before the extension anchor,
    unless the pattern already tolerates trailing text (has '.*') or is empty."""
    if not pat:
        return pat
    tail = r"\.(xlsx|xls)$"
    if ".*" + tail in pat:
        return pat  # already tolerant (e.g. Credit Card)
    if pat.endswith(tail) and SUFFIX not in pat:
        return pat[: -len(tail)] + SUFFIX + tail
    return pat

state = wizard_drafts.load_draft(WS, SHORT, model="migration")
if state is None:
    raise SystemExit("no draft")

print("BEFORE -> AFTER")
old_top = state.get("file_pattern")
state["file_pattern"] = loosen(old_top)
print(f"top: {old_top}\n  -> {state['file_pattern']}\n")
for e in (state.get("sample_uploads") or {}).get("loan_data_files") or []:
    op = e.get("file_pattern")
    e["file_pattern"] = loosen(op)
    print(f"{e.get('name')}:\n  {op}\n  -> {e['file_pattern']}\n")

dpath = wizard_drafts._file_for(WS, SHORT, "migration")
if dpath.exists():
    shutil.copy2(dpath, str(dpath) + f".bak.phase_patternfix_{STAMP}")
wizard_drafts.save_draft(WS, state, active_step=state.get("_active_step") or "sample", model="migration")

cpath = config_service.configs_dir(WS) / f"{SHORT}.yaml"
if cpath.exists():
    shutil.copy2(cpath, str(cpath) + f".bak.phase_patternfix_{STAMP}")
cfg = config_service.build_yaml_from_wizard(state)
config_service.save_client_config(WS, SHORT, cfg, overwrite=True)
print("=== YAML loan_data_extracts patterns now ===")
print("top:", cfg.get("file_pattern"))
for e in cfg.get("loan_data_extracts") or []:
    print(f"  {e.get('label')}: {e.get('file_pattern')}")
