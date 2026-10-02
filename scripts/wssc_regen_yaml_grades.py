"""Regenerate WSSC YAML from the current draft so the new credit_grades (from the
Step 9 risk-tier upload) reach the config the report reads. Backs up first and
verifies no regression to pools / provenance."""
import os, shutil, datetime
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
WS = os.environ["CECL_WORKSPACE_ROOT"]
from cecl_ui.services import config_service, wizard_drafts

SHORT = "wssc_fcu"
STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
state = wizard_drafts.load_draft(WS, SHORT, model="migration")
if state is None:
    raise SystemExit("no draft")

cfg_path = config_service.configs_dir(WS) / f"{SHORT}.yaml"
if cfg_path.exists():
    shutil.copy2(cfg_path, str(cfg_path) + f".bak.phase_grades_upload_{STAMP}")

cfg = config_service.build_yaml_from_wizard(state)
config_service.save_client_config(WS, SHORT, cfg, overwrite=True)

print("credit_grades now in YAML:")
for g in cfg.get("credit_grades") or []:
    print("  ", g)
print("not_risk_rated:", cfg.get("not_risk_rated"))
print("hist_balance_provenance:", bool(cfg.get("hist_balance_provenance")))
print("co_recov_provenance:", list((cfg.get("co_recov_provenance") or {}).keys()))
print("no_score_label:", cfg.get("no_score_label"))
