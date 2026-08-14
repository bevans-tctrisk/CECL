"""Regenerate WSSC FCU YAML from its draft so the persisted config includes the
hist_balance_provenance block. Idempotent; backs up the YAML first."""
import os, shutil, datetime
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
WS = os.environ["CECL_WORKSPACE_ROOT"]
import yaml
from cecl_ui.services import config_service, wizard_drafts

SHORT = "wssc_fcu"
STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

state = wizard_drafts.load_draft(WS, SHORT, model="migration")
if state is None:
    raise SystemExit("No WSSC migration draft found")

cfg_path = config_service.configs_dir(WS) / f"{SHORT}.yaml"
if cfg_path.exists():
    bak = str(cfg_path) + f".bak.phase_hist_provenance_yaml_{STAMP}"
    shutil.copy2(cfg_path, bak)
    print("YAML backup ->", bak)

cfg = config_service.build_yaml_from_wizard(state)
target = config_service.save_client_config(WS, SHORT, cfg, overwrite=True)
print("Config saved ->", target)
print("\nhist_balance_provenance in YAML:")
print(yaml.safe_dump(cfg.get("hist_balance_provenance", {}), sort_keys=False, allow_unicode=True))
