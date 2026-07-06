"""Persist WSSC FCU charge-off / recovery provenance: compute the CO and
recovery provenance records, store them on the wizard draft, and regenerate
the YAML so the saved config carries the validator-facing trail. Idempotent;
backs up draft and YAML first."""
import os, shutil, datetime
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
WS = os.environ["CECL_WORKSPACE_ROOT"]
import yaml
from cecl_ui.services import config_service, wizard_drafts
from cecl_ui.routes.setup import _co_recov_provenance_view

SHORT = "wssc_fcu"
STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

state = wizard_drafts.load_draft(WS, SHORT, model="migration")
if state is None:
    raise SystemExit("No WSSC migration draft found")

state["co_recov_provenance"] = {
    "chargeoff": _co_recov_provenance_view(state, "co"),
    "recovery": _co_recov_provenance_view(state, "recov"),
}
print("CO method:", state["co_recov_provenance"]["chargeoff"].get("method"))
print("Recovery method:", state["co_recov_provenance"]["recovery"].get("method"))

draft_path = wizard_drafts._file_for(WS, SHORT, "migration")
if draft_path.exists():
    shutil.copy2(draft_path, str(draft_path) + f".bak.phase_co_recov_prov_{STAMP}")
wizard_drafts.save_draft(WS, state, active_step=state.get("_active_step") or "co_history", model="migration")
print("Draft saved.")

cfg_path = config_service.configs_dir(WS) / f"{SHORT}.yaml"
if cfg_path.exists():
    shutil.copy2(cfg_path, str(cfg_path) + f".bak.phase_co_recov_prov_{STAMP}")
cfg = config_service.build_yaml_from_wizard(state)
target = config_service.save_client_config(WS, SHORT, cfg, overwrite=True)
print("Config saved ->", target)
print("\nco_recov_provenance in YAML:")
print(yaml.safe_dump(cfg.get("co_recov_provenance", {}), sort_keys=False, allow_unicode=True))
