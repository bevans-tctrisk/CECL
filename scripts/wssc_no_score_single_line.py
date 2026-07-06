"""WSSC FCU: default pools with NO credit scores to single-line (not broken out
by grade). Based on 2025-12-31 audit, these pools have zero (or effectively zero)
scored balance -> risk_rated: false:

  - Credit Card         (0 / 210 loans scored)
  - Overdraft Privilege (0 /  81 loans scored)
  - Holiday 365         (1 / 425 loans scored -> $1,840 of $551k, 0.3%)

Updates BOTH the wizard draft (pool_settings) AND regenerates the YAML config via
the same serializer the wizard uses, so the change is consistent and sticks.
Idempotent; backs up draft and YAML first."""
import os, json, shutil, datetime
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
WS = os.environ["CECL_WORKSPACE_ROOT"]
from cecl_ui.services import config_service, wizard_drafts

SHORT = "wssc_fcu"
SINGLE_LINE = {"Credit Card", "Overdraft Privilege", "Holiday 365"}
STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

# ---- 1. Load draft state ----
state = wizard_drafts.load_draft(WS, SHORT, model="migration")
if state is None:
    raise SystemExit("No WSSC migration draft found")

# ---- 2. Flip risk_rated -> False for the no-score pools in pool_settings ----
changed = []
for p in state.get("pool_settings") or []:
    if p.get("name") in SINGLE_LINE and p.get("risk_rated") is not False:
        p["risk_rated"] = False
        p["brr"] = False
        changed.append(p["name"])
print("Flipped to single-line (risk_rated=False):", changed or "(already set)")

# ---- 3. Backup + save draft ----
draft_path = wizard_drafts._file_for(WS, SHORT, "migration")
if draft_path.exists():
    bak = str(draft_path) + f".bak.phase_no_score_single_line_{STAMP}"
    shutil.copy2(draft_path, bak)
    print("Draft backup ->", bak)
wizard_drafts.save_draft(WS, state, active_step=state.get("_active_step") or "review", model="migration")
print("Draft saved ->", draft_path)

# ---- 4. Regenerate YAML from corrected draft (same path as wizard save) ----
cfg_path = config_service.configs_dir(WS) / f"{SHORT}.yaml"
if cfg_path.exists():
    ybak = str(cfg_path) + f".bak.phase_no_score_single_line_{STAMP}"
    shutil.copy2(cfg_path, ybak)
    print("YAML backup ->", ybak)
cfg = config_service.build_yaml_from_wizard(state)
target = config_service.save_client_config(WS, SHORT, cfg, overwrite=True)
print("Config saved ->", target)

# ---- 5. Verify ----
pools = {p["name"]: p for p in cfg.get("pools", [])}
print("\nVerification (risk_rated per pool):")
for name in [p["name"] for p in cfg.get("pools", [])]:
    print(f"  {name:<26} risk_rated={pools[name]['risk_rated']}")
print("\nnot_risk_rated:", cfg.get("not_risk_rated"))
