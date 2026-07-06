"""Update WSSC hist_balance_provenance to reflect the extended, hybrid monthly
balance history (Loan Totals workbook for 2023-07..2024-01 + imported loan
extracts for 2024-03..2025-12 + interpolated 2024-02 & 2025-05). Persists to
draft + YAML. Idempotent; backs up first."""
import os, shutil, datetime
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
WS = os.environ["CECL_WORKSPACE_ROOT"]
import yaml
from cecl_ui.services import config_service, wizard_drafts

SHORT = "wssc_fcu"
STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
state = wizard_drafts.load_draft(WS, SHORT, model="migration")
if state is None:
    raise SystemExit("no draft")

prov = {
    "method": "derived_from_loan_extracts",
    "label": "Loan extracts + CU balance summary + NCUA 5300 backfill",
    "summary": (
        "Each pool's monthly balance is the sum of current_balance for that pool "
        "in that month's imported loan-data extract (Jan 2024 - Dec 2025). The "
        "trend is extended to Jul-Dec 2023 using the CU's "
        "'07-2023 - 01-2024 Loan Totals.xlsx' summary, and back to Oct 2018 from "
        "NCUA 5300 Call Report quarter totals distributed across pools by the "
        "Jul-2023 pool mix (flat within each quarter)."
    ),
    "inputs": [
        "Imported monthly loan-data extracts (monthly_loan_data), Jan 2024 - Dec 2025",
        "07-2023 - 01-2024 Loan Totals.xlsx (Jul 2023 - Dec 2023, by loan type)",
        "NCUA 5300 Call Report, distributed (Oct 2018 - Jun 2023, quarterly)",
    ],
    "coverage": {"start": "2018-10-31", "end": "2025-12-31", "months": 87, "pools": 9},
    "generated_by": "Auto-derived during setup (not a single uploaded balance file)",
    "validation_hint": (
        "2024-2025: open that month's loan-data extract, filter to the pool and "
        "sum current_balance. Jul-Dec 2023: tie to '07-2023 - 01-2024 Loan "
        "Totals.xlsx'. Oct 2018 - Jun 2023: each quarter = the CU's 5300 total "
        "loan balance x the pool's Jul-2023 share (same value across the 3 months "
        "of a quarter)."
    ),
    "generated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
}
state["hist_balance_provenance"] = prov

dpath = wizard_drafts._file_for(WS, SHORT, "migration")
if dpath.exists():
    shutil.copy2(dpath, str(dpath) + f".bak.phase_histbal_extend_{STAMP}")
wizard_drafts.save_draft(WS, state, active_step=state.get("_active_step") or "historical", model="migration")

cpath = config_service.configs_dir(WS) / f"{SHORT}.yaml"
if cpath.exists():
    shutil.copy2(cpath, str(cpath) + f".bak.phase_histbal_extend_{STAMP}")
cfg = config_service.build_yaml_from_wizard(state)
config_service.save_client_config(WS, SHORT, cfg, overwrite=True)
print("Updated hist_balance_provenance:")
print(yaml.safe_dump(cfg.get("hist_balance_provenance"), sort_keys=False, allow_unicode=True))
print("credit_grades top label:", (cfg.get("credit_grades") or [{}])[0].get("label"),
      "| not_risk_rated:", cfg.get("not_risk_rated"))
