"""Backfill WSSC FCU's historical-balance provenance so Step 4 honestly shows
that the monthly balances were DERIVED from the imported loan-data extracts
(via scripts/build_wssc_monthly_bal.py), not uploaded as a balance workbook.

Sets state['hist_balance_provenance'] on the wizard draft. Idempotent; backs
up the draft first."""
import os, shutil, datetime
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
WS = os.environ["CECL_WORKSPACE_ROOT"]
from cecl_ui.services import wizard_drafts
from cecl_ui.routes.setup import build_derived_loan_extract_provenance

SHORT = "wssc_fcu"
CU = "WSSC FCU"
STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

state = wizard_drafts.load_draft(WS, SHORT, model="migration")
if state is None:
    raise SystemExit("No WSSC migration draft found")

prov = build_derived_loan_extract_provenance(CU)
state["hist_balance_provenance"] = prov
print("Provenance record:")
for k, v in prov.items():
    print(f"  {k}: {v}")

draft_path = wizard_drafts._file_for(WS, SHORT, "migration")
if draft_path.exists():
    bak = str(draft_path) + f".bak.phase_hist_provenance_{STAMP}"
    shutil.copy2(draft_path, bak)
    print("Draft backup ->", bak)
wizard_drafts.save_draft(WS, state, active_step=state.get("_active_step") or "historical", model="migration")
print("Draft saved ->", draft_path)
