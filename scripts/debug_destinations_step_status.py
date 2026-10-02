"""Print step completion + HIL state for the Destinations CU draft."""
import json
from pathlib import Path

from cecl_ui.services import auto_setup

draft = Path("wizard_drafts/destinations_cu__migration.json")
state = json.loads(draft.read_text(encoding="utf-8"))

print("CU =", state.get("credit_union"))
print("report_period =", state.get("report_period"))
print("pool_settings count =", len(state.get("pool_settings") or []))
named = [p.get("name") for p in (state.get("pool_settings") or []) if (p.get("name") or "").strip()]
print("named pools (first 5):", named[:5])

pm = state.get("pool_map") or {}
ignore_count = sum(1 for v in pm.values() if str(v).strip().lower() in ("", "ignore"))
print(f"pool_map size = {len(pm)}, Ignore/blank = {ignore_count}, mapped = {len(pm) - ignore_count}")

print()
print("--- compute_step_completion ---")
done = auto_setup.compute_step_completion(state)
print(sorted(done))

print()
print("--- compute_hil_needs ---")
needs = auto_setup.compute_hil_needs(state)
for n in needs:
    reason = n["reasons"][0] if n.get("reasons") else ""
    print(f"  {n['step_key']:<14s} {n['severity']:<12s} {reason[:80]}")

print()
print("--- expected stepper status (NO_WARM path) ---")
no_warm_keys = [
    "identity", "loan_pools", "pools", "historical", "co_history",
    "recov_history", "dq_hist", "monthly_bal", "grades", "credit_pull",
    "orig_score", "sample", "columns", "balance_check", "co_recov",
    "impaired", "files", "economic", "mgmt_adj", "reports", "review",
]
need_keys = {n["step_key"] for n in needs}
required_keys = {n["step_key"] for n in needs if str(n.get("severity")).lower() == "required"}
report = state.get("_auto_scan_report") or {}
for key in no_warm_keys:
    if key in done and key not in required_keys:
        status = "AUTO  (green)"
    elif key in need_keys:
        status = "HIL   (amber)"
    elif key in report:
        status = "AUTO  (green legacy)"
    else:
        status = "pending"
    print(f"  {key:<14s} -> {status}")
