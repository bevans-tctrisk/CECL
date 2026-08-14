"""Smoke test for auto_setup.apply_findings_to_state pool_seed wiring."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from cecl_ui.services import auto_setup

FOLDER = (
    r"Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients"
    r"\Destinations CU 66333\Portfolio Management\CECL Migration IDLR"
)

findings = auto_setup.scan_folder_for_setup(FOLDER, "202512")
state = {"credit_union": "Destinations CU", "short_name": "destinations_cu"}
report = auto_setup.apply_findings_to_state(state, findings, r"c:\Dev\CECL")

print("=== Report ===")
for k, v in report.items():
    for m in v:
        print(f"  {k}: {m}")

ps_rows = state.get("pool_settings") or []
print(f"\n=== pool_settings ({len(ps_rows)} rows) ===")
for p in ps_rows:
    print(
        f"  {p['name']!r}  acl_months={p['acl_months']}  "
        f"risk_rated={p['risk_rated']}  excluded={p['excluded']}"
    )

mb = state.get("monthly_bal") or {}
pm = mb.get("pool_map") or {}
print(f"\n=== monthly_bal.pool_map ({len(pm)} entries) ===")
items = list(pm.items())
for label, pool in items[:10]:
    print(f"  {label!r} -> {pool!r}")
if len(items) > 10:
    print(f"  ... ({len(items) - 10} more)")

print(f"\nwarm.pools: {(state.get('warm') or {}).get('pools')}")
print(f"\n_pool_seed.source_file: {state.get('_pool_seed', {}).get('source_file')}")
