"""Re-run Nova CU DQ backfill with overwrite=True to populate Phase 9.30 per-code balances."""
import os
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")

from cecl_ui.services.solr_5300_delq_backfill import (
    backfill_missing_delinquency_quarters,
)

# Nova CU charter 63425, target current report period 2026-03-31
# Use ~96 months history to cover the existing 29 quarters (2018-12 → 2025-12)
# plus the recent additions (2026-03)
SOLR_URL = "http://searchserver1.tctrisk.com:8983/solr"
CORE = "ncua"

result = backfill_missing_delinquency_quarters(
    cu="NOVA CU",
    charter=63425,
    solr_url=SOLR_URL,
    core=CORE,
    target_period="2026-03-31",
    history_months=96,
    overwrite=True,
)

print(f"ok={result.get('ok')}")
print(f"error={result.get('error')}")
print(f"stale_rows_removed={result.get('stale_rows_removed')}")
print(f"rows_written={result.get('rows_written')}")
print(f"expected={len(result.get('expected') or [])}")
print(f"months_filled={len(result.get('months_filled') or [])}")
print(f"months_skipped={len(result.get('months_skipped') or [])}")
print(f"months_no_data={len(result.get('months_no_data') or [])}")

if result.get("months_skipped"):
    print("\nFirst 5 skipped:")
    for s in result["months_skipped"][:5]:
        print(f"  {s}")

if result.get("months_filled"):
    print("\nFirst 3 filled:")
    for f in result["months_filled"][:3]:
        print(f"  {f}")
    print("Last 3 filled:")
    for f in result["months_filled"][-3:]:
        print(f"  {f}")
