"""Force-rerun ER DQ backfill with overwrite=True against production Solr."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from cecl_ui.services.solr_5300_delq_backfill import (
    backfill_missing_delinquency_quarters,
)

CU = "Emergency Responders CU"
CHARTER = 66354
SOLR_URL = "http://searchserver1.tctrisk.com:8983/solr"
SOLR_CORE = "ncua"
TARGET = "2025-12-31"
HISTORY = 84

print(f"Running backfill cu={CU} charter={CHARTER} target={TARGET} months={HISTORY} overwrite=True")
res = backfill_missing_delinquency_quarters(
    CU, CHARTER, SOLR_URL, SOLR_CORE, TARGET, HISTORY, overwrite=True,
)
print("ok:", res.get("ok"))
print("error:", res.get("error"))
print("expected:", len(res.get("expected") or []))
print("months_filled:", len(res.get("months_filled") or []))
print("months_skipped:", len(res.get("months_skipped") or []))
print("months_no_data:", len(res.get("months_no_data") or []))
print("rows_written:", res.get("rows_written"))
print("stale_rows_removed:", res.get("stale_rows_removed"))
mf = res.get("months_filled") or []
if mf:
    print("first 3 filled:", mf[:3])
ms = res.get("months_skipped") or []
if ms:
    print("first 3 skipped:", ms[:3])
