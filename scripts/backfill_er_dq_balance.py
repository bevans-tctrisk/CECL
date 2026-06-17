"""Phase 9.30 retroactive backfill — populate total_balance per NCUA code
on existing loan_code_delinquency_history rows by re-running the 5300DQ
backfill against the live Solr core for a given CU.

Usage:
    python scripts/backfill_er_dq_balance.py emergency_responders_cu
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import yaml  # type: ignore

from cecl_ui.services.solr_5300_delq_backfill import (
    backfill_missing_delinquency_quarters,
)


WS = Path(r"Z:\Shared\TCT Files\CECL - CM Files")


def main(short: str) -> int:
    cfg_path = WS / "client_configs" / f"{short}.yaml"
    if not cfg_path.exists():
        print(f"Config not found: {cfg_path}")
        return 1
    cfg = yaml.safe_load(cfg_path.read_text(encoding="utf-8"))
    cu = cfg.get("credit_union")
    charter_raw = cfg.get("charter_number")
    try:
        charter = int(str(charter_raw).strip())
    except (TypeError, ValueError):
        print(f"Bad charter_number: {charter_raw!r}")
        return 1

    hist_cfg = (cfg.get("hist_extracts") or {})
    solr_cfg = (hist_cfg.get("solr_backfill") or {})
    solr_url = solr_cfg.get("solr_url") or "http://localhost:8983/solr"
    core = solr_cfg.get("core") or "ncua"
    target_period = cfg.get("report_period") or "2025-12"
    # _parse_yyyy_mm_dd in solr_5300_backfill expects YYYY-MM-DD; if the
    # YAML stores YYYY-MM (the wizard's canonical Identity-step form),
    # expand to the last day of that month.
    if len(target_period) == 7 and target_period.count("-") == 1:
        import calendar as _cal
        y, m = target_period.split("-")
        d = _cal.monthrange(int(y), int(m))[1]
        target_period = f"{int(y):04d}-{int(m):02d}-{d:02d}"
    history_months = int(cfg.get("history_months", 84))

    print(f"CU: {cu} (charter {charter})")
    print(f"Solr: {solr_url}/{core}")
    print(f"Period: {target_period} ({history_months} months back)")
    print("Re-running DQ backfill with overwrite=True ...")

    result = backfill_missing_delinquency_quarters(
        cu, charter, solr_url, core, target_period, history_months,
        overwrite=True,
    )
    print(f"ok: {result.get('ok')}")
    print(f"error: {result.get('error')}")
    print(f"rows_written: {result.get('rows_written')}")
    print(f"stale_rows_removed: {result.get('stale_rows_removed')}")
    print(f"months_filled: {len(result.get('months_filled') or [])}")
    print(f"months_no_data: {len(result.get('months_no_data') or [])}")
    print(f"months_skipped: {len(result.get('months_skipped') or [])}")
    for m in (result.get("months_skipped") or [])[:5]:
        print(f"  skipped sample: {m}")
    for m in (result.get("months_filled") or [])[:3]:
        print(f"  filled sample: {m}")
    expected = result.get("expected") or []
    print(f"expected quarters: {len(expected)} -> first 3: {expected[:3]}, last 3: {expected[-3:]}")
    return 0 if result.get("ok") else 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "emergency_responders_cu"))
