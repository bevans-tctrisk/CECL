"""Refill DQ total_balance for all CUs with pre-Phase-9.30 NULL-balance rows.

Reads each CU's YAML, extracts charter + report_period, and calls
backfill_missing_delinquency_quarters with overwrite=True. Skips test clones.
"""
import calendar
import os
import sys
from datetime import date
from pathlib import Path

import yaml

os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")

from cecl_ui.services.solr_5300_delq_backfill import (
    backfill_missing_delinquency_quarters,
)

SOLR_URL = "http://searchserver1.tctrisk.com:8983/solr"
CORE = "ncua"
WS = Path(os.environ["CECL_WORKSPACE_ROOT"])
CONFIGS = WS / "client_configs"

EXCLUDE = {"Census FCU (Scratch)", "Test Nova CU", "Test Nucor Emp CU"}

# CUs identified by audit_zero_balance_dq.py
TARGETS = [
    "Census FCU",
    "Central Keystone FCU",
    "Destinations CU",
    "Emergency Responders CU",
    "Jackson River Community CU",
    "Mountain CU",
    "Nucor Emp CU",
    "Shuford FCU",
    "TCP CU",
    "United Community FCU",
    "Utah Community FCU",
]


def find_yaml_for_cu(cu_name: str) -> Path | None:
    """Find the YAML file whose 'credit_union' matches cu_name."""
    for p in CONFIGS.glob("*.yaml"):
        try:
            with p.open("r", encoding="utf-8") as fh:
                cfg = yaml.safe_load(fh) or {}
        except Exception:
            continue
        if (cfg.get("credit_union") or "").strip() == cu_name:
            return p
    return None


def coerce_target_period(rp: str) -> str | None:
    """YYYY-MM -> YYYY-MM-DD (last day). YYYY-MM-DD -> as-is."""
    if not rp:
        return None
    rp = str(rp).strip()
    parts = rp.split("-")
    if len(parts) == 3:
        return rp
    if len(parts) == 2:
        y, m = int(parts[0]), int(parts[1])
        d = calendar.monthrange(y, m)[1]
        return date(y, m, d).isoformat()
    return None


def main():
    results = []
    for cu in TARGETS:
        print(f"\n=== {cu} ===")
        if cu in EXCLUDE:
            print("  SKIP (test clone)")
            continue
        yp = find_yaml_for_cu(cu)
        if not yp:
            print(f"  ERROR: no YAML found for cu='{cu}'")
            results.append((cu, False, "no YAML"))
            continue
        with yp.open("r", encoding="utf-8") as fh:
            cfg = yaml.safe_load(fh) or {}
        charter_raw = cfg.get("charter_number")
        rp = cfg.get("report_period")
        target = coerce_target_period(rp)
        if not target:
            # Fall back to latest quarter-end <= today (2026-03-31 as of 2026-06).
            target = "2026-03-31"
            print(f"  (no report_period in YAML; defaulting to {target})")
        if not charter_raw:
            msg = f"missing charter ({charter_raw})"
            print(f"  ERROR: {msg}")
            results.append((cu, False, msg))
            continue
        try:
            charter = int(str(charter_raw).strip())
        except Exception:
            msg = f"bad charter: {charter_raw!r}"
            print(f"  ERROR: {msg}")
            results.append((cu, False, msg))
            continue
        print(f"  charter={charter} target={target} yaml={yp.name}")
        try:
            res = backfill_missing_delinquency_quarters(
                cu=cu,
                charter=charter,
                solr_url=SOLR_URL,
                core=CORE,
                target_period=target,
                history_months=96,
                overwrite=True,
            )
        except Exception as exc:  # noqa: BLE001
            msg = f"{type(exc).__name__}: {exc}"
            print(f"  EXCEPTION: {msg}")
            results.append((cu, False, msg))
            continue
        ok = bool(res.get("ok"))
        n_filled = len(res.get("months_filled") or [])
        n_skipped = len(res.get("months_skipped") or [])
        n_nodata = len(res.get("months_no_data") or [])
        rows = res.get("rows_written") or 0
        stale = res.get("stale_rows_removed") or 0
        err = res.get("error")
        print(
            f"  ok={ok} stale_removed={stale} rows_written={rows} "
            f"filled={n_filled} skipped={n_skipped} no_data={n_nodata}"
        )
        if err:
            print(f"  error: {err}")
        # Show first 3 skip reasons if any
        for s in (res.get("months_skipped") or [])[:3]:
            print(f"    skip: {s}")
        results.append((cu, ok, f"rows={rows} filled={n_filled}"))

    print("\n\n=== SUMMARY ===")
    for cu, ok, msg in results:
        flag = "OK " if ok else "FAIL"
        print(f"  {flag}  {cu:<40} {msg}")


if __name__ == "__main__":
    main()
