"""Refill DQ history for any CU whose 5300DQ rows are all-zero."""
from __future__ import annotations
import os
import yaml
from pathlib import Path
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")

from sqlalchemy import text  # noqa: E402
from cecl_ui.app import create_app  # noqa: E402

WS = Path(os.environ["CECL_WORKSPACE_ROOT"])
SOLR = "http://searchserver1.tctrisk.com:8983/solr"

app = create_app()
with app.app_context():
    from cecl_ui.services import extract_hist_processor, solr_5300_delq_backfill
    eng = extract_hist_processor._engine_lazy()
    with eng.connect() as conn:
        rows = conn.execute(text(
            "SELECT cu, COUNT(DISTINCT as_of_date), SUM(dq_amount) "
            "FROM loan_code_delinquency_history "
            "WHERE source LIKE '5300DQ:%' "
            "GROUP BY cu HAVING SUM(dq_amount) = 0"
        )).fetchall()
    suspects = [r[0] for r in rows]
    print(f"Suspect CUs (all-zero 5300DQ): {suspects}")

    # Map CU -> YAML to find charter
    yaml_dir = WS / "client_configs"
    cu_to_yaml = {}
    for y in yaml_dir.glob("*.yaml"):
        try:
            cfg = yaml.safe_load(y.read_text(encoding="utf-8"))
        except Exception:
            continue
        if not isinstance(cfg, dict):
            continue
        cu = (cfg.get("credit_union") or "").strip()
        if cu:
            cu_to_yaml.setdefault(cu, []).append((y.stem, cfg))

    for cu in suspects:
        entries = cu_to_yaml.get(cu, [])
        if not entries:
            print(f"  {cu}: no YAML found, skip")
            continue
        short, cfg = entries[0]
        charter = cfg.get("charter_number")
        if not charter:
            print(f"  {cu}: no charter_number in {short}.yaml, skip")
            continue
        period = cfg.get("report_period") or "2026-03-31"
        # Coerce YYYY-MM -> YYYY-MM-DD
        if len(str(period)) == 7:
            import calendar as _cal
            y, m = int(str(period)[:4]), int(str(period)[5:7])
            period = f"{y:04d}-{m:02d}-{_cal.monthrange(y, m)[1]:02d}"
        print(f"\n>>> Refilling {cu} (charter {charter}, period {period})...")
        res = solr_5300_delq_backfill.backfill_missing_delinquency_quarters(
            cu, int(charter), SOLR, "ncua", period, 84, overwrite=True,
        )
        print(f"    ok={res.get('ok')} rows_written={res.get('rows_written')} "
              f"filled={len(res.get('months_filled') or [])} "
              f"skipped={len(res.get('months_skipped') or [])} "
              f"error={res.get('error')}")
        if res.get('months_skipped'):
            print(f"    first skip: {res['months_skipped'][0]}")
