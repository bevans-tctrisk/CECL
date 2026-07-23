"""Impaired-loans verification for the run-new-quarter flow.

Mirrors wizard Step 16's impaired-loans validation surface at report-run
time so users can confirm each impaired loan resolves to a real pool
BEFORE the standard reports are generated.

Design parallels :mod:`balance_check.compare_run` (Phase 9.22):

    * Reads YAML ``cfg`` directly (no wizard ``state`` needed) so the
      figures shown match what the reports will use at run time.
    * Delegates file discovery to
      :func:`generate_report.find_standalone_impaired_file` so both
      paths agree on which workbook is authoritative for the period.
    * Uses :func:`impaired_parser.parse_file` +
      :func:`impaired_parser.recompute_all` to enrich rows against the
      staged loan-data extracts (via
      :func:`balance_check._build_state_for_run`).
    * Builds an ``impaired_validation`` dict shaped identically to the
      one built in ``cecl_ui/routes/setup.py step_impaired`` so the
      ``run/new_quarter.html`` template can reuse Step 16's macros
      without translation.

Callers: ``cecl_ui/routes/run.py new_quarter()`` between the balance-
check intercept and the report block. Skips silently when no impaired
workbook is on disk for the period — most CUs don't ship one.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any


def verify_for_run(cfg: dict[str, Any], snapshot_iso: str,
                   *, short_name: str | None = None) -> dict[str, Any]:
    """Discover + parse + enrich the impaired workbook for ``snapshot_iso``.

    Returns::

        {
            "ok": bool,
            "error": str | None,
            "found": bool,                # True if a workbook was located
            "filename": str,              # basename of the workbook, or ""
            "path": str,                  # absolute path, or ""
            "cu_name": str,
            "period_ending": str | None,
            "data_rows": list[dict],      # enriched impaired rows
            "matched": int,
            "unmatched": int,
            "source": str | None,         # names of loan-data extracts used
            "totals": {
                "current_balance": float,
                "provision_amount": float,
                "balance_removed": float,
            },
            "validation": {
                "unmapped_codes": list[str],
                "counts": dict[str, int],
                "pool_choices": list[str],
            },
        }

    When no workbook is found, returns ``{"ok": True, "found": False,
    ...}`` — callers should skip the intercept in that case (nothing to
    show the user).
    """
    result: dict[str, Any] = {
        "ok": False,
        "error": None,
        "found": False,
        "filename": "",
        "path": "",
        "cu_name": "",
        "period_ending": None,
        "data_rows": [],
        "matched": 0,
        "unmatched": 0,
        "source": None,
        "totals": {
            "current_balance": 0.0,
            "provision_amount": 0.0,
            "balance_removed": 0.0,
        },
        "validation": {
            "unmapped_codes": [],
            "counts": {},
            "pool_choices": [],
        },
        "distribution": {},
    }

    if not snapshot_iso:
        result["error"] = "Missing snapshot date."
        return result

    # 1) Locate the workbook using the shared discovery helper so we
    # match the exact file the report engine will read.
    try:
        import generate_report  # local: keeps service decoupled at import
        impaired_path = generate_report.find_standalone_impaired_file(
            cfg, snapshot_iso)
    except Exception as exc:  # noqa: BLE001
        result["error"] = f"Impaired file discovery failed: {exc}"
        return result

    if not impaired_path:
        # No impaired file for this period — most CUs don't ship one.
        # Return ok=True/found=False so the caller can silently skip.
        result["ok"] = True
        return result

    result["found"] = True
    result["path"] = str(impaired_path)
    result["filename"] = Path(impaired_path).name

    # 2) Parse the workbook.
    try:
        from cecl_ui.services import impaired_parser
        parsed = impaired_parser.parse_file(impaired_path)
    except Exception as exc:  # noqa: BLE001
        result["error"] = f"Could not parse impaired workbook: {exc}"
        return result

    if not parsed.get("ok"):
        result["error"] = parsed.get("error") or (
            "Impaired workbook did not parse cleanly.")
        result["cu_name"] = parsed.get("cu_name") or ""
        result["period_ending"] = parsed.get("period_ending")
        return result

    result["cu_name"] = parsed.get("cu_name") or ""
    result["period_ending"] = parsed.get("period_ending")

    # 3) Build the state-like dict for recompute_all. Reuse
    # balance_check's helper so the loan-data extract discovery matches
    # the balance-check intercept exactly.
    try:
        from cecl_ui.services import balance_check
        state = balance_check._build_state_for_run(  # noqa: SLF001
            cfg, short_name, snapshot_iso, include_source_folder=True)
    except Exception as exc:  # noqa: BLE001
        result["error"] = f"Could not build extract state: {exc}"
        return result

    # 4) Recompute rows: computes K:Q columns + performs loan-data
    # lookup for R:U columns.
    try:
        lookup_status = impaired_parser.recompute_all(parsed, state)
    except Exception as exc:  # noqa: BLE001
        result["error"] = f"Recompute failed: {exc}"
        return result

    data_rows = parsed.get("data_rows") or []
    result["data_rows"] = data_rows
    result["matched"] = int(lookup_status.get("matched") or 0)
    result["unmatched"] = int(lookup_status.get("unmatched") or 0)
    result["source"] = lookup_status.get("source")

    # 5) Totals for the summary strip.
    tot_cb = 0.0
    tot_prov = 0.0
    tot_rem = 0.0
    for row in data_rows:
        for key, target in (
            ("current_balance", "current_balance"),
            ("provision_amount", "provision_amount"),
            ("balance_removed", "balance_removed"),
        ):
            v = row.get(key)
            try:
                if v is not None:
                    if target == "current_balance":
                        tot_cb += float(v)
                    elif target == "provision_amount":
                        tot_prov += float(v)
                    else:
                        tot_rem += float(v)
            except (TypeError, ValueError):
                continue
    result["totals"] = {
        "current_balance": tot_cb,
        "provision_amount": tot_prov,
        "balance_removed": tot_rem,
    }

    # 6) Build the validation dict (mirrors the block in
    # ``cecl_ui/routes/setup.py step_impaired`` view).
    pool_map = state.get("pool_map") or {}
    pool_code_split = (state.get("pool_code_split") or "").strip()

    def _is_ignore_pool(v: Any) -> bool:
        s = str(v or "").strip().lower()
        return s in ("", "ignore")

    def _truncate_code(code: str) -> str:
        code = str(code or "").strip()
        if not code:
            return ""
        if pool_code_split and pool_code_split in code:
            return code.split(pool_code_split, 1)[0]
        return code

    code_counts: dict[str, int] = {}
    for r in data_rows:
        if not _is_ignore_pool(r.get("loan_pool")):
            continue
        raw_code = r.get("loan_pool_code_resolved") or r.get("loan_type") or ""
        code = _truncate_code(str(raw_code))
        if not code:
            continue
        code_counts[code] = code_counts.get(code, 0) + 1

    # Pool choices: pool_map values + pool_settings names (from cfg), minus
    # Ignore/Exclude. cfg's ``pools`` list is the canonical pool registry.
    pool_names: set[str] = set()
    for p in (cfg.get("pools") or []):
        if isinstance(p, dict):
            n = str(p.get("name") or "").strip()
        else:
            n = str(p or "").strip()
        if n:
            pool_names.add(n)
    for v in pool_map.values():
        s = str(v or "").strip()
        if s:
            pool_names.add(s)
    pool_choices = sorted(
        (n for n in pool_names if n.lower() not in ("ignore", "exclude")),
        key=lambda s: s.lower(),
    )

    result["validation"] = {
        "unmapped_codes": sorted(code_counts.keys(), key=lambda s: s.lower()),
        "counts": code_counts,
        "pool_choices": pool_choices,
    }

    # 7) Grade × pool distribution — lets the user verify that every
    # impaired loan lands in the expected pool and credit grade as
    # resolved from the AIRES loan extract (grade is derived from the
    # extract's FICO score; pool from the extract's loan-code map). Rows
    # that never matched the extract fall back to grade "Not Reported"
    # and their entered loan-type code, and are counted under
    # ``unmatched`` so they surface for review.
    result["distribution"] = _build_distribution(cfg, data_rows)

    result["ok"] = True
    return result


def _build_distribution(cfg: dict[str, Any],
                        data_rows: list[dict[str, Any]]) -> dict[str, Any]:
    """Aggregate impaired rows into a pool × grade matrix for the review UI.

    Returns ``{"grades": [...], "grade_totals": {...}, "by_pool": [...],
    "totals": {...}}`` where ``by_pool`` is ordered by the CU's configured
    ``pool_order`` and grade columns follow the configured ``credit_grades``
    order (with any fallback grade such as "Not Reported" appended).
    """
    def _num(v: Any) -> float:
        try:
            return float(v) if v is not None else 0.0
        except (TypeError, ValueError):
            return 0.0

    # Configured grade order (+ the no-score fallback label).
    grade_order: list[str] = []
    for g in (cfg.get("credit_grades") or []):
        lbl = str((g.get("label") if isinstance(g, dict) else g) or "").strip()
        if lbl and lbl not in grade_order:
            grade_order.append(lbl)
    no_score = str(cfg.get("no_score_label") or "Not Reported").strip()
    if no_score and no_score not in grade_order:
        grade_order.append(no_score)

    # Configured pool order (fall back to the pools registry).
    pool_order: list[str] = []
    for p in (cfg.get("pool_order") or []):
        s = str(p or "").strip()
        if s and s not in pool_order:
            pool_order.append(s)
    if not pool_order:
        for p in (cfg.get("pools") or []):
            n = str((p.get("name") if isinstance(p, dict) else p) or "").strip()
            if n and n not in pool_order:
                pool_order.append(n)

    pools_acc: dict[str, dict[str, Any]] = {}
    grades_present: list[str] = []
    grade_totals: dict[str, int] = {}
    tot = {"count": 0, "matched": 0, "unmatched": 0,
           "current_balance": 0.0, "provision_amount": 0.0}

    for r in data_rows:
        pool = str(r.get("loan_pool") or "").strip() or "Ignore"
        grade = str(r.get("credit_grade") or "").strip() or no_score
        is_matched = not r.get("unmatched_in_loan_data")
        cb = _num(r.get("current_balance"))
        prov = _num(r.get("provision_amount"))
        p = pools_acc.setdefault(pool, {
            "pool": pool, "count": 0, "matched": 0, "unmatched": 0,
            "current_balance": 0.0, "provision_amount": 0.0, "grades": {},
            "is_ignore": pool.strip().lower() in ("ignore", "exclude"),
        })
        p["count"] += 1
        p["matched"] += 1 if is_matched else 0
        p["unmatched"] += 0 if is_matched else 1
        p["current_balance"] += cb
        p["provision_amount"] += prov
        p["grades"][grade] = p["grades"].get(grade, 0) + 1
        if grade not in grades_present:
            grades_present.append(grade)
        grade_totals[grade] = grade_totals.get(grade, 0) + 1
        tot["count"] += 1
        tot["matched"] += 1 if is_matched else 0
        tot["unmatched"] += 0 if is_matched else 1
        tot["current_balance"] += cb
        tot["provision_amount"] += prov

    ordered_grades = [g for g in grade_order if g in grades_present]
    for g in grades_present:
        if g not in ordered_grades:
            ordered_grades.append(g)

    ordered_pools = [p for p in pool_order if p in pools_acc]
    for p in pools_acc:
        if p not in ordered_pools:
            ordered_pools.append(p)

    return {
        "grades": ordered_grades,
        "grade_totals": grade_totals,
        "by_pool": [pools_acc[p] for p in ordered_pools],
        "totals": tot,
    }
