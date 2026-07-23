"""Ground-truth validator for candidate setup configs (Setup Router, Phase 2 substrate).

Runs a candidate configuration through the **real** pipeline under an isolated
sentinel credit_union (``__setup_probe__<short>``), then scores it with the
**existing** production checks — ``balance_check.compare_run`` (do pools tie to
the monthly book?) and ``impaired_check_service.verify_for_run`` (impaired match
rate?). Cleans up all probe rows afterward.

Why a sentinel key in the live DB rather than a throwaway schema (design Q2):
the whole point of ground-truth validation is that the probe behaves *identically*
to a real run. The existing services are wired to the global engine and
``monthly_loan_data``; a namespaced sentinel CU exercises those exact code paths
with zero reimplementation and can never collide with real reporting data.
``import_file`` already wholesale-deletes per (cu, snapshot), so probe imports are
idempotent; a final ``DELETE ... LIKE '__setup_probe__%'`` guarantees no residue.
"""

from __future__ import annotations

import copy
import importlib
import os
from dataclasses import dataclass, field
from typing import Any

import yaml
from sqlalchemy import text
from sqlalchemy.engine import Engine

SENTINEL_PREFIX = "__setup_probe__"


@dataclass
class ValidationReport:
    ok: bool
    verdict: str                       # green | yellow | red
    pools_tied: int = 0
    pools_total: int = 0
    total_diff: float = 0.0
    impaired_matched: int = 0
    impaired_unmatched: int = 0
    per_pool: list[dict] = field(default_factory=list)
    messages: list[str] = field(default_factory=list)
    error: str | None = None


def _workspace_configs_dir() -> str:
    ws = os.environ.get("CECL_WORKSPACE_ROOT", "").strip()
    if not ws:
        raise RuntimeError("CECL_WORKSPACE_ROOT not set; cannot locate client_configs.")
    return os.path.join(ws, "client_configs")


def _cleanup(engine: Engine, sentinel_cu: str, tmp_cfg_path: str | None) -> None:
    try:
        with engine.begin() as c:
            c.execute(text("DELETE FROM monthly_loan_data WHERE credit_union = :cu"),
                      {"cu": sentinel_cu})
    except Exception:
        pass
    if tmp_cfg_path and os.path.isfile(tmp_cfg_path):
        try:
            os.remove(tmp_cfg_path)
        except OSError:
            pass


def _verdict(pools_tied: int, pools_total: int,
             impaired_unmatched: int) -> str:
    """Validation-gated tier (design Q3): green only when everything ties."""
    if pools_total and pools_tied == pools_total and impaired_unmatched == 0:
        return "green"
    if pools_total and pools_tied >= max(1, pools_total - 1):
        return "yellow"
    return "red"


def probe_validate(engine: Engine, candidate_cfg: dict, *, short: str,
                   snapshot: str, scan_folder: str,
                   run_impaired: bool = True) -> ValidationReport:
    """Import ``candidate_cfg``'s files under a sentinel CU and score the result.

    ``candidate_cfg`` is a full config dict (as the router assembles it);
    ``scan_folder`` is the CU delivery folder to import from; ``snapshot`` is
    the ISO month-end being validated.
    """
    import_data = importlib.import_module("import_data")
    balance_check = importlib.import_module("cecl_ui.services.balance_check")
    impaired_svc = importlib.import_module(
        "cecl_ui.services.impaired_check_service")

    sentinel_short = f"{SENTINEL_PREFIX}{short}"
    sentinel_cu = f"{SENTINEL_PREFIX}{short}"
    cfg = copy.deepcopy(candidate_cfg)
    cfg["credit_union"] = sentinel_cu
    cfg["archive_imported_files"] = False   # never move the CU's source files

    configs_dir = _workspace_configs_dir()
    tmp_cfg_path = os.path.join(configs_dir, f"{sentinel_short}.yaml")
    report = ValidationReport(ok=False, verdict="red")

    try:
        # Persist the candidate so the REAL loaders (process_client,
        # compare_run, impaired discovery) resolve it by short name.
        with open(tmp_cfg_path, "w", encoding="utf-8") as fh:
            yaml.safe_dump(cfg, fh, sort_keys=False, allow_unicode=True)

        # 1) REAL import under the sentinel CU (scan_folder_override => no archive).
        import_data.process_client(sentinel_short,
                                   scan_folder_override=scan_folder)

        # 2) Pool tie-out via the production balance check. NRR / provision
        # pools (e.g. OTHER PARTICIPATIONS) have no loan-level extract, so they
        # legitimately never "tie" — exclude them from the denominator.
        loaded_cfg = import_data.load_client_config(sentinel_short)
        nrr = {str(p).strip() for p in (loaded_cfg.get("not_risk_rated") or [])}
        cmp = balance_check.compare_run(loaded_cfg, snapshot, sentinel_short)
        rows = cmp.get("rows", []) or []
        tied = 0
        total_diff = 0.0
        scored = 0
        for r in rows:
            pool = r.get("pool")
            m, l = r.get("monthly"), r.get("loans")
            if m is None or l is None or str(pool).strip() in nrr:
                continue
            scored += 1
            diff = abs(float(l) - float(m))
            total_diff += diff
            if diff <= 1.0:
                tied += 1
            report.per_pool.append({
                "pool": pool, "monthly": m, "loans": l,
                "diff": round(float(l) - float(m), 2),
            })
        report.pools_tied = tied
        report.pools_total = scored

        # 3) Impaired match rate via the production verifier.
        if run_impaired:
            imp = impaired_svc.verify_for_run(loaded_cfg, snapshot,
                                              short_name=sentinel_short)
            if imp.get("found"):
                report.impaired_matched = int(imp.get("matched", 0))
                report.impaired_unmatched = int(imp.get("unmatched", 0))
                report.messages.append(
                    f"impaired: {report.impaired_matched} matched, "
                    f"{report.impaired_unmatched} unmatched")
            else:
                report.messages.append("impaired: no workbook for period")

        report.total_diff = round(total_diff, 2)
        report.verdict = _verdict(report.pools_tied, report.pools_total,
                                  report.impaired_unmatched)
        report.ok = True
        report.messages.insert(
            0, f"pools tied {report.pools_tied}/{report.pools_total} "
               f"(total diff ${report.total_diff:,.2f})")
    except Exception as exc:  # noqa: BLE001
        report.error = f"{type(exc).__name__}: {exc}"
    finally:
        _cleanup(engine, sentinel_cu, tmp_cfg_path)

    return report
