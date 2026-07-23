"""Assemble a full client config from a parsed WARM workbook (Setup Router —
WARM ingestion path).

``parse_warm`` extracts the config *knowledge* from a CU's WARM workbook. This
module turns that knowledge into (1) a ready-to-run client config, (2) a
standalone credit-pull file the pipeline can consume, and (3) high-trust
registry patterns seeded from the WARM.

For WARM CUs the WARM is authoritative, so seeded patterns are
``source='warm', trust='trusted'``.
"""

from __future__ import annotations

import os
from typing import Any

import openpyxl

from . import setup_registry as reg


# ---------------------------------------------------------------------------
# 1) config assembly
# ---------------------------------------------------------------------------
def build_config_from_warm(
    warm: dict, *,
    short_name: str,
    delivery_folder: str,
    loan_file_pattern: str,
    snapshot_date: str,
    core_extract_label: str = "AIRES loan extract",
    date_pattern: str = r"(?i)(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[-_ .]+(20\d{2})",
    date_format: str = "YYYY-MM",
    credit_pull_filename: str | None = None,
    credit_pull_folder: str | None = None,
    warm_folder: str | None = None,
    reports: dict | None = None,
) -> dict:
    """Build a client-config dict from a ``parse_warm`` result.

    ``delivery_folder`` is the root that holds the quarterly upload files;
    ``loan_file_pattern`` matches the raw core (AIRES) extract.
    """
    meta = warm.get("cu_meta") or {}
    econ = meta.get("economic_stress") or {}
    cmap = dict(warm.get("column_mappings") or {})
    ma = dict(warm.get("member_account") or {})
    ma_clean = {k: ma[k] for k in ("mode", "suffix_length") if k in ma}

    period = (snapshot_date or "")[:7]

    cfg: dict[str, Any] = {
        "credit_union": meta.get("cu_name") or short_name,
        "charter_number": str(meta.get("charter_number") or "") or None,
        "report_period": period,
        "source": "warm",
        "file_pattern": loan_file_pattern,
        "date_pattern": date_pattern,
        "date_format": date_format,
        "data_directory": delivery_folder,
        "loan_file_folder": delivery_folder,
        "loan_file_recursive": True,
        "has_header": True,
        "member_account": ma_clean or {"mode": "fixed_suffix", "suffix_length": 0},
        "column_mappings": cmap,
        "pool_map": dict(warm.get("pool_map") or {}),
        "default_pool": "Ignore",
        "credit_grades": list(warm.get("credit_grades") or []),
        "no_score_label": "Not Reported",
        "reports": reports or {"tct": True, "vizo": False,
                               "vizo_supp": False, "impdet": False},
    }

    mgmt = warm.get("mgmt_adj_params") or {}
    if mgmt.get("ltv_baseline") is not None or mgmt.get("probability_factor") is not None:
        cfg["mgmt_adj"] = {"ltv_baseline": mgmt.get("ltv_baseline"),
                           "probability_factor": mgmt.get("probability_factor")}

    if econ:
        cfg["economic_data"] = {
            "state": econ.get("state"),
            "county": econ.get("county"),
            "unemployment_rate": econ.get("unemployment"),
        }

    # monthly balances (the book) + ACL history
    mb = warm.get("monthly_balance")
    if mb:
        cfg["monthly_balance"] = {
            "source": "manual",
            "months": list(mb.get("months") or []),
            "entries": dict(mb.get("entries") or {}),
        }
        if warm.get("alll"):
            cfg["monthly_balance"]["alll"] = dict(warm["alll"])

    # credit pull (standalone file the pipeline consumes)
    cp = warm.get("credit_pull") or {}
    if credit_pull_filename:
        import re as _re
        cfg["credit_pull"] = {
            "use_standalone_file": True,
            "file_pattern": r"(?i)^" + _re.escape(credit_pull_filename) + r"$",
            "uploaded_filename": credit_pull_filename,
            "source_folder": credit_pull_folder or delivery_folder,
            "member_column": "Member Number",
            "score_column": "FICO",
            "pull_as_of_date": cp.get("period"),
            # Date-aware original/current score logic (analyst methodology):
            # original = pull before origination (else AIRES); current = latest
            # pull, or the newest post-pull loan's origination score. Reads the
            # WARM's dated pull tabs via fallback_report_folder below.
            "origination_aware": True,
        }
        # SSN-keyed pull: the pull is keyed by SSN (not member number), so the
        # loans must join it on their SSN column (detected by warm_parser from
        # the Data tab's raw headers, e.g. 'MBR SSN STRING'). Emit the join key
        # and the pull's key-column name so import_data joins correctly.
        if cp.get("ssn_keyed") and cp.get("loan_join_column"):
            cfg["credit_pull"]["member_column"] = cp.get("member_column") or "SSN"
            cfg["credit_pull"]["loan_join_column"] = cp["loan_join_column"]
        # The report's load_impaired_data() builds the per-pool ACL data
        # (allow_before etc. used by warm_allowance_pools) from the WARM file;
        # point it at the WARM's folder so those values come from THIS quarter's
        # WARM, not a stale prior report.
        if warm_folder:
            cfg["credit_pull"]["fallback_report_folder"] = warm_folder

    # core loan extract
    cfg["loan_data_extracts"] = [{
        "label": core_extract_label,
        "file_pattern": loan_file_pattern,
        "column_mappings": cmap,
        "member_account": ma_clean or {"mode": "fixed_suffix", "suffix_length": 0},
        "has_header": True,
    }]

    # pools (balance-only participation pools included; carried by monthly book)
    cfg["pools"] = _pools_from_settings(warm.get("pool_settings") or [])
    # Pool display order = the WARM's own pool order. Emitting an explicit
    # pool_order makes every report tab list pools exactly as the analyst's
    # WARM does (report ordering otherwise falls back to alphabetical). Applies
    # to all WARM-onboarded CUs.
    if cfg["pools"]:
        cfg["pool_order"] = [p["name"] for p in cfg["pools"]]
    acl_rates = warm.get("acl_reserve_rates") or {}
    # NRR = balance-only pools (no per-grade loan detail — overdraft /
    # participations). The report reserves these from the book balance x a
    # pool-level rate instead of a grade distribution. NOTE: this is NOT the
    # WARM's "Risk Rated Yes/No" flag — many graded pools (Commercial RE,
    # Business Boat) are flagged "No" yet still carry full per-grade detail.
    # Exclude placeholder HIDE-* / Exclude rows: they carry no grade balances
    # (so they read as "balance_only") but are not real reporting pools and
    # must not be emitted as not_risk_rated.
    def _real_pool(name: str) -> bool:
        low = str(name or "").strip().lower()
        return bool(low) and not low.startswith("hide") and low != "exclude"

    nrr = [pool for pool, d in acl_rates.items()
           if d.get("balance_only") and _real_pool(pool)]
    if nrr:
        cfg["not_risk_rated"] = nrr

    reserve = derive_reserve_config(acl_rates)
    if reserve.get("warm_allowance_pools"):
        cfg["warm_allowance_pools"] = reserve["warm_allowance_pools"]
    if reserve.get("base_loss_rate_by_pool_grade"):
        cfg["base_loss_rate_by_pool_grade"] = reserve["base_loss_rate_by_pool_grade"]
    return cfg


def derive_reserve_config(acl_rates: dict) -> dict:
    """Derive the WARM-pinned reserve config from parsed ACL reserve rates.

    Returns a dict with the keys the report engine reads:
      * ``not_risk_rated`` — balance-only pools (overdraft / participations).
      * ``warm_allowance_pools`` — graded pools whose allowance is analyst
        mgmt-adjustment driven (base rate ~0 but non-zero allowance); the
        firm-wide model can't reproduce the WARM's methodology, so use the
        WARM's computed allowance verbatim (needs the WARM as report baseline).
      * ``base_loss_rate_by_pool_grade`` — per-grade full allowance factor
        (base + mgmt adj, col G) for risk-rated pools, plus a 'Total' blended
        rate for balance-only NRR pools.
    Shared by ``build_config_from_warm`` and the reserve-config resolver.
    """
    out: dict = {}
    nrr = [pool for pool, d in acl_rates.items() if d.get("balance_only")]
    if nrr:
        out["not_risk_rated"] = nrr
    warm_allow = [
        pool for pool, d in acl_rates.items()
        if not d.get("balance_only")
        and (d.get("max_base") or 0.0) <= 1e-9
        and (d.get("allow_before") or 0.0) > 0.005
    ]
    if warm_allow:
        out["warm_allowance_pools"] = warm_allow
    blr: dict = {}
    for pool, d in acl_rates.items():
        entry = {g: r for g, r in (d.get("grades") or {}).items()}
        if d.get("blended") is not None:
            entry["Total"] = d["blended"]
        if entry:
            blr[pool] = entry
    if blr:
        out["base_loss_rate_by_pool_grade"] = blr
    return out


def _pools_from_settings(settings: list[dict]) -> list[dict]:
    pools = []
    for p in settings:
        name = p.get("pool")
        if not name or name.lower() == "ignore":
            continue
        basis = (p.get("grade_basis") or "").upper()
        pools.append({
            "name": name,
            "risk_rated": bool(p.get("risk_rated")),
            "acl_months": int(p["acl_months"]) if p.get("acl_months") is not None else 36,
            "excluded": name.lower() == "exclude",
            "use_default_mgmt_adj": False,
            "brr": basis == "RR",
        })
    return pools


# ---------------------------------------------------------------------------
# 2) credit-pull emission
# ---------------------------------------------------------------------------
def emit_credit_pull_file(warm: dict, out_path: str) -> dict:
    """Write the WARM's embedded credit-pull scores to a standalone xlsx with
    the (Member Number, FICO) columns the pipeline expects."""
    cp = warm.get("credit_pull") or {}
    scores = cp.get("scores") or {}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Credit Pull"
    ws.append(["Member Number", "FICO"])
    for member, fico in scores.items():
        ws.append([member, fico])
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)
    return {"path": out_path, "count": len(scores), "period": cp.get("period")}


# ---------------------------------------------------------------------------
# 3) registry seeding (WARM = authoritative -> trusted)
# ---------------------------------------------------------------------------
def seed_from_warm(engine, warm: dict, cu: str) -> dict[str, int]:
    """Seed high-trust patterns from a WARM into the setup registry."""
    reg.ensure_tables(engine)
    counts: dict[str, int] = {}

    def bump(k):
        counts[k] = counts.get(k, 0) + 1

    cmap = warm.get("column_mappings") or {}
    ma = warm.get("member_account") or {}

    # member_account
    if ma.get("mode"):
        reg.save_pattern(
            engine, "member_account",
            {"loan_suffix_present": bool(cmap.get("loan_suffix"))},
            {"mode": ma["mode"]}, source="warm",
            evidence={"suffix_lengths": [ma.get("suffix_length")]}, cu=cu)
        bump("member_account")

    # pool_code_column
    if cmap.get("loan_pool_code"):
        col = cmap["loan_pool_code"]
        reg.save_pattern(
            engine, "pool_code_column",
            {"normalized": _norm(col)},
            {"role": "loan_pool_code", "column": col}, source="warm", cu=cu)
        bump("pool_code_column")

    # score sources
    ss = warm.get("score_sources") or {}
    for role in ("original", "current"):
        col = ss.get(role) or cmap.get(f"{role}_fico_score")
        if col:
            reg.save_pattern(
                engine, "score_source",
                {"role": role, "normalized": _norm(col)},
                {"role": role, "column": col}, source="warm", cu=cu)
            bump("score_source")

    # pool_map priors (authoritative from the CU's own WARM)
    for code, pool in (warm.get("pool_map") or {}).items():
        reg.save_pattern(
            engine, "pool_map_code",
            {"code": _norm_code(code)},
            {"pool": str(pool)}, source="warm", cu=cu)
        bump("pool_map_code")

    # monthly-balance archetype
    if warm.get("monthly_balance"):
        reg.save_pattern(engine, "monthly_balance_source", {},
                         {"source": "manual"}, source="warm", cu=cu)
        bump("monthly_balance_source")

    # credit-pull keys
    if warm.get("credit_pull"):
        reg.save_pattern(engine, "credit_pull_keys", {},
                         {"member_column": "Member Number",
                          "score_column": "FICO"}, source="warm", cu=cu)
        bump("credit_pull_keys")

    return counts


def _norm(s: Any) -> str:
    import re
    s = str(s or "").strip().lower()
    s = re.sub(r"[_\-]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _norm_code(x: Any) -> str:
    x = str(x).strip()
    if x.replace(".", "", 1).isdigit():
        try:
            return str(int(float(x)))
        except (TypeError, ValueError):
            return x
    return x
