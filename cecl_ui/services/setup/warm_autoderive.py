"""WARM auto-derive bridge for the setup wizard.

Runs the *validated* resolver-path derivation (``parse_warm`` ->
``config_assembler.build_config_from_warm``) and returns both the assembled
config draft and a per-decision list (value + tier + evidence) that the wizard's
auto-derive review surface renders as green / yellow / red cards.

This replaces the wizard's dependence on the older ``services/warm_parser.
analyse_warm_file`` for everything the resolver path derives more completely
(column mappings via Data-tab formula tracing, reserve col-G pinning, SSN
credit-pull detection, keep-first pool_map, WARM pool order, member_account).

Delivery-specific inputs (loan-file pattern, delivery folder) are NOT known at
WARM-upload time; they are filled by the later loan-extract step. This module
derives everything the WARM alone determines and marks the rest ``needs_input``.
"""

from __future__ import annotations

import os
import re
from typing import Any

from . import warm_parser as wp
from . import config_assembler as ca


# Decision kinds surfaced on the review screen, in display order.
_DECISION_ORDER = [
    "identity",
    "pools",
    "loan_code_map",
    "credit_grades",
    "column_mappings",
    "member_account",
    "monthly_balance",
    "reserve_config",
    "credit_pull",
    "economic",
    "pool_order",
]


def _tier(ok_green: bool, ok_yellow: bool = False) -> str:
    if ok_green:
        return "green"
    if ok_yellow:
        return "yellow"
    return "red"


def _decisions(warm: dict, cfg: dict) -> list[dict[str, Any]]:
    """Build the ordered per-decision review list from the parsed WARM + cfg."""
    meta = warm.get("cu_meta") or {}
    econ = (meta.get("economic_stress") or {})
    cmap = cfg.get("column_mappings") or {}
    ma = cfg.get("member_account") or {}
    pool_map = cfg.get("pool_map") or {}
    grades = cfg.get("credit_grades") or []
    pools = cfg.get("pools") or []
    mb = cfg.get("monthly_balance") or {}
    cp = warm.get("credit_pull") or {}
    reserve_keys = ("not_risk_rated", "warm_allowance_pools",
                    "base_loss_rate_by_pool_grade")
    n_reserve = sum(1 for k in reserve_keys if cfg.get(k))
    distinct_pools = len({v for v in pool_map.values() if v})

    out: list[dict[str, Any]] = []

    def add(kind, label, value, tier, evidence):
        out.append({"kind": kind, "label": label, "value": value,
                    "tier": tier, "source": "warm", "evidence": evidence})

    add("identity", "CU Identity",
        {"credit_union": cfg.get("credit_union"),
         "charter_number": cfg.get("charter_number"),
         "report_period": cfg.get("report_period")},
        _tier(bool(cfg.get("credit_union"))),
        f"name={cfg.get('credit_union')!r}, charter={cfg.get('charter_number')}, "
        f"period={cfg.get('report_period')}")

    add("pools", "Loan Pools",
        [p.get("name") for p in pools],
        _tier(len(pools) > 0),
        f"{len(pools)} pool(s) from WARM 'BS CO DQ Data Enter'")

    add("loan_code_map", "Loan Code Mapping",
        {"codes": len(pool_map), "pools": distinct_pools},
        _tier(len(pool_map) > 0),
        f"{len(pool_map)} code(s) -> {distinct_pools} pool(s) "
        "(Grade Ranges & Loan Codes cols S:T, first-wins)")

    add("credit_grades", "Credit Grades / BRR",
        [(g.get("label"), g.get("min_score"), g.get("max_score")) for g in grades],
        _tier(len(grades) > 0),
        f"{len(grades)} grade band(s)"
        + (f"; BRR labels: {len(warm.get('brr_labels') or {})}" if warm.get("brr_labels") else ""))

    add("column_mappings", "Column Mappings",
        cmap,
        _tier(bool(cmap.get("current_balance") and cmap.get("loan_pool_code")),
              bool(cmap)),
        f"{len(cmap)} field(s) traced from the WARM Data-tab formulas"
        if cmap else "no Data-tab formula trace — map on the loan-extract step")

    # Member/account key: the WARM Data-tab formula trace can land on a
    # computed "Member" helper column instead of the raw account column
    # (common for delimiter/split CUs). Flag that case yellow so the user
    # confirms the real account column + delimiter on the loan-extract step.
    _ms = str((warm.get("member_account") or {}).get("member_source") or "").strip().lower()
    _ma_helper = _ms in ("member", "member#", "member #", "member number",
                         "member-sfx", "member #-suffix", "member-suffix")
    _ma_green = bool(ma.get("mode")) and not _ma_helper
    add("member_account", "Member / Account Key",
        ma,
        _tier(_ma_green, bool(ma)),
        (f"mode={ma.get('mode')!r}"
         + (f", suffix_length={ma.get('suffix_length')}"
            if ma.get("suffix_length") is not None else ""))
        + (" &mdash; traced to a WARM helper column; confirm the raw account "
           "column + delimiter on the loan-extract step" if _ma_helper else ""))

    n_months = len(mb.get("months") or [])
    add("monthly_balance", "Monthly Balances (book)",
        {"source": mb.get("source"), "months": n_months,
         "pools": len(mb.get("entries") or {})},
        _tier(n_months > 0),
        f"{n_months} month(s), {len(mb.get('entries') or {})} pool(s) from BS Data"
        if n_months else "no monthly balances parsed")

    add("reserve_config", "Reserve (base loss + mgmt adj)",
        {"not_risk_rated": cfg.get("not_risk_rated"),
         "warm_allowance_pools": cfg.get("warm_allowance_pools"),
         "base_loss_pools": len(cfg.get("base_loss_rate_by_pool_grade") or {})},
        _tier(n_reserve >= 1),
        f"col-G per-grade pins for {len(cfg.get('base_loss_rate_by_pool_grade') or {})} pool(s); "
        f"NRR={len(cfg.get('not_risk_rated') or [])}, "
        f"warm-allowance={len(cfg.get('warm_allowance_pools') or [])}")

    ssn = bool(cp.get("ssn_keyed"))
    add("credit_pull", "Credit Pull",
        {"member_header": cp.get("member_header"),
         "ssn_keyed": ssn,
         "loan_join_column": cp.get("loan_join_column"),
         "latest_tab": cp.get("tab"), "count": cp.get("count")},
        _tier(bool(cp.get("member_header") or cp.get("count")), bool(cp)),
        (f"SSN-keyed -> join on {cp.get('loan_join_column')!r}; " if ssn else "member-keyed; ")
        + f"latest tab {cp.get('tab')!r} ({cp.get('count')} scores)")

    add("economic", "Economic Data",
        cfg.get("economic_data") or {},
        _tier(bool(econ.get("state"))),
        f"state={econ.get('state')}, county={econ.get('county')}, "
        f"unemployment={econ.get('unemployment')}")

    add("pool_order", "Pool Display Order",
        cfg.get("pool_order"),
        _tier(bool(cfg.get("pool_order"))),
        "WARM pool order applied to every report tab"
        if cfg.get("pool_order") else "no pool order")

    order = {k: i for i, k in enumerate(_DECISION_ORDER)}
    out.sort(key=lambda d: order.get(d["kind"], 99))
    return out


def derive(warm_path: str, *, short_name: str = "",
           snapshot_date: str | None = None,
           reports: dict | None = None,
           headerless: bool = False) -> dict[str, Any]:
    """Parse a WARM workbook and assemble the config draft + review decisions.

    Returns ``{"ok", "warm", "config", "decisions", "overall_tier", "error"}``.
    Delivery-specific fields (loan-file pattern / folders) are left blank for
    the loan-extract step to fill. Pass ``headerless=True`` when the delivered
    loan extract has no header row so column_mappings are emitted as 0-based
    integer positions (from the WARM Data-tab column trace) instead of names.
    """
    try:
        warm = wp.parse_warm(warm_path)
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "error": f"WARM parse failed: {exc}",
                "warm": None, "config": None, "decisions": [],
                "overall_tier": "red"}

    snap = snapshot_date or warm.get("snapshot_date") or ""
    warm_folder = os.path.dirname(warm_path)

    cfg = ca.build_config_from_warm(
        warm,
        short_name=short_name or (warm.get("cu_meta") or {}).get("cu_name") or "client",
        delivery_folder="",
        loan_file_pattern="",
        snapshot_date=snap,
        warm_folder=warm_folder,
        reports=reports,
        headerless=headerless,
    )

    # ``build_config_from_warm`` only emits a credit_pull block when a standalone
    # pull filename is supplied (not known at WARM-upload time). Build a
    # fallback-only credit_pull block here so the report can load the WARM's
    # credit scores + ACL/CO/DQ at report time, and surface the SSN-keyed join.
    cp = warm.get("credit_pull") or {}
    warm_name = os.path.basename(warm_path)
    _m = re.match(r"(\d{4}-\d{2})", warm_name)
    period_rx = (r"^" + re.escape(_m.group(1)) + r" CECL-Migration.*\.xlsx$"
                 if _m else r"CECL-Migration.*\.xlsx$")
    cp_block: dict[str, Any] = {
        "use_standalone_file": False,
        "member_column": cp.get("member_column") or "Member Number",
        "score_column": "FICO",
        "fallback_report_folder": warm_folder,
        "fallback_report_pattern": period_rx,
        "fallback_sheet_pattern": "Credit Pull",
        "fallback_member_col": 0,
        "fallback_score_col": 1,
    }
    if snap:
        cp_block["pull_as_of_date"] = snap
    if cp.get("ssn_keyed") and cp.get("loan_join_column"):
        cp_block["loan_join_column"] = cp["loan_join_column"]
        cp_block["member_column"] = cp.get("member_column") or "SSN"
    cfg["credit_pull"] = cp_block

    decisions = _decisions(warm, cfg)
    tiers = {d["tier"] for d in decisions}
    overall = "red" if "red" in tiers else ("yellow" if "yellow" in tiers else "green")

    return {"ok": True, "error": None, "warm": warm, "config": cfg,
            "decisions": decisions, "overall_tier": overall}
