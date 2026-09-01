"""Compute-time populator: build the ReportModel from report DATA.

The counterpart to :mod:`cecl_report_web.from_workbook`, and the whole point
of the PDF migration: populate the *same* model dataclasses the templates
already consume, but from ``(client_name, snapshot_date, config, grades,
hist, df)`` -- the inputs the report engine itself uses -- instead of by
scraping a generated ``.xlsx``. No openpyxl, no cell coordinates, no reading
business meaning out of fills or fonts.

Pages are added here one archetype at a time; each returns the identical node
``from_workbook`` returns, so ``render.py`` / the Jinja templates are unchanged.
"""

from __future__ import annotations

import base64
import datetime as _dt
import os
from typing import Any

from .model import (
    AclEnvPage,
    AclPoolRow,
    AdjustmentRow,
    CoverPage,
    KeyValueRow,
    MatrixCell,
    MatrixRow,
    RiskChangeMatrix,
    RiskChangePage,
)


def _snap_parts(snapshot_date: str) -> tuple[int, int, int]:
    d = _dt.date.fromisoformat(str(snapshot_date)[:10])
    return d.month, d.day, d.year


def _date_text(snapshot_date: str) -> str:
    """Match the workbook cover's ``m/d/yyyy`` display (no leading zeros)."""
    m, d, y = _snap_parts(snapshot_date)
    return f"{m}/{d}/{y}"


def _logo_data_uri(path: str) -> str | None:
    """Base64-encode a PNG asset as a ``data:`` URI for inline <img> use."""
    try:
        if not path or not os.path.isfile(path):
            return None
        with open(path, "rb") as fh:
            raw = fh.read()
        return "data:image/png;base64," + base64.b64encode(raw).decode("ascii")
    except OSError:
        return None


# Verbatim from report_vizo.sheet_cover_vizo (B21) so the from-data cover
# reads identically to the workbook cover.
_DISCLAIMER = (
    "The following analysis and all parts thereof (\u2018analysis\u2019) are based upon "
    "information obtained by Vizo Financial Corporate Credit Union (Vizo Financial) from "
    "the credit union that is subject of this analysis and other sources that Vizo Financial "
    "believes to be reliable and utilized in models using methods and assumptions which Vizo "
    "Financial believes to be reasonable.  However, actual performance compared to estimated "
    "performance of the subject credit union may be different and cannot be guaranteed.  This "
    "analysis is for informational purposes only and is intended only for the use of the "
    "subject credit union.  The analysis does not constitute either legal or tax advice. \n"
    "All reports are confidential."
)


def build_cover(client_name: str, snapshot_date: str, config: dict,
                *, supplemental: bool = False) -> CoverPage:
    """Populate the Vizo cover from config/period alone (no workbook).

    Mirrors ``report_vizo.sheet_cover_vizo``: title, CU name, period date,
    the confidentiality paragraph, the copyright footer, and the two bundled
    logos resolved from ``report_vizo``'s asset paths.
    """
    cu = (config or {}).get("credit_union") or client_name

    # Reuse the report engine's own logo paths so the asset resolves the same
    # way it does for the workbook; import lazily to avoid a heavy import at
    # module load.
    try:
        import report_vizo as _rv
        top = _logo_data_uri(getattr(_rv, "LOGO_VIZO", ""))
        bottom = _logo_data_uri(getattr(_rv, "LOGO_TCT", ""))
    except Exception:  # noqa: BLE001 - logos are optional; cover still renders
        top = bottom = None

    date_text = _date_text(snapshot_date)
    return CoverPage(
        credit_union=cu,
        period_ending=str(snapshot_date)[:10],
        title="CECL Credit Migration Report",
        subtitle="Supplemental Reports" if supplemental else None,
        date_text=date_text,
        paragraph=_DISCLAIMER,
        footer=f"\u00a9 {_snap_parts(snapshot_date)[2]} TCT Risk Solutions",
        top_logo=top,
        bottom_logo=bottom,
    )


def _cell_state(i: int, j: int, cur_grade: str, orig_grade: str,
                no_score: str, n_top: int) -> str:
    """Improved / deteriorated / plain, computed from grade ORDERING.

    Mirrors ``report_vizo._sheet_risk_change`` exactly -- the state that the
    workbook encodes in the cell fill is derived here from position instead:
    a current grade worse than original deteriorates, unless the original is
    among the top ``n_top`` grades and the drop is a single band (the WARM
    "top grades need a 2+ drop" rule). Not-Reported never migrates.
    """
    if cur_grade == no_score or orig_grade == no_score:
        return "plain"
    if i > j:  # current grade ranked below original -> potential deterioration
        if j < n_top and (i - j) < 2:
            return "plain"
        return "deteriorated"
    if i < j:  # current grade ranked above original -> improvement
        return "improved"
    return "plain"


def build_risk_change(client_name: str, snapshot_date: str, df: Any,
                      config: dict, grades: Any,
                      hist: dict | None = None) -> RiskChangePage:
    """Populate the "Risk Change Total" page from the loan frame.

    Reuses the report engine's own compute (``cecl_engine.risk_change_matrix``
    plus report_vizo's grade helpers) so the numbers are identical to the
    workbook; the improved/deteriorated state is computed from grade ordering,
    not read from a cell fill.
    """
    import report_vizo as _rv
    from cecl_engine import risk_change_matrix

    no_score = (config or {}).get("no_score_label", "Not Reported")
    n_top = int((config or {}).get("top_grades_double_drop", 3))
    gl = [g for g in _rv._all_grades(grades, no_score) if not _rv._is_hidden(g)]
    matrix = risk_change_matrix(df, grades, no_score)
    rng = _rv._grade_ranges(grades, no_score)
    total = float(df["current_balance"].sum())
    cu = (config or {}).get("credit_union") or client_name

    def _mv(cur: str, og: str) -> float:
        return float(_rv._matrix_val(matrix, cur, og) or 0.0)

    col_totals = {og: sum(_mv(g2, og) for g2 in gl) for og in gl}

    def _build(is_pct: bool) -> RiskChangeMatrix:
        rows: list[MatrixRow] = []
        for i, g in enumerate(gl):
            cells: list[MatrixCell] = []
            rtotal = 0.0
            for j, og in enumerate(gl):
                v = _mv(g, og)
                rtotal += v
                val = (v / col_totals[og] if col_totals[og] else 0.0) if is_pct else v
                cells.append(MatrixCell(
                    value=val, state=_cell_state(i, j, g, og, no_score, n_top),
                    is_pct=is_pct))
            tot = (rtotal / total if total else 0.0) if is_pct else rtotal
            rows.append(MatrixRow(
                label=g, range_label=rng.get(g, ""), cells=cells,
                total=MatrixCell(value=tot, is_pct=is_pct, bold=True)))
        # Grand Total row: column sums ($) / 100% (pct).
        gt_cells = [MatrixCell(value=(1.0 if is_pct else col_totals[og]),
                               is_pct=is_pct, bold=True) for og in gl]
        rows.append(MatrixRow(
            label="Grand Total", range_label="", cells=gt_cells,
            total=MatrixCell(value=(1.0 if is_pct else total),
                             is_pct=is_pct, bold=True)))
        return RiskChangeMatrix(
            corner=("% Current Grade" if is_pct else "$ Current Grade"),
            col_headers=list(gl), rows=rows, is_pct=is_pct)

    _imp = (hist or {}).get("impaired", {}) or {}
    bal_adj = float(_imp.get("total_balance_adjustment", 0.0) or 0.0)
    tip = _imp.get("total_in_portfolio") or (total + bal_adj)
    summary = [
        KeyValueRow(label="Balance Adjustment", value=bal_adj),
        KeyValueRow(label="Total in Portfolio", value=float(tip)),
    ]
    heading = [
        "Executive Summary Total Loans",
        "Risk Change By Credit Score",
        f"For Quarter Ending {_rv._snap_display(snapshot_date)}",
    ]
    return RiskChangePage(credit_union=cu, heading_lines=heading,
                          matrices=[_build(False), _build(True)],
                          summary=summary)


_ACL_COL_HEADERS = [
    "Current Grade", "Balance", "Specific Identification",
    "Loan Loss Calc Balance", "ACL Base Loss Rate", "Mgmt Adj",
    "Allowance Factor", "Allowance before Env", "Env Factor",
    "Env Allowance", "Total Allowance",
]
_ACL_ADJ_KEYS = [
    ("Total Specifically Identified Allowance", "total_spec_allow"),
    ("Total Allowance Needed", "total_allow_needed"),
    ("Allowance for Credit Loss Balance", "acl_balance"),
    ("Adjustment", "adjustment"),
]


def build_acl_env(client_name: str, snapshot_date: str, config: dict,
                  hist: dict | None = None) -> AclEnvPage | None:
    """Populate the "ACL Env by Pool Mgmt Adj" page from the ACL data dicts.

    Consumes ``hist['impaired']['acl_pools' | 'acl_summary' | 'acl_impaired']``
    -- the per-pool/per-grade allowance structure ``load_impaired_data``
    produces for WARM-fed CUs. Returns ``None`` when ``acl_pools`` is absent
    (wizard-onboarded CUs, whose ACL env is still computed inside
    ``report_vizo._sheet_acl_reserve``; extracting that compute to populate
    ``acl_pools`` is the remaining Step-1a work -- this consumer is then ready).
    """
    import report_vizo as _rv

    _imp = (hist or {}).get("impaired", {}) or {}
    # Prefer the values report_vizo._sheet_acl_reserve computes and publishes
    # (isolated underscore keys) -- available for every CU, wizard-onboarded
    # included -- over the WARM-parsed dicts (present only for WARM CUs).
    acl_pools = _imp.get("_acl_pools_computed") or _imp.get("acl_pools") or {}
    if not acl_pools:
        return None
    acl_summary = _imp.get("_acl_summary_computed") or _imp.get("acl_summary") or {}
    acl_impaired = _imp.get("_acl_impaired_computed") or _imp.get("acl_impaired") or {}
    pool_order = list(acl_pools.keys()) or _imp.get("pool_order") or []
    cu = (config or {}).get("credit_union") or client_name

    def _match(pool: str) -> dict | None:
        if pool in acl_pools:
            return acl_pools[pool]
        lc = pool.strip().lower()
        return next((v for k, v in acl_pools.items()
                     if k.strip().lower() == lc), None)

    pool_rows: list[AclPoolRow] = []
    for pool in pool_order:
        pdata = _match(pool)
        if not pdata:
            continue
        pool_rows.append(AclPoolRow(pool=pool, kind="header"))
        for g, gv in (pdata.get("grades") or {}).items():
            if str(g).upper().startswith("HIDE"):
                continue
            pool_rows.append(AclPoolRow(
                pool=g, kind="grade",
                balance=gv.get("balance"), specific_id=gv.get("spec_id"),
                llc_balance=gv.get("calc_bal"), base_loss_rate=gv.get("base_rate"),
                mgmt_adj=gv.get("mgmt_adj"), allowance_factor=gv.get("factor"),
                allowance_before_env=gv.get("allow_before")))
        t = pdata.get("total") or {}
        _bal, _spec = t.get("balance"), t.get("spec_id")
        pool_rows.append(AclPoolRow(
            pool="Total", kind="total",
            balance=_bal, specific_id=_spec,
            llc_balance=((_bal or 0) - (_spec or 0)) if _bal is not None else None,
            base_loss_rate=t.get("base_rate"), mgmt_adj=t.get("mgmt_adj"),
            allowance_factor=t.get("factor"),
            allowance_before_env=t.get("allow_before"),
            env_factor=t.get("env_factor"), env_allowance=t.get("env_allow"),
            total_allowance=t.get("total_allow")))

    _pb = acl_summary.get("pooled_balance")
    _ps = acl_summary.get("pooled_spec_id")
    pooled = AclPoolRow(
        pool="Pooled Totals", is_total=True,
        balance=_pb, specific_id=_ps,
        llc_balance=((_pb or 0) - (_ps or 0)) if _pb is not None else None,
        allowance_before_env=acl_summary.get("pooled_allow_before"),
        env_allowance=acl_summary.get("pooled_env_allow"),
        total_allowance=acl_summary.get("pooled_total_allow"))

    impaired_rows = [AdjustmentRow(label=k, value=v)
                     for k, v in acl_impaired.items()]
    adjustment_rows = [
        AdjustmentRow(label=lbl, value=acl_summary.get(key), bold=True)
        for lbl, key in _ACL_ADJ_KEYS if key in acl_summary
    ]
    heading = [
        "Allowance & Provision for Credit Loss Reserve Analysis",
        f"For Quarter Ending {_rv._snap_display(snapshot_date)}",
    ]
    return AclEnvPage(
        credit_union=cu, heading_lines=heading, col_headers=list(_ACL_COL_HEADERS),
        pool_rows=pool_rows, pooled_totals=pooled,
        impaired_rows=impaired_rows, adjustment_rows=adjustment_rows)


def build_report_model(client_name: str, snapshot_date: str, config: dict,
                       grades: Any = None, hist: dict | None = None,
                       df: Any = None, *, supplemental: bool = False) -> dict:
    """Build the render-ready page set from report data.

    Returns ``{"cover": CoverPage, "pages": [ (template, ctx, landscape), ... ]}``.
    Archetypes are added here incrementally (cover, Risk Change so far), each
    reusing the report engine's own pure compute functions.
    """
    cover = build_cover(client_name, snapshot_date, config,
                        supplemental=supplemental)
    pages: list[tuple[str, dict, bool]] = [
        ("cover.html", {"cover": cover}, False),
    ]
    if df is not None and not supplemental:
        rc = build_risk_change(client_name, snapshot_date, df, config, grades, hist)
        pages.append(("risk_change.html", {"page": rc, "charts": []}, True))
    if not supplemental:
        acl = build_acl_env(client_name, snapshot_date, config, hist)
        if acl is not None:
            pages.append(("acl_env.html", {"page": acl, "charts": []}, True))
    return {"cover": cover, "pages": pages}
