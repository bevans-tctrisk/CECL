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
    ChartSpec,
    CoverPage,
    ImprDeterPage,
    KeyValueRow,
    MatrixCell,
    MatrixRow,
    NarrativePage,
    NarrativeSection,
    RiskChangeMatrix,
    RiskChangePage,
    SummaryVarianceBlock,
    SummaryVariancePage,
    TableCell,
    TablePage,
    TableSection,
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
        import os as _os
        import report_vizo as _rv
        _vizo = getattr(_rv, "LOGO_VIZO", "") or ""
        # Prefer a shadow-free variant next to the source logo when present, so
        # the PDF cover shows the Vizo mark without the baked-in drop shadow.
        _noshadow = (_os.path.splitext(_vizo)[0] + "_noshadow.png") if _vizo else ""
        top = _logo_data_uri(_noshadow if _os.path.exists(_noshadow) else _vizo)
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
                      hist: dict | None = None,
                      pool_name: str | None = None) -> RiskChangePage:
    """Populate a Risk Change page from the loan frame -- the total portfolio
    when ``pool_name`` is None, otherwise a single pool (pass the pool-filtered
    ``df``). One page per matrix set, mirroring the workbook's Risk Change tabs.

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
    if pool_name:
        bal_adj = float((_imp.get("balance_adjustments") or {}).get(
            pool_name, 0.0) or 0.0)
        tip = total + bal_adj
        heading = [
            "Risk Change By Credit Score",
            pool_name,
            f"For Quarter Ending {_rv._snap_display(snapshot_date)}",
        ]
    else:
        bal_adj = float(_imp.get("total_balance_adjustment", 0.0) or 0.0)
        tip = _imp.get("total_in_portfolio") or (total + bal_adj)
        heading = [
            "Executive Summary Total Loans",
            "Risk Change By Credit Score",
            f"For Quarter Ending {_rv._snap_display(snapshot_date)}",
        ]
    summary = [
        KeyValueRow(label="Balance Adjustment", value=bal_adj),
        KeyValueRow(label="Total in Portfolio", value=float(tip)),
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


def _acl_data(config: dict, snapshot_date: str, hist: dict | None,
              df: Any, grades: Any) -> tuple[dict, dict, dict]:
    """(acl_pools, acl_summary, acl_impaired): the values report_vizo publishes,
    then the WARM-parsed dicts, then a standalone compute when df/grades exist."""
    import report_vizo as _rv

    _imp = (hist or {}).get("impaired", {}) or {}
    acl_pools = _imp.get("_acl_pools_computed") or _imp.get("acl_pools") or {}
    acl_summary = _imp.get("_acl_summary_computed") or _imp.get("acl_summary") or {}
    acl_impaired = _imp.get("_acl_impaired_computed") or _imp.get("acl_impaired") or {}
    if not acl_pools and df is not None:
        computed = _rv.compute_acl_environmental(df, grades, config, hist, snapshot_date)
        acl_pools = computed.get("acl_pools") or {}
        acl_summary = computed.get("acl_summary") or acl_summary
        acl_impaired = computed.get("acl_impaired") or acl_impaired
    return acl_pools, acl_summary, acl_impaired


def build_acl_env(client_name: str, snapshot_date: str, config: dict,
                  hist: dict | None = None, df: Any = None,
                  grades: Any = None) -> AclEnvPage | None:
    """Populate the "ACL Env by Pool Mgmt Adj" page from the ACL data dicts.

    Prefers the values ``report_vizo._sheet_acl_reserve`` publishes when the
    report was composed (isolated underscore keys), then the WARM-parsed dicts,
    and finally -- when neither is present and ``df``/``grades`` are supplied --
    computes them standalone via ``report_vizo.compute_acl_environmental`` so
    the page renders without the workbook being built at all.
    """
    import report_vizo as _rv

    acl_pools, acl_summary, acl_impaired = _acl_data(
        config, snapshot_date, hist, df, grades)
    if not acl_pools:
        return None
    pool_order = list(acl_pools.keys())
    cu = (config or {}).get("credit_union") or client_name
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


def build_impr_deter(client_name: str, snapshot_date: str, config: dict,
                     hist: dict | None = None, df: Any = None,
                     grades: Any = None) -> ImprDeterPage:
    """Populate the "Impr Deter" page's CECL Adjustment box from data.

    The four headline allowance figures are exactly the ACL summary the ACL
    Env page already computes, so the box ties out to that tab. Reuses the
    published/computed ``acl_summary`` (standalone-computes it when absent).
    The improved/deteriorated charts are added once the ChartSpec node lands.
    """
    import report_vizo as _rv

    _imp = (hist or {}).get("impaired", {}) or {}
    acl_summary = _imp.get("_acl_summary_computed") or _imp.get("acl_summary") or {}
    if not acl_summary.get("total_allow_needed") and df is not None:
        acl_summary = (_rv.compute_acl_environmental(
            df, grades, config, hist, snapshot_date).get("acl_summary")
            or acl_summary)

    cu = (config or {}).get("credit_union") or client_name
    adj = float(acl_summary.get("adjustment") or 0.0)
    adj_label = ("Adjustment (Underfunded)" if adj >= 0
                 else "Adjustment (Overfunded)")
    cecl = [
        KeyValueRow(label="Total Specifically Identified Allowance",
                    value=acl_summary.get("total_spec_allow")),
        KeyValueRow(label="Total Allowance Needed",
                    value=acl_summary.get("total_allow_needed")),
        KeyValueRow(label=f"Allowance for Credit Loss Balance as of {snapshot_date}",
                    value=acl_summary.get("acl_balance")),
        KeyValueRow(label=adj_label, value=adj),
    ]
    heading = [
        "Improved & Deteriorated Loans",
        f"For Quarter Ending {_rv._snap_display(snapshot_date)}",
    ]
    return ImprDeterPage(
        credit_union=cu, period_ending=str(snapshot_date)[:10],
        heading_lines=heading, cecl_adjustment=cecl)


# ── Charts (rendered from data via cecl_report_web.charts) ───────────
# Migration-status slice colours, matching report_vizo's DQ pie / CO bar /
# Net-Credit-Change doughnut (olive / maroon / teal / gold).
# Single semantic migration palette, shared by every chart AND the matrix/table
# CSS (.rc-matrix .improved/.deteriorated) so "Improved" is olive everywhere,
# "Deteriorated" maroon, "Unchanged"/"Net" neutral teal, "Not Reported" gold.
_C_IMPROVED = "829901"
_C_DETERIORATED = "873A3A"
_C_UNCHANGED = "0D4D5E"
_C_NOT_REPORTED = "FFC000"
_C_NET = "0D4D5E"
_MIG_LABELS = ("Improved", "Deteriorated", "Unchanged", "Not Reported")
_MIG_COLORS = (_C_IMPROVED, _C_DETERIORATED, _C_UNCHANGED, _C_NOT_REPORTED)
_NCC_COLORS = (_C_IMPROVED, _C_DETERIORATED, _C_UNCHANGED)


def _chartspec_to_render_dict(cs: ChartSpec) -> dict:
    """Adapt a semantic ChartSpec to the dict charts.render_chart_svg consumes."""
    lbl_fmt = "0.0%" if cs.value_format == "pct" else None
    if cs.kind in ("pie", "doughnut"):
        s0 = cs.series[0] if cs.series else {}
        ctype = "DoughnutChart" if cs.kind == "doughnut" else "PieChart"
        return {
            "type": ctype, "title": cs.title,
            "series": [{
                "name": s0.get("name"), "values": s0.get("values") or [],
                "cats": list(cs.categories),
                "point_colors": s0.get("colors"),
                "label_fmt": lbl_fmt, "show_labels": True,
            }],
        }
    bar_dir = "bar" if cs.kind in ("bar_h", "diverging_bar") else "col"
    grouping = "stacked" if cs.kind == "diverging_bar" else "clustered"
    return {
        "type": "BarChart", "bar_dir": bar_dir, "grouping": grouping,
        "title": cs.title,
        "series": [{
            "name": s.get("name"), "values": s.get("values") or [],
            "cats": list(cs.categories),
            "point_colors": s.get("colors"),
            "color": (s.get("colors") or [None])[0], "filled": True,
            "label_fmt": lbl_fmt, "show_labels": True,
        } for s in cs.series],
    }


def render_chart_specs(specs: list[ChartSpec]) -> list[str]:
    """Render a list of ChartSpecs to inline SVG strings."""
    from .charts import render_chart_svg
    out: list[str] = []
    for cs in specs:
        try:
            out.append(render_chart_svg(_chartspec_to_render_dict(cs)))
        except Exception:  # noqa: BLE001 - a bad chart must not sink the page
            continue
    return out


def _mig_status_series(data: dict, use_pct: bool = True) -> list[float] | None:
    """[Improved, Deteriorated, Unchanged, Not Reported] from a status dict."""
    if not data:
        return None
    key = "pct" if use_pct else "balance"
    vals = [float((data.get(l) or {}).get(key) or 0.0) for l in _MIG_LABELS]
    return vals if any(vals) else None


def risk_change_charts(hist: dict | None, pool_name: str | None = None) -> list[ChartSpec]:
    """DQ pie + CO bar for a Risk Change tab, from the migration-status dicts."""
    _imp = (hist or {}).get("impaired", {}) or {}
    if pool_name:
        pl = pool_name.strip().lower()
        dq = next((v for k, v in (_imp.get("dq_by_pool") or {}).items()
                   if k.strip().lower() == pl), {})
        co = next((v for k, v in (_imp.get("co_by_pool") or {}).items()
                   if k.strip().lower() == pl), {})
    else:
        dq = _imp.get("dq_by_status") or {}
        co = _imp.get("co_by_status") or {}
    specs: list[ChartSpec] = []
    dq_vals = _mig_status_series(dq)
    if dq_vals is not None:
        specs.append(ChartSpec(
            kind="pie", title="Delinquency by Credit Grade Migration",
            categories=list(_MIG_LABELS),
            series=[{"name": "DQ", "values": dq_vals, "colors": list(_MIG_COLORS)}],
            value_format="pct"))
    co_vals = _mig_status_series(co)
    if co_vals is not None:
        specs.append(ChartSpec(
            kind="bar_h", title="Charge off by Credit Grade Migration",
            categories=list(_MIG_LABELS),
            series=[{"name": "CO", "values": co_vals, "colors": list(_MIG_COLORS)}],
            value_format="pct"))
    return specs


def _ncc_totals(df: Any, grades: Any, config: dict) -> tuple[float, float, float, float]:
    """(improved, deteriorated, unchanged, total) balances from the matrix,
    using the same migration-state rule as the Risk Change grid."""
    import report_vizo as _rv
    from cecl_engine import risk_change_matrix

    no_score = (config or {}).get("no_score_label", "Not Reported")
    n_top = int((config or {}).get("top_grades_double_drop", 3))
    gl = [g for g in _rv._all_grades(grades, no_score) if not _rv._is_hidden(g)]
    matrix = risk_change_matrix(df, grades, no_score)
    total = float(df["current_balance"].sum())
    imp = det = 0.0
    for i, g in enumerate(gl):
        for j, og in enumerate(gl):
            v = float(_rv._matrix_val(matrix, g, og) or 0.0)
            st = _cell_state(i, j, g, og, no_score, n_top)
            if st == "improved":
                imp += v
            elif st == "deteriorated":
                det += v
    return imp, det, max(0.0, total - imp - det), total


def risk_change_ncc_chart(df: Any, grades: Any, config: dict) -> list[ChartSpec]:
    """Net Credit Change doughnut (Improved / Deteriorated / Unchanged)."""
    imp, det, unc, total = _ncc_totals(df, grades, config)
    if total <= 0:
        return []
    return [ChartSpec(
        kind="doughnut", title="Net Credit Change",
        categories=["Improved", "Deteriorated", "Unchanged"],
        series=[{"name": "NCC",
                 "values": [imp / total, det / total, unc / total],
                 "colors": list(_NCC_COLORS)}],
        value_format="pct")]


def impr_deter_charts(df: Any, grades: Any, config: dict,
                      hist: dict | None = None) -> list[ChartSpec]:
    """The four Impr Deter charts: Improved/Deteriorated by grade (%), the
    Improved/Deteriorated diverging bar by pool, and Net Change by pool."""
    import report_vizo as _rv
    from cecl_engine import risk_change_matrix

    no_score = (config or {}).get("no_score_label", "Not Reported")
    _imp = (hist or {}).get("impaired", {}) or {}
    gl = [g for g in _rv._all_grades(grades, no_score) if not _rv._is_hidden(g)]
    chart_grades = [g for g in gl if g != no_score]

    # Grade-level improved/deteriorated: WARM Executive Summary (3) if present,
    # else derived from the migration matrix (same rule as the Risk Change grid).
    es3 = _imp.get("exec_summary_3") or {}
    if es3:
        grade_imp = {g: float((es3.get("improved") or {}).get(g, 0) or 0)
                     for g in chart_grades}
        grade_det = {g: float((es3.get("deteriorated") or {}).get(g, 0) or 0)
                     for g in chart_grades}
    else:
        matrix = risk_change_matrix(df, grades, no_score)
        n_top = int((config or {}).get("top_grades_double_drop", 3))
        grade_imp = {g: 0.0 for g in chart_grades}
        grade_det = {g: 0.0 for g in chart_grades}
        for j, og in enumerate(gl):
            for i, cg in enumerate(gl):
                v = float(_rv._matrix_val(matrix, cg, og) or 0.0)
                if i > j:
                    if not (j < n_top and (i - j) < 2) and og in grade_det:
                        grade_det[og] += v
                elif i < j and og in grade_imp:
                    grade_imp[og] += v
    imp_tot = sum(grade_imp.values()) or 0.0
    det_tot = sum(grade_det.values()) or 0.0
    imp_pct_g = [(grade_imp[g] / imp_tot if imp_tot else 0.0) for g in chart_grades]
    det_pct_g = [(grade_det[g] / det_tot if det_tot else 0.0) for g in chart_grades]

    # Pool-level improved / deteriorated (negative) / net, risk-rated pools only.
    _rr = _imp.get("risk_rated", {}) or {}
    names, p_imp, p_det, p_net = [], [], [], []
    for pool in _rv._ordered_pools(df, hist):
        pdf = df[df["loan_pool"] == pool]
        if _rr.get(pool, True):
            ip, dp, npct = _rv._ncc(pdf, grades, config)
        else:
            ip, dp, npct = 0.0, 0.0, 0.0
        names.append(pool)
        p_imp.append(float(ip)); p_det.append(-float(dp)); p_net.append(float(npct))

    specs: list[ChartSpec] = []
    if chart_grades:
        specs.append(ChartSpec(
            kind="column", title="Improved Loans", categories=chart_grades,
            series=[{"name": "Improved", "values": imp_pct_g, "colors": [_C_IMPROVED]}],
            value_format="pct"))
        specs.append(ChartSpec(
            kind="column", title="Deteriorated Loans", categories=chart_grades,
            series=[{"name": "Deteriorated", "values": det_pct_g, "colors": [_C_DETERIORATED]}],
            value_format="pct"))
    if names:
        specs.append(ChartSpec(
            kind="diverging_bar", title="Improved / Deteriorated Loans",
            categories=names,
            series=[{"name": "Improved", "values": p_imp, "colors": [_C_IMPROVED]},
                    {"name": "Deteriorated", "values": p_det, "colors": [_C_DETERIORATED]}],
            value_format="pct"))
        specs.append(ChartSpec(
            kind="bar_h", title="Net Change", categories=names,
            series=[{"name": "Net", "values": p_net, "colors": [_C_NET]}],
            value_format="pct"))
    return specs


def build_acl_summary(client_name: str, snapshot_date: str, config: dict,
                      hist: dict | None = None, df: Any = None,
                      grades: Any = None) -> TablePage | None:
    """ACL Summary: one line per pool (the pool Total rows), pooled totals, and
    the impaired/adjustment lines -- a view of the ACL Env data."""
    import report_vizo as _rv

    acl_pools, acl_summary, acl_impaired = _acl_data(
        config, snapshot_date, hist, df, grades)
    if not acl_pools:
        return None
    cu = (config or {}).get("credit_union") or client_name

    cols = ["Portfolio Segment", "Balance", "Specific Identification",
            "Loan Loss Calc Balance", "Allowance before Env Factor",
            "Env Factor", "Env Factor Allowance", "Total Allowance"]
    rows: list[list[TableCell]] = []
    for pool, pdata in acl_pools.items():
        t = pdata.get("total") or {}
        bal, spec = t.get("balance"), t.get("spec_id")
        calc = t.get("calc_bal")
        if calc is None and bal is not None:
            calc = (bal or 0) - (spec or 0)
        rows.append([
            TableCell(pool, "text", align="left"),
            TableCell(bal, "currency"), TableCell(spec, "currency2"),
            TableCell(calc, "currency"), TableCell(t.get("allow_before"), "currency"),
            TableCell(t.get("env_factor"), "pct"),
            TableCell(t.get("env_allow"), "currency"),
            TableCell(t.get("total_allow"), "currency"),
        ])
    pb, ps = acl_summary.get("pooled_balance"), acl_summary.get("pooled_spec_id")
    rows.append([
        TableCell("Pooled Totals", "text", bold=True, align="left"),
        TableCell(pb, "currency", bold=True), TableCell(ps, "currency2", bold=True),
        TableCell(((pb or 0) - (ps or 0)) if pb is not None else None, "currency", bold=True),
        TableCell(acl_summary.get("pooled_allow_before"), "currency", bold=True),
        TableCell(None), TableCell(acl_summary.get("pooled_env_allow"), "currency", bold=True),
        TableCell(acl_summary.get("pooled_total_allow"), "currency", bold=True),
    ])
    sections = [TableSection(columns=cols, rows=rows)]

    adj_rows: list[list[TableCell]] = []
    for k, v in (acl_impaired or {}).items():
        adj_rows.append([TableCell(k, "text", align="left"), TableCell(v, "currency")])
    for lbl, key in _ACL_ADJ_KEYS:
        if key in acl_summary:
            adj_rows.append([TableCell(lbl, "text", bold=True, align="left"),
                             TableCell(acl_summary.get(key), "currency", bold=True)])
    if adj_rows:
        sections.append(TableSection(
            title="Impaired Loans & Adjustment", rows=adj_rows))

    return TablePage(
        credit_union=cu,
        title="Allowance for Credit Loss - Summary by Pool",
        heading_lines=[f"For Quarter Ending {_rv._snap_display(snapshot_date)}"],
        sections=sections)


def build_mgmt_adj_summary(client_name: str, snapshot_date: str, config: dict,
                           hist: dict | None = None, df: Any = None,
                           grades: Any = None) -> TablePage | None:
    """Mgmt Adj Summary: per-pool the grades that carry a management adjustment
    plus the pool's environmental factor -- from the ACL Env data."""
    import report_vizo as _rv

    acl_pools, acl_summary, _ = _acl_data(config, snapshot_date, hist, df, grades)
    if not acl_pools:
        return None
    cu = (config or {}).get("credit_union") or client_name
    heading = [f"For Quarter Ending {_rv._snap_display(snapshot_date)}"]
    title = "Management & Environmental Adjustments"

    def blank() -> TableCell:
        return TableCell(None)

    cols = ["Portfolio Segment", "Grade", "Balance", "ACL Base Loss Rate",
            "Mgmt Adj", "Allowance Factor", "Allowance before Env Factor",
            "Env Factor", "Env Factor Allowance"]
    rows: list[list[TableCell]] = []
    any_adj = False
    for pool, pdata in acl_pools.items():
        grades_d = pdata.get("grades") or {}
        total = pdata.get("total") or {}
        adj_grades = [(g, gv) for g, gv in grades_d.items()
                      if (gv.get("mgmt_adj") or 0)]
        env_factor = total.get("env_factor") or 0
        if not adj_grades and not env_factor:
            continue
        any_adj = True
        rows.append([TableCell(pool, "text", bold=True, align="left")]
                    + [blank() for _ in range(8)])
        for g, gv in adj_grades:
            rows.append([
                blank(), TableCell(g, "text", align="left"),
                TableCell(gv.get("balance"), "currency"),
                TableCell(gv.get("base_rate"), "pct4"),
                TableCell(gv.get("mgmt_adj"), "pct4"),
                TableCell(gv.get("factor"), "pct4"),
                TableCell(gv.get("allow_before"), "currency"), blank(), blank()])
        rows.append([
            blank(), TableCell("Total", "text", bold=True, align="left"),
            TableCell(total.get("balance"), "currency", bold=True),
            blank(), blank(), blank(),
            TableCell(total.get("allow_before"), "currency", bold=True),
            TableCell(total.get("env_factor"), "pct", bold=True),
            TableCell(total.get("env_allow"), "currency", bold=True)])

    if not any_adj:
        return TablePage(
            credit_union=cu, title=title, heading_lines=heading,
            sections=[TableSection(rows=[[TableCell(
                "No management or environmental adjustments were applied "
                "this period.", "text", align="left")]])])

    rows.append([
        TableCell("Pooled Totals", "text", bold=True, align="left"), blank(),
        TableCell(acl_summary.get("pooled_balance"), "currency", bold=True),
        blank(), blank(), blank(),
        TableCell(acl_summary.get("pooled_allow_before"), "currency", bold=True),
        blank(),
        TableCell(acl_summary.get("pooled_env_allow"), "currency", bold=True)])
    return TablePage(credit_union=cu, title=title, heading_lines=heading,
                     sections=[TableSection(columns=cols, rows=rows)])


def build_impaired_loans(client_name: str, snapshot_date: str, config: dict,
                         hist: dict | None = None, df: Any = None,
                         grades: Any = None) -> TablePage | None:
    """Impaired Loans - ASC 310-10: the specifically identified allowance by
    impairment category, from the ACL Env impaired data."""
    import report_vizo as _rv

    acl_pools, acl_summary, acl_impaired = _acl_data(
        config, snapshot_date, hist, df, grades)
    if not acl_impaired:
        return None
    cu = (config or {}).get("credit_union") or client_name
    rows = [[TableCell(k, "text", align="left"), TableCell(v, "currency")]
            for k, v in acl_impaired.items()]
    if "total_spec_allow" in acl_summary:
        rows.append([
            TableCell("Total Specifically Identified Allowance", "text",
                      bold=True, align="left"),
            TableCell(acl_summary.get("total_spec_allow"), "currency", bold=True)])
    return TablePage(
        credit_union=cu, title="Impaired Loans - ASC 310-10",
        heading_lines=[f"For Quarter Ending {_rv._snap_display(snapshot_date)}"],
        sections=[TableSection(columns=["Impairment Category", "Allowance"],
                               rows=rows)])


def build_report_index(client_name: str, config: dict,
                       supplemental: bool = False) -> NarrativePage:
    """Static Report Index / overview page (ports report_vizo._sheet_report_index)."""
    cu = (config or {}).get("credit_union") or client_name
    if supplemental:
        return NarrativePage(
            credit_union=cu, title="Report Index",
            sections=[
                NarrativeSection("Report Overview", (
                    "The CECL Credit Migration Supplemental Reports from TCT, Inc. "
                    "presents the historical details of the changing nature of risk in "
                    "the credit union\u2019s loan portfolio.")),
                NarrativeSection("Supplemental Reporting Package:", (
                    "Historical Loan Balances by Credit Score\n"
                    "Loss Factor Historical Detail\n"
                    "Charge off and Recoveries Historical Detail\n"
                    "Balance Adjustment Detail")),
            ])
    return NarrativePage(
        credit_union=cu, title="Report Index",
        sections=[
            NarrativeSection("Report Overview", (
                "The CECL Credit Migration Reports from TCT, Inc. presents a comprehensive "
                "picture of the changing nature of risk in the credit union\u2019s loan "
                "portfolio. Credit migration is measured by the improvement or "
                "deterioration of risk, measured by the credit score, from the date of "
                "loan funding to the most recent data pull.  New credit scores are "
                "typically pulled twice per year.  Migration may still be measured on a "
                "quarterly basis to take into account new loans and changing loan "
                "balances.")),
            NarrativeSection("Executive Summary", (
                "CECL Adjustment  & Improved/Deteriorated\n"
                "Improved & Deteriorated Loans Risk Change By Credit Score")),
            NarrativeSection("Detailed Reporting", (
                "Allowance & Provision for Credit Loss Reserve Analysis\n"
                "Risk Change by Credit Score - Total Loans\n"
                "Risk Change by Credit Score - Loan Pools\n"
                "Environmental Factor Provision for Loan Loss\n"
                "Loss Factor Calculation\n"
                "Delinquency Calculation\n\n"
                "Additional detailed reporting located in the Supplemental Reporting "
                "Package")),
        ])


def build_introduction(client_name: str, config: dict) -> NarrativePage:
    """Static Appendix - Credit Migration / CECL methodology narrative."""
    cu = (config or {}).get("credit_union") or client_name
    return NarrativePage(
        credit_union=cu, title="Appendix - Credit Migration",
        sections=[
            NarrativeSection("Credit Migration", (
                "Credit Migration describes the movement of individual loans through the "
                "credit scoring system. Each loan is assigned a risk grade based on the "
                "borrower's credit score at origination and the most recent credit score. "
                "When the current score differs from the original score, the loan has "
                "\"migrated\" - either improving (higher score) or deteriorating (lower "
                "score). This migration forms the basis for assessing changes in portfolio "
                "risk.")),
            NarrativeSection("CECL Methodology", (
                "Under the Current Expected Credit Losses (CECL) standard, institutions must "
                "estimate lifetime expected credit losses on financial assets measured at "
                "amortized cost. The Credit Migration methodology uses the Weighted Average "
                "Remaining Maturity (WARM) approach to estimate these losses, incorporating "
                "historical loss experience, current conditions, and reasonable and "
                "supportable forecasts.")),
        ])


def build_exec_summary_narrative(client_name: str, config: dict) -> NarrativePage:
    """Static Appendix - Executive Summary narrative."""
    cu = (config or {}).get("credit_union") or client_name
    return NarrativePage(
        credit_union=cu, title="Appendix - Executive Summary",
        sections=[NarrativeSection("Executive Summary", (
            "The Executive Summary provides an overview of the credit union's current "
            "portfolio risk position. It includes the CECL Adjustment calculation showing "
            "the relationship between pooled allowance, specifically identified allowance, "
            "total allowance needed, and the current ACL balance. The summary also presents "
            "improved and deteriorated loan totals by portfolio segment."))])


def build_env_factor(client_name: str, snapshot_date: str, config: dict,
                     hist: dict | None = None, df: Any = None,
                     grades: Any = None) -> TablePage | None:
    """Environmental Factor by Pool: the economic-stress index inputs and the
    per-pool Net Credit / Delinquency / Economic Stress scores that combine
    into each pool's environmental factor -- computed from data, not the .xlsx.
    """
    import report_vizo as _rv

    if df is None:
        return None
    cfg = config or {}
    cu = cfg.get("credit_union") or client_name
    ed = cfg.get("economic_data", {}) or {}
    _imp = (hist or {}).get("impaired", {}) or {}
    if _imp.get("economic_data"):
        ed = _imp["economic_data"]
    econ_stress = _rv._eco_stress(cfg, ed_override=ed)
    ncc_r, dq_r, es_r = _rv._env_ranges(hist)
    pools = _rv._ordered_pools(df, hist)
    if not pools:
        return None
    dq_var = _rv._pool_dq_variance(pools, hist, snapshot_date)
    risk_rated_map = _imp.get("risk_rated", {})

    pop = ed.get("population", 1) or 1
    state_cols = ["State", "Unemployment Rate", "Foreclosures per Person",
                  "Bankruptcies", "Population"]
    state_rows = [[
        TableCell(ed.get("state", ""), "text", align="left"),
        TableCell(ed.get("unemployment_rate", 0), "pct2"),
        TableCell(ed.get("foreclosures", 0), "currency"),
        TableCell(ed.get("bankruptcies", 0), "currency"),
        TableCell(pop, "currency"),
    ]]
    bk_pct = (ed.get("bankruptcies", 0) / pop) if pop else 0
    fc_pct = (ed.get("foreclosures", 0) / pop) if pop else 0
    county_cols = ["County", "Unemployment Rate", "Bankruptcy %",
                   "Foreclosure %", "Economic Stress Index"]
    county_rows = [[
        TableCell(ed.get("county", ""), "text", align="left"),
        TableCell(ed.get("unemployment_rate", 0), "pct2"),
        TableCell(bk_pct, "pct2"),
        TableCell(fc_pct, "pct2"),
        TableCell(econ_stress / 100.0, "pct2"),
    ]]

    pool_cols = ["Portfolio Segment", "Net Credit Change", "Net Credit Score",
                 "Delinquency Variance from Ave.", "Delinquency Score",
                 "Economic Stress Actual", "Economic Stress Score",
                 "Environmental Factor"]
    pool_rows: list[list[TableCell]] = []
    for pool in pools:
        pdf = df[df["loan_pool"] == pool]
        is_rr = risk_rated_map.get(pool, True)
        ncc_pct = _rv._ncc(pdf, grades, cfg)[2] if is_rr else 0.0
        ncc_score = _rv._score(ncc_pct * 100, ncc_r) / 100.0
        dqv = dq_var.get(pool, 0)
        dq_score = _rv._score(dqv * 100, dq_r) / 100.0
        es_score = _rv._score(econ_stress, es_r) / 100.0
        env_f = ncc_score + dq_score + es_score
        pool_rows.append([
            TableCell(pool, "text", align="left"),
            TableCell(ncc_pct, "pct2", align="center"),
            TableCell(ncc_score, "pct2", align="center"),
            TableCell(dqv, "pct2", align="center"),
            TableCell(dq_score, "pct2", align="center"),
            TableCell(econ_stress / 100.0, "pct2", align="center"),
            TableCell(es_score, "pct2", align="center"),
            TableCell(env_f, "pct2", align="center"),
        ])

    return TablePage(
        credit_union=cu,
        title="Environmental Factor for PLL",
        heading_lines=[f"For Quarter Ending {_rv._snap_display(snapshot_date)}"],
        sections=[
            TableSection(title="Economic Stress Index Calculation",
                         columns=state_cols, rows=state_rows),
            TableSection(columns=county_cols, rows=county_rows),
            TableSection(columns=pool_cols, rows=pool_rows),
        ])


def build_co_recov_dq(client_name: str, snapshot_date: str, config: dict,
                      hist: dict | None = None, df: Any = None,
                      grades: Any = None) -> TablePage | None:
    """Display CO-Recov-DQ: Charge-offs, Recoveries, Net Charge-offs and
    Delinquency %, each by pool across the WARM look-back years -- computed
    from ``hist`` (windowing logic ported from report_vizo._sheet_co_recov_dq).
    """
    import report_vizo as _rv

    if df is None:
        return None
    cfg = config or {}
    cu = cfg.get("credit_union") or client_name
    pools = _rv._ordered_pools(df, hist)
    if not pools:
        return None
    h = hist or {}
    co_data = h.get("chargeoffs", {})
    rc_data = h.get("recoveries", {})
    dq_pct = h.get("dq_pct", {})
    years = h.get("years", []) or list(range(2019, int(snapshot_date[:4]) + 1))
    _imp = h.get("impaired", {}) or {}
    acl_months_map = _imp.get("acl_months", {})
    snap_year = int(snapshot_date[:4])
    snap_month = int(snapshot_date[5:7])
    if pools and years:
        _max_lol = max(acl_months_map.get(p, 36) for p in pools)
        _abs_first = (snap_year * 12 + snap_month) - _max_lol + 1
        _cutoff = (_abs_first - 1) // 12
        years = [y for y in years if y >= _cutoff]
    year_strs = [str(y) for y in years]

    warm_co = _imp.get("warm_co", {})
    warm_rc = _imp.get("warm_rc", {})
    use_warm = bool(warm_co)
    warm_co_monthly = _imp.get("warm_co_monthly", {}) or h.get("co_monthly", {})
    warm_rc_monthly = _imp.get("warm_rc_monthly", {}) or h.get("rc_monthly", {})
    co_monthly = h.get("co_monthly", {})
    rc_monthly = h.get("rc_monthly", {})

    def _window_start(pool):
        pool_acl = acl_months_map.get(pool, 36)
        abs_first = (snap_year * 12 + snap_month) - pool_acl + 1
        ey = (abs_first - 1) // 12
        return ey, abs_first - ey * 12

    def _windowed(monthly_data, yearly_data, pool, year, ey, em):
        if year != ey:
            return yearly_data.get(year, {}).get(pool, 0)
        partial = 0
        has_window = False
        for m in range(em, 13):
            v = monthly_data.get((year, m), {}).get(pool, 0)
            if v:
                has_window = True
            partial += v
        has_any = has_window or any(
            monthly_data.get((year, m), {}).get(pool, 0) for m in range(1, em))
        if has_any:
            full_year = yearly_data.get(year, {}).get(pool, 0)
            if full_year and partial and (full_year > 0) != (partial > 0):
                partial = -partial
            return partial
        full = yearly_data.get(year, {}).get(pool, 0)
        return full * (12 - em + 1) / 12 if full else 0

    def _warm_months(pool):
        return acl_months_map.get(pool, cfg.get("warm_months", {}).get(pool, 36))

    def _year_labels():
        labels = list(year_strs)
        if labels:
            labels[-1] = f"YTD {year_strs[-1]}"
        return labels

    def _flow_section(title, total_label, yearly, monthly, net=False):
        cols = [title] + _year_labels() + [total_label, "WARM Months"]
        rows: list[list[TableCell]] = []
        for pool in pools:
            ey, em = _window_start(pool)
            cells = [TableCell(pool, "text", bold=True, align="left")]
            total = 0
            for y in years:
                if y < ey:
                    cells.append(TableCell(None))
                    continue
                if net:
                    cv = _windowed(warm_co_monthly if use_warm else co_monthly,
                                   warm_co if use_warm else co_data, pool, y, ey, em)
                    rv = _windowed(warm_rc_monthly if use_warm else rc_monthly,
                                   warm_rc if use_warm else rc_data, pool, y, ey, em)
                    val = abs(cv) - abs(rv)
                else:
                    val = abs(_windowed(monthly, yearly, pool, y, ey, em) or 0)
                cells.append(TableCell(val, "currency"))
                total += val
            cells.append(TableCell(total, "currency", bold=True))
            cells.append(TableCell(_warm_months(pool), "text", align="center"))
            rows.append(cells)
        return TableSection(title=title, columns=cols, rows=rows)

    sections = [
        _flow_section("Charge offs", "ACL Charge offs",
                      warm_co if use_warm else co_data,
                      warm_co_monthly if use_warm else co_monthly),
        _flow_section("Recoveries", "ACL Recoveries",
                      warm_rc if use_warm else rc_data,
                      warm_rc_monthly if use_warm else rc_monthly),
        _flow_section("Net Charge offs", "Net Charge offs", None, None, net=True),
    ]

    warm_dq = _imp.get("warm_dq_pct", {})
    use_dq = warm_dq if warm_dq else dq_pct
    dq_cols = ["DQ %"] + _year_labels() + ["Average", "Variance"]
    dq_rows: list[list[TableCell]] = []
    for pool in pools:
        ey = _window_start(pool)[0]
        cells = [TableCell(pool, "text", bold=True, align="left")]
        rates = []
        for y in years:
            if y < ey:
                cells.append(TableCell(None))
                continue
            val = use_dq.get(y, {}).get(pool, 0)
            cells.append(TableCell(val, "pct2"))
            rates.append(val)
        avg = sum(rates) / len(rates) if rates else 0
        var = rates[-1] - avg if len(rates) > 1 else 0
        cells.append(TableCell(avg, "pct2", bold=True))
        cells.append(TableCell(var, "pct2", bold=True))
        dq_rows.append(cells)
    sections.append(TableSection(title="Delinquency", columns=dq_cols, rows=dq_rows))

    return TablePage(
        credit_union=cu,
        title="Delinquency Calculation",
        heading_lines=[f"For Quarter Ending {_rv._snap_display(snapshot_date)}"],
        sections=sections)


def build_loss_factor(client_name: str, snapshot_date: str, config: dict,
                      hist: dict | None = None, df: Any = None,
                      grades: Any = None) -> TablePage | None:
    """Display HIst Bal -- Loss Factor Calculation.  Per-grade annual average
    balances across the WARM window plus each grade's Life Loss Rate,
    Distribution Factor, ACL Base Loss Rate and % of Loans.  Ported from
    report_vizo._sheet_loss_factor (left balance grid + right rate summary
    combined into one wide table for the PDF).
    """
    import report_vizo as _rv

    if df is None:
        return None
    cfg = config or {}
    cu = cfg.get("credit_union") or client_name
    no_score = cfg.get("no_score_label", "Not Reported")
    gl = [g for g in _rv._all_grades(grades, no_score) if not _rv._is_hidden(g)]
    brr_labels = _rv._brr_grade_labels(cfg, no_score)
    brr_pool_lcs = _rv._brr_pools_set(cfg) if brr_labels else set()

    pools = _rv._ordered_pools(df, hist)
    if not pools:
        return None
    h = hist or {}
    co_data = h.get("chargeoffs", {})
    rc_data = h.get("recoveries", {})
    avg_bals = h.get("avg_balances", {})
    years = h.get("years", [])
    _imp = h.get("impaired", {}) or {}
    acl_months_map = _imp.get("acl_months", {})
    snap_year = int(snapshot_date[:4])
    snap_month = int(snapshot_date[5:7])
    if pools and years:
        _max_lol = max(acl_months_map.get(p, 36) for p in pools)
        _abs_first = (snap_year * 12 + snap_month) - _max_lol + 1
        _cutoff = (_abs_first - 1) // 12
        years = [y for y in years if y >= _cutoff]
    num_years = len(years)

    hbd = _imp.get("hist_bal_data", {})
    annual_grade_avg: dict = {}
    for _pk, pdata in hbd.items():
        _dates = pdata.get("dates", [])
        _grades_data = pdata.get("grades", {})
        annual_grade_avg[_pk] = {}
        for _gk, _vals in _grades_data.items():
            if _gk.upper().startswith("HIDE"):
                continue
            yr_sums: dict = {}
            yr_cnts: dict = {}
            for _i, _d in enumerate(_dates):
                if _i < len(_vals) and _vals[_i] > 0:
                    yr_sums[_d.year] = yr_sums.get(_d.year, 0) + _vals[_i]
                    yr_cnts[_d.year] = yr_cnts.get(_d.year, 0) + 1
            for _y in yr_sums:
                annual_grade_avg[_pk].setdefault(_y, {})
                annual_grade_avg[_pk][_y][_gk] = yr_sums[_y] / yr_cnts[_y]

    def _pool_earliest_year(pool):
        pool_acl = acl_months_map.get(pool, 36)
        abs_first = (snap_year * 12 + snap_month) - pool_acl + 1
        return (abs_first - 1) // 12

    warm_net_co = _imp.get("warm_net_co", {})
    pool_life_rates: dict = {}
    pool_avg_totals: dict = {}
    for pool in pools:
        pe = _pool_earliest_year(pool)
        pa = annual_grade_avg.get(pool, {})
        yr_tots = []
        for y in years:
            if y < pe:
                continue
            yt = sum(pa.get(y, {}).values())
            if not yt:
                yt = avg_bals.get(y, {}).get(pool, 0)
            if yt:
                yr_tots.append(yt)
        avg_tot = sum(yr_tots) / len(yr_tots) if yr_tots else 0
        pool_avg_totals[pool] = avg_tot
        pool_stripped = pool.strip()
        net_co_match = warm_net_co.get(pool_stripped, warm_net_co.get(pool, None))
        if net_co_match is not None:
            total_net = net_co_match
        else:
            total_net = 0
            for y in years:
                if y < pe:
                    continue
                total_net += abs(co_data.get(y, {}).get(pool, 0) or 0) \
                    - abs(rc_data.get(y, {}).get(pool, 0) or 0)
        pool_life_rates[pool] = total_net / avg_tot if avg_tot > 0 else 0

    year_strs = [str(y) for y in years]
    year_labels = list(year_strs)
    if year_labels:
        year_labels[-1] = f"YTD {year_strs[-1]}"
    cols = (["Current Grade"] + year_labels
            + ["Average Balance", "Life Loss Rate", "Distribution Factor",
               "ACL Base Loss Rate", "% of Loans", "WARM Months"])
    width = len(cols)

    def _row(first, year_cells, tail, bold=False):
        cells = [first]
        cells += year_cells + [TableCell(None)] * (num_years - len(year_cells))
        cells += tail
        cells += [TableCell(None)] * (width - len(cells))
        return cells

    risk_rated_map = _imp.get("risk_rated", {})
    rows: list[list[TableCell]] = []
    for pool in pools:
        pool_earliest = _pool_earliest_year(pool)
        # Pool title spans the whole row so it never wraps in the label column.
        rows.append([TableCell(pool, "text", bold=True, align="left",
                               colspan=width)])
        pdf = df[df["loan_pool"] == pool]
        pool_total = pdf["current_balance"].sum()
        pool_ll = pool_life_rates.get(pool, 0)
        is_rr = risk_rated_map.get(pool, True)
        pool_annual = annual_grade_avg.get(pool, {})

        if not is_rr:
            yc = []
            for yi in range(num_years):
                if years[yi] < pool_earliest:
                    yc.append(TableCell(None))
                    continue
                yt = sum(pool_annual.get(years[yi], {}).values()) \
                    or avg_bals.get(years[yi], {}).get(pool, 0)
                yc.append(TableCell(yt, "currency") if yt else TableCell(None))
            nrr_avg = pool_avg_totals.get(pool, 0)
            warm = acl_months_map.get(pool, cfg.get("warm_months", {}).get(pool, 36))
            rows.append(_row(
                TableCell("Total", "text", bold=True, align="left"), yc,
                [TableCell(nrr_avg, "currency", bold=True),
                 TableCell(pool_ll, "pct2", bold=True),
                 TableCell(None), TableCell(None),
                 TableCell(1.0, "pct2", bold=True),
                 TableCell(warm, "text", align="center")]))
            continue

        pool_grade_labels = (
            brr_labels if (brr_labels and _rv._is_brr_pool(pool, brr_pool_lcs)) else gl)
        for gi, g in enumerate(pool_grade_labels):
            g_df = pdf[pdf["current_grade"] == g]
            balance = g_df["current_balance"].sum()
            yr_vals = []
            yc = []
            for yi in range(num_years):
                if years[yi] < pool_earliest:
                    yc.append(TableCell(None))
                    continue
                grade_avg = pool_annual.get(years[yi], {}).get(g, 0)
                if not grade_avg:
                    avg = avg_bals.get(years[yi], {}).get(pool, 0)
                    grade_avg = avg * (balance / pool_total) if pool_total and avg else 0
                if grade_avg:
                    yr_vals.append(grade_avg)
                    yc.append(TableCell(grade_avg, "currency"))
                else:
                    yc.append(TableCell(None))
            avg_bal = sum(yr_vals) / len(yr_vals) if yr_vals else 0
            dist = (_rv._dist_factor(len(_rv.DIST_FACTORS) - 1)
                    if g == no_score else _rv._dist_factor(gi))
            base_rate = max(0, pool_ll * dist)
            pct_pool = balance / pool_total if pool_total else 0
            warm_cell = (TableCell(
                acl_months_map.get(pool, cfg.get("warm_months", {}).get(pool, 36)),
                "text", align="center") if gi == 0 else TableCell(None))
            rows.append(_row(
                TableCell(g, "text", align="left"), yc,
                [TableCell(avg_bal, "currency"),
                 TableCell(pool_ll, "pct2"), TableCell(dist, "pct2"),
                 TableCell(base_rate, "pct2"), TableCell(pct_pool, "pct2"),
                 warm_cell]))

        yc = []
        for yi in range(num_years):
            if years[yi] < pool_earliest:
                yc.append(TableCell(None))
                continue
            yr_total = sum(pool_annual.get(years[yi], {}).values()) \
                or avg_bals.get(years[yi], {}).get(pool, 0)
            yc.append(TableCell(yr_total, "currency", bold=True) if yr_total else TableCell(None))
        rr_avg = pool_avg_totals.get(pool, 0)
        rows.append(_row(
            TableCell("Total", "text", bold=True, align="left"), yc,
            [TableCell(rr_avg, "currency", bold=True),
             TableCell(pool_ll, "pct2", bold=True),
             TableCell(None), TableCell(None),
             TableCell(1.0, "pct2", bold=True), TableCell(None)]))

    yc = []
    for yi in range(num_years):
        ytot = sum(sum(annual_grade_avg.get(p, {}).get(years[yi], {}).values())
                   for p in pools)
        yc.append(TableCell(ytot, "currency", bold=True) if ytot else TableCell(None))
    grand_avg = sum(pool_avg_totals.get(p, 0) for p in pools)
    rows.append(_row(
        TableCell("Grand Total", "text", bold=True, align="left"), yc,
        [TableCell(grand_avg, "currency", bold=True),
         TableCell(None), TableCell(None), TableCell(None),
         TableCell(1.0, "pct2", bold=True), TableCell(None)]))

    # Group rows into per-pool blocks (+ a Grand Total block) so no pool splits
    # across a page. A pool starts at its bold header row (label only, no data).
    groups: list = []
    cur: list | None = None
    for r in rows:
        c0 = r[0]
        is_header = (getattr(c0, "value", None) and getattr(c0, "bold", False)
                     and all(getattr(c, "value", None) is None for c in r[1:]))
        is_grand = getattr(c0, "value", None) == "Grand Total"
        if is_header or is_grand or cur is None:
            cur = [r]
            groups.append(cur)
        else:
            cur.append(r)

    return TablePage(
        credit_union=cu,
        title="Loss Factor Calculation",
        heading_lines=[f"For Quarter Ending {_rv._snap_display(snapshot_date)}"],
        sections=[TableSection(columns=cols, row_groups=groups)],
        css_class="loss-factor")


def _acl_current_shape(acl_pools: dict, acl_summary: dict, acl_impaired: dict) -> dict:
    """Shape the current-period ACL data into the same dict layout
    change_analysis._parse_acl_sheet produces, so the diff + commentary code
    path is byte-identical to the workbook's."""
    pools: dict = {}
    order: list = []
    for name, pdata in (acl_pools or {}).items():
        t = pdata.get("total") or {}
        pools[name] = {
            "balance": float(t.get("balance") or 0),
            "spec_id": float(t.get("spec_id") or 0),
            "total_allow": float(t.get("total_allow") or 0),
        }
        order.append(name)
    impaired = {k: float(v or 0) for k, v in (acl_impaired or {}).items()}
    s = acl_summary or {}
    totals = {k: float(s.get(k) or 0) for k in (
        "pooled_balance", "pooled_total_allow", "total_spec_allow",
        "total_allow_needed", "acl_balance", "adjustment")}
    return {"pools": pools, "order": order, "impaired": impaired, "totals": totals}


def _load_prior_acl(cu: str, snapshot_date: str, config: dict):
    """Locate + parse the prior quarter's ACL data. Prefers the JSON sidecar
    (no .xlsx dependency); falls back to reading the prior workbook's ACL Env
    tab for periods generated before sidecars existed. Returns
    (prior_dict | None, prior_snap | None)."""
    cfg = config or {}
    rpt_dir = (cfg.get("report_dir") or cfg.get("output_dir")
               or os.path.join(os.environ.get("CECL_WORKSPACE_ROOT")
                               or os.getcwd(), "Reports"))
    # 1) Prefer the sidecar.
    try:
        from . import acl_store
        shape, prior_snap = acl_store.load_prior_snapshot(
            rpt_dir, cu, snapshot_date)
        if shape and shape.get("totals"):
            return shape, prior_snap
    except Exception as exc:  # noqa: BLE001
        print(f"  prior ACL sidecar unavailable: {exc}")
    # 2) Fall back to the prior workbook.
    try:
        import openpyxl
        from change_analysis import _find_prior_report, _parse_acl_sheet
        from report_vizo import ACL_SHEET
        safe_cu = cu.replace(" ", "_").replace("/", "-")
        path, prior_snap = _find_prior_report(rpt_dir, safe_cu, "Vizo_Model",
                                              snapshot_date)
        if not path:
            return None, None
        pwb = openpyxl.load_workbook(path, data_only=True)
        prior = (_parse_acl_sheet(pwb[ACL_SHEET])
                 if ACL_SHEET in pwb.sheetnames else None)
        pwb.close()
        return prior, prior_snap
    except Exception as exc:  # noqa: BLE001 - never break the PDF path
        print(f"  prior ACL unavailable: {exc}")
        return None, None


_NO_PRIOR_NOTE = ("No prior report is available for comparison - this is the "
                  "earliest report on file for this credit union.")


def build_summary_variance(client_name: str, snapshot_date: str, config: dict,
                           hist: dict | None = None, df: Any = None,
                           grades: Any = None) -> SummaryVariancePage | None:
    """Summary Variance -- the SCALE 'Executive Summary-Vizo' banded card:
    a centred 3-line title over Current / Prior / Change blocks of the four
    ACL measures."""
    from . import format as _fmt

    if df is None:
        return None
    acl_pools, acl_summary, acl_impaired = _acl_data(
        config, snapshot_date, hist, df, grades)
    if not acl_pools:
        return None
    cu = (config or {}).get("credit_union") or client_name
    cur = _acl_current_shape(acl_pools, acl_summary, acl_impaired)
    prior, prior_snap = _load_prior_acl(cu, snapshot_date, config)

    def _md(snap):
        s = str(snap)[:10]
        try:
            y, m, d = s.split("-")
            return f"{int(m)}/{int(d)}/{int(y)}"
        except Exception:  # noqa: BLE001
            return s

    def _ratio(t):
        pb = t.get("pooled_balance") or 0
        return (t.get("total_allow_needed", 0) / pb) if pb else None

    labels = ["Total Expected Losses on Loans", "Current ACL Balance",
              "Adjustment", "ACL/Total Loans"]

    def _measures(vals):
        out = []
        for i, v in enumerate(vals):
            if v is None:
                out.append("")
            else:
                out.append(_fmt.pct2(v) if i == 3 else _fmt.acct2(v))
        return list(zip(labels, out))

    ct = cur["totals"]
    cur_vals = [ct["total_allow_needed"], ct["acl_balance"], ct["adjustment"],
                _ratio(ct)]
    blocks = [SummaryVarianceBlock("Current ACL", _md(snapshot_date),
                                   _measures(cur_vals))]
    note = ""
    if prior:
        pt = prior["totals"]
        prior_vals = [pt.get("total_allow_needed", 0), pt.get("acl_balance", 0),
                      pt.get("adjustment", 0), _ratio(pt)]
        change_vals = [(c - p) if (c is not None and p is not None) else None
                       for c, p in zip(cur_vals, prior_vals)]
        blocks.append(SummaryVarianceBlock("Prior ACL", _md(prior_snap),
                                           _measures(prior_vals)))
        blocks.append(SummaryVarianceBlock("Change", "", _measures(change_vals)))
    else:
        note = _NO_PRIOR_NOTE

    return SummaryVariancePage(credit_union=cu, quarter_ended=_md(snapshot_date),
                               blocks=blocks, note=note)


def build_change_analysis(client_name: str, snapshot_date: str, config: dict,
                          hist: dict | None = None, df: Any = None,
                          grades: Any = None) -> TablePage | None:
    """Change Analysis: per-pool ACL variance, impaired variance, summary
    metrics, and the auto-generated significant-change commentary -- current
    from data, prior from the prior quarter's workbook."""
    from change_analysis import _explain_impaired, _explain_pool, _is_significant

    if df is None:
        return None
    acl_pools, acl_summary, acl_impaired = _acl_data(
        config, snapshot_date, hist, df, grades)
    if not acl_pools:
        return None
    cfg = config or {}
    cu = cfg.get("credit_union") or client_name
    cur = _acl_current_shape(acl_pools, acl_summary, acl_impaired)
    prior, prior_snap = _load_prior_acl(cu, snapshot_date, config)

    if not prior:
        return TablePage(
            credit_union=cu, title="Change Analysis - Period over Period",
            heading_lines=[f"Current period: {snapshot_date}"],
            notes=[_NO_PRIOR_NOTE])

    all_pools = list(cur["order"]) + [p for p in prior["order"]
                                      if p not in cur["order"]]
    rows_data = []
    for p in all_pools:
        c = cur["pools"].get(p, {})
        pr = prior["pools"].get(p, {})
        ca, pa = c.get("total_allow", 0.0), pr.get("total_allow", 0.0)
        rows_data.append({
            "pool": p, "cur": ca, "prior": pa, "delta": ca - pa,
            "pct": (ca - pa) / pa if pa else None,
            "cur_bal": c.get("balance", 0.0), "prior_bal": pr.get("balance", 0.0),
            "cur_spec": c.get("spec_id", 0.0), "prior_spec": pr.get("spec_id", 0.0),
        })

    pool_rows = []
    for rd in rows_data:
        pool_rows.append([
            TableCell(rd["pool"], "text", align="left"),
            TableCell(rd["cur"], "currency"), TableCell(rd["prior"], "currency"),
            TableCell(rd["delta"], "currency"),
            TableCell(rd["pct"], "pct") if rd["pct"] is not None else TableCell(None)])
    ctot = cur["totals"].get("pooled_total_allow", 0.0)
    ptot = prior["totals"].get("pooled_total_allow", 0.0)
    pool_rows.append([
        TableCell("Pooled Total Allowance", "text", bold=True, align="left"),
        TableCell(ctot, "currency", bold=True), TableCell(ptot, "currency", bold=True),
        TableCell(ctot - ptot, "currency", bold=True),
        TableCell((ctot - ptot) / ptot if ptot else None, "pct", bold=True)])
    sec1 = TableSection(title="Pool ACL Allowance Variance",
                        columns=["Pool", "Current", "Prior", "$ Change", "% Change"],
                        rows=pool_rows)

    cats = list(cur["impaired"].keys()) + [c for c in prior["impaired"]
                                           if c not in cur["impaired"]]
    imp_rows = []
    for cat in cats:
        cv = cur["impaired"].get(cat, 0.0)
        pv = prior["impaired"].get(cat, 0.0)
        imp_rows.append([TableCell(cat, "text", align="left"),
                         TableCell(cv, "currency"), TableCell(pv, "currency"),
                         TableCell(cv - pv, "currency")])
    cimp = cur["totals"].get("total_spec_allow", sum(cur["impaired"].values()))
    pimp = prior["totals"].get("total_spec_allow", sum(prior["impaired"].values()))
    imp_rows.append([
        TableCell("Total Specifically Identified", "text", bold=True, align="left"),
        TableCell(cimp, "currency", bold=True), TableCell(pimp, "currency", bold=True),
        TableCell(cimp - pimp, "currency", bold=True)])
    sec2 = TableSection(title="Impaired Loan Reserve Variance",
                        columns=["Impairment Type", "Current", "Prior", "$ Change"],
                        rows=imp_rows)

    sum_rows = []
    for label, key in (("Total Allowance Needed", "total_allow_needed"),
                       ("Allowance for Credit Loss Balance", "acl_balance"),
                       ("Adjustment (Over)/Under-funded", "adjustment")):
        cv = cur["totals"].get(key, 0.0)
        pv = prior["totals"].get(key, 0.0)
        sum_rows.append([TableCell(label, "text", align="left"),
                         TableCell(cv, "currency"), TableCell(pv, "currency"),
                         TableCell(cv - pv, "currency")])
    sec3 = TableSection(title="Summary",
                        columns=["Metric", "Current", "Prior", "$ Change"],
                        rows=sum_rows)

    notes = []
    specific_pools = {str(p).strip().lower()
                      for p in (cfg.get("warm_allowance_pools") or [])}
    for rd in sorted(rows_data, key=lambda x: -abs(x["delta"])):
        if _is_significant(rd):
            notes.append(_explain_pool(
                rd, specific=rd["pool"].strip().lower() in specific_pools))
    imp_note = _explain_impaired(cur["impaired"], prior["impaired"], cimp, pimp)
    if imp_note:
        notes.append(imp_note)
    tot_d = cur["totals"].get("total_allow_needed", 0.0) \
        - prior["totals"].get("total_allow_needed", 0.0)
    if abs(tot_d) >= 1000:
        notes.insert(0, f"Total Allowance Needed "
                        f"{'increased' if tot_d > 0 else 'decreased'} "
                        f"${abs(tot_d):,.0f} versus the prior report, to "
                        f"${cur['totals'].get('total_allow_needed', 0.0):,.0f}.")
    if not notes:
        notes.append("No pool reserve moved materially versus the prior report; "
                     "changes were within normal quarter-to-quarter variation.")

    return TablePage(
        credit_union=cu, title="Change Analysis - Period over Period",
        heading_lines=[f"Current: {snapshot_date}    |    Prior: {prior_snap}"],
        sections=[sec1, sec2, sec3],
        notes_title="Analysis of Significant Changes", notes=notes)


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
    if not supplemental:
        pages.append(("narrative.html",
                      {"page": build_report_index(client_name, config)}, False))
        if df is not None:
            var = build_summary_variance(client_name, snapshot_date, config, hist,
                                         df=df, grades=grades)
            if var is not None:
                pages.append(("summary_variance.html", {"page": var}, False))
        impd = build_impr_deter(client_name, snapshot_date, config, hist,
                                df=df, grades=grades)
        impd_charts = (render_chart_specs(impr_deter_charts(df, grades, config, hist))
                       if df is not None else [])
        pages.append(("impr_deter.html", {"page": impd, "charts": impd_charts}, True))
    if df is not None and not supplemental:
        import report_vizo as _rv
        rc = build_risk_change(client_name, snapshot_date, df, config, grades, hist)
        rc_charts = render_chart_specs(
            risk_change_ncc_chart(df, grades, config) + risk_change_charts(hist))
        pages.append(("risk_change.html", {"page": rc, "charts": rc_charts}, True))
        # One page per risk-rated pool, mirroring the workbook's Risk Chg tabs.
        _rr = ((hist or {}).get("impaired", {}) or {}).get("risk_rated", {})
        for _pool in _rv._ordered_pools(df, hist):
            if not _rr.get(_pool, True):
                continue
            _pdf = df[df["loan_pool"] == _pool]
            if _pdf.empty:
                continue
            _prc = build_risk_change(client_name, snapshot_date, _pdf, config,
                                     grades, hist, pool_name=_pool)
            _pcharts = render_chart_specs(
                risk_change_ncc_chart(_pdf, grades, config)
                + risk_change_charts(hist, _pool))
            pages.append(("risk_change.html",
                          {"page": _prc, "charts": _pcharts}, True))
    if not supplemental:
        acl = build_acl_env(client_name, snapshot_date, config, hist,
                           df=df, grades=grades)
        if acl is not None:
            pages.append(("acl_env.html", {"page": acl, "charts": []}, False))
        env = build_env_factor(client_name, snapshot_date, config, hist,
                               df=df, grades=grades)
        if env is not None:
            pages.append(("table_page.html", {"page": env}, True))
        loss = build_loss_factor(client_name, snapshot_date, config, hist,
                                 df=df, grades=grades)
        if loss is not None:
            pages.append(("table_page.html", {"page": loss}, False))
        codq = build_co_recov_dq(client_name, snapshot_date, config, hist,
                                 df=df, grades=grades)
        if codq is not None:
            pages.append(("table_page.html", {"page": codq}, True))
        acl_sum = build_acl_summary(client_name, snapshot_date, config, hist,
                                    df=df, grades=grades)
        if acl_sum is not None:
            pages.append(("table_page.html", {"page": acl_sum}, True))
        mgmt = build_mgmt_adj_summary(client_name, snapshot_date, config, hist,
                                      df=df, grades=grades)
        if mgmt is not None:
            pages.append(("table_page.html", {"page": mgmt}, True))
        impaired = build_impaired_loans(client_name, snapshot_date, config, hist,
                                        df=df, grades=grades)
        if impaired is not None:
            pages.append(("table_page.html", {"page": impaired}, False))
        chg = build_change_analysis(client_name, snapshot_date, config, hist,
                                    df=df, grades=grades)
        if chg is not None:
            pages.append(("table_page.html", {"page": chg}, False))
        pages.append(("narrative.html",
                      {"page": build_introduction(client_name, config)}, False))
        pages.append(("narrative.html",
                      {"page": build_exec_summary_narrative(client_name, config)}, False))
    return {"cover": cover, "pages": pages}
