"""Adapter: read a generated report workbook into the ReportModel.

Phase-0 data source. Populating the model from the already-generated
xlsx guarantees the PDF numbers match the Excel report exactly, and it
doubles as the "extract a page" capability. Later phases can add a
compute-time populator with the same model as output — templates don't
change.

Scanning is label-anchored (find the row by its text, not a fixed
coordinate) so it survives minor row-position drift between report
versions.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from .model import (
    AclEnvPage,
    AclPoolRow,
    AdjustmentRow,
    GradeMigrationRow,
    ImprDeterPage,
    KeyValueRow,
    MatrixCell,
    MatrixRow,
    PoolMigrationRow,
    RiskChangeMatrix,
    RiskChangePage,
)

_IMPR_DETER_TAB_HINTS = ("impr deter", "improved", "deteriorated")
_CECL_LABELS = (
    "Total Specifically Identified Allowance",
    "Total Allowance Needed",
    "Allowance for Credit Loss Balance",
    "Adjustment (Overfunded)",
)


def _find_tab(wb, hints: tuple[str, ...]) -> str | None:
    for sn in wb.sheetnames:
        low = sn.strip().lower()
        if any(h in low for h in hints):
            return sn
    return None


def _num(v):
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _cell_str(v) -> str:
    return "" if v is None else str(v).strip()


def load_impr_deter(report_path: str | Path) -> ImprDeterPage:
    """Read the 'Impr Deter' tab of a Vizo Model workbook into the model."""
    wb = load_workbook(report_path, data_only=True)
    tab = _find_tab(wb, _IMPR_DETER_TAB_HINTS)
    if tab is None:
        raise ValueError(
            f"No Improved/Deteriorated tab found in {Path(report_path).name} "
            f"(sheets: {wb.sheetnames})"
        )
    ws = wb[tab]

    # ── Header lines: first 4 populated rows of column A ──
    heading: list[str] = []
    cu = ""
    period = ""
    for r in range(1, 6):
        v = ws.cell(r, 1).value
        if v is None or v == "":
            continue
        if r == 1:
            cu = _cell_str(v)
        elif hasattr(v, "strftime"):
            period = v.strftime("%m-%d-%y")
            heading.append(period)
        else:
            heading.append(_cell_str(v))

    # ── CECL Adjustment box: label in a left column, value ~3 cols right ──
    cecl: list[KeyValueRow] = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            label = _cell_str(ws.cell(r, c).value)
            if not label:
                continue
            for want in _CECL_LABELS:
                if label.startswith(want) or want.split(" as of")[0] in label:
                    # value: first numeric cell to the right on this row
                    val = None
                    for cc in range(c + 1, ws.max_column + 1):
                        n = _num(ws.cell(r, cc).value)
                        if n is not None:
                            val = n
                            break
                    cecl.append(KeyValueRow(label=label, value=val))
                    break

    # ── Improved/Deteriorated by pool: anchor on a "Loan Type" header ──
    by_pool: list[PoolMigrationRow] = []
    anchor = _find_label(ws, "Loan Type")
    if anchor:
        hr, hc = anchor
        r = hr + 1
        while r <= ws.max_row:
            pool = _cell_str(ws.cell(r, hc).value)
            if not pool:
                break
            by_pool.append(PoolMigrationRow(
                pool=pool,
                improved=_num(ws.cell(r, hc + 1).value),
                deteriorated=_num(ws.cell(r, hc + 2).value),
                net_change=_num(ws.cell(r, hc + 3).value),
            ))
            r += 1

    # ── Improved/Deteriorated by grade: anchor on a "Grade" header ──
    by_grade: list[GradeMigrationRow] = []
    ganchor = _find_label(ws, "Grade")
    if ganchor:
        hr, hc = ganchor
        r = hr + 1
        while r <= ws.max_row:
            grade = _cell_str(ws.cell(r, hc).value)
            if not grade:
                break
            by_grade.append(GradeMigrationRow(
                grade=grade,
                balance=_num(ws.cell(r, hc + 1).value),
                improved=_num(ws.cell(r, hc + 3).value),
                deteriorated=_num(ws.cell(r, hc + 4).value),
            ))
            r += 1

    return ImprDeterPage(
        credit_union=cu,
        period_ending=period,
        heading_lines=heading,
        cecl_adjustment=cecl,
        by_pool=by_pool,
        by_grade=by_grade,
    )


def _find_label(ws, text: str) -> tuple[int, int] | None:
    """Return (row, col) of the first cell whose stripped value == text."""
    low = text.strip().lower()
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            if _cell_str(ws.cell(r, c).value).lower() == low:
                return (r, c)
    return None


# ── Risk Change matrices ─────────────────────────────────────────────

_RISK_TAB_HINTS = ("risk change",)
# Fill colors (last 6 hex) -> cell state.
_FILL_STATE = {
    "0D4D5E": "header",
    "829901": "improved",
    "873A3A": "deteriorated",
}


def _fill_state(styled_cell) -> str:
    try:
        fill = styled_cell.fill
        if fill and fill.patternType:
            rgb = fill.fgColor.rgb
            if isinstance(rgb, str):
                return _FILL_STATE.get(rgb[-6:].upper(), "plain")
    except Exception:  # noqa: BLE001
        pass
    return "plain"


def _read_matrix(wsv, wsf, hr: int, *, is_pct: bool) -> RiskChangeMatrix:
    """Read one matrix whose corner ('$/% Current Grade') header is at
    row ``hr``, column 1. Columns C..K (3..11) are original-grade
    buckets, L (12) is Grand Total, N/O/P (14..16) are the side panel.
    """
    corner = _cell_str(wsv.cell(hr, 1).value)
    col_headers = [
        _cell_str(wsv.cell(hr, c).value) for c in range(3, 12)
        if _cell_str(wsv.cell(hr, c).value)
    ]
    side_headers = [
        _cell_str(wsv.cell(hr, c).value) for c in range(14, 17)
        if _cell_str(wsv.cell(hr, c).value)
    ]
    rows: list[MatrixRow] = []
    r = hr + 1
    while r <= wsv.max_row:
        label = _cell_str(wsv.cell(r, 1).value)
        if not label or label in ("Balance Adjustment", "Total in Portfolio"):
            break
        is_total_row = label.lower() == "grand total"
        cells = [
            MatrixCell(
                value=_num(wsv.cell(r, c).value),
                state=_fill_state(wsf.cell(r, c)),
                is_pct=is_pct,
                bold=is_total_row,
            )
            for c in range(3, 3 + len(col_headers))
        ]
        total = MatrixCell(
            value=_num(wsv.cell(r, 12).value), is_pct=is_pct, bold=True,
        )
        side = [
            MatrixCell(value=_num(wsv.cell(r, c).value), is_pct=is_pct)
            for c in range(14, 14 + len(side_headers))
        ] if side_headers else []
        rows.append(MatrixRow(label=label, cells=cells, total=total, side=side))
        r += 1
    return RiskChangeMatrix(
        corner=corner, col_headers=col_headers, rows=rows,
        side_headers=side_headers, is_pct=is_pct,
    )


def load_risk_change(report_path: str | Path) -> RiskChangePage:
    """Read the 'Risk Change Total' tab into the model (both matrices)."""
    wbv = load_workbook(report_path, data_only=True)
    wbf = load_workbook(report_path, data_only=False)
    tab = _find_tab(wbv, _RISK_TAB_HINTS)
    if tab is None:
        raise ValueError(
            f"No Risk Change tab in {Path(report_path).name} "
            f"(sheets: {wbv.sheetnames})"
        )
    wsv, wsf = wbv[tab], wbf[tab]

    heading: list[str] = []
    cu = _cell_str(wsv.cell(1, 1).value)
    for r in range(2, 5):
        v = _cell_str(wsv.cell(r, 1).value)
        if v:
            heading.append(v)

    matrices: list[RiskChangeMatrix] = []
    for r in range(1, wsv.max_row + 1):
        label = _cell_str(wsv.cell(r, 1).value)
        if label in ("$ Current Grade", "% Current Grade"):
            matrices.append(_read_matrix(
                wsv, wsf, r, is_pct=label.startswith("%")))

    summary: list[KeyValueRow] = []
    for name in ("Balance Adjustment", "Total in Portfolio"):
        loc = _find_label(wsv, name)
        if loc:
            r, _c = loc
            val = _num(wsv.cell(r, 12).value)  # column L
            summary.append(KeyValueRow(label=name, value=val))

    return RiskChangePage(
        credit_union=cu, heading_lines=heading,
        matrices=matrices, summary=summary,
    )


# ── ACL Env by Pool ──────────────────────────────────────────────────

_ACL_TAB_HINTS = ("acl env",)
_ACL_ADJ_LABELS = (
    "Total Specifically Identified Allowance",
    "Total Allowance Needed",
    "Allowance for Credit Loss Balance",
    "Adjustment (Overfunded)",
)


def _bold(wsf, r: int, c: int) -> bool:
    try:
        return bool(wsf.cell(r, c).font.bold)
    except Exception:  # noqa: BLE001
        return False


def load_acl_env(report_path: str | Path) -> AclEnvPage:
    """Read the 'ACL Env by Pool Mgmt Adj' tab into the model."""
    wbv = load_workbook(report_path, data_only=True)
    wbf = load_workbook(report_path, data_only=False)
    tab = _find_tab(wbv, _ACL_TAB_HINTS)
    if tab is None:
        raise ValueError(
            f"No ACL Env tab in {Path(report_path).name} "
            f"(sheets: {wbv.sheetnames})"
        )
    wsv, wsf = wbv[tab], wbf[tab]

    cu = _cell_str(wsv.cell(1, 1).value)
    heading = [_cell_str(wsv.cell(r, 1).value)
               for r in (2, 3) if _cell_str(wsv.cell(r, 1).value)]

    hdr = _find_label(wsv, "Current Grade")
    hr = hdr[0] if hdr else 5
    col_headers = [
        _cell_str(wsv.cell(hr, c).value).replace("\n", " ").strip()
        for c in range(1, 12) if _cell_str(wsv.cell(hr, c).value)
    ]

    def _row_vals(r: int) -> dict:
        g = lambda c: _num(wsv.cell(r, c).value)  # noqa: E731
        return dict(
            balance=g(2), specific_id=g(3), llc_balance=g(4),
            base_loss_rate=g(5), mgmt_adj=g(6), allowance_factor=g(7),
            allowance_before_env=g(8), env_factor=g(9),
            env_allowance=g(10), total_allowance=g(11),
        )

    pool_rows: list[AclPoolRow] = []
    pooled_totals: AclPoolRow | None = None
    impaired_rows: list[AdjustmentRow] = []
    adjustment_rows: list[AdjustmentRow] = []
    current_pool: str | None = None
    section = "pools"

    for r in range(hr + 1, wsv.max_row + 1):
        a = _cell_str(wsv.cell(r, 1).value)
        if not a:
            continue
        if a == "Pooled Totals":
            pooled_totals = AclPoolRow(pool=a, is_total=True, **_row_vals(r))
            section = "await_impaired"
        elif a == "Impaired Loans":
            section = "impaired"
        elif section == "pools":
            if a == "Total" and current_pool:
                pool_rows.append(AclPoolRow(pool=current_pool, **_row_vals(r)))
            else:
                current_pool = a
        elif section == "impaired":
            val = _num(wsv.cell(r, 11).value)  # column K
            bold = _bold(wsf, r, 1)
            row = AdjustmentRow(label=a, value=val, bold=bold)
            if any(a.startswith(lbl) for lbl in _ACL_ADJ_LABELS):
                adjustment_rows.append(row)
            else:
                impaired_rows.append(row)

    return AclEnvPage(
        credit_union=cu, heading_lines=heading, col_headers=col_headers,
        pool_rows=pool_rows, pooled_totals=pooled_totals,
        impaired_rows=impaired_rows, adjustment_rows=adjustment_rows,
    )


