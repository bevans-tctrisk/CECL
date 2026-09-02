"""Report data model — the contract between compute and rendering.

Every page the browser renderer knows how to draw has a small, typed
node here. The compute engine (or, during migration, an adapter that
reads the already-generated workbook) populates these; the Jinja2
templates and the xlsx "extract page" writers both consume them.

Keep these dataclasses render-agnostic: no HTML, no openpyxl, no
formatting decisions. Formatting (currency, %, colors) lives in the
templates / CSS and in :mod:`cecl_report_web.format`.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class ChartSpec:
    """A chart described as DATA, so a page carries its charts instead of
    pulling them back out of a workbook.

    ``kind`` is semantic; the adapter in :mod:`cecl_report_web.from_data`
    maps it onto the normalized dict :func:`cecl_report_web.charts.render_chart_svg`
    consumes. ``series`` entries are ``{"name": str, "values": [float],
    "colors": [hex]}`` -- ``colors`` are per-slice for pie/doughnut and
    per-series for bar/column.
    """

    kind: str  # 'pie' | 'doughnut' | 'bar_h' | 'column' | 'diverging_bar'
    title: str | None = None
    categories: list[str] = field(default_factory=list)
    series: list[dict] = field(default_factory=list)
    value_format: str | None = None  # 'pct' | 'currency' | None


@dataclass
class CoverPage:
    """The report cover (mirrors ``sheet_cover_vizo``)."""

    credit_union: str
    period_ending: str  # display string, e.g. "2026-06-30"
    title: str = "CECL Credit Migration Report"
    subtitle: str | None = None  # e.g. "Supplemental Reports"
    firm: str = "TCT Risk Solutions"
    confidentiality: str = "All reports are confidential."
    date_text: str | None = None       # date as shown in the workbook (e.g. "6/30/2026")
    paragraph: str | None = None        # full confidentiality/disclaimer paragraph
    footer: str | None = None           # e.g. "© 2026 TCT Risk Solutions"
    top_logo: str | None = None         # data: URI (Vizo logo)
    bottom_logo: str | None = None      # data: URI (TCT logo)


@dataclass
class KeyValueRow:
    """One labelled currency line (e.g. a CECL Adjustment box row)."""

    label: str
    value: float | None


@dataclass
class PoolMigrationRow:
    """Improved/Deteriorated by loan type/pool (fractions, e.g. 0.12)."""

    pool: str
    improved: float | None
    deteriorated: float | None
    net_change: float | None


@dataclass
class GradeMigrationRow:
    """Improved/Deteriorated by credit grade."""

    grade: str
    balance: float | None
    improved: float | None
    deteriorated: float | None


@dataclass
class ImprDeterPage:
    """The "Impr Deter" tab — CECL adjustment box + migration tables.

    Mirrors the numeric Executive Summary the CU cares about: the
    headline allowance figures plus improved/deteriorated migration by
    pool and by grade.
    """

    credit_union: str
    period_ending: str  # display "MM-DD-YY" style handled in template
    heading_lines: list[str] = field(default_factory=list)
    cecl_adjustment: list[KeyValueRow] = field(default_factory=list)
    by_pool: list[PoolMigrationRow] = field(default_factory=list)
    by_grade: list[GradeMigrationRow] = field(default_factory=list)


@dataclass
class MatrixCell:
    """One cell of a risk-change matrix.

    ``state`` drives the background color and is read straight from the
    workbook fill so the rendered grid matches Excel exactly (including
    the template's diagonal-band quirks):
      'plain' | 'improved' (green) | 'deteriorated' (maroon) | 'header'
    """

    value: float | None
    state: str = "plain"
    is_pct: bool = False
    bold: bool = False


@dataclass
class MatrixRow:
    label: str
    cells: list[MatrixCell] = field(default_factory=list)   # original-grade cols
    total: MatrixCell | None = None
    side: list[MatrixCell] = field(default_factory=list)     # Det/Imp/Unch
    range_label: str = ""                                     # score range (e.g. "730+")


@dataclass
class RiskChangeMatrix:
    """A migration matrix (dollar or percent)."""

    corner: str                                    # "$ Current Grade" / "% Current Grade"
    col_headers: list[str] = field(default_factory=list)
    rows: list[MatrixRow] = field(default_factory=list)
    side_headers: list[str] = field(default_factory=list)  # Deteriorated/Improved/Unchanged
    is_pct: bool = False


@dataclass
class RiskChangePage:
    """The "Risk Change Total" tab — dollar + percent migration matrices."""

    credit_union: str
    heading_lines: list[str] = field(default_factory=list)
    matrices: list[RiskChangeMatrix] = field(default_factory=list)
    summary: list[KeyValueRow] = field(default_factory=list)  # Balance Adj / Total in Portfolio


@dataclass
class AclPoolRow:
    """One row on the ACL Env by Pool tab.

    ``kind`` distinguishes a pool ``header`` (name only, no values), a
    per-grade ``grade`` row, and a ``total`` row — so grade-rated CUs
    render their full breakdown, not just single-line pool totals.
    """

    pool: str
    kind: str = "total"  # "header" | "grade" | "total"
    balance: float | None = None
    specific_id: float | None = None
    llc_balance: float | None = None
    base_loss_rate: float | None = None
    mgmt_adj: float | None = None
    allowance_factor: float | None = None
    allowance_before_env: float | None = None
    env_factor: float | None = None
    env_allowance: float | None = None
    total_allowance: float | None = None
    is_total: bool = False  # the "Pooled Totals" summary row


@dataclass
class AdjustmentRow:
    """A labelled currency line in the impaired/adjustment section."""

    label: str
    value: float | None
    bold: bool = False


@dataclass
class AclEnvPage:
    """The "ACL Env by Pool Mgmt Adj" tab — the core allowance calc."""

    credit_union: str
    heading_lines: list[str] = field(default_factory=list)
    col_headers: list[str] = field(default_factory=list)
    pool_rows: list[AclPoolRow] = field(default_factory=list)
    pooled_totals: AclPoolRow | None = None
    impaired_rows: list[AdjustmentRow] = field(default_factory=list)
    adjustment_rows: list[AdjustmentRow] = field(default_factory=list)


@dataclass
class GridCell:
    """One faithfully-rendered cell for the generic grid renderer.

    ``text`` is already formatted (via :func:`format.excel_format`);
    styling is copied straight from the workbook so any labelled-grid tab
    reproduces without a bespoke model/template.
    """

    text: str = ""
    align: str = "left"      # left | center | right
    bold: bool = False
    italic: bool = False
    fill: str | None = None  # 6-hex background
    color: str | None = None  # 6-hex text color
    size: float | None = None
    colspan: int = 1
    rowspan: int = 1
    wrap: bool = False       # wrap long text / honor embedded newlines


@dataclass
class GridPage:
    """A whole tab rendered generically as a styled grid of cells."""

    credit_union: str
    sheet_name: str
    rows: list[list[GridCell]] = field(default_factory=list)
    landscape: bool = False


@dataclass
class ReportModel:
    """Top-level model for one report (Vizo Model or Supplemental).

    Pages are added incrementally as the migration progresses. Only the
    cover is wired in the first slice; later phases attach exec summary,
    risk-change matrices, ACL/env tables, historical data, and charts.
    """

    credit_union: str
    period_ending: str
    report_type: str = "vizo"  # "vizo" | "vizo_supp"
    cover: CoverPage | None = None
    # Forward-looking: additional page nodes keyed by page id. Kept as a
    # generic bag so we can add pages without churning this class while
    # the schema stabilizes.
    pages: dict[str, Any] = field(default_factory=dict)


@dataclass
class TableCell:
    """One cell of a generic table page. ``fmt`` drives display formatting."""

    value: Any = None
    fmt: str = "text"  # text | currency | currency2 | pct | pct1 | pct2
    bold: bool = False
    align: str = "right"  # left | right | center


@dataclass
class TableSection:
    """A titled block of a TablePage: optional column headers + data rows."""

    columns: list[str] = field(default_factory=list)  # "" = blank header cell
    rows: list[list[TableCell]] = field(default_factory=list)
    title: str | None = None


@dataclass
class TablePage:
    """A generic titled, sectioned table page.

    Backs the workbook's table-only tabs (ACL Summary, Mgmt Adj Summary,
    Impaired Loans, ...) with real data instead of the cell-copying grid.
    """

    credit_union: str
    title: str
    heading_lines: list[str] = field(default_factory=list)
    sections: list[TableSection] = field(default_factory=list)
    notes_title: str = ""
    notes: list[str] = field(default_factory=list)


@dataclass
class NarrativeSection:
    """A heading + prose paragraph on a narrative page."""

    heading: str
    body: str


@dataclass
class NarrativePage:
    """A text-only appendix page (Introduction, Executive Summary, ...)."""

    credit_union: str
    title: str
    sections: list[NarrativeSection] = field(default_factory=list)


@dataclass
class SummaryVarianceBlock:
    """One Current / Prior / Change block on the Summary Variance card."""

    label: str
    period: str
    measures: list = field(default_factory=list)  # [(label, formatted value)]


@dataclass
class SummaryVariancePage:
    """SCALE 'Executive Summary-Vizo' banded card: a centred 3-line title over
    Current / Prior / Change blocks of the four ACL measures."""

    credit_union: str
    quarter_ended: str
    blocks: list = field(default_factory=list)  # [SummaryVarianceBlock]
    note: str = ""


