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
class CoverPage:
    """The report cover (mirrors ``sheet_cover_vizo``)."""

    credit_union: str
    period_ending: str  # display string, e.g. "2026-06-30"
    title: str = "CECL Credit Migration Report"
    subtitle: str | None = None  # e.g. "Supplemental Reports"
    firm: str = "TCT Risk Solutions"
    confidentiality: str = "All reports are confidential."


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
    """One pool's Total row on the ACL Env by Pool tab."""

    pool: str
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
