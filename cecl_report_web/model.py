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
