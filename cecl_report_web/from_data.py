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

from .model import CoverPage


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


def build_report_model(client_name: str, snapshot_date: str, config: dict,
                       grades: Any = None, hist: dict | None = None,
                       df: Any = None, *, supplemental: bool = False) -> dict:
    """Build the render-ready page set from report data.

    Returns ``{"cover": CoverPage, "pages": [ (template, ctx, landscape), ... ]}``.
    Only the cover is populated today; further archetypes (Risk Change, ACL
    Env, Impr Deter) are added here incrementally, each reusing the report
    engine's own pure compute functions.
    """
    cover = build_cover(client_name, snapshot_date, config,
                        supplemental=supplemental)
    pages: list[tuple[str, dict, bool]] = [
        ("cover.html", {"cover": cover}, False),
    ]
    return {"cover": cover, "pages": pages}
