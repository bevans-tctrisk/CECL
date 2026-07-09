"""Phase 1 POC: render the WNC report Cover page to PDF via the browser
pipeline (ReportModel -> Jinja2 HTML -> headless Chromium PDF).

This proves the full toolchain end-to-end with pixel control before the
larger model-extraction work. Run from the repo root:

    python scripts/report_web_poc.py

Writes both the HTML (for eyeballing in a browser) and the PDF next to
the workspace Reports folder.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(_REPO))
os.environ.setdefault(
    "CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files"
)

from cecl_report_web.model import CoverPage  # noqa: E402
from cecl_report_web import render  # noqa: E402


def main() -> int:
    cover = CoverPage(
        credit_union="WNC Community CU",
        period_ending="2026-06-30",
    )

    out_dir = Path(os.environ["CECL_WORKSPACE_ROOT"]) / "Reports" / "_web_poc"
    out_dir.mkdir(parents=True, exist_ok=True)

    html = render.render_html("cover.html", cover=cover)
    html_path = out_dir / "WNC_Cover.html"
    html_path.write_text(html, encoding="utf-8")
    print(f"HTML written: {html_path}")

    pdf = render.render_pdf(html)
    pdf_path = out_dir / "WNC_Cover.pdf"
    pdf_path.write_bytes(pdf)
    print(f"PDF written : {pdf_path}  ({len(pdf):,} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
