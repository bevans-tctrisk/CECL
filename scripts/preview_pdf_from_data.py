"""Render a report page to PDF from DATA (no workbook), for the PDF migration.

Proves the compute-time path end to end: config -> cecl_report_web.from_data
-> Jinja template -> headless Chromium -> PDF, with no generated .xlsx read.

    set CECL_WORKSPACE_ROOT=Z:\\Shared\\TCT Files\\CECL - CM Files
    python scripts/preview_pdf_from_data.py <config> [<snapshot YYYY-MM-DD>] [<out.pdf>]
"""

from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault("CECL_WORKSPACE_ROOT",
                      r"Z:\Shared\TCT Files\CECL - CM Files")

import generate_report as gr  # noqa: E402
from cecl_report_web import from_data, render as R  # noqa: E402


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print(__doc__)
        return 2
    name = argv[1]
    snapshot = argv[2] if len(argv) > 2 else "2026-06-30"
    out = argv[3] if len(argv) > 3 else f"_from_data_{name}_{snapshot}.pdf"

    config = gr.load_config(name)
    cover = from_data.build_cover(name, snapshot, config)
    print(f"cover: CU={cover.credit_union!r} date={cover.date_text!r} "
          f"logos: top={'yes' if cover.top_logo else 'no'} "
          f"bottom={'yes' if cover.bottom_logo else 'no'}")

    html = R.render_html("cover.html", cover=cover)
    pdf = R.render_pdf(html, landscape=False)
    with open(out, "wb") as fh:
        fh.write(pdf)
    ok = pdf[:5] == b"%PDF-"
    print(f"wrote {out} ({len(pdf):,} bytes) valid_pdf={ok}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
