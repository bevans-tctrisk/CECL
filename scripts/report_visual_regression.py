"""Visual-regression runner for the browser report renderer.

Fixtures cover the shapes that matter: single-line (WNC), grade-rated with
real migration (Jackson River), and many pools (Tongass). Reference images
live in the workspace (Reports/_web_refs) — not committed, since they are
large binaries derived from the fixture workbooks.

    # capture/refresh reference images (after an intentional visual change)
    python scripts/report_visual_regression.py --update

    # compare current renders to references (CI / pre-commit check)
    python scripts/report_visual_regression.py

Exit code is non-zero if any page regressed beyond the pixel threshold.
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(_REPO))
os.environ.setdefault(
    "CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")

from cecl_report_web import regression  # noqa: E402

WS = os.environ["CECL_WORKSPACE_ROOT"]
_REPORTS = Path(WS) / "Reports"

FIXTURES = [
    {
        "label": "WNC single-line",
        "cu": "WNC Community CU",
        "snap": "2026-06-30",
        "report_path": _REPORTS
        / "2026-06-30_CECL_Migration_WNC_Community_CU_Vizo_Model.xlsx",
    },
    {
        "label": "Jackson River grade-rated migration",
        "cu": "Jackson River Community CU",
        "snap": "2026-06-30",
        "report_path": _REPORTS
        / "2026-06-30_CECL_Migration_Jackson_River_Community_CU_Vizo_Model.xlsx",
    },
    {
        "label": "Tongass many pools",
        "cu": "Tongass FCU",
        "snap": "2025-12-31",
        "report_path": _REPORTS
        / "2025-12-31_CECL_Migration_Tongass_FCU_Vizo_Model.xlsx",
    },
]


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--update", action="store_true",
                    help="Capture/refresh reference images instead of comparing.")
    ap.add_argument("--refs-dir", default=str(_REPORTS / "_web_refs"))
    args = ap.parse_args()

    summary = regression.run(FIXTURES, args.refs_dir, update=args.update)

    print(f"\n{'=' * 66}")
    print(f"Visual regression — {'UPDATE' if args.update else 'COMPARE'} "
          f"(refs: {summary['refs_dir']})")
    print("=" * 66)
    captured = changed = okc = 0
    for r in summary["results"]:
        st = r["status"]
        if st == "captured":
            captured += 1
        elif st == "ok":
            okc += 1
        elif st == "CHANGED":
            changed += 1
            print(f"  CHANGED  {r['fixture']} / {r['sheet']}  "
                  f"({r['changed_fraction']*100:.3f}% px)  -> {r.get('diff','')}")
        else:
            print(f"  {st.upper()}  {r['fixture']}  {r.get('report','')}")
    print("-" * 66)
    print(f"  captured={captured}  ok={okc}  CHANGED={changed}")
    print(f"  RESULT: {'PASS' if summary['ok'] else 'FAIL'}")
    return 0 if summary["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
