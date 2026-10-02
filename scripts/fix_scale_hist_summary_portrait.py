"""Switch the SCALE "Historical Summary" tab from landscape to portrait.

The tab is only 5 columns wide (A:E) by ~52 rows, so landscape wasted the
page and shrank the fit-to-one-page scale factor. Portrait matches the
other narrow Vizo tabs.

Orientation rides along from the template: ``vizo_layout._deep_copy_sheet``
copies ``page_setup.orientation`` when it replaces the TEMPLATE_SOURCED
sheets, so fixing the templates fixes every newly generated report. Older
carried workbooks are covered by the portrait enforcement in
``vizo_layout.normalize_vizo_layout`` (step 6b).

Applied to EVERY SCALE template; specific generated output workbooks can
also be passed as CLI args.
"""
import glob
import os
import shutil
import sys
from datetime import datetime

import openpyxl

TAB = "Historical Summary"
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_GLOB = os.path.join(
    HERE, "cecl_ui", "data", "scale", "templates", "*_CECL_SCALE_template.xlsx"
)


def patch(path):
    wb = openpyxl.load_workbook(path)
    if TAB not in wb.sheetnames:
        print("  skip (no %s tab): %s" % (TAB, os.path.basename(path)))
        return False
    ws = wb[TAB]
    if ws.page_setup.orientation == "portrait":
        print("  already portrait: %s" % os.path.basename(path))
        return False

    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
    shutil.copy2(path, "%s.bak.histsummportrait_%s" % (path, stamp))

    ws.page_setup.orientation = "portrait"
    # Keep the existing fit-to-one-page behavior intact.
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.scale = None
    wb.save(path)
    print("  landscape -> portrait: %s" % os.path.basename(path))
    return True


def main():
    targets = sys.argv[1:] or sorted(glob.glob(TEMPLATE_GLOB))
    if not targets:
        print("No workbooks found.")
        return 1
    changed = 0
    print("Patching %d workbook(s):" % len(targets))
    for t in targets:
        if patch(t):
            changed += 1
    print("Done. %d changed." % changed)
    return 0


if __name__ == "__main__":
    sys.exit(main())
