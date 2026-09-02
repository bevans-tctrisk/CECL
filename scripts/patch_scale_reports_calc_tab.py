"""Patch the Calc tab running-total fill error in already-generated
SCALE report workbooks (the same defect fixed in the templates).

For each report:
  1. Locate the worksheet part that holds the buggy Calc tab formula.
  2. Replace the four AC75/AD75/AC76/AD76 formulas with the standard
     running-total pattern used by every other cell in the block.
  3. Set ``fullCalcOnLoad`` so Excel recomputes all dependents (incl.
     the Vizo report tabs) on open.
  4. Back up the original once (``.bak_calcfix``) and rewrite the zip,
     preserving every other part byte-for-byte.

Files that don't contain the buggy formula (already correct / different
layout) are skipped and reported. Recalc of on-disk cached values is
handled separately by the caller via excel_recalc.recalc_and_save.
"""
from __future__ import annotations

import re
import shutil
import sys
import zipfile
from pathlib import Path

REPLACEMENTS = {
    "AC58+SUM(AD75:$AE75)": "AC58+AD75",
    "AD58+SUM(AE75:$AE75)": "AD58+AE75",
    "AC59+SUM(AD76:$AE76)": "AC59+AD76",
    "AD59+SUM(AE76:$AE76)": "AD59+AE76",
}
_MARKER = "AC58+SUM(AD75:$AE75)"


def _calc_tab_part(zf: zipfile.ZipFile) -> str | None:
    for name in zf.namelist():
        if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
            if _MARKER in zf.read(name).decode("utf8"):
                return name
    return None


def _set_full_calc_on_load(wbxml: str) -> str:
    if "<calcPr" not in wbxml:
        return wbxml
    m = re.search(r"<calcPr\b[^>]*/>", wbxml)
    if not m:
        return wbxml
    tag = m.group(0)
    if "fullCalcOnLoad" in tag:
        new = re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', tag)
    else:
        new = tag[:-2] + ' fullCalcOnLoad="1"/>'
    return wbxml.replace(tag, new, 1)


def patch_report(path: Path) -> str:
    with zipfile.ZipFile(path, "r") as zf:
        part = _calc_tab_part(zf)
        if part is None:
            return "skip-no-bug"
        sheet_xml = zf.read(part).decode("utf8")
        wb_xml = zf.read("xl/workbook.xml").decode("utf8")

    for old, new in REPLACEMENTS.items():
        n = sheet_xml.count(old)
        if n != 1:
            return f"skip-unexpected({old}={n})"
        sheet_xml = sheet_xml.replace(old, new)
    wb_xml = _set_full_calc_on_load(wb_xml)

    backup = path.with_suffix(path.suffix + ".bak_calcfix")
    if not backup.exists():
        shutil.copy2(path, backup)

    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(
        tmp, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            if item.filename == part:
                zout.writestr(item, sheet_xml.encode("utf8"))
            elif item.filename == "xl/workbook.xml":
                zout.writestr(item, wb_xml.encode("utf8"))
            else:
                zout.writestr(item, zin.read(item.filename))
    tmp.replace(path)
    return "fixed"


def main(argv: list[str]) -> None:
    if len(argv) < 2:
        raise SystemExit("usage: patch_scale_reports_calc_tab.py <file> [<file> ...]")
    results: dict[str, list[str]] = {}
    for arg in argv[1:]:
        p = Path(arg)
        try:
            status = patch_report(p)
        except Exception as exc:  # noqa: BLE001
            status = f"error:{exc}"
        results.setdefault(status, []).append(p.name)
        print(f"[{status}] {p.name}")
    print("\n--- SUMMARY ---")
    for status, names in sorted(results.items()):
        print(f"{status}: {len(names)}")


if __name__ == "__main__":
    main(sys.argv)
