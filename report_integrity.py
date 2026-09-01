"""Structural validation for generated report workbooks.

Why this exists
---------------
On 2026-08-31 a namespace bug in ``report_vizo.patch_impdet_charts`` wrote a
drawing whose chart reference was ``<chart r:id="rId1"/>`` instead of
``<c:chart r:id="rId1"/>``. Unprefixed, it bound to the spreadsheetDrawing
namespace instead of the chart namespace, Excel could not resolve the chart,
and it refused to open the workbook at all.

Nothing caught it. ``openpyxl`` loaded the file happily, the zip was intact and
every XML part was well-formed. The corruption was only visible to Excel — so a
client could have received an unopenable report.

This module is the gate. ``validate_workbook`` runs a set of pure-Python
structural checks (no Excel, fast enough to run on every generated report and
in CI). ``excel_can_open`` is an optional stronger check for environments that
have Excel and pywin32.

Standalone use:

    python report_integrity.py <file.xlsx> [more.xlsx ...]

exits non-zero if any file fails.
"""
from __future__ import annotations

import os
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_R_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _rels_path_for(part: str) -> str:
    d, name = posixpath.split(part)
    return posixpath.join(d, "_rels", name + ".rels")


def _resolve(part: str, target: str) -> str:
    """Resolve a relationship Target against the part that declares it."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def _read_rels(z: zipfile.ZipFile, part: str) -> dict:
    rels_path = _rels_path_for(part)
    if rels_path not in z.namelist():
        return {}
    try:
        root = ET.fromstring(z.read(rels_path))
    except ET.ParseError:
        return {}
    out = {}
    for rel in root:
        rid, target = rel.get("Id"), rel.get("Target")
        if not rid or not target:
            continue
        if (rel.get("TargetMode") or "").lower() == "external":
            continue
        out[rid] = _resolve(part, target)
    return out


def _check_xml_well_formed(z, errors):
    for name in z.namelist():
        if not (name.endswith(".xml") or name.endswith(".rels")):
            continue
        try:
            ET.fromstring(z.read(name))
        except ET.ParseError as exc:
            errors.append(f"malformed XML in {name}: {exc}")


def _check_rels_targets_exist(z, errors):
    names = set(z.namelist())
    for name in list(names):
        if not name.endswith(".rels"):
            continue
        owner = name.replace("/_rels/", "/").removesuffix(".rels")
        for rid, target in _read_rels(z, owner).items():
            if target not in names:
                errors.append(
                    f"{name}: {rid} points at missing part {target!r}")


def _check_drawing_chart_refs(z, errors):
    """Every chart reference in a drawing must be a c:chart in the chart
    namespace, and its rId must resolve to a chart part that exists.

    This is the check that would have caught the 2026-08-31 corruption.
    """
    for name in z.namelist():
        if not re.fullmatch(r"xl/drawings/drawing\d+\.xml", name):
            continue
        try:
            root = ET.fromstring(z.read(name))
        except ET.ParseError:
            continue  # reported by the well-formed check
        rels = _read_rels(z, name)
        parts = set(z.namelist())

        for gd in root.iter(f"{{{A_NS}}}graphicData"):
            if (gd.get("uri") or "") != C_NS:
                continue
            children = list(gd)
            if not children:
                errors.append(f"{name}: chart graphicData has no child element")
                continue
            for child in children:
                tag = child.tag
                if tag != f"{{{C_NS}}}chart":
                    bad_ns = tag.split("}")[0].lstrip("{") if "}" in tag else "(none)"
                    errors.append(
                        f"{name}: chart reference is <{tag}> - expected "
                        f"{{{C_NS}}}chart. It is bound to namespace {bad_ns!r}, "
                        f"so Excel cannot resolve the chart and will refuse to "
                        f"open the workbook.")
                    continue
                rid = child.get(f"{{{R_NS}}}id")
                if not rid:
                    errors.append(f"{name}: c:chart has no r:id")
                elif rid not in rels:
                    errors.append(
                        f"{name}: c:chart r:id={rid!r} not found in "
                        f"{_rels_path_for(name)}")
                elif rels[rid] not in parts:
                    errors.append(
                        f"{name}: c:chart r:id={rid!r} -> missing part "
                        f"{rels[rid]!r}")


def _check_sheet_count(z, errors):
    try:
        root = ET.fromstring(z.read("xl/workbook.xml"))
    except (KeyError, ET.ParseError) as exc:
        errors.append(f"cannot read xl/workbook.xml: {exc}")
        return
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    sheets = root.findall(".//s:sheet", ns)
    if not sheets:
        errors.append("workbook declares no sheets")


MIGRATION_STATES = ("Improved", "Deteriorated", "Unchanged")


def _check_dq_co_blocks(z, warnings):
    """Warn when a Risk Change tab's delinquency / charge-off block is all zero.

    These blocks feed the DQ pie and the charge-off bar. When the upstream
    source is missing they are written as literal zeros and the charts render
    as a titled, legended, empty frame -- which reads as broken rather than as
    "no data". 19 client-facing workbooks shipped that way before anyone
    noticed (docs/pdf_migration/04_blank_charts.md).

    A warning, not an error: unlike malformed XML this is "almost certainly
    wrong" rather than "definitely broken", and the caller should not refuse to
    deliver a report over it. Column positions differ between credit unions, so
    the blocks are located by their row labels rather than by coordinates.
    """
    try:
        import openpyxl  # noqa: PLC0415
    except ImportError:
        return
    try:
        wb = openpyxl.load_workbook(z.filename, data_only=True, read_only=True)
    except Exception:  # noqa: BLE001
        return

    empty_blocks = 0
    sheets_hit = set()
    try:
        for ws in wb.worksheets:
            if not ws.title.startswith("Risk Ch"):
                continue
            rows = list(ws.iter_rows())
            for r_idx, row in enumerate(rows):
                for cell in row:
                    if cell.value != "Improved" or getattr(
                            cell, "column", None) is None:
                        continue
                    # Confirm this is a migration-status block, not a stray word.
                    below = []
                    for off in (1, 2):
                        if r_idx + off < len(rows):
                            nxt = rows[r_idx + off]
                            # read_only mode yields EmptyCell, which has no
                            # .column -- getattr keeps the scan cheap.
                            hit = [c.value for c in nxt
                                   if getattr(c, "column", None) == cell.column]
                            below.append(hit[0] if hit else None)
                    if below[:2] != list(MIGRATION_STATES[1:]):
                        continue
                    # Collect the numbers to the right of the four label rows.
                    nums = []
                    for off in range(4):
                        if r_idx + off >= len(rows):
                            break
                        for c in rows[r_idx + off]:
                            col = getattr(c, "column", None)
                            if col is not None and col > cell.column                                     and isinstance(c.value, (int, float)):
                                nums.append(c.value)
                    if nums and not any(nums):
                        empty_blocks += 1
                        sheets_hit.add(ws.title)
    finally:
        wb.close()

    if empty_blocks:
        warnings.append(
            f"{empty_blocks} delinquency/charge-off block(s) across "
            f"{len(sheets_hit)} Risk Change tab(s) are entirely zero, so "
            f"{empty_blocks} chart(s) will render as empty frames. Usually "
            f"means the upstream source is missing for this credit union - "
            f"see docs/pdf_migration/04_blank_charts.md")


CHECKS = (
    _check_xml_well_formed,
    _check_rels_targets_exist,
    _check_drawing_chart_refs,
    _check_sheet_count,
)


def validate_workbook(path) -> dict:
    """Structurally validate a generated .xlsx. No Excel required.

    Returns ``{'ok': bool, 'errors': [str], 'path': str}``. Never raises for a
    bad workbook — a caller decides how loud to be.
    """
    path = str(path)
    errors: list[str] = []
    warnings: list[str] = []
    try:
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
            if bad is not None:
                errors.append(f"corrupt zip entry: {bad}")
            else:
                for check in CHECKS:
                    check(z, errors)
                _check_dq_co_blocks(z, warnings)
    except (zipfile.BadZipFile, OSError) as exc:
        errors.append(f"cannot open as xlsx: {exc}")
    return {"ok": not errors, "errors": errors, "warnings": warnings,
            "path": path}


def excel_can_open(path) -> dict:
    """Stronger check: ask Excel itself to open the file, read-only.

    Best-effort, matching cecl_ui/services/scale/excel_recalc.py: when pywin32
    or Excel is unavailable this reports ``checked=False`` rather than failing,
    so it never blocks a machine that simply has no Excel.
    """
    result = {"ok": None, "checked": False, "error": "", "path": str(path)}
    try:
        import win32com.client as win32  # noqa: PLC0415
    except ImportError:
        result["error"] = "pywin32 not installed; Excel open-check skipped"
        return result

    excel = wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        wb = excel.Workbooks.Open(os.path.abspath(str(path)),
                                  UpdateLinks=0, ReadOnly=True)
        result.update(ok=True, checked=True)
    except Exception as exc:  # noqa: BLE001
        result.update(ok=False, checked=True, error=str(exc)[:300])
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:  # noqa: BLE001
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:  # noqa: BLE001
            pass
    return result


def check_and_report(path, label="", use_excel=False) -> bool:
    """Validate and print a one-line verdict. Returns True when sound."""
    name = label or os.path.basename(str(path))
    res = validate_workbook(path)
    if not res["ok"]:
        print(f"  INTEGRITY FAILURE in {name}:")
        for err in res["errors"][:10]:
            print(f"    - {err}")
        extra = len(res["errors"]) - 10
        if extra > 0:
            print(f"    ... and {extra} more")
        return False

    if use_excel:
        xl = excel_can_open(path)
        if xl["checked"] and not xl["ok"]:
            print(f"  INTEGRITY FAILURE in {name}: Excel refused to open it "
                  f"({xl['error']})")
            return False
    for warn in res.get("warnings", []):
        print(f"  Integrity WARNING in {name}: {warn}")
    print(f"  Integrity OK: {name}")
    return True


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    use_excel = "--excel" in argv
    files = [a for a in argv[1:] if not a.startswith("--")]
    ok = True
    for f in files:
        ok &= check_and_report(f, use_excel=use_excel)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
