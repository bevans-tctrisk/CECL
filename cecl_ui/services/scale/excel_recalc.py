"""Force Excel to evaluate every formula in a workbook and persist the
cached values to disk.

Why this exists
---------------
``openpyxl`` cannot evaluate formulas; it can only read the *cached*
result that Excel stamps into the file the last time the workbook was
opened-and-saved. Reports generated entirely by openpyxl therefore
have ``None`` for every formula cell when re-read with
``data_only=True`` -- which broke the carry-history "Prior ACL" flow
(``runs_service.read_prior_acl_values`` couldn't pull values out of an
un-opened prior workbook, so AY2..AY5 in the new run stayed as live
formulas pointing at the *current* quarter's Scale Calculation tab).

This module spawns a dedicated, invisible Excel.exe via COM,
recalculates the whole workbook, saves it, and quits. All subsequent
``openpyxl`` ``data_only=True`` reads then see real numbers.

Best-effort by design: if pywin32 isn't installed or Excel COM is
unavailable, recalc is skipped with a structured error and the caller
proceeds (workbook is still on disk, the user can open + save it
manually as a last resort).
"""
from __future__ import annotations

from pathlib import Path
from typing import Any


def recalc_and_save(workbook_path: str | Path) -> dict[str, Any]:
    """Open ``workbook_path`` in a headless Excel instance, recalc all
    formulas, save, and close.

    Returns ``{"ok": bool, "skipped": bool, "error": str, "path": str}``.
    ``skipped=True`` means pywin32/Excel isn't usable -- not a failure.
    """
    result: dict[str, Any] = {
        "ok": False, "skipped": False, "error": "",
        "path": str(workbook_path),
    }
    p = Path(workbook_path)
    if not p.exists():
        result["error"] = f"Workbook not found: {p}"
        return result

    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client as win32  # type: ignore[import-not-found]
    except ImportError as exc:
        result["skipped"] = True
        result["error"] = f"pywin32 not installed ({exc}); skipping recalc."
        return result

    excel = None
    wb = None
    pythoncom.CoInitialize()
    try:
        # DispatchEx forces a brand-new Excel.exe so we never clobber an
        # Excel session the analyst already has open with other files.
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        try:
            excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        except Exception:  # noqa: BLE001
            pass

        # UpdateLinks=0 -> do not refresh external links (avoid network
        # round-trips and credential prompts).
        wb = excel.Workbooks.Open(
            str(p.resolve()),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
        )
        try:
            excel.CalculateFullRebuild()
        except Exception:
            # Older Excel: fall back to a plain recalc.
            excel.Calculate()
        wb.Save()
        result["ok"] = True
    except Exception as exc:  # noqa: BLE001
        result["error"] = f"Excel recalc failed for {p.name}: {exc}"
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
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass
    return result
