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

import gc
import subprocess
import time
from pathlib import Path
from typing import Any


def _excel_pids() -> set[int]:
    """Return the set of currently-running EXCEL.EXE PIDs.

    Uses ``tasklist`` (always available on Windows, no extra deps) so we
    can diff before/after ``DispatchEx`` to identify the PID we spawned
    and force-kill it during cleanup if Excel's normal Quit() hangs.
    """
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=10, check=False,
        ).stdout
    except Exception:  # noqa: BLE001
        return set()
    pids: set[int] = set()
    for line in out.splitlines():
        line = line.strip()
        if not line or line.startswith('"INFO:'):
            continue
        parts = [p.strip().strip('"') for p in line.split('","')]
        # CSV: "EXCEL.EXE","<pid>","Console","<sess>","<mem>"
        if len(parts) >= 2:
            try:
                pids.add(int(parts[1].strip('"')))
            except ValueError:
                continue
    return pids


def _taskkill(pid: int) -> None:
    """Best-effort force-kill of an EXCEL.EXE PID (with child tree).

    Used as a last-resort safety net after Quit() to guarantee the COM
    server releases its file lock -- otherwise the user can't reopen
    the freshly-generated workbook ("locked for editing by 'Brian
    Evans'").
    """
    try:
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(pid)],
            capture_output=True, text=True, timeout=10, check=False,
        )
    except Exception:  # noqa: BLE001
        pass


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
    spawned_pid: int | None = None
    pythoncom.CoInitialize()
    try:
        # Snapshot existing EXCEL.EXE PIDs so we can identify (and, if
        # necessary, force-kill) the one we're about to spawn.
        before_pids = _excel_pids()

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

        new_pids = _excel_pids() - before_pids
        if len(new_pids) == 1:
            spawned_pid = next(iter(new_pids))

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
        # Order matters: close every open workbook (not just the one we
        # opened -- belt and braces), drop COM refs so pythoncom can
        # release the proxy, run gc to actually collect them, then Quit.
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:  # noqa: BLE001
            pass
        try:
            if excel is not None:
                # Close any workbooks that might still be open in this
                # hidden instance (defensive against partial failures).
                try:
                    while excel.Workbooks.Count > 0:
                        excel.Workbooks(1).Close(SaveChanges=False)
                except Exception:  # noqa: BLE001
                    pass
                excel.Quit()
        except Exception:  # noqa: BLE001
            pass

        # CRITICAL: drop strong refs *before* CoUninitialize, then force
        # a gc pass so the COM proxy is actually released. Without this
        # the EXCEL.EXE we spawned can outlive this function and keep an
        # exclusive lock on the workbook, producing
        # "locked for editing by 'Brian Evans'" the next time the user
        # double-clicks the file.
        wb = None
        excel = None
        try:
            gc.collect()
        except Exception:  # noqa: BLE001
            pass

        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass

        # Safety net: if our spawned EXCEL.EXE is still alive a moment
        # after Quit(), force-kill it. This handles the case where Excel
        # hangs on a recalculation chain, a stuck add-in load, or a
        # license/activation prompt.
        if spawned_pid is not None:
            for _ in range(10):
                if spawned_pid not in _excel_pids():
                    break
                time.sleep(0.2)
            if spawned_pid in _excel_pids():
                _taskkill(spawned_pid)
                if not result["error"]:
                    result["error"] = (
                        f"Excel PID {spawned_pid} did not exit after Quit(); "
                        "force-killed."
                    )
    return result
