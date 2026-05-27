"""End-to-end SCALE run.

``run_single_quarter(state, workspace_root)`` resolves the template and
mapping for the chosen period, fetches the Solr doc, writes the cells,
and saves an output workbook under
``Generated_Reports/<short>/<YYYY-MM>/CECL_SCALE_<short>.xlsx``.

``run_multi_quarter(state, workspace_root, quarters)`` repeats the fill
for ``quarters`` consecutive quarter-ends (newest -> oldest), writing
each quarter's data into the same workbook. After the first iteration
the output file is used as the template so writes accumulate (the
period-specific mapping CSVs point to different cells per quarter so
prior writes are preserved). Quarters with no mapping CSV are skipped
and reported.
"""
from __future__ import annotations

from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any
import shutil
import zipfile

import openpyxl

from . import (
    env_factor_writer, impaired_loader, mapping_loader, mgmt_adj_writer,
    qfactor_loader, runs_service, solr_fetcher, template_loader,
)


_THEME_ZIP_PATH = "xl/theme/theme1.xml"
_DEFAULT_OFFICE_THEME_CACHE: bytes | None = None


def _default_office_theme_bytes() -> bytes:
    """Return ``xl/theme/theme1.xml`` from a fresh openpyxl workbook.

    Fresh openpyxl workbooks ship with the standard Office theme (the
    one VS Excel uses for blank workbooks). Cached after first call.
    """
    global _DEFAULT_OFFICE_THEME_CACHE
    if _DEFAULT_OFFICE_THEME_CACHE is not None:
        return _DEFAULT_OFFICE_THEME_CACHE
    buf = BytesIO()
    wb = openpyxl.Workbook()
    wb.save(buf)
    buf.seek(0)
    with zipfile.ZipFile(buf, "r") as zf:
        _DEFAULT_OFFICE_THEME_CACHE = zf.read(_THEME_ZIP_PATH)
    return _DEFAULT_OFFICE_THEME_CACHE


def _read_workbook_theme(path: str | Path) -> bytes | None:
    """Return ``xl/theme/theme1.xml`` bytes from an xlsx file, or None."""
    try:
        with zipfile.ZipFile(path, "r") as zf:
            return zf.read(_THEME_ZIP_PATH)
    except (KeyError, zipfile.BadZipFile, FileNotFoundError):
        return None


def _replace_workbook_theme(path: str | Path, theme_bytes: bytes) -> None:
    """Rewrite ``xl/theme/theme1.xml`` inside an xlsx zip in-place."""
    src = Path(path)
    tmp = src.with_suffix(src.suffix + ".themetmp")
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(
        tmp, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = (
                theme_bytes if item.filename == _THEME_ZIP_PATH
                else zin.read(item.filename)
            )
            zout.writestr(item, data)
    tmp.replace(src)


VALID_Q_MONTHS = {3, 6, 9, 12}


def prev_quarter(ym: str) -> str:
    """``2025-12`` -> ``2025-09`` (matches multi_quarter.py CLI)."""
    y_str, m_str = ym.split("-")
    y, m = int(y_str), int(m_str)
    if m not in VALID_Q_MONTHS:
        raise ValueError("period month must be one of 03, 06, 09, 12")
    if m == 3:
        return f"{y - 1}-12"
    if m == 6:
        return f"{y}-03"
    if m == 9:
        return f"{y}-06"
    return f"{y}-09"


def make_periods(start_period: str, quarters: int) -> list[str]:
    if quarters < 1:
        raise ValueError("quarters must be >= 1")
    periods = [start_period]
    while len(periods) < quarters:
        periods.append(prev_quarter(periods[-1]))
    return periods


def _output_path(workspace_root: str, short: str, period: str) -> Path:
    safe_short = short or "cu"
    out_dir = Path(workspace_root) / "Generated_Reports" / safe_short / period
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir / f"CECL_SCALE_{safe_short}.xlsx"


def apply_qfactors(
    workbook_path: str | Path, entries: list[dict],
) -> dict:
    """Write Q-factor values (bps -> decimal) into ``workbook_path``.

    ``entries`` come from :func:`qfactor_loader.merge_with_overrides`
    (each has ``sheet``, ``cell``, ``effective_bps``).
    """
    result = {
        "applied": 0,
        "total": len(entries),
        "missing_sheets": [],
        "issues": [],
    }
    if not entries:
        return result
    wb = openpyxl.load_workbook(workbook_path)
    for row in entries:
        sheet = row["sheet"]
        cell = row["cell"]
        if sheet not in wb.sheetnames:
            result["missing_sheets"].append(f"{sheet} (for {cell})")
            continue
        try:
            wb[sheet][cell].value = float(row.get("effective_bps", 0.0)) / 10000.0
            result["applied"] += 1
        except Exception as exc:  # noqa: BLE001
            result["issues"].append(f"{sheet}!{cell}: {exc}")
    wb.save(workbook_path)
    return result


def _qfactor_entries_from_state(state: dict) -> list[dict]:
    scale = state.get("scale") or {}
    defaults = qfactor_loader.load_defaults()
    return qfactor_loader.merge_with_overrides(
        defaults, scale.get("qfactor_overrides") or {}
    )


def _impaired_rows_from_state(state: dict) -> list[dict]:
    scale = state.get("scale") or {}
    imp = scale.get("impaired_file") or {}
    parsed = imp.get("parsed") or {}
    if not parsed.get("ok"):
        return []
    return list(parsed.get("rows") or [])


def _normalize_mgmt_adj_state(
    mgmt_state: dict, template_path: str | Path,
) -> dict:
    """Re-align the wizard's name-keyed mgmt_adj rows to template order.

    Wizard state stores per-pool rows in
    ``mgmt_state["pool_rows"]`` as ``{pool_name: {hard_code_pct,
    use_default}}``. The writer needs an ordered list matching the
    template's pool rows. Pool names missing from saved state get a
    zero/off default. Returns the shape ``mgmt_adj_writer.apply_mgmt_adj``
    expects.
    """
    if not isinstance(mgmt_state, dict):
        return {}
    try:
        names = mgmt_adj_writer.list_pool_names(template_path)
    except Exception:  # noqa: BLE001
        names = []
    saved = mgmt_state.get("pool_rows") or {}
    if not isinstance(saved, dict):
        saved = {}
    ordered: list[dict] = []
    for name in names:
        row = saved.get(name) or {}
        ordered.append({
            "name": name,
            "hard_code_pct": float(row.get("hard_code_pct") or 0.0),
            "use_default": bool(row.get("use_default")),
        })
    return {
        "default_pct": float(mgmt_state.get("default_pct") or 0.0),
        "pool_rows": ordered,
        "portfolio": mgmt_state.get("portfolio") or {},
    }


# Tab sets per report variant. Shared tabs (Management Adjustment,
# Historical Data, Industry Data) are NOT listed below and are kept
# visible for every variant. ``DNU `` sheets are always hidden.
_TCT_ONLY_SHEETS = {
    "Cover", "Introduction", "Scale Calculation", "Env Factor by Pool",
    "Environmental Factor Ranges", " Impaired Loans ASC 310-10", "Calc tab",
}
_VIZO_ONLY_SHEETS = {
    "Cover-Vizo", "Executive Summary-Vizo", "Scale Calculation-Vizo",
    "Env Factor by Pool-Vizo", " Impaired Loans-Vizo", "Introduction-Vizo",
    "Explanation of ACL Calc-Vizo", "Envir Factor Ranges-Vizo",
    "New Report Calc-Vizo",
}


def _hide_other_variant(
    workbook_path: str | Path, keep_variant: str,
) -> tuple[list[str], list[str]]:
    """Hide the sheets that don't belong to ``keep_variant`` ('tct' or 'vizo').

    ``DNU `` sheets are always hidden. Returns (hidden, kept) name lists.
    Shared tabs (Management Adjustment, Historical Data, Industry Data)
    are kept visible for both variants.
    """
    hidden: list[str] = []
    kept: list[str] = []
    wb = openpyxl.load_workbook(workbook_path)
    for name in list(wb.sheetnames):
        ws = wb[name]
        if name.startswith("DNU "):
            hide = True
        elif keep_variant == "tct" and name in _VIZO_ONLY_SHEETS:
            hide = True
        elif keep_variant == "vizo" and name in _TCT_ONLY_SHEETS:
            hide = True
        else:
            hide = False
        if hide:
            ws.sheet_state = "hidden"
            hidden.append(name)
        else:
            ws.sheet_state = "visible"
            kept.append(name)
    visible = [s for s in wb.sheetnames if wb[s].sheet_state == "visible"]
    if visible:
        wb.active = wb.sheetnames.index(visible[0])
    wb.save(workbook_path)
    return hidden, kept


def apply_report_variant(
    workbook_path: str | Path, variant: str,
) -> dict:
    """Produce one themed output file per requested variant.

    A workbook can only carry a single ``xl/theme/theme1.xml``. The
    master template ships with the Vizo theme; the TCT variant uses the
    default Office theme. To honour both, we write a per-variant copy of
    ``workbook_path`` (suffixed ``_TCT.xlsx`` / ``_Vizo.xlsx``), hide the
    other variant's sheets in the copy, and inject the correct theme.

    Hidden (not deleted) so cross-sheet formulas (e.g.
    ``='Scale Calculation'!C9``) keep evaluating.

    ``variant`` in {``tct``, ``vizo``, ``both``}. When at least one
    themed file is written the master ``workbook_path`` is removed so
    the user is not confused by an extra single-theme file.

    Returns ``{variant, outputs:[{path,label,hidden,kept,theme}],
    hidden, kept}`` where ``hidden``/``kept`` mirror the last output for
    back-compat with existing call sites.
    """
    v = (variant or "both").strip().lower()
    if v not in ("tct", "vizo", "both"):
        v = "both"
    master = Path(workbook_path)
    targets: list[str] = []
    if v in ("tct", "both"):
        targets.append("tct")
    if v in ("vizo", "both"):
        targets.append("vizo")

    vizo_theme = _read_workbook_theme(master)
    tct_theme = _default_office_theme_bytes()

    outputs: list[dict] = []
    last_hidden: list[str] = []
    last_kept: list[str] = []
    for which in targets:
        suffix = "_TCT" if which == "tct" else "_Vizo"
        out_path = master.with_name(master.stem + suffix + master.suffix)
        shutil.copy2(master, out_path)
        hidden, kept = _hide_other_variant(out_path, which)
        theme_bytes = tct_theme if which == "tct" else vizo_theme
        theme_label = "Office default" if which == "tct" else "Vizo"
        if theme_bytes is not None:
            try:
                _replace_workbook_theme(out_path, theme_bytes)
            except Exception:  # noqa: BLE001
                # Theme swap is best-effort; the file is still usable.
                theme_label += " (theme swap failed)"
        outputs.append({
            "path": str(out_path),
            "label": "TCT" if which == "tct" else "Vizo",
            "hidden": hidden,
            "kept": kept,
            "theme": theme_label,
        })
        last_hidden, last_kept = hidden, kept

    # Remove the un-themed master once at least one themed copy exists.
    if outputs and master.exists():
        try:
            master.unlink()
        except OSError:
            pass

    return {
        "variant": v,
        "outputs": outputs,
        "hidden": last_hidden,
        "kept": last_kept,
    }


def _report_variant_from_state(state: dict) -> str:
    scale = state.get("scale") or {}
    return (scale.get("report_variant") or "both").strip().lower()


def fill_template(
    template_path: str | Path,
    out_path: str | Path,
    mapping_rows: list[dict],
    fields: dict[str, Any],
) -> dict:
    wb = openpyxl.load_workbook(template_path)
    applied = 0
    issues: list[str] = []
    missing_fields: list[str] = []
    missing_sheets: list[str] = []
    for row in mapping_rows:
        code = row["field_code"]
        sheet = row["sheet"]
        cell = row["cell"]
        if code not in fields:
            missing_fields.append(code)
            continue
        if sheet not in wb.sheetnames:
            missing_sheets.append(f"{sheet} (for {code})")
            continue
        try:
            wb[sheet][cell].value = solr_fetcher.coerce_numeric(fields[code])
            applied += 1
        except Exception as exc:  # noqa: BLE001
            issues.append(f"Failed to write {code} to {sheet}!{cell}: {exc}")
    wb.save(out_path)
    return {
        "applied": applied,
        "issues": issues,
        "missing_fields": missing_fields,
        "missing_sheets": missing_sheets,
    }


def run_single_quarter(state: dict, workspace_root: str) -> dict:
    """Drive a one-quarter SCALE fill from wizard state.

    Returns a result dict suitable for stashing back into
    ``state["scale"]["last_run"]`` and rendering on the review page.
    """
    scale = state.get("scale") or {}
    period = (scale.get("period") or "").strip()
    solr_url = (scale.get("solr_url") or "").strip()
    solr_core = (scale.get("solr_core") or "").strip()
    charter_raw = state.get("charter_number") or ""
    short = state.get("short_name") or ""

    errors: list[str] = []
    if not period:
        errors.append("Target period (YYYY-MM) is required.")
    if not solr_url:
        errors.append("Solr URL is required.")
    if not solr_core:
        errors.append("Solr core is required.")
    if not charter_raw:
        errors.append("Charter number is missing on the Identity step.")
    if errors:
        return {"ok": False, "errors": errors, "ran_at": "", "output_path": ""}
    try:
        charter = int(charter_raw)
    except (TypeError, ValueError):
        return {
            "ok": False,
            "errors": [f"Charter number {charter_raw!r} is not numeric."],
            "ran_at": "",
            "output_path": "",
        }

    tmpl = template_loader.resolve_template(
        period, scale.get("template_override_path") or None
    )
    if not tmpl["ok"]:
        return {"ok": False, "errors": [tmpl["message"]], "ran_at": "",
                "output_path": ""}
    mp = template_loader.resolve_map(
        period, scale.get("map_override_path") or None
    )
    if not mp["ok"]:
        return {"ok": False, "errors": [mp["message"]], "ran_at": "",
                "output_path": ""}

    try:
        doc = solr_fetcher.fetch_doc(
            solr_url=solr_url,
            core=solr_core,
            charter=charter,
            period=period,
            username=scale.get("solr_user") or None,
            password=scale.get("solr_pass") or None,
        )
    except LookupError as exc:
        return {"ok": False, "errors": [str(exc)], "ran_at": "",
                "output_path": ""}
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "errors": [f"Solr fetch failed: {exc}"],
                "ran_at": "", "output_path": ""}

    try:
        rows = mapping_loader.load_rows(mp["path"])
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "errors": [f"Mapping load failed: {exc}"],
                "ran_at": "", "output_path": ""}

    out_path = _output_path(workspace_root, short, period)
    fill = fill_template(tmpl["path"], out_path, rows, doc)

    qf_entries = _qfactor_entries_from_state(state)
    qf_result = apply_qfactors(out_path, qf_entries)

    mgmt_state = (state.get("scale") or {}).get("mgmt_adj") or {}
    mgmt_state_norm = _normalize_mgmt_adj_state(mgmt_state, tmpl["path"])
    mgmt_result = mgmt_adj_writer.apply_mgmt_adj(out_path, mgmt_state_norm)

    imp_rows = _impaired_rows_from_state(state)
    imp_result = impaired_loader.apply_impaired_rows(out_path, imp_rows)

    env_result = env_factor_writer.apply_env_factor_ranges(out_path)

    variant = _report_variant_from_state(state)
    variant_result = apply_report_variant(out_path, variant)
    primary_output = (
        variant_result["outputs"][0]["path"]
        if variant_result["outputs"] else str(out_path)
    )

    prior_acl = _inject_prior_acl(
        workspace_root, short, period, variant_result["outputs"]
    )

    return {
        "ok": True,
        "ran_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "output_path": primary_output,
        "output_files": variant_result["outputs"],
        "period": period,
        "charter": charter,
        "prior_acl": prior_acl,
        "template_path": tmpl["path"],
        "template_source": tmpl["source"],
        "template_message": tmpl.get("message", ""),
        "map_path": mp["path"],
        "map_source": mp["source"],
        "applied": fill["applied"],
        "total_rows": len(rows),
        "issues": fill["issues"],
        "missing_fields": fill["missing_fields"],
        "missing_sheets": fill["missing_sheets"],
        "qfactor_applied": qf_result["applied"],
        "qfactor_total": qf_result["total"],
        "qfactor_missing_sheets": qf_result["missing_sheets"],
        "qfactor_issues": qf_result["issues"],
        "mgmt_adj_ok": mgmt_result["ok"],
        "mgmt_adj_pools_written": mgmt_result["pools_written"],
        "mgmt_adj_default_written": mgmt_result["default_written"],
        "mgmt_adj_portfolio_written": mgmt_result["portfolio_written"],
        "mgmt_adj_error": mgmt_result["error"],
        "impaired_applied": imp_result["applied"],
        "impaired_cleared": imp_result["cleared"],
        "impaired_error": imp_result["error"],
        "env_factor_ok": env_result["ok"],
        "env_factor_applied_delq": env_result["applied_delq"],
        "env_factor_applied_econ": env_result["applied_econ"],
        "env_factor_error": env_result["error"],
        "report_variant": variant_result["variant"],
        "report_variant_hidden": variant_result["hidden"],
        "errors": [],
    }


def run_multi_quarter(
    state: dict,
    workspace_root: str,
    quarters: int = 32,
) -> dict:
    """Drive a multi-quarter SCALE fill from wizard state.

    Iterates newest -> oldest. Resolves the starting template once
    (period-exact if available, else newest <= start). After the first
    successful iteration, the output file becomes the template for
    subsequent quarters so writes accumulate. Each period MUST have an
    exact-match mapping CSV in ``cecl_ui/data/scale/maps/`` or that
    quarter is skipped (recorded in ``skipped``).

    Returns ``{ok, ran_at, output_path, periods, iterations, skipped,
    errors}`` where ``iterations`` is one entry per attempted quarter.
    """
    scale = state.get("scale") or {}
    start_period = (scale.get("period") or "").strip()
    solr_url = (scale.get("solr_url") or "").strip()
    solr_core = (scale.get("solr_core") or "").strip()
    charter_raw = state.get("charter_number") or ""
    short = state.get("short_name") or ""

    errors: list[str] = []
    if not start_period:
        errors.append("Target period (YYYY-MM) is required.")
    if not solr_url:
        errors.append("Solr URL is required.")
    if not solr_core:
        errors.append("Solr core is required.")
    if not charter_raw:
        errors.append("Charter number is missing on the Identity step.")
    if errors:
        return {
            "ok": False, "errors": errors, "ran_at": "",
            "output_path": "", "periods": [], "iterations": [], "skipped": [],
        }
    try:
        charter = int(charter_raw)
    except (TypeError, ValueError):
        return {
            "ok": False,
            "errors": [f"Charter number {charter_raw!r} is not numeric."],
            "ran_at": "", "output_path": "", "periods": [],
            "iterations": [], "skipped": [],
        }

    try:
        periods = make_periods(start_period, quarters)
    except Exception as exc:  # noqa: BLE001
        return {
            "ok": False, "errors": [f"Period list failed: {exc}"],
            "ran_at": "", "output_path": "", "periods": [],
            "iterations": [], "skipped": [],
        }

    # Resolve the starting template once. Per-CU override always wins.
    tmpl = template_loader.resolve_template(
        start_period, scale.get("template_override_path") or None
    )
    if not tmpl["ok"]:
        return {
            "ok": False, "errors": [tmpl["message"]], "ran_at": "",
            "output_path": "", "periods": periods, "iterations": [],
            "skipped": [],
        }

    out_path = _output_path(workspace_root, short, start_period)
    iterations: list[dict] = []
    skipped: list[dict] = []
    successes = 0
    current_template: str = tmpl["path"]

    for idx, p in enumerate(periods, start=1):
        # Per-quarter mapping override only applies to the starting
        # period (the user uploaded it for that period); older quarters
        # always use canonical maps.
        override_map = (
            scale.get("map_override_path")
            if p == start_period else None
        )
        mp = template_loader.resolve_map(p, override_map or None)
        if not mp["ok"]:
            skipped.append({"period": p, "reason": mp["message"]})
            continue

        try:
            doc = solr_fetcher.fetch_doc(
                solr_url=solr_url,
                core=solr_core,
                charter=charter,
                period=p,
                username=scale.get("solr_user") or None,
                password=scale.get("solr_pass") or None,
            )
        except LookupError as exc:
            skipped.append({"period": p, "reason": f"Solr: {exc}"})
            continue
        except Exception as exc:  # noqa: BLE001
            skipped.append({"period": p, "reason": f"Solr fetch failed: {exc}"})
            continue

        try:
            rows = mapping_loader.load_rows(mp["path"])
        except Exception as exc:  # noqa: BLE001
            skipped.append({"period": p, "reason": f"Map load failed: {exc}"})
            continue

        fill = fill_template(current_template, out_path, rows, doc)
        successes += 1
        iterations.append({
            "period": p,
            "index": idx,
            "map_path": mp["path"],
            "map_source": mp["source"],
            "applied": fill["applied"],
            "total_rows": len(rows),
            "missing_fields_count": len(fill["missing_fields"]),
            "missing_sheets_count": len(fill["missing_sheets"]),
            "issues_count": len(fill["issues"]),
        })
        # After first successful write, accumulate into the output file.
        current_template = str(out_path)

    # Apply Q-factor overlays once at the end so they sit on top of the
    # most recent 5300 writes (won't be overwritten by accumulation).
    qf_result = {"applied": 0, "total": 0, "missing_sheets": [], "issues": []}
    imp_result = {"applied": 0, "cleared": 0, "error": ""}
    mgmt_result = {"ok": False, "pools_written": 0, "default_written": False,
                   "portfolio_written": False, "error": ""}
    variant_result: dict = {
        "variant": "both", "hidden": [], "kept": [], "outputs": [],
    }
    env_result = {"ok": False, "applied_delq": 0, "applied_econ": 0,
                  "skipped": [], "error": ""}
    if successes > 0:
        qf_entries = _qfactor_entries_from_state(state)
        qf_result = apply_qfactors(out_path, qf_entries)
        mgmt_state = (state.get("scale") or {}).get("mgmt_adj") or {}
        mgmt_state_norm = _normalize_mgmt_adj_state(mgmt_state, tmpl["path"])
        mgmt_result = mgmt_adj_writer.apply_mgmt_adj(out_path, mgmt_state_norm)
        imp_rows = _impaired_rows_from_state(state)
        imp_result = impaired_loader.apply_impaired_rows(out_path, imp_rows)
        env_result = env_factor_writer.apply_env_factor_ranges(out_path)
        variant_result = apply_report_variant(
            out_path, _report_variant_from_state(state)
        )

    primary_output = (
        variant_result["outputs"][0]["path"]
        if variant_result.get("outputs") else str(out_path)
    )

    prior_acl = _inject_prior_acl(
        workspace_root, short, start_period, variant_result.get("outputs") or []
    ) if successes > 0 else {"prior_period": "", "prior_path": "", "applied": [], "error": ""}
    return {
        "ok": successes > 0,
        "ran_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "output_path": primary_output,
        "output_files": variant_result.get("outputs", []),
        "period": start_period,
        "charter": charter,
        "quarters_requested": quarters,
        "quarters_written": successes,
        "periods": periods,
        "iterations": iterations,
        "skipped": skipped,
        "template_path": tmpl["path"],
        "template_source": tmpl["source"],
        "template_message": tmpl.get("message", ""),
        "qfactor_applied": qf_result["applied"],
        "qfactor_total": qf_result["total"],
        "qfactor_missing_sheets": qf_result["missing_sheets"],
        "qfactor_issues": qf_result["issues"],
        "mgmt_adj_ok": mgmt_result["ok"],
        "mgmt_adj_pools_written": mgmt_result["pools_written"],
        "mgmt_adj_default_written": mgmt_result["default_written"],
        "mgmt_adj_portfolio_written": mgmt_result["portfolio_written"],
        "mgmt_adj_error": mgmt_result["error"],
        "impaired_applied": imp_result["applied"],
        "impaired_cleared": imp_result["cleared"],
        "impaired_error": imp_result["error"],
        "env_factor_ok": env_result["ok"],
        "env_factor_applied_delq": env_result["applied_delq"],
        "env_factor_applied_econ": env_result["applied_econ"],
        "env_factor_error": env_result["error"],
        "report_variant": variant_result["variant"],
        "report_variant_hidden": variant_result["hidden"],
        "prior_acl": prior_acl,
        "errors": [] if successes > 0 else ["No quarters were successfully written."],
    }


def _inject_prior_acl(workspace_root: str, short: str, period: str,
                      output_files: list[dict]) -> dict:
    """Hard-code Prior ACL values from the previous quarter's report
    into Historical Data!AY2..AY5 of every output file.

    Looks up `Generated_Reports/<short>/<prev_quarter>/` for the
    newest SCALE workbook (prefers Vizo variant) and copies its
    Total Expected Losses / ACL Balance / Adjustment / ACL Ratio
    cells. Silently no-ops when the prior report is missing -- e.g.
    first ever run for this CU.
    """
    result: dict[str, Any] = {
        "prior_period": "", "prior_path": "",
        "applied": [], "error": "",
    }
    try:
        prev = prev_quarter(period)
    except Exception as exc:  # noqa: BLE001
        result["error"] = f"Could not compute prior quarter: {exc}"
        return result
    result["prior_period"] = prev
    prior_path = runs_service.find_prior_report(workspace_root, short, prev)
    if prior_path is None:
        # Fallback: newest report strictly older than target period.
        fallback = runs_service.find_newest_prior_report(
            workspace_root, short, period
        )
        if fallback is None:
            return result
        prev, prior_path = fallback
        result["prior_period"] = prev
    result["prior_path"] = str(prior_path)
    read = runs_service.read_prior_acl_values(prior_path)
    if not read["ok"]:
        result["error"] = read["error"]
        return result
    values = read["values"]
    for entry in output_files or []:
        path = entry.get("path") or ""
        if not path:
            continue
        write = runs_service.write_prior_acl_values(path, values)
        result["applied"].append({
            "path": path,
            "ok": write["ok"],
            "written": write["written"],
            "error": write["error"],
        })
    return result


def run_quarter_carry_history(state: dict, workspace_root: str) -> dict:
    """Run only the target quarter, carrying historical data from a
    prior quarter's report (the "normal process" mode).

    Copies the most recent prior SCALE workbook (`*_Vizo.xlsx`
    preferred) to the new output path, unhides any tabs the variant
    apply step previously hid, then overlays the target quarter's
    Solr mapping, Q-factors, mgmt-adj, impaired, and env factors --
    leaving every earlier quarter's data and the Historical Data tab
    intact.

    Returns a dict shaped like `run_single_quarter` plus extra
    `carry_*` provenance keys.
    """
    scale = state.get("scale") or {}
    period = (scale.get("period") or "").strip()
    solr_url = (scale.get("solr_url") or "").strip()
    solr_core = (scale.get("solr_core") or "").strip()
    charter_raw = state.get("charter_number") or ""
    short = state.get("short_name") or ""

    errors: list[str] = []
    if not period:
        errors.append("Target period (YYYY-MM) is required.")
    if not solr_url:
        errors.append("Solr URL is required.")
    if not solr_core:
        errors.append("Solr core is required.")
    if not charter_raw:
        errors.append("Charter number is missing on the Identity step.")
    if not short:
        errors.append("Short name is missing on the Identity step.")
    if errors:
        return {"ok": False, "errors": errors, "ran_at": "",
                "output_path": ""}
    try:
        charter = int(charter_raw)
    except (TypeError, ValueError):
        return {
            "ok": False,
            "errors": [f"Charter number {charter_raw!r} is not numeric."],
            "ran_at": "", "output_path": "",
        }

    try:
        prev_period = prev_quarter(period)
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "errors": [str(exc)], "ran_at": "",
                "output_path": ""}

    prior_path = runs_service.find_prior_report(workspace_root, short, prev_period)
    if prior_path is None:
        # Search further back -- the user may have skipped a quarter.
        fallback = runs_service.find_newest_prior_report(
            workspace_root, short, period
        )
        if fallback is None:
            return {
                "ok": False,
                "errors": [
                    f"No prior SCALE report found for {short}. Run the "
                    "'Re-pull all from 5300' mode for the first run, "
                    "or generate a baseline through the wizard.",
                ],
                "ran_at": "", "output_path": "",
            }
        prev_period, prior_path = fallback

    mp = template_loader.resolve_map(
        period, scale.get("map_override_path") or None
    )
    if not mp["ok"]:
        return {"ok": False, "errors": [mp["message"]], "ran_at": "",
                "output_path": ""}

    try:
        doc = solr_fetcher.fetch_doc(
            solr_url=solr_url,
            core=solr_core,
            charter=charter,
            period=period,
            username=scale.get("solr_user") or None,
            password=scale.get("solr_pass") or None,
        )
    except LookupError as exc:
        return {"ok": False, "errors": [str(exc)], "ran_at": "",
                "output_path": ""}
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "errors": [f"Solr fetch failed: {exc}"],
                "ran_at": "", "output_path": ""}

    try:
        rows = mapping_loader.load_rows(mp["path"])
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "errors": [f"Mapping load failed: {exc}"],
                "ran_at": "", "output_path": ""}

    out_path = _output_path(workspace_root, short, period)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(prior_path, out_path)
    runs_service.unhide_all_sheets(out_path)

    # Now overlay the target quarter's data on top of the carried workbook.
    fill = fill_template(out_path, out_path, rows, doc)

    qf_entries = _qfactor_entries_from_state(state)
    qf_result = apply_qfactors(out_path, qf_entries)

    mgmt_state = (state.get("scale") or {}).get("mgmt_adj") or {}
    mgmt_state_norm = _normalize_mgmt_adj_state(mgmt_state, str(out_path))
    mgmt_result = mgmt_adj_writer.apply_mgmt_adj(out_path, mgmt_state_norm)

    imp_rows = _impaired_rows_from_state(state)
    imp_result = impaired_loader.apply_impaired_rows(out_path, imp_rows)

    env_result = env_factor_writer.apply_env_factor_ranges(out_path)

    variant = _report_variant_from_state(state)
    variant_result = apply_report_variant(out_path, variant)
    primary_output = (
        variant_result["outputs"][0]["path"]
        if variant_result["outputs"] else str(out_path)
    )

    prior_acl = _inject_prior_acl(
        workspace_root, short, period, variant_result["outputs"]
    )

    return {
        "ok": True,
        "ran_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "output_path": primary_output,
        "output_files": variant_result["outputs"],
        "period": period,
        "charter": charter,
        "carry_mode": True,
        "carry_from_period": prev_period,
        "carry_from_path": str(prior_path),
        "map_path": mp["path"],
        "map_source": mp["source"],
        "applied": fill["applied"],
        "total_rows": len(rows),
        "issues": fill["issues"],
        "missing_fields": fill["missing_fields"],
        "missing_sheets": fill["missing_sheets"],
        "qfactor_applied": qf_result["applied"],
        "qfactor_total": qf_result["total"],
        "qfactor_missing_sheets": qf_result["missing_sheets"],
        "qfactor_issues": qf_result["issues"],
        "mgmt_adj_ok": mgmt_result["ok"],
        "mgmt_adj_pools_written": mgmt_result["pools_written"],
        "mgmt_adj_default_written": mgmt_result["default_written"],
        "mgmt_adj_portfolio_written": mgmt_result["portfolio_written"],
        "mgmt_adj_error": mgmt_result["error"],
        "impaired_applied": imp_result["applied"],
        "impaired_cleared": imp_result["cleared"],
        "impaired_error": imp_result["error"],
        "env_factor_ok": env_result["ok"],
        "env_factor_applied_delq": env_result["applied_delq"],
        "env_factor_applied_econ": env_result["applied_econ"],
        "env_factor_error": env_result["error"],
        "report_variant": variant_result["variant"],
        "report_variant_hidden": variant_result["hidden"],
        "prior_acl": prior_acl,
        "errors": [],
    }
