"""Auto-setup orchestrator for the Migration wizard.

Given a folder full of credit-union raw data, this module:

  1. Classifies the files (loan extracts, balance sheets, impaired loans,
     credit pulls, WARM workbooks, charge-off & recovery files, etc.)
     by leaning on ``warm_parser.scan_historical_folder``.
  2. Runs the appropriate per-file parser on the most-representative
     file of each type (sample_parser for loan extracts, monthly_bal_parser
     for balance sheet workbooks, etc.).
  3. Mutates a wizard ``state`` dict to populate as much as it can
     auto-derive (file paths, file patterns, column mappings, pool codes,
     mgmt-adj baselines, report defaults, etc.).
  4. Returns a per-step "what was filled" report and a list of HIL steps
     that still need user attention.

The wizard's existing per-step handlers continue to own all editing UX;
this module is only an upfront populater.

Patterns learned from the Census FCU greenfield script
(``scripts/setup_census_scratch.py``) are baked in:

  * mgmt_adj baselines (ltv_baseline=0.9, probability_factor=0.35).
  * reports default: Vizo-only when ``"vizo financial"`` is in the folder
    path (the Vizo Financial corporate-CU client convention); otherwise TCT.
  * impaired-loans block emitted only when an impaired-loans file is
    discovered for the target snapshot.
  * credit-pull block emitted only when a matching credit-pull file
    is discovered for the target snapshot.
  * pool_code_split default ``"/"``.
  * credit_grades default 5-grade industry standard with reserve_rate on
    every entry (so ``cecl_engine.get_reserve_rate`` doesn't crash).
"""
from __future__ import annotations

import re
import shutil
import tempfile
from pathlib import Path
from typing import Any

from cecl_ui.services import (
    config_service,
    monthly_bal_parser,
    sample_parser,
    warm_parser,
)


# --------------------------------------------------------------------------
# File classification
# --------------------------------------------------------------------------

# Extra patterns layered on top of ``warm_parser.scan_historical_folder``.
# That function already buckets WARM / CO / recov / impaired / credit-pull
# / monthly-bal / loan-data, but we want to flag annual balance-sheet
# workbooks separately (file names like ``2025 Balance Sheets.xlsx``).
_ANNUAL_BAL_RX = re.compile(
    r"(?:^|[^a-z])(?P<yr>20\d{2})\s*balance\s*sheets?",
    re.IGNORECASE,
)

# A "monthly detailed balance sheet" file (per-month workbook).
# Pattern: name contains "balance" + "sheet" + an 8-digit date.
_MONTHLY_DETAIL_RX = re.compile(
    r"detailed?\s*balance\s*sheet.*?(?P<dt>\d{8})",
    re.IGNORECASE,
)

# Recognise a 5300 / call-report file (informational only).
_5300_RX = re.compile(r"(?:^|[^a-z])(?:5300|call[\s_\-]*report)", re.IGNORECASE)

# Broader loan-extract patterns. warm_parser's _LOAN_DATA_RX only matches
# "aires" / "loan data" / "loan file" / "lndn" — too narrow for CUs like
# Census FCU whose monthly extracts are named "Dec 2025 Loans V2.xlsx" or
# "Dec 2025 Cuma Loans.xlsx". We re-classify ``other_files`` against this
# richer pattern, but ONLY when the file ALSO has a month/year fingerprint
# in its name (so we don't accidentally grab "Loan Type Codes.xlsx" etc).
_LOAN_EXTRA_RX = re.compile(
    r"(?:^|[\s_\-])(?:loans?(?:[\s_\-]*v?\d+)?|"
    r"cuma[\s_\-]*loans?|"
    r"active[\s_\-]*loans?|"
    r"loan[\s_\-]*portfolio|"
    r"loan[\s_\-]*detail|"
    r"loan[\s_\-]*list)"
    r"(?:[\s_\-]|\.|$)",
    re.IGNORECASE,
)

# A name "looks dated" when it has a YYYY-MM, YYYYMMDD, MM-YYYY, or
# month-name + year combo.
_DATE_HINT_RX = re.compile(
    r"(20\d{2}[-_]?\d{2})"
    r"|(\d{2}[-_]?20\d{2})"
    r"|(\d{8})"
    r"|((?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*[-_\s]+20\d{2})",
    re.IGNORECASE,
)

# Filenames that explicitly are NOT loan extracts even when they contain
# the word "loans" (mapping / reference workbooks, not data).
_LOAN_FALSE_POSITIVE_RX = re.compile(
    r"(loan[\s_\-]*type[\s_\-]*code|loan[\s_\-]*code[\s_\-]*map|"
    r"loan[\s_\-]*pool[\s_\-]*map|collateral[\s_\-]*code|"
    r"impaired[\s_\-]*loan|standalone[\s_\-]*impaired|"
    r"collateral[\s_\-]*dependent|"
    r"improved[\s_\-]*deteriorated|"
    r"impr[\s_\-]*deter|"
    r"watchlist|risk[\s_\-]*rating|"
    r"loan[\s_\-]*code|loan[\s_\-]*pool[\s_\-]*name)",
    re.IGNORECASE,
)


def _looks_like_loan_extract(name: str) -> bool:
    """True when a filename looks like a monthly loan-data extract."""
    if _LOAN_FALSE_POSITIVE_RX.search(name):
        return False
    if not _LOAN_EXTRA_RX.search(name):
        return False
    # Require either a date fingerprint OR the file is named generically
    # with "AIRES"/"LNDN" (those are already pre-classified by
    # warm_parser, so a hit here means the file has SOME month/year hint).
    return bool(_DATE_HINT_RX.search(name))


def _bucket_extra(entry: dict[str, Any], folder_path: str) -> str | None:
    """Return an extra bucket name for an ``other_files`` entry, or None."""
    name = entry.get("name", "")
    if _ANNUAL_BAL_RX.search(name):
        m = _ANNUAL_BAL_RX.search(name)
        try:
            entry["year"] = int(m.group("yr"))
        except (ValueError, AttributeError, TypeError):
            entry["year"] = None
        return "annual_balance_files"
    if _MONTHLY_DETAIL_RX.search(name):
        return "monthly_detail_balance_files"
    if _5300_RX.search(name):
        return "five_thirtythousand_files"
    if _looks_like_loan_extract(name):
        # Promote to loan_data_files. period_from_filename was already
        # set by warm_parser when it had a chance; re-derive here for
        # safety.
        if not entry.get("period"):
            entry["period"] = warm_parser._period_from_filename(name)  # noqa: SLF001
        return "loan_data_files"
    return None


def classify_folder(folder: str | Path) -> dict[str, Any]:
    """Walk *folder* and bucket every supported file by purpose.

    Wraps ``warm_parser.scan_historical_folder`` (which already does WARM /
    CO / recov / impaired / credit-pull / monthly-bal / loan-data) and
    layers on annual / monthly balance-sheet workbook detection plus 5300
    detection. Returns the same shape as ``scan_historical_folder`` plus
    additional bucket keys.
    """
    base = warm_parser.scan_historical_folder(folder)
    if not base.get("ok"):
        return base

    base.setdefault("annual_balance_files", [])
    base.setdefault("monthly_detail_balance_files", [])
    base.setdefault("five_thirtythousand_files", [])

    # Re-route files that fell into ``other_files`` into our extra buckets.
    leftover: list[dict[str, Any]] = []
    folder_str = str(folder)
    for entry in base.get("other_files") or []:
        extra = _bucket_extra(entry, folder_str)
        if extra:
            base[extra].append(entry)
        else:
            leftover.append(entry)
    base["other_files"] = leftover

    return base


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

_AUTOSCAN_TEMP = Path(tempfile.gettempdir()) / "cecl_ui_autoscan"


def _copy_into_tmp(src: Path, subfolder: str) -> Path:
    """Copy *src* into the wizard's working temp area and return the new path.

    The wizard's per-step upload handlers (Step 2 Sample, Step 15 Impaired,
    Step 10 Credit Pull, Step 5 Monthly Balance, etc.) all store uploaded
    files in dedicated subfolders of ``%TEMP%`` so the parsers can read
    them on subsequent requests without holding open handles. We mirror
    that pattern here so post-scan navigation between steps remains
    consistent.
    """
    dest_dir = _AUTOSCAN_TEMP / subfolder
    dest_dir.mkdir(parents=True, exist_ok=True)
    target = dest_dir / src.name
    try:
        shutil.copy2(src, target)
    except OSError:
        pass
    return target


def _safe_period(entry: dict[str, Any]) -> str | None:
    """Return ``YYYY-MM`` period string from a classifier entry, or None."""
    period = entry.get("period")
    if not period:
        return None
    period = str(period)
    if len(period) >= 7 and period[4] == "-":
        return period[:7]
    return None


def _pick_latest_for_period(
    entries: list[dict[str, Any]], snapshot_yyyymm: str | None
) -> dict[str, Any] | None:
    """Pick the entry whose period matches snapshot, else the latest period.

    When multiple entries share the target period, prefer those whose
    filename does NOT contain a subsystem keyword (cuma/mortgage/premier
    etc.) — those are usually the broader / primary extract for a CU
    that runs more than one loan system.

    Files with no period sort to the end. Returns None when *entries* is empty.
    """
    if not entries:
        return None

    def _is_subsystem(name: str) -> bool:
        # These keywords typically denote a partial / subsystem extract.
        return bool(
            re.search(
                r"(cuma|mortgage|premier|redicash|line[\s_\-]*of[\s_\-]*credit|"
                r"share[\s_\-]*secured)",
                name or "",
                re.IGNORECASE,
            )
        )

    if snapshot_yyyymm:
        matches = [e for e in entries if _safe_period(e) == snapshot_yyyymm]
        if matches:
            # Prefer non-subsystem files first.
            primary = [e for e in matches if not _is_subsystem(e.get("name", ""))]
            if primary:
                return primary[0]
            return matches[0]
    # Fall back to most-recent period (entries are NOT guaranteed sorted).
    sortable = sorted(
        entries,
        key=lambda e: (
            _safe_period(e) or "0000-00",
            0 if _is_subsystem(e.get("name", "")) else 1,  # non-subsystem wins ties
        ),
        reverse=True,
    )
    return sortable[0] if sortable else None


# --------------------------------------------------------------------------
# Findings runner
# --------------------------------------------------------------------------


def scan_folder_for_setup(
    folder: str | Path,
    snapshot_yyyymm: str | None = None,
) -> dict[str, Any]:
    """Scan *folder* and produce a ``findings`` dict for state population.

    Calls the existing per-file parsers (sample_parser, monthly_bal_parser)
    on the most-representative file of each type. Does not mutate any
    state; use ``apply_findings_to_state`` for that.

    Returns::

        {
            "ok": bool,
            "error": str | None,
            "folder": str,
            "classification": <classify_folder result>,
            "sample": <sample_parser.analyse_sample_file result> | None,
            "annual_bal": <monthly_bal_parser.analyse_per_year_file result> | None,
            "monthly_bal": <monthly_bal_parser.analyse_per_month_file result> | None,
            "warm": dict | None,
            "impaired_file": dict | None,   # entry that will be copied to Raw_Uploads
            "credit_pull_file": dict | None,
            "co_file": dict | None,
            "recov_file": dict | None,
            "messages": list[str],            # human-readable narration
            "errors":   list[str],
        }
    """
    folder = Path(folder)
    out: dict[str, Any] = {
        "ok": False,
        "error": None,
        "folder": str(folder),
        "classification": None,
        "sample": None,
        "annual_bal": None,
        "monthly_bal": None,
        "warm": None,
        "impaired_file": None,
        "credit_pull_file": None,
        "co_file": None,
        "recov_file": None,
        "messages": [],
        "errors": [],
    }

    cls = classify_folder(folder)
    out["classification"] = cls
    if not cls.get("ok"):
        out["error"] = cls.get("error") or "folder scan failed"
        return out

    out["ok"] = True
    out["messages"].append(
        f"Classified {len(cls.get('loan_data_files') or [])} loan extract(s), "
        f"{len(cls.get('annual_balance_files') or [])} annual balance sheet(s), "
        f"{len(cls.get('monthly_files') or [])} monthly balance file(s), "
        f"{len(cls.get('impaired_files') or [])} impaired-loan file(s), "
        f"{len(cls.get('credit_pull_files') or [])} credit-pull file(s), "
        f"{len(cls.get('co_files') or [])} CO file(s), "
        f"{len(cls.get('recov_files') or [])} recovery file(s), "
        f"{len(cls.get('warm_files') or [])} WARM file(s)."
    )

    # ----- Loan extract sample (drives column_mappings + pool_codes) -----
    loan_files = cls.get("loan_data_files") or []
    sample_entry = _pick_latest_for_period(loan_files, snapshot_yyyymm)
    sample_candidates = []
    if sample_entry:
        sample_candidates.append(sample_entry)
        # If multiple files share the snapshot period, queue them as
        # fallbacks in case the primary one analyses poorly.
        if snapshot_yyyymm:
            for e in loan_files:
                if e is sample_entry:
                    continue
                if _safe_period(e) == snapshot_yyyymm:
                    sample_candidates.append(e)

    best_analysis: dict[str, Any] | None = None
    best_entry: dict[str, Any] | None = None
    best_score = -1
    for cand in sample_candidates:
        src = Path(cand["path"])
        try:
            tmp = _copy_into_tmp(src, "sample")
        except Exception as exc:  # noqa: BLE001
            out["errors"].append(f"copy failed for {cand['name']}: {exc}")
            continue
        try:
            analysis = sample_parser.analyse_sample_file(str(tmp), cand["name"])
        except Exception as exc:  # noqa: BLE001
            out["errors"].append(
                f"sample_parser failed on {cand['name']}: {exc}"
            )
            continue
        if not analysis.get("ok"):
            out["errors"].append(
                f"sample analysis returned not-ok for "
                f"{cand['name']}: {analysis.get('error')}"
            )
            continue
        # Score = #column mappings + #pool codes. Higher = better
        # foothold for the wizard.
        score = (
            len(analysis.get("column_suggestions") or {})
            + len(analysis.get("pool_code_suggestions") or [])
        )
        if score > best_score:
            best_score = score
            analysis["saved_path"] = str(tmp)
            best_analysis = analysis
            best_entry = cand

    if best_analysis and best_entry:
        out["sample"] = best_analysis
        out["messages"].append(
            f"Analysed sample loan extract: {best_entry['name']} "
            f"({len(best_analysis.get('headers') or [])} columns, "
            f"{len(best_analysis.get('pool_code_suggestions') or [])} pool codes)"
        )

    # ----- Annual balance-sheet workbooks (drives monthly_bal per_year) ---
    annual_entries = sorted(
        cls.get("annual_balance_files") or [],
        key=lambda e: e.get("year") or 0,
        reverse=True,
    )
    if annual_entries:
        # Analyse the most-recent one to capture the layout (sheet, cols, etc).
        recent = annual_entries[0]
        try:
            an = monthly_bal_parser.analyse_per_year_file(recent["path"])
        except Exception as exc:  # noqa: BLE001
            out["errors"].append(
                f"monthly_bal_parser.analyse_per_year_file failed on "
                f"{recent['name']}: {exc}"
            )
        else:
            if an.get("ok"):
                out["annual_bal"] = {
                    "layout": {
                        "sheet": an.get("sheet"),
                        "header_row": an.get("header_row"),
                        "label_col": an.get("label_col"),
                        "period_columns": an.get("period_columns") or [],
                    },
                    "pool_labels": an.get("pool_labels") or [],
                    "files": annual_entries,
                }
                out["messages"].append(
                    f"Analysed annual balance-sheet layout from "
                    f"{recent['name']}: sheet={an.get('sheet')}, "
                    f"{len(an.get('pool_labels') or [])} pool labels."
                )

    # ----- Per-month detailed balance sheet (drives monthly_bal per_month) -
    if not out["annual_bal"] and cls.get("monthly_detail_balance_files"):
        # Pick the one matching snapshot if possible.
        m_entries = cls.get("monthly_detail_balance_files") or []
        recent = _pick_latest_for_period(m_entries, snapshot_yyyymm)
        if recent:
            try:
                an = monthly_bal_parser.analyse_per_month_file(recent["path"])
            except Exception as exc:  # noqa: BLE001
                out["errors"].append(
                    f"monthly_bal_parser.analyse_per_month_file failed on "
                    f"{recent['name']}: {exc}"
                )
            else:
                if an.get("ok"):
                    out["monthly_bal"] = {
                        "layout": {
                            "sheet": an.get("sheet"),
                            "header_row": an.get("header_row"),
                            "label_col": an.get("label_col"),
                            "balance_col": an.get("balance_col"),
                            "as_of_cell": an.get("as_of_cell"),
                            "stop_row": an.get("stop_row"),
                        },
                        "pool_labels": an.get("pool_labels") or [],
                        "files": m_entries,
                        "representative": recent,
                    }
                    out["messages"].append(
                        f"Analysed per-month balance layout from "
                        f"{recent['name']}."
                    )

    # ----- WARM (if any) -------------------------------------------------
    warm_entry = _pick_latest_for_period(
        cls.get("warm_files") or [], snapshot_yyyymm
    )
    if warm_entry:
        src = Path(warm_entry["path"])
        tmp = _copy_into_tmp(src, "warm")
        try:
            analysis = warm_parser.analyse_warm_file(tmp, warm_entry["name"])
        except Exception as exc:  # noqa: BLE001
            out["errors"].append(
                f"warm_parser failed on {warm_entry['name']}: {exc}"
            )
        else:
            out["warm"] = {
                "entry": warm_entry,
                "saved_path": str(tmp),
                "analysis": analysis,
            }
            out["messages"].append(
                f"Found WARM workbook for {warm_entry.get('period')}: "
                f"{warm_entry['name']}"
            )

    # ----- Companion files: impaired, credit pull, CO, recov -------------
    out["impaired_file"] = _pick_latest_for_period(
        cls.get("impaired_files") or [], snapshot_yyyymm
    )
    out["credit_pull_file"] = _pick_latest_for_period(
        cls.get("credit_pull_files") or [], snapshot_yyyymm
    )
    out["co_file"] = _pick_latest_for_period(
        cls.get("co_files") or [], snapshot_yyyymm
    )
    out["recov_file"] = _pick_latest_for_period(
        cls.get("recov_files") or [], snapshot_yyyymm
    )

    return out


# --------------------------------------------------------------------------
# State population
# --------------------------------------------------------------------------


# Heuristic: when the folder path contains "vizo financial" we know the CU
# is hosted under Vizo Financial Corporate CU's analyst tree. That client
# convention turns OFF TCT reports and turns ON Vizo + Vizo Supplemental
# + ImpDet. Pure-TCT setup is the safer fallback for any other CU.
_VIZO_FOLDER_HINT = re.compile(r"vizo[\s_\-]*financial", re.IGNORECASE)


def _set_reports_default_from_folder(state: dict[str, Any], folder: str) -> None:
    if _VIZO_FOLDER_HINT.search(folder or ""):
        state["reports"] = {
            "tct": False,
            "vizo": True,
            "vizo_supp": True,
            "impdet": True,
        }
    # else leave whatever default was already in state (TCT-only by default).


def _apply_mgmt_adj_defaults(state: dict[str, Any]) -> None:
    """Apply the Census-learned analyst-standard mgmt_adj baselines.

    Only writes when the current values are still the in-code default.
    """
    cur = state.get("mgmt_adj") or {}
    # In-code default: ltv_baseline=0.9, probability_factor=0.35. If the
    # state already matches those, that's because nothing's been edited
    # yet — re-writing is a no-op. We still write them explicitly so the
    # auto-fill report can show "mgmt_adj baselines applied".
    cur.setdefault("ltv_baseline", 0.9)
    cur.setdefault("probability_factor", 0.35)
    state["mgmt_adj"] = cur


def _apply_sample_to_state(
    state: dict[str, Any],
    sample: dict[str, Any],
    representative_name: str,
) -> list[str]:
    """Mutate state with the sample-file analysis. Returns step-level messages."""
    msgs: list[str] = []

    # state["sample"] is what Step 2 Sample / Step 12 Sample (no-WARM)
    # reads in subsequent renders.
    state["sample"] = {
        "filename": representative_name,
        "saved_path": sample.get("saved_path"),
        "has_header": bool(sample.get("has_header")),
        "header_row": sample.get("header_row"),
        "headers": sample.get("headers") or [],
        "sample_rows": sample.get("sample_rows") or [],
        "column_suggestions": sample.get("column_suggestions") or {},
        "pool_code_suggestions": sample.get("pool_code_suggestions") or [],
        "file_pattern": sample.get("file_pattern"),
        "date_pattern": sample.get("date_pattern"),
    }
    msgs.append(
        f"sample analysis loaded ({len(state['sample']['headers'])} columns, "
        f"{len(state['sample']['pool_code_suggestions'])} pool codes)"
    )

    # Column mappings: apply suggestions over any blank-or-default cells only.
    cm = state.setdefault("column_mappings", {})
    suggestions = sample.get("column_suggestions") or {}
    filled = []
    for field, header in suggestions.items():
        if not header:
            continue
        cur = cm.get(field, "")
        if not cur or cur == "MEMBER_ID" or cur == "BALANCE" or cur == "FICO_SCORE" \
                or cur == "LOAN_TYPE" or cur == "DQ_DAYS" or cur == "INT_RATE" \
                or cur == "OPEN_DATE" or cur == "ORIG_AMT":
            # blank or in-code default placeholder -> overwrite
            cm[field] = header
            filled.append(field)
    if filled:
        msgs.append(f"column_mappings filled: {', '.join(sorted(filled))}")

    # File pattern + date pattern (drives importer file discovery).
    if sample.get("file_pattern"):
        state["file_pattern"] = sample["file_pattern"]
        msgs.append(f"file_pattern set to {sample['file_pattern']}")
    if sample.get("date_pattern"):
        state["date_pattern"] = sample["date_pattern"]
        msgs.append(f"date_pattern set to {sample['date_pattern']}")
    # has_header
    state["has_header"] = bool(sample.get("has_header"))

    # Pool map seed: every detected code -> "Ignore" so user only has to
    # rename pools, not type the codes. We keep existing entries untouched
    # so re-running the scan doesn't clobber the user's renames.
    pm = state.setdefault("pool_map", {})
    pool_codes = sample.get("pool_code_suggestions") or []
    seeded = 0
    for code in pool_codes:
        if not code:
            continue
        if code not in pm:
            pm[code] = "Ignore"
            seeded += 1
    if seeded:
        msgs.append(
            f"pool_map seeded with {seeded} code(s) (all set to 'Ignore' — "
            "rename on the Loan Code Mapping step)"
        )

    # Sensible split-char default (learned from Census FCU greenfield).
    if not state.get("pool_code_split"):
        state["pool_code_split"] = "/"

    return msgs


def _apply_annual_bal_to_state(
    state: dict[str, Any], annual: dict[str, Any]
) -> list[str]:
    """Mutate state.monthly_bal with annual-balance-sheet findings."""
    msgs: list[str] = []
    mb = state.setdefault("monthly_bal", {})
    mb["source"] = "per_year"
    mb["per_year_layout"] = annual.get("layout") or {}
    files = []
    for entry in annual.get("files") or []:
        files.append({
            "year": entry.get("year"),
            "filename": entry.get("name"),
            "saved_path": entry.get("path"),
        })
    mb["year_files"] = files
    mb["parsed_pool_labels"] = annual.get("pool_labels") or []
    # User still needs to map labels -> pool names on Step 5.
    msgs.append(
        f"monthly_bal source=per_year with {len(files)} workbook(s), "
        f"{len(mb['parsed_pool_labels'])} pool labels detected"
    )
    return msgs


def _apply_monthly_bal_to_state(
    state: dict[str, Any], mbal: dict[str, Any]
) -> list[str]:
    """Mutate state.monthly_bal with per-month-detail findings."""
    msgs: list[str] = []
    mb = state.setdefault("monthly_bal", {})
    mb["source"] = "per_month"
    mb["per_month_layout"] = mbal.get("layout") or {}
    files = []
    for entry in mbal.get("files") or []:
        files.append({
            "filename": entry.get("name"),
            "saved_path": entry.get("path"),
            "period": entry.get("period"),
        })
    mb["monthly_files"] = files
    mb["parsed_pool_labels"] = mbal.get("pool_labels") or []
    msgs.append(
        f"monthly_bal source=per_month with {len(files)} file(s), "
        f"{len(mb['parsed_pool_labels'])} pool labels detected"
    )
    return msgs


def _apply_companion_file(
    state: dict[str, Any],
    kind: str,
    entry: dict[str, Any] | None,
    state_list_key: str,
    tmp_subfolder: str,
) -> list[str]:
    """Stage a companion file (impaired / credit-pull / CO / recov / etc.)
    into ``state['sample_uploads'][state_list_key]`` so the Review-step
    file copier picks it up.
    """
    if not entry:
        return []
    src = Path(entry["path"])
    if not src.exists():
        return []
    tmp = _copy_into_tmp(src, tmp_subfolder)
    su = state.setdefault("sample_uploads", {})
    bucket = su.setdefault(state_list_key, [])
    # Don't add duplicate names.
    if any(e.get("name") == src.name for e in bucket):
        return [f"{kind} file already staged: {src.name}"]
    bucket.append({"name": src.name, "path": str(tmp)})
    return [f"{kind} file staged: {src.name}"]


def _apply_credit_pull_to_state(
    state: dict[str, Any], entry: dict[str, Any] | None
) -> list[str]:
    """Stage credit-pull file (similar to companion_file but writes to
    state.credit_pull as well so Step 10 can render it correctly)."""
    if not entry:
        return []
    src = Path(entry["path"])
    if not src.exists():
        return []
    tmp = _copy_into_tmp(src, "credit_pull")
    cp = state.setdefault("credit_pull", {})
    cp["uploaded_filename"] = src.name
    cp["source_folder"] = str(src.parent)
    cp["use_standalone_file"] = True
    # Also stage into sample_uploads so Review-step copies into Raw_Uploads.
    su = state.setdefault("sample_uploads", {})
    bucket = su.setdefault("credit_pull_files", [])
    if not any(e.get("name") == src.name for e in bucket):
        bucket.append({"name": src.name, "path": str(tmp)})
    return [f"credit_pull file staged: {src.name}"]


def _apply_data_directory(
    state: dict[str, Any],
    workspace_root: str | Path,
) -> list[str]:
    """Point state.data_directory at the per-CU Raw_Uploads folder so the
    importer + report engines can find any companion files we'll copy
    there on Review submit.
    """
    sn = state.get("short_name") or config_service.slugify(
        state.get("credit_union") or ""
    )
    if not sn:
        return []
    raw = config_service.raw_uploads_dir(workspace_root) / sn
    state["data_directory"] = str(raw)
    return [f"data_directory set to {raw}"]


def apply_findings_to_state(
    state: dict[str, Any],
    findings: dict[str, Any],
    workspace_root: str | Path,
) -> dict[str, list[str]]:
    """Mutate *state* with everything *findings* can supply.

    Returns a dict ``{step_key: [messages]}`` describing what was filled
    on each wizard step. Step keys mirror the keys in
    ``WIZARD_STEPS_NO_WARM`` / ``WIZARD_STEPS_WARM`` so the stepper UI
    can render per-step status badges off the same dict.
    """
    report: dict[str, list[str]] = {}

    if not findings.get("ok"):
        report.setdefault("identity", []).append(
            f"Scan failed: {findings.get('error') or 'unknown error'}"
        )
        return report

    folder = findings.get("folder") or ""

    # ----- Sample / columns / pools / files steps ----------------------
    sample = findings.get("sample")
    if sample:
        rep = findings.get("classification", {}).get("loan_data_files") or []
        name = (rep and rep[0].get("name")) or ""
        msgs = _apply_sample_to_state(state, sample, name)
        report.setdefault("sample", []).extend(msgs)
        # column_mappings and file_pattern/date_pattern affect multiple
        # steps; mirror to those keys for stepper UX.
        if any("column_mappings" in m for m in msgs):
            report.setdefault("columns", []).append("column_mappings auto-filled from sample headers")
        if any("file_pattern" in m or "date_pattern" in m for m in msgs):
            report.setdefault("files", []).append("file_pattern and date_pattern auto-filled")
        if any("pool_map seeded" in m for m in msgs):
            report.setdefault("pools", []).append(
                "pool_map seeded with detected codes (all 'Ignore' — rename to your real pool names)"
            )

    # ----- monthly_bal step (annual or per-month) ----------------------
    if findings.get("annual_bal"):
        msgs = _apply_annual_bal_to_state(state, findings["annual_bal"])
        report.setdefault("monthly_bal", []).extend(msgs)
        # When we found annual workbooks, surface that on Step 3 historical
        # too because that step has a related "hist_balance_source" toggle.
        state["hist_balance_source"] = "annual_balance_sheets"
        report.setdefault("historical", []).append(
            "hist_balance_source set to annual_balance_sheets"
        )
    elif findings.get("monthly_bal"):
        msgs = _apply_monthly_bal_to_state(state, findings["monthly_bal"])
        report.setdefault("monthly_bal", []).extend(msgs)
        state["hist_balance_source"] = "monthly_balance_sheets"
        report.setdefault("historical", []).append(
            "hist_balance_source set to monthly_balance_sheets"
        )

    # ----- WARM presence ------------------------------------------------
    if findings.get("warm"):
        state["has_warm_files"] = "yes"
        report.setdefault("warm", []).append(
            f"WARM workbook detected — WARM-path wizard steps will be used."
        )
    else:
        # Don't override an explicit prior answer; only set when unanswered.
        if state.get("has_warm_files") is None:
            state["has_warm_files"] = "no"
            report.setdefault("loan_pools", []).append(
                "No WARM workbook found — non-WARM wizard path active."
            )

    # ----- Companion files: impaired, credit pull, CO, recov ----------
    cls = findings.get("classification") or {}
    if findings.get("impaired_file"):
        msgs = _apply_companion_file(
            state, "impaired", findings["impaired_file"],
            "impaired_files", "impaired"
        )
        report.setdefault("impaired", []).extend(msgs)
    if findings.get("credit_pull_file"):
        msgs = _apply_credit_pull_to_state(state, findings["credit_pull_file"])
        report.setdefault("credit_pull", []).extend(msgs)
    if findings.get("co_file"):
        msgs = _apply_companion_file(
            state, "charge-off", findings["co_file"],
            "co_files", "co"
        )
        report.setdefault("co_recov", []).extend(msgs)
        report.setdefault("co_history", []).extend(msgs)
    if findings.get("recov_file"):
        msgs = _apply_companion_file(
            state, "recovery", findings["recov_file"],
            "recov_files", "recov"
        )
        report.setdefault("co_recov", []).extend(msgs)
        report.setdefault("recov_history", []).extend(msgs)

    # ----- Folder-derived defaults -------------------------------------
    _set_reports_default_from_folder(state, folder)
    report.setdefault("reports", []).append(
        "reports defaults applied (Vizo-client convention if folder is under "
        "Vizo Financial; else TCT)"
    )

    _apply_mgmt_adj_defaults(state)
    report.setdefault("mgmt_adj", []).append(
        "mgmt_adj baselines applied (ltv_baseline=0.9, probability_factor=0.35)"
    )

    # ----- Raw_Uploads pointer -----------------------------------------
    msgs = _apply_data_directory(state, workspace_root)
    if msgs:
        report.setdefault("identity", []).extend(msgs)

    # ----- Persist the report on state so UI can read it back ----------
    state["_auto_scan_completed"] = True
    state["_auto_scan_folder"] = folder
    state["_auto_scan_report"] = report
    state["_auto_scan_messages"] = list(findings.get("messages") or [])
    state["_auto_scan_errors"] = list(findings.get("errors") or [])

    return report


# --------------------------------------------------------------------------
# HIL detection
# --------------------------------------------------------------------------


def compute_hil_needs(state: dict[str, Any]) -> list[dict[str, Any]]:
    """Walk *state* and return ordered list of steps still needing user input.

    Each entry: ``{step_key, severity, reasons: list[str]}`` where severity
    is ``"required"`` (the wizard cannot proceed without user input) or
    ``"recommended"`` (we filled a default but the user should review).

    The order matches the step list ordering so callers can pick the
    first entry as "next step to visit".
    """
    needs: list[dict[str, Any]] = []

    def add(key: str, severity: str, *reasons: str) -> None:
        needs.append({
            "step_key": key,
            "severity": severity,
            "reasons": list(reasons),
        })

    # identity — required CU name + report period
    if not state.get("credit_union"):
        add("identity", "required", "Credit Union name is blank")
    if not state.get("report_period"):
        add("identity", "recommended", "report_period (target YYYY-MM) is blank")

    # economic — state + county
    econ = state.get("economic_data") or {}
    if not econ.get("state"):
        add("economic", "required", "Economic state is blank")
    elif not econ.get("county"):
        add("economic", "recommended", "Economic county is blank")

    # pools — ALL seeded codes default to 'Ignore' which means user has not
    # renamed any. That's the most common 'must review' step after a scan.
    pm = state.get("pool_map") or {}
    if pm:
        ignore_count = sum(1 for v in pm.values() if str(v).lower() == "ignore")
        named_count = len(pm) - ignore_count
        if ignore_count == len(pm):
            add("pools", "required",
                f"All {len(pm)} detected loan codes are set to 'Ignore' — "
                "rename each to your real pool names")
        elif ignore_count:
            add("pools", "recommended",
                f"{ignore_count} of {len(pm)} loan codes still set to 'Ignore'")
    else:
        add("pools", "recommended", "No loan codes detected — upload a sample on the Loan Data Extracts step")

    # monthly_bal — at least one source must be picked AND mapped
    mb = state.get("monthly_bal") or {}
    if mb.get("source") in ("per_year", "per_month"):
        labels = mb.get("parsed_pool_labels") or []
        mapping = mb.get("pool_map") or {}
        if labels and not mapping:
            add("monthly_bal", "required",
                f"{len(labels)} balance-sheet labels detected but none mapped to pools")
        elif labels:
            unmapped = [lab for lab in labels if not mapping.get(lab)]
            if unmapped:
                add("monthly_bal", "recommended",
                    f"{len(unmapped)} balance-sheet label(s) still unmapped")
    else:
        # No source picked yet.
        add("monthly_bal", "required", "Monthly Balance source not yet chosen")

    # grades — only HIL when user hasn't customised reserve_rate yet.
    grades = state.get("credit_grades") or []
    if not grades:
        add("grades", "required", "Credit grades are empty")
    else:
        defaults = config_service.DEFAULT_CREDIT_GRADES
        # Defaults dict shape: list of {label, min_fico, max_fico, reserve_rate}
        is_default = (len(grades) == len(defaults)) and all(
            grades[i].get("reserve_rate") == defaults[i].get("reserve_rate")
            for i in range(len(grades))
        )
        if is_default:
            add("grades", "recommended",
                "Credit grades using factory default reserve rates — review "
                "against your analyst's CECL methodology")

    # credit_pull — recommended unless the user explicitly opted out.
    cp = state.get("credit_pull") or {}
    if not cp.get("uploaded_filename") and not cp.get("use_configured_report"):
        # Was a credit-pull file found in the auto-scan?
        cp_file_present = bool(
            (state.get("sample_uploads") or {}).get("credit_pull_files")
        )
        if not cp_file_present:
            add("credit_pull", "recommended",
                "No credit-pull file found for the target period — current "
                "FICO will equal original FICO (no grade movement)")

    # impaired — recommended review when a file was staged.
    if (state.get("sample_uploads") or {}).get("impaired_files"):
        add("impaired", "recommended",
            "Impaired-loans file staged — review parsed rows before final save")

    # files — recommended when we auto-filled from sample
    if not state.get("file_pattern"):
        add("files", "required", "File pattern is blank")

    # columns — required when member_number unmapped
    cm = state.get("column_mappings") or {}
    if not cm.get("member_number"):
        add("columns", "required", "member_number column is unmapped")

    # mgmt_adj — recommended review (we applied baseline defaults)
    if state.get("_auto_scan_completed"):
        add("mgmt_adj", "recommended",
            "ltv_baseline=0.9 and probability_factor=0.35 applied — review "
            "against your analyst's CECL methodology")

    # review — always shown last; the user must click Save to write YAML.
    add("review", "required", "Confirm and save the final YAML")

    return needs


def first_hil_step_key(state: dict[str, Any], step_list: list[tuple[str, str]]) -> str | None:
    """Pick the first HIL step that exists in *step_list* (in display order)."""
    needs = compute_hil_needs(state)
    need_keys = {n["step_key"] for n in needs}
    for key, _label in step_list:
        if key in need_keys:
            return key
    return None
