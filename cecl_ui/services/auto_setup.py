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

# A consolidated SINGLE historical balance workbook (one file, wide
# layout with one column per month-end). Filename hints:
#   * "Historical Loan Balances"
#   * "Balance History"
#   * "Balance_History_-_<CU>"
#   * "Loan Balance History"
#   * "Historical Balance(s)"
# When found, we set hist_balance_source = "single_workbook" and
# analyse_file() detects sheet / header_row / pool_name_col automatically.
_SINGLE_HIST_BAL_RX = re.compile(
    r"(?:historical[\s_\-]*(?:loan[\s_\-]*)?balances?"
    r"|balance[\s_\-]*history"
    r"|loan[\s_\-]*balance[\s_\-]*history)",
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
    # CUMA MTG Servicing reports are the mortgage-portfolio loan
    # extract for CUs that ship them (Vizo client convention, e.g.
    # ``CUMA MTG Servicing Report 12-31-2025 V2.xls``). Treat as a
    # loan-data extract.
    r"cuma[\s_\-]*mtg(?:[\s_\-]*servicing)?|"
    r"cuma[\s_\-]*servicing|"
    r"active[\s_\-]*loans?|"
    r"loan[\s_\-]*portfolio|"
    r"loan[\s_\-]*detail|"
    r"loan[\s_\-]*list|"
    # Symitar/Episys CECL bucket extracts: ceclce (Closed-End),
    # cecloe (Open-End), ceclcc / ceclcc1 (Credit Card). Some CUs
    # space-separate ("cecl ce"), some run together ("ceclce").
    r"cecl[\s_\-]*(?:ce|oe|cc\d*))"
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
    r"aires[\s_\-]*shares|shares[\s_\-]*aires|"
    r"loan[\s_\-]*code|loan[\s_\-]*pool[\s_\-]*name)",
    re.IGNORECASE,
)

# Aires/AIRES files come in TWO varieties: "Aires Loans" (the loan
# portfolio extract this tool consumes) and "Aires Shares" (the member
# share/deposit export, NOT loan data). Any filename matching this
# pattern is excluded from ``loan_data_files`` even if warm_parser's
# ``_LOAN_DATA_RX`` already matched on the bare "aires" token.
_AIRES_SHARES_RX = re.compile(
    r"aires[\s_\-]*shares|shares[\s_\-]*aires",
    re.IGNORECASE,
)

# Phase 9.26: filename heuristic for loan-data extracts whose loan codes
# legitimately contain ``/`` (so the global ``pool_code_split = "/"`` would
# corrupt them). CUMA / mortgage servicing exports ship codes like
# ``15/15 ARM``, ``30/30 FIXED`` where the ``/`` is part of the label, not a
# prefix:description separator. When a loan-extract candidate's filename
# matches this regex the auto-scan calls ``analyse_sample_file`` with
# ``split_char=""`` AND records ``pool_code_split=""`` on the staged extract
# entry so the importer's per-file override (Phase 9.22) takes effect at
# runtime. Filename keywords: ``CUMA``, ``MTG``, ``Mortgage``, ``Servicing``.
_NO_SPLIT_FILENAME_RX = re.compile(
    r"(?:^|[\s_\-])(?:cuma|mtg|mortgage|servicing)(?:[\s_\-]|$|\.)",
    re.IGNORECASE,
)

# Phase 9.26: value-shape detector for composite mortgage codes. Matches
# raw values like ``15/15 ARM``, ``30/30 FIXED``, ``5/1 ARM``, ``15 YR FIXED``
# — anything where a ``/`` separator is followed by alphabetic content
# rather than a pure prefix:description split. Used as a fallback when the
# filename heuristic above misses (some CUs ship CUMA-style data in a
# generically-named workbook).
_COMPOSITE_CODE_VALUE_RX = re.compile(
    r"^\s*\d{1,3}\s*/\s*\d{1,3}\s+[A-Za-z]",
)


def _looks_like_no_split_extract(
    name: str,
    sample_values: list[str] | None = None,
) -> bool:
    """Return True if this loan-extract candidate carries composite codes
    where ``/`` is part of the label (CUMA mortgage convention).

    Two independent triggers — either flips it to no-split:

    * Filename matches CUMA/MTG/Mortgage/Servicing regex, OR
    * ANY of the supplied raw pool-code values matches the
      ``\\d+/\\d+\\s+[A-Z]`` composite-code shape.
    """
    if name and _NO_SPLIT_FILENAME_RX.search(str(name)):
        return True
    if sample_values:
        for val in sample_values:
            if val and _COMPOSITE_CODE_VALUE_RX.match(str(val)):
                return True
    return False

# A consolidated "Balance Sheets <CU>" pool-grouping workbook (Vizo
# Financial analyst convention). Filename starts with "Balance Sheets"
# (or "Balance Sheet") followed by a NON-year token (the CU's short
# name). Structure: col A holds NCUA loan codes ("NNN-NN") on sub-cat
# rows and an integer on pool-header rows; col B holds labels with the
# pool-header rows bolded. See ``extract_pools_from_balance_workbook``.
_CONSOLIDATED_POOL_BAL_RX = re.compile(
    r"^balance\s*sheets?\s+(?!20\d{2}\b)\S",
    re.IGNORECASE,
)

# Sub-category code in col A (e.g. "708-01", "702-03"). Used by
# ``extract_pools_from_balance_workbook``.
_SUBCAT_CODE_RX = re.compile(r"^\d{3}-\d{2}$")

# Grand-total row in col B (terminates the pool walk).
_GRAND_TOTAL_RX = re.compile(r"^\s*total\s+loans\s*$", re.IGNORECASE)


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


# Fallback period extractor: warm_parser._period_from_filename only
# handles YYYY-MM and month-name+year; CUMA MTG Servicing reports use
# MM-DD-YYYY (e.g. "CUMA MTG Servicing Report 12-31-2025 V2.xls"). This
# helper layers on the missing shapes so ``_pick_latest_for_period``
# can still respect snapshot for those files.
_MD_YYYY_RX = re.compile(r"(?<!\d)(\d{1,2})[-_/](\d{1,2})[-_/](20\d{2})(?!\d)")
_MM_YYYY_RX = re.compile(r"(?<!\d)(\d{1,2})[-_/](20\d{2})(?!\d)")


def _fallback_period_from_filename(name: str) -> str | None:
    """Return ``YYYY-MM`` from MM-DD-YYYY / MM-YYYY shapes."""
    m = _MD_YYYY_RX.search(name)
    if m:
        mo = int(m.group(1))
        if 1 <= mo <= 12:
            return f"{m.group(3)}-{mo:02d}"
    m = _MM_YYYY_RX.search(name)
    if m:
        mo = int(m.group(1))
        if 1 <= mo <= 12:
            return f"{m.group(2)}-{mo:02d}"
    return None


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
    if _SINGLE_HIST_BAL_RX.search(name):
        return "single_hist_bal_files"
    if _CONSOLIDATED_POOL_BAL_RX.search(name):
        return "consolidated_pool_balance_files"
    if _5300_RX.search(name):
        return "five_thirtythousand_files"
    if _looks_like_loan_extract(name):
        # Promote to loan_data_files. period_from_filename was already
        # set by warm_parser when it had a chance; re-derive here for
        # safety.
        if not entry.get("period"):
            entry["period"] = warm_parser._period_from_filename(name)  # noqa: SLF001
        # warm_parser._period_from_filename only handles YYYY-MM and
        # month-name+year shapes. CUMA MTG Servicing reports are named
        # like "CUMA MTG Servicing Report 12-31-2025 V2.xls" (MM-DD-YYYY)
        # and a few other CU exports use MM/YYYY or M-YYYY. Layer those
        # on as a fallback so multi-extract pickup respects snapshot.
        if not entry.get("period"):
            entry["period"] = _fallback_period_from_filename(name)
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
    base.setdefault("single_hist_bal_files", [])
    base.setdefault("five_thirtythousand_files", [])
    base.setdefault("consolidated_pool_balance_files", [])

    # warm_parser._LOAN_DATA_RX matches the bare "aires" token, so it
    # promotes BOTH "Aires Loans <date>.xlsx" (the loan extract this
    # tool consumes) AND "Aires Shares <date>.xlsx" (the member
    # share/deposit export — NOT loan data). Demote any "Aires Shares"
    # entry from ``loan_data_files`` to ``other_files`` so Step 13 and
    # the sample picker only see real loan extracts. The same overlay
    # also runs in ``_bucket_extra`` via ``_LOAN_FALSE_POSITIVE_RX`` so
    # files routed through ``other_files`` first never come back.
    _ldf = base.get("loan_data_files") or []
    if _ldf:
        _kept: list[dict[str, Any]] = []
        _demoted: list[dict[str, Any]] = []
        for _e in _ldf:
            if _AIRES_SHARES_RX.search(_e.get("name") or ""):
                _demoted.append(_e)
            else:
                _kept.append(_e)
        base["loan_data_files"] = _kept
        if _demoted:
            base.setdefault("other_files", []).extend(_demoted)

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

    # warm_parser._CO_FILE_RX is strict ("charge off" / "co hist"). Many
    # CUs (Destinations, Shuford-style) name their combined CO+Recovery
    # files "Recoveries - Charge-offs", "Recoveries and Charged off",
    # "Recoveries-Chg-offs" — none of which start with the CO keyword and
    # so end up exclusively in recov_files (or vice-versa). Cross-list any
    # entry whose filename matches the BROADER pattern so Step 5 and
    # Step 6 both see combined files.
    _broad_co_rx = re.compile(
        r"(charge[ds]?[\s_\-]*off|chg[\s_\-]*off|co[\s_\-]*hist|charge_off_track)",
        re.IGNORECASE,
    )
    _broad_recov_rx = re.compile(
        r"(recov(?:er(?:y|ies))?|historical[\s_\-]*recov)",
        re.IGNORECASE,
    )

    def _names_in(bucket: list[dict[str, Any]]) -> set[str]:
        return {(e.get("name") or "").lower() for e in bucket}

    co_list = base.setdefault("co_files", [])
    rc_list = base.setdefault("recov_files", [])
    co_names = _names_in(co_list)
    rc_names = _names_in(rc_list)

    for entry in list(rc_list):
        nm = (entry.get("name") or "")
        if _broad_co_rx.search(nm) and nm.lower() not in co_names:
            co_list.append(entry)
            co_names.add(nm.lower())
    for entry in list(co_list):
        nm = (entry.get("name") or "")
        if _broad_recov_rx.search(nm) and nm.lower() not in rc_names:
            rc_list.append(entry)
            rc_names.add(nm.lower())

    return base


# --------------------------------------------------------------------------
# Consolidated "Balance Sheets <CU>" pool extractor
# --------------------------------------------------------------------------


def _cell_text(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _is_bold(cell: Any) -> bool:
    try:
        return bool(cell.font and cell.font.bold)
    except Exception:
        return False


def extract_pools_from_balance_workbook(
    xlsx_path: str | Path,
) -> dict[str, Any]:
    """Walk a consolidated 'Balance Sheets <CU>' workbook.

    Convention learned from Destinations CU 66333 (Vizo Financial client):

      * Col A holds either an INTEGER (pool number, 1/4/5/...), a
        ``"NNN-NN"`` NCUA loan code (sub-category), or is empty.
      * Col B holds the label (pool name OR sub-category name).
      * **BOLD col B with INTEGER col A** => POOL row.
      * **BOLD col B with EMPTY col A** and leading whitespace in label
        => CONTINUATION of the previous pool's name (some pool names
        span two rows, e.g. ``"Total Loans/Lines of Credit Secured" +
        "   by 1st Lien 1-4 Family Residential Properties"``).
      * Sub-category rows (``NNN-NN`` col A) flow into the NEXT pool row.
      * ``"Total Loans"`` row (col B only, no children) is the grand
        total: terminates the walk.

    Returns::

        {
            "ok": True,
            "sheet": str,
            "pools": [
                {"name": str,
                 "subcategories": [{"code": "NNN-NN", "label": str}, ...]},
                ...
            ],
            "subcategory_to_pool": {label: pool_name, ...},
            "code_to_pool":        {code:  pool_name, ...},
        }
    """
    try:
        import openpyxl  # local import — heavy
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "error": f"openpyxl not available: {exc}"}

    path = Path(xlsx_path)
    if not path.exists():
        return {"ok": False, "error": f"file not found: {path}"}
    try:
        wb = openpyxl.load_workbook(str(path), data_only=True)
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "error": f"openpyxl load failed: {exc}"}

    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    pools: list[dict[str, Any]] = []
    subcat_buf: list[dict[str, str]] = []
    pending_pool_name: str | None = None

    max_row = ws.max_row or 0
    for r in range(1, max_row + 1):
        a_cell = ws.cell(row=r, column=1)
        b_cell = ws.cell(row=r, column=2)
        a_raw = a_cell.value
        b_raw = b_cell.value

        a_txt = _cell_text(a_raw)
        b_txt_raw = "" if b_raw is None else str(b_raw)
        b_txt = b_txt_raw.strip()

        if not a_txt and not b_txt:
            continue

        # Grand-total row -> stop processing.
        if not a_txt and _GRAND_TOTAL_RX.match(b_txt):
            break

        # Sub-category row (NCUA loan code in col A).
        if _SUBCAT_CODE_RX.match(a_txt):
            if b_txt:
                subcat_buf.append({"code": a_txt, "label": b_txt})
            continue

        # Pool row: BOLD col B, with integer in col A OR (empty col A
        # AND leading whitespace in label = continuation of prior pool).
        if _is_bold(b_cell) and b_txt:
            is_continuation = (
                not a_txt
                and (b_txt_raw != b_txt_raw.lstrip())
            )
            is_integer_a = False
            if a_txt:
                try:
                    int(float(a_txt))
                    is_integer_a = True
                except (ValueError, TypeError):
                    is_integer_a = False

            if is_continuation and pending_pool_name and pools:
                pending_pool_name = f"{pending_pool_name} {b_txt}".strip()
                pools[-1]["name"] = pending_pool_name
                continue

            if is_integer_a:
                pool = {
                    "name": b_txt,
                    "subcategories": list(subcat_buf),
                }
                pools.append(pool)
                pending_pool_name = b_txt
                subcat_buf = []
                continue
        # Anything else: ignore (header rows, blank labels, formulas).

    # Build flat lookup dicts.
    subcat_to_pool: dict[str, str] = {}
    code_to_pool: dict[str, str] = {}
    for p in pools:
        for sc in p.get("subcategories") or []:
            subcat_to_pool[sc["label"]] = p["name"]
            code_to_pool[sc["code"]] = p["name"]

    return {
        "ok": True,
        "sheet": sheet_name,
        "pools": pools,
        "subcategory_to_pool": subcat_to_pool,
        "code_to_pool": code_to_pool,
    }


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
        "sample_extracts": [],
        "annual_bal": None,
        "monthly_bal": None,
        "single_hist_bal": None,
        "warm": None,
        "impaired_file": None,
        "credit_pull_file": None,
        "co_file": None,
        "recov_file": None,
        "pool_seed": None,
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
    sample_candidates: list[dict[str, Any]] = []
    if sample_entry:
        sample_candidates.append(sample_entry)
        # If multiple files share the snapshot period, queue them so
        # multi-extract CUs (e.g. AIRES + CUMA) stage all their
        # snapshot extracts on Step 13 and the best-scoring one
        # still drives the top-level state.
        if snapshot_yyyymm:
            for e in loan_files:
                if e is sample_entry:
                    continue
                if _safe_period(e) == snapshot_yyyymm:
                    sample_candidates.append(e)
        else:
            # No snapshot pinned -- group remaining files by file_pattern
            # shape and add the latest from each shape distinct from the
            # primary so multi-extract CUs without an anchor still stage
            # all their unique extracts.
            try:
                primary_pat = sample_parser.guess_filename_patterns(
                    sample_entry["name"]
                ).get("file_pattern", "") or ""
            except Exception:  # noqa: BLE001
                primary_pat = ""
            seen_patterns: set[str] = {primary_pat} if primary_pat else set()
            remaining = sorted(
                (e for e in loan_files if e is not sample_entry),
                key=lambda e: _safe_period(e) or "0000-00",
                reverse=True,
            )
            for e in remaining:
                try:
                    pat = sample_parser.guess_filename_patterns(
                        e["name"]
                    ).get("file_pattern", "") or ""
                except Exception:  # noqa: BLE001
                    pat = ""
                if pat and pat not in seen_patterns:
                    seen_patterns.add(pat)
                    sample_candidates.append(e)

    best_analysis: dict[str, Any] | None = None
    best_entry: dict[str, Any] | None = None
    best_score = -1
    extracts: list[dict[str, Any]] = []
    for cand in sample_candidates:
        src = Path(cand["path"])
        try:
            tmp = _copy_into_tmp(src, "sample")
        except Exception as exc:  # noqa: BLE001
            out["errors"].append(f"copy failed for {cand['name']}: {exc}")
            continue
        # Phase 9.26: detect "no-split" CUMA/mortgage extracts BEFORE
        # analyse_sample_file runs, so the seeded pool_code_suggestions
        # carry the full composite labels (e.g. ``15/15 ARM``) instead of
        # being silently truncated to ``"15"`` by the default ``"/"`` split.
        # First-pass uses filename-only; if that misses we re-check the
        # actual sample values after analysis and re-run if needed.
        no_split = _looks_like_no_split_extract(cand["name"])
        try:
            analysis = sample_parser.analyse_sample_file(
                str(tmp), cand["name"],
                split_char=("" if no_split else "/"),
            )
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
        # Second-pass detection on raw sample-rows values (catches CUs
        # that ship CUMA-style codes in generically-named workbooks).
        if not no_split:
            pool_col = (analysis.get("column_suggestions") or {}).get(
                "loan_pool_code"
            )
            sample_rows = analysis.get("sample_rows") or []
            if pool_col and sample_rows:
                raw_vals = []
                for row in sample_rows:
                    if isinstance(row, dict):
                        v = row.get(pool_col)
                        if v is not None:
                            raw_vals.append(str(v))
                if _looks_like_no_split_extract(cand["name"], raw_vals):
                    no_split = True
                    try:
                        analysis = sample_parser.analyse_sample_file(
                            str(tmp), cand["name"], split_char="",
                        )
                    except Exception:  # noqa: BLE001
                        pass
        analysis["saved_path"] = str(tmp)
        ex_entry = {
            "name": cand["name"],
            "saved_path": str(tmp),
            "analysis": analysis,
        }
        if no_split:
            ex_entry["pool_code_split"] = ""
        extracts.append(ex_entry)
        # Score = #column mappings + #pool codes. Higher = better
        # foothold for the wizard.
        score = (
            len(analysis.get("column_suggestions") or {})
            + len(analysis.get("pool_code_suggestions") or [])
        )
        if score > best_score:
            best_score = score
            best_analysis = analysis
            best_entry = cand

    if best_analysis and best_entry:
        out["sample"] = best_analysis
        out["messages"].append(
            f"Analysed sample loan extract: {best_entry['name']} "
            f"({len(best_analysis.get('headers') or [])} columns, "
            f"{len(best_analysis.get('pool_code_suggestions') or [])} pool codes)"
        )

    out["sample_extracts"] = extracts
    if len(extracts) > 1:
        names = ", ".join(e["name"] for e in extracts)
        out["messages"].append(
            f"Staged {len(extracts)} loan-extract sample(s) for Step 13: {names}"
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

    # ----- Consolidated pool-grouping balance workbook ------------------
    # (Vizo Financial convention: ``Balance Sheets <CU>.xlsx`` carries the
    # authoritative pool list AND the sub-category -> pool mapping.)
    cpb_entries = cls.get("consolidated_pool_balance_files") or []
    if cpb_entries:
        recent = cpb_entries[0]
        try:
            seed = extract_pools_from_balance_workbook(recent["path"])
        except Exception as exc:  # noqa: BLE001
            out["errors"].append(
                f"extract_pools_from_balance_workbook failed on "
                f"{recent['name']}: {exc}"
            )
        else:
            if seed.get("ok") and seed.get("pools"):
                seed["source_file"] = recent.get("name")
                seed["source_path"] = recent.get("path")
                out["pool_seed"] = seed
                out["messages"].append(
                    f"Extracted {len(seed['pools'])} pool(s) and "
                    f"{len(seed.get('subcategory_to_pool') or {})} "
                    f"sub-category mapping(s) from {recent['name']}"
                )
            elif not seed.get("ok"):
                out["errors"].append(
                    f"pool extraction returned not-ok for "
                    f"{recent['name']}: {seed.get('error')}"
                )

    # ----- Single consolidated historical balance workbook --------------
    # (Used when neither annual nor monthly-detail workbooks were found.
    # File names like "Historical Loan Balances.xlsx" or
    # "Balance_History_-_<CU>.xlsx" -- a single wide-format workbook with
    # one column per month-end.)
    if (
        not out["annual_bal"]
        and not out["monthly_bal"]
        and cls.get("single_hist_bal_files")
    ):
        shb_entries = cls.get("single_hist_bal_files") or []
        # Prefer the largest/newest file if multiple exist.
        recent = shb_entries[0]
        try:
            an = monthly_bal_parser.analyse_file(recent["path"])
        except Exception as exc:  # noqa: BLE001
            out["errors"].append(
                f"monthly_bal_parser.analyse_file failed on "
                f"{recent['name']}: {exc}"
            )
        else:
            if an.get("ok"):
                out["single_hist_bal"] = {
                    "entry": recent,
                    "sheet": an.get("sheet"),
                    "header_row": an.get("header_row"),
                    "pool_name_col": an.get("pool_name_col"),
                    "pool_labels": an.get("parsed_pool_labels") or [],
                    "dates": an.get("dates") or [],
                }
                out["messages"].append(
                    f"Analysed single historical balance workbook "
                    f"{recent['name']}: sheet={an.get('sheet')}, "
                    f"{len(an.get('parsed_pool_labels') or [])} pool labels, "
                    f"{len(an.get('dates') or [])} date columns."
                )
            else:
                out["errors"].append(
                    f"analyse_file failed on {recent['name']}: "
                    f"{an.get('error')}"
                )

    # ----- Fallback: consolidated pool-grouping workbook as hist source -
    # Many CUs (especially Vizo Financial clients) ship a single
    # ``Balance Sheets <CU>.xlsx`` workbook that is BOTH the pool-seed
    # AND the historical balance time-series (loan codes in col A,
    # category names in col B, one column per month-end in cols C+).
    # When no annual / per-month / single-hist workbook was found but
    # we DID find a consolidated pool workbook, reuse it as the single
    # historical balance source.
    if (
        not out["annual_bal"]
        and not out["monthly_bal"]
        and not out["single_hist_bal"]
        and cls.get("consolidated_pool_balance_files")
    ):
        cpb = (cls.get("consolidated_pool_balance_files") or [])[0]
        try:
            an = monthly_bal_parser.analyse_file(cpb["path"])
        except Exception as exc:  # noqa: BLE001
            out["errors"].append(
                f"monthly_bal_parser.analyse_file (consolidated fallback) "
                f"failed on {cpb['name']}: {exc}"
            )
        else:
            if an.get("ok") and (an.get("dates") or []):
                # ``analyse_file`` always picks col A as the label column,
                # but the Vizo consolidated workbooks put loan codes in
                # col A and descriptive labels in col B. Pick whichever
                # column overlaps best with the pool-seed sub-category
                # mapping we extracted from the same workbook.
                pool_name_col = an.get("pool_name_col") or "A"
                seed = out.get("pool_seed") or {}
                sub_to_pool = (seed.get("subcategory_to_pool") or {})
                if sub_to_pool:
                    sub_keys = {
                        str(k).strip().casefold()
                        for k in sub_to_pool
                        if str(k).strip()
                    }
                    try:
                        from openpyxl import load_workbook as _lw
                        _wb = _lw(cpb["path"], read_only=True, data_only=True)
                        try:
                            _ws = _wb[an.get("sheet")]
                            best_col = pool_name_col
                            best_hits = -1
                            for col_letter in ("A", "B", "C"):
                                col_idx = ord(col_letter) - ord("A") + 1
                                hits = 0
                                for row in _ws.iter_rows(
                                    min_row=int(an.get("header_row") or 1) + 1,
                                    max_row=int(an.get("header_row") or 1) + 80,
                                    min_col=col_idx,
                                    max_col=col_idx,
                                    values_only=True,
                                ):
                                    val = row[0]
                                    if val is None:
                                        continue
                                    if (
                                        str(val).strip().casefold()
                                        in sub_keys
                                    ):
                                        hits += 1
                                if hits > best_hits:
                                    best_hits = hits
                                    best_col = col_letter
                            if best_hits > 0:
                                pool_name_col = best_col
                        finally:
                            _wb.close()
                    except Exception:  # noqa: BLE001
                        pass
                # Re-read parsed pool labels from the elected column so
                # the wizard's Step 4 displays the descriptive names that
                # match the user's pool map (rather than loan codes).
                pool_labels = an.get("parsed_pool_labels") or []
                if pool_name_col != (an.get("pool_name_col") or "A"):
                    try:
                        from openpyxl import load_workbook as _lw
                        _wb = _lw(cpb["path"], read_only=True, data_only=True)
                        try:
                            _ws = _wb[an.get("sheet")]
                            col_idx = (
                                ord(pool_name_col.upper()) - ord("A") + 1
                            )
                            new_labels = []
                            for row in _ws.iter_rows(
                                min_row=int(an.get("header_row") or 1) + 1,
                                max_row=int(an.get("header_row") or 1) + 200,
                                min_col=col_idx,
                                max_col=col_idx,
                                values_only=True,
                            ):
                                val = row[0]
                                if val is None:
                                    continue
                                txt = str(val).strip()
                                if txt and txt not in new_labels:
                                    new_labels.append(txt)
                            if new_labels:
                                pool_labels = new_labels
                        finally:
                            _wb.close()
                    except Exception:  # noqa: BLE001
                        pass
                out["single_hist_bal"] = {
                    "entry": cpb,
                    "sheet": an.get("sheet"),
                    "header_row": an.get("header_row"),
                    "pool_name_col": pool_name_col,
                    "pool_labels": pool_labels,
                    "dates": an.get("dates") or [],
                }
                out["messages"].append(
                    f"Using consolidated balance workbook {cpb['name']!r} "
                    f"as single historical balance source "
                    f"(sheet={an.get('sheet')}, pool col={pool_name_col}, "
                    f"{len(pool_labels)} labels, "
                    f"{len(an.get('dates') or [])} date columns)."
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


def _apply_single_hist_bal_to_state(
    state: dict[str, Any], shb: dict[str, Any]
) -> list[str]:
    """Mutate state.monthly_bal with single-historical-workbook findings.

    Sets the ``single_workbook`` layout fields (saved_path / sheet /
    header_row / pool_name_col) so the runtime balance-preview helper +
    Step 4's green confidence card can render immediately without the
    user re-uploading the file or re-saving the layout.
    """
    msgs: list[str] = []
    mb = state.setdefault("monthly_bal", {})
    mb["source"] = "single"
    entry = shb.get("entry") or {}
    saved_path = entry.get("path") or ""
    mb["saved_path"] = saved_path
    mb["filename"] = entry.get("name") or ""
    mb["sheet"] = shb.get("sheet") or ""
    mb["header_row"] = shb.get("header_row") or 0
    mb["pool_name_col"] = shb.get("pool_name_col") or ""
    mb["parsed_pool_labels"] = shb.get("pool_labels") or []
    msgs.append(
        f"monthly_bal source=single with workbook {entry.get('name')!r} "
        f"(sheet={shb.get('sheet')!r}, header row {shb.get('header_row')}, "
        f"pool label col {shb.get('pool_name_col')!r}, "
        f"{len(mb['parsed_pool_labels'])} pool labels, "
        f"{len(shb.get('dates') or [])} date columns)"
    )
    return msgs


# Pool-name keywords that imply a long life-of-loan (mortgages, HE, real
# estate). Used to pick a sensible ACL-months default when seeding pools
# from a 'Balance Sheets <CU>' workbook. Everything else defaults to 36.
_LONG_LOL_HINTS = re.compile(
    r"mortgage|first\s*lien|junior\s*lien|home\s*equity|heloc|"
    r"real\s*estate|residential|mobile\s*home|land\s*loan|"
    r"business\s*loan|commercial",
    re.IGNORECASE,
)

# Aggregate / total / grand-total row labels that should be excluded
# when seeding pool_settings from a flat hist-balance workbook (col A
# typically ends with a "Total" row that sums the pools above it).
_AGG_LABEL_RX = re.compile(
    r"^\s*(?:grand[\s_\-]*)?total(?:\s+loans?)?\s*$|"
    r"^\s*sub[\s_\-]*total\s*$|"
    r"^\s*all[\s_\-]*loans?\s*$",
    re.IGNORECASE,
)


def _default_acl_months_for_pool(name: str) -> int:
    """Sensible ACL-months default based on pool-name keywords."""
    if not name:
        return 36
    return 84 if _LONG_LOL_HINTS.search(name) else 36


def _purge_stale_pool_map_values(state: dict[str, Any]) -> int:
    """Rewrite any ``state["pool_map"]`` value that doesn't match a
    current ``pool_settings`` pool name (or the sentinels ``Ignore`` /
    ``Exclude``) to ``"Ignore"``.

    Defends against cross-CU pool-map pollution: pool_map entries from a
    previous CU's pool-map upload survive in the active Flask session
    after the user starts a new CU (no code path clears them on CU
    change). When auto-scan replaces ``pool_settings`` with the new
    CU's pools, leftover values become "unrecognized pool names" on
    Step 3. Run this AFTER ``pool_settings`` has been populated.

    Returns the number of entries rewritten.
    """
    pm = state.get("pool_map")
    if not isinstance(pm, dict) or not pm:
        return 0
    valid_names: set[str] = {
        (p.get("name") or "").strip()
        for p in (state.get("pool_settings") or [])
        if isinstance(p, dict) and p.get("name")
    }
    if not valid_names:
        # No pool_settings yet -> can't validate; do nothing.
        return 0
    valid_names |= {"Ignore", "Exclude"}
    purged = 0
    for code, value in list(pm.items()):
        if not value:
            continue
        if str(value).strip() in valid_names:
            continue
        pm[code] = "Ignore"
        purged += 1
    return purged


def _apply_flat_pool_labels_to_state(
    state: dict[str, Any],
    labels: list[str],
    src: str,
) -> dict[str, list[str]]:
    """Seed Step 2 (pool_settings) + Step 8 (monthly_bal.pool_map) from a
    FLAT historical-balance workbook (one pool per row in col A; no
    sub-category hierarchy).

    Mirrors ``_apply_pool_seed_to_state`` but for the flat layout: each
    label IS a canonical pool name, so ``pool_map`` is seeded with each
    label mapping to itself. Aggregate rows ("Total", "Grand Total",
    "Sub-total", "All Loans") are filtered out.

    Never overwrites a user's existing entries: ``pool_settings`` is
    only populated when empty; ``monthly_bal.pool_map`` entries are
    only added for labels not yet keyed in the map.
    """
    out: dict[str, list[str]] = {}
    cleaned: list[str] = []
    seen_norm: set[str] = set()
    for raw in labels or []:
        nm = (raw or "").strip()
        if not nm:
            continue
        if _AGG_LABEL_RX.match(nm):
            continue
        key = nm.lower()
        if key in seen_norm:
            continue
        seen_norm.add(key)
        cleaned.append(nm)
    if not cleaned:
        return out

    # ---- state["pool_settings"] (Step 2 "Loan Pools" table) ------------
    if not state.get("pool_settings"):
        ps_rows: list[dict[str, Any]] = []
        for nm in cleaned:
            ps_rows.append({
                "name": nm,
                "risk_rated": False,
                "brr": False,
                "acl_months": _default_acl_months_for_pool(nm),
                "use_default_mgmt_adj": False,
                "excluded": False,
            })
        state["pool_settings"] = ps_rows
        out.setdefault("loan_pools", []).append(
            f"pool_settings seeded with {len(ps_rows)} pool(s) from {src} "
            f"(review the risk-rated / ACL-months / excluded toggles)"
        )

    # Mirror to state.warm.pools so downstream steps and the final YAML
    # see the same ordered list.
    warm = state.get("warm") or {}
    if not warm.get("pools"):
        warm["pools"] = [p["name"] for p in (state.get("pool_settings") or [])]
        state["warm"] = warm

    # ---- state["monthly_bal"]["pool_map"] (label -> itself) -----------
    mb = state.setdefault("monthly_bal", {})
    pm = mb.setdefault("pool_map", {})
    added = 0
    for nm in cleaned:
        if nm not in pm or not pm.get(nm):
            pm[nm] = nm
            added += 1
    if added:
        out.setdefault("monthly_bal", []).append(
            f"pool_map seeded with {added} label -> pool mapping(s) from "
            f"{src} (each label adopted as its own pool name)"
        )

    # Defensive: clear any pool_map values left over from a prior CU's
    # pool-map upload that don't match the freshly-seeded pool_settings.
    purged = _purge_stale_pool_map_values(state)
    if purged:
        out.setdefault("pools", []).append(
            f"Cleared {purged} stale pool_map value(s) (carried over from "
            f"a prior CU) — re-map on Step 3."
        )

    return out


def _apply_pool_seed_to_state(
    state: dict[str, Any], seed: dict[str, Any]
) -> dict[str, list[str]]:
    """Mutate state.pool_settings / state.warm.pools / state.monthly_bal.pool_map
    from a consolidated 'Balance Sheets <CU>' workbook extraction.

    Never overwrites a user's existing entries: ``pool_settings`` is only
    populated when empty; ``monthly_bal.pool_map`` entries are only
    added for sub-category labels that are not yet keyed in the map.

    Returns a ``{step_key: [messages]}`` dict ready to merge into the
    main report.
    """
    out: dict[str, list[str]] = {}
    pools = seed.get("pools") or []
    if not pools:
        return out
    src = seed.get("source_file") or "balance-sheet workbook"

    # ---- state["pool_settings"] (Step 2 "Loan Pools" table) ------------
    if not state.get("pool_settings"):
        ps_rows: list[dict[str, Any]] = []
        for p in pools:
            nm = (p.get("name") or "").strip()
            if not nm:
                continue
            ps_rows.append({
                "name": nm,
                "risk_rated": False,
                "brr": False,
                "acl_months": _default_acl_months_for_pool(nm),
                "use_default_mgmt_adj": False,
                "excluded": False,
            })
        state["pool_settings"] = ps_rows
        out.setdefault("loan_pools", []).append(
            f"pool_settings seeded with {len(ps_rows)} pool(s) from {src} "
            f"(review the risk-rated / ACL-months / excluded toggles)"
        )

    # Mirror to state.warm.pools so downstream steps and the final YAML
    # see the same ordered list (the no-WARM step2 handler does the same
    # sync on save — we do it eagerly here).
    warm = state.get("warm") or {}
    if not warm.get("pools"):
        warm["pools"] = [p["name"] for p in (state.get("pool_settings") or [])]
        state["warm"] = warm

    # ---- state["monthly_bal"]["pool_map"] (Step 8 sub-cat -> pool) ----
    mb = state.setdefault("monthly_bal", {})
    pm = mb.setdefault("pool_map", {})
    subcat_to_pool = seed.get("subcategory_to_pool") or {}
    added = 0
    for label, pool_name in subcat_to_pool.items():
        if not label:
            continue
        if label not in pm or not pm.get(label):
            pm[label] = pool_name
            added += 1
    if added:
        out.setdefault("monthly_bal", []).append(
            f"pool_map seeded with {added} sub-category mapping(s) from {src} "
            f"(parent pools applied automatically)"
        )
    # Mirror sub-cat labels into parsed_pool_labels so the Step 8 mapping
    # table renders the right row set when the auto-scan ran before the
    # user uploaded a balance file.
    parsed = mb.get("parsed_pool_labels") or []
    for label in subcat_to_pool:
        if label and label not in parsed:
            parsed.append(label)
    mb["parsed_pool_labels"] = parsed

    # Hierarchical 'Balance Sheets <CU>' workbooks have BOTH parent pool
    # rows (showing the sum) AND sub-category rows underneath them. The
    # monthly_bal_parser walks every label row, so without an exclusion
    # list the parent rows would double-count their own sub-categories
    # whenever the parent label is also keyed in the pool_map. Stash the
    # parent pool names here so balance_check / preview / 5300-backfill
    # can pass them as exclude_labels into the parser.
    excl_existing = mb.get("exclude_labels") or []
    excl_existing_norm = {
        (s or "").strip().lower() for s in excl_existing
        if (s or "").strip()
    }
    for p in pools:
        nm = (p.get("name") or "").strip()
        if not nm:
            continue
        if nm.lower() in excl_existing_norm:
            continue
        excl_existing.append(nm)
        excl_existing_norm.add(nm.lower())
    mb["exclude_labels"] = excl_existing

    # ---- Stash full seed on state for UI inspection -------------------
    state["_pool_seed"] = {
        "source_file": src,
        "pools": pools,
        "subcategory_to_pool": subcat_to_pool,
        "code_to_pool": seed.get("code_to_pool") or {},
    }

    # Defensive: clear any pool_map values left over from a prior CU's
    # pool-map upload that don't match the freshly-seeded pool_settings.
    purged = _purge_stale_pool_map_values(state)
    if purged:
        out.setdefault("pools", []).append(
            f"Cleared {purged} stale pool_map value(s) (carried over from "
            f"a prior CU) — re-map on Step 3."
        )

    return out


def apply_pool_grouping_choice(
    state: dict[str, Any],
    choices: dict[str, str],
) -> dict[str, Any]:
    """Rebuild ``state["pool_settings"]`` + ``state["warm"]["pools"]`` +
    ``state["monthly_bal"]["pool_map"]`` from ``state["_pool_seed"]``
    using the user's per-pool keep/split choices.

    Args:
        state:    Wizard state dict (mutated in place).
        choices:  ``{pool_name: "keep" | "split"}``. Pools absent from
                  this map default to ``"keep"``. Pools with no
                  sub-categories are always kept (split would be a
                  no-op).

    Semantics:
        * **keep**  — Parent pool stays a single ``pool_settings`` row;
          all of its sub-category labels point to it in
          ``monthly_bal.pool_map``.
        * **split** — Parent pool is REMOVED from ``pool_settings``;
          each sub-category becomes its OWN ``pool_settings`` row, and
          each sub-cat label points to itself in
          ``monthly_bal.pool_map``.

    User-edited per-pool settings (``risk_rated`` / ``brr`` /
    ``acl_months`` / ``excluded`` / ``use_default_mgmt_adj``) are
    preserved by NAME when the rebuilt list contains a pool with the
    same name as an existing row.

    Returns ``{"ok": bool, "pools_count": int, "mappings_count": int,
    "split_count": int, "kept_count": int}`` or
    ``{"ok": False, "error": str}``.
    """
    seed = state.get("_pool_seed") or {}
    pools_seed = seed.get("pools") or []
    if not pools_seed:
        return {"ok": False, "error": "No pool seed available on state."}

    # Preserve user-edited per-row settings by name.
    existing = {
        (p.get("name") or "").strip(): p
        for p in (state.get("pool_settings") or [])
        if (p.get("name") or "").strip()
    }

    def _row_for(name: str) -> dict[str, Any]:
        if name in existing:
            return dict(existing[name])
        return {
            "name": name,
            "risk_rated": False,
            "brr": False,
            "acl_months": _default_acl_months_for_pool(name),
            "use_default_mgmt_adj": False,
            "excluded": False,
        }

    new_settings: list[dict[str, Any]] = []
    managed_map: dict[str, str] = {}  # sub-cat label -> pool name
    split_count = 0
    kept_count = 0

    for pool in pools_seed:
        pname = (pool.get("name") or "").strip()
        if not pname:
            continue
        subcats = pool.get("subcategories") or []
        choice = (choices.get(pname) or "keep").strip().lower()
        # Split-with-no-subcats is meaningless — fall back to keep.
        if choice == "split" and subcats:
            for sc in subcats:
                label = (sc.get("label") or "").strip()
                if not label:
                    continue
                new_settings.append(_row_for(label))
                managed_map[label] = label
            split_count += 1
        else:
            new_settings.append(_row_for(pname))
            for sc in subcats:
                label = (sc.get("label") or "").strip()
                if label:
                    managed_map[label] = pname
            kept_count += 1

    state["pool_settings"] = new_settings
    warm = state.get("warm") or {}
    warm["pools"] = [p["name"] for p in new_settings]
    state["warm"] = warm

    mb = state.setdefault("monthly_bal", {})
    pm = mb.setdefault("pool_map", {})
    # Only overwrite entries that came from our seed — labels outside
    # the seed (e.g. user-added rows on Step 8) are left alone.
    for label, pool_name in managed_map.items():
        pm[label] = pool_name
    parsed = mb.get("parsed_pool_labels") or []
    for label in managed_map:
        if label not in parsed:
            parsed.append(label)
    mb["parsed_pool_labels"] = parsed

    state["_pool_grouping_choices"] = {
        (pool.get("name") or "").strip(): (
            "split"
            if (
                (choices.get(pool.get("name") or "") or "keep").lower()
                == "split"
                and (pool.get("subcategories") or [])
            )
            else "keep"
        )
        for pool in pools_seed
        if (pool.get("name") or "").strip()
    }

    return {
        "ok": True,
        "pools_count": len(new_settings),
        "mappings_count": len(managed_map),
        "split_count": split_count,
        "kept_count": kept_count,
    }


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


def _stage_hist_scan_files(
    state: dict[str, Any],
    entries: list[dict[str, Any]] | None,
    hist_scan_key: str,
    tmp_subfolder: str,
) -> list[str]:
    """Bulk-stage a list of classifier entries into ``state['hist_scan'][key]``.

    Each entry is shaped ``{name, period, path}`` (the warm_parser/classifier
    output). Files are copied into the auto-scan temp area so subsequent
    requests can read them, and registered under ``hist_scan[<hist_scan_key>]``
    in the same shape that ``routes/setup._add_hist_file_from_path`` writes.
    Existing ``hist_scan`` entries with the same name are NOT duplicated;
    user-uploaded files are preserved.

    Used so Step 5 (Historical Charge-Offs) and Step 6 (Historical Recoveries)
    pick up every monthly CO/Recov file from the folder scan, not just the
    single representative file kept in ``sample_uploads``.
    """
    if not entries:
        return []
    hs = state.get("hist_scan")
    if not isinstance(hs, dict):
        hs = {}
        state["hist_scan"] = hs
    # Match the keyset that `_add_hist_file*` writes so downstream code
    # never trips over a missing list.
    for k in (
        "warm_files", "co_files", "recov_files", "impaired_files",
        "credit_pull_files", "monthly_files", "loan_data_files", "other_files",
        "monthly_co_files", "monthly_recov_files",
    ):
        hs.setdefault(k, [])
    hs.setdefault("ok", True)
    hs.setdefault("folder", state.get("_auto_scan_folder") or "Auto-scanned folder")
    rows = list(hs.get(hist_scan_key) or [])
    existing_names = {e.get("name") for e in rows}
    staged = 0
    for entry in entries:
        try:
            src = Path(entry.get("path") or "")
        except (TypeError, ValueError):
            continue
        if not src.exists():
            continue
        if src.name in existing_names:
            continue
        tmp = _copy_into_tmp(src, tmp_subfolder)
        rows.append({
            "name": src.name,
            "period": entry.get("period"),
            "path": str(tmp),
        })
        existing_names.add(src.name)
        staged += 1
    hs[hist_scan_key] = rows
    if staged:
        return [f"{staged} {hist_scan_key.replace('_', ' ')} staged from folder scan"]
    return []


def _stage_loan_data_extracts(
    state: dict[str, Any],
    extracts: list[dict[str, Any]] | None,
) -> list[str]:
    """Bulk-stage analysed loan-data extracts into
    ``state['sample_uploads']['loan_data_files']`` so the Column Mappings
    step (``step3_columns``) renders dropdowns for each detected extract
    without the user having to re-upload them on Step 12.

    Each ``extracts`` entry is shaped ``{name, saved_path, analysis}``
    where ``analysis`` is the full ``sample_parser.analyse_sample_file``
    return dict. Builds Step-12-shaped entries (mirrors the upload
    handler in ``routes/setup.step2_sample``) and seeds per-file
    ``column_mappings`` from the file's own ``column_suggestions`` plus
    a ``member_account`` block from the parser's per-file inference.
    Skips entries whose name is already present in ``loan_data_files``
    so user-uploaded files are preserved on a re-scan.
    """
    if not extracts:
        return []
    su = state.setdefault("sample_uploads", {})
    bucket = su.setdefault("loan_data_files", [])
    existing_names = {e.get("name") for e in bucket}
    msgs: list[str] = []
    for ex in extracts:
        name = ex.get("name") or ""
        path = ex.get("saved_path") or ""
        analysis = ex.get("analysis") or {}
        if not name or not path or not analysis:
            continue
        if name in existing_names:
            msgs.append(f"loan extract already staged: {name}")
            continue
        headers = list(analysis.get("headers") or [])
        suggestions = dict(analysis.get("column_suggestions") or {})
        # Slim ``analysis`` matches the shape produced by
        # ``routes/setup._loan_data_entry_analysis`` so Step 13 renders
        # dropdowns immediately without lazy hydrate on first GET.
        slim = {
            "headers": headers,
            "column_suggestions": suggestions,
            "pool_code_suggestions": list(
                analysis.get("pool_code_suggestions") or []
            ),
            "sample_rows": list(analysis.get("sample_rows") or [])[:5],
            "has_header": bool(analysis.get("has_header")),
            "header_row": int(analysis.get("header_row") or 0),
            "member_account_suggestion": (
                dict(analysis.get("member_account_suggestion") or {})
                if analysis.get("member_account_suggestion") else None
            ),
            "date_format": str(analysis.get("date_format") or "YYYY-MM"),
        }
        # Per-file member_account: prefer parser inference, fall back
        # to the historical fixed_suffix=3 default. Strip metadata keys.
        ma_sug = analysis.get("member_account_suggestion") or {}
        if ma_sug.get("mode"):
            member_account = {
                "mode": ma_sug.get("mode"),
                "suffix_length": int(ma_sug.get("suffix_length") or 0),
                "delimiter": ma_sug.get("delimiter") or "-",
            }
        else:
            member_account = {
                "mode": "fixed_suffix", "suffix_length": 3, "delimiter": "-",
            }
        # Per-file column_mappings seeded from this file's own header
        # suggestions. Step 13 renders these as dropdown defaults; user
        # can override per-file. Only include suggestions whose value
        # is actually one of the file's headers (avoids ghost options).
        col_map = {
            field: header
            for field, header in suggestions.items()
            if header and header in headers
        }
        entry = {
            "name": name,
            "path": path,
            "has_header": bool(analysis.get("has_header")),
            "header_row": int(analysis.get("header_row") or 0),
            "analysis": slim,
            "column_mappings": col_map,
            "member_account": member_account,
            "file_pattern": analysis.get("file_pattern") or "",
        }
        # Phase 9.26: per-extract pool_code_split override (CUMA/mortgage
        # files where ``/`` is part of the label, not a prefix:desc split).
        if "pool_code_split" in ex:
            entry["pool_code_split"] = ex.get("pool_code_split") or ""
        bucket.append(entry)
        existing_names.add(name)
        msgs.append(
            f"loan extract staged: {name} "
            f"({len(headers)} cols, {len(col_map)} suggested mappings)"
        )
    return msgs


def selfheal_hist_scan_from_folder(state: dict[str, Any]) -> list[str]:
    """Proactively bulk-stage CO/Recov files into ``hist_scan`` for any draft
    whose Step-1 auto-scan completed but whose ``monthly_co_files`` /
    ``monthly_recov_files`` lists are still empty.

    Safe to call from any wizard GET hook. Returns the staging messages so
    callers can decide whether to flash them.

    Two reasons this self-heal is needed:

    1. Drafts saved BEFORE the Phase 9.4 bulk-stage code was wired into
       ``apply_findings_to_state`` only have the single representative
       file in ``sample_uploads``; the per-month list is empty.
    2. Combined CO+Recov files (Destinations-style "Recoveries -
       Charge-offs") are routed into BOTH ``classify_folder['co_files']``
       and ``classify_folder['recov_files']`` via the broadened cross-list
       overlay, but ``apply_findings_to_state`` ran on a draft before that
       overlay existed -- so only one bucket got bulk-staged.

    The old step3_historical lazy self-heal only fired when the user
    landed on Step 5 or Step 6, so the OTHER section's badge stayed
    stale until the user navigated there. Running this on every wizard
    GET (from ``_wizard_ctx``) makes the breadcrumb truthful across
    every step.
    """
    if not state.get("_auto_scan_completed"):
        return []
    folder = state.get("_auto_scan_folder") or ""
    hs = state.get("hist_scan") or {}
    need_co = not (hs.get("monthly_co_files") or hs.get("co_files"))
    need_recov = not (hs.get("monthly_recov_files") or hs.get("recov_files"))
    msgs: list[str] = []
    # If files are already staged but the source picker is still on its
    # bare default, just flip the picker -- no need to reclassify.
    if not (need_co or need_recov):
        msgs.extend(_auto_flip_hist_source(state))
        return msgs
    if not folder or not Path(folder).exists():
        # Folder unreachable: still try to flip the source picker for
        # the side that already has files staged.
        msgs.extend(_auto_flip_hist_source(state))
        return msgs
    try:
        cls = classify_folder(folder)
    except Exception:  # noqa: BLE001 -- never let self-heal block the page
        msgs.extend(_auto_flip_hist_source(state))
        return msgs
    if need_co:
        msgs.extend(_stage_hist_scan_files(
            state, cls.get("co_files") or [], "monthly_co_files", "co",
        ))
    if need_recov:
        msgs.extend(_stage_hist_scan_files(
            state, cls.get("recov_files") or [], "monthly_recov_files", "recov",
        ))
    # When per-month files are staged, flip the section's source picker
    # from the default "single_workbook" to "monthly_files" so the user
    # actually SEES those files on Step 5/6 instead of an empty single-
    # workbook upload field. Respects user choice: only flips when the
    # current value is still the bare default and no single workbook
    # has been uploaded.
    msgs.extend(_auto_flip_hist_source(state))
    return msgs


def selfheal_flat_pool_seed_from_state(state: dict[str, Any]) -> list[str]:
    """Proactively seed Step 2 ``pool_settings`` + Step 8
    ``monthly_bal.pool_map`` from a previously-detected flat
    ``Historical Balance Sheets`` workbook when those state keys are
    still empty.

    Targets drafts saved BEFORE the flat-pool-labels seeding wired into
    ``apply_findings_to_state`` for ``single_hist_bal`` findings. Reads
    only state (``monthly_bal.parsed_pool_labels`` /
    ``monthly_bal.filename``) — no folder I/O — so it's safe to call
    from any wizard GET hook.

    Skips silently when:
      * Auto-scan never completed
      * ``pool_settings`` already populated (preserves user edits)
      * The hist-balance source is not the flat single workbook shape
        (consolidated pool-seed workbooks have their own seeding path)
      * No pool labels were parsed
    """
    if not state.get("_auto_scan_completed"):
        return []
    if state.get("pool_settings"):
        return []
    mb = state.get("monthly_bal") or {}
    if (mb.get("source") or "") != "single":
        return []
    labels = mb.get("parsed_pool_labels") or []
    if not labels:
        return []
    src = mb.get("filename") or "single historical balance workbook"
    flat_report = _apply_flat_pool_labels_to_state(state, labels, src)
    msgs: list[str] = []
    for key in ("loan_pools", "monthly_bal"):
        for m in flat_report.get(key, []):
            msgs.append(m)
    return msgs


def _auto_flip_hist_source(state: dict[str, Any]) -> list[str]:
    """If monthly CO/Recov files have been staged but the user-facing
    source picker on Step 5/6 still says ``single_workbook`` (the bare
    default), flip it to ``monthly_files`` so the per-file table
    renders. Only fires when the corresponding ``co_files`` /
    ``recov_files`` (legacy single-workbook uploads) list is empty,
    so a user who deliberately uploaded a single workbook is preserved.
    Returns human-readable messages for the auto-scan report.
    """
    msgs: list[str] = []
    hs = state.get("hist_scan") or {}
    # CO side
    if (
        state.get("hist_co_source") in (None, "", "single_workbook")
        and (hs.get("monthly_co_files") or [])
        and not (hs.get("co_files") or [])
    ):
        if state.get("hist_co_source") != "monthly_files":
            state["hist_co_source"] = "monthly_files"
            msgs.append(
                "Charge-off source set to 'Monthly charge-off files' "
                "(per-month files detected by folder scan)."
            )
    # Recov side
    if (
        state.get("hist_recov_source") in (None, "", "single_workbook")
        and (hs.get("monthly_recov_files") or [])
        and not (hs.get("recov_files") or [])
    ):
        if state.get("hist_recov_source") != "monthly_files":
            state["hist_recov_source"] = "monthly_files"
            msgs.append(
                "Recovery source set to 'Monthly recovery files' "
                "(per-month files detected by folder scan)."
            )
    return msgs


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


# --------------------------------------------------------------------------
# 5300 distributed backfill (auto-run at end of scan)
# --------------------------------------------------------------------------


def _compute_earliest_month_pool_distribution(
    state: dict[str, Any],
) -> dict[str, Any]:
    """Service-side mirror of ``routes/setup._earliest_month_pool_distribution``.

    Only handles the two ``hist_balance_source`` values that the
    auto-scan can ever produce -- ``annual_balance_sheets`` and
    ``monthly_balance_sheets``. Single-workbook + monthly-loan-extract
    flavours are picked by the user manually on Step 4 and the
    auto-runner just skips them.

    Returns a dict with keys ``ok``, ``error``, ``period``, ``total``,
    ``pool_distribution`` (ratios that sum to 1.0), ``by_pool`` (raw $
    per pool at the earliest month), and ``source``.
    """
    out: dict[str, Any] = {
        "ok": False, "error": "", "period": "", "total": 0.0,
        "pool_distribution": {}, "by_pool": {}, "source": "",
    }
    mb_state = state.get("monthly_bal") or {}
    label_to_pool = mb_state.get("pool_map") or {}
    source = state.get("hist_balance_source") or ""
    by_period: dict[str, Any] = {}
    if source == "annual_balance_sheets":
        out["source"] = "annual"
        ann = monthly_bal_parser.pool_balances_for_per_year_files(
            mb_state.get("year_files") or [],
            mb_state.get("per_year_layout") or {},
            label_to_pool,
        )
        if not ann.get("ok") and not (ann.get("by_period") or {}):
            out["error"] = (
                f"Could not read annual balance workbook(s): "
                f"{ann.get('error') or 'unknown error'}"
            )
            return out
        by_period = ann.get("by_period") or {}
    elif source == "monthly_balance_sheets":
        out["source"] = "per_month"
        pm = monthly_bal_parser.pool_balances_for_per_month_files(
            mb_state.get("monthly_files") or [],
            mb_state.get("per_month_layout") or {},
            label_to_pool,
        )
        if not pm.get("ok") and not (pm.get("by_period") or {}):
            out["error"] = (
                f"Could not read monthly balance file(s): "
                f"{pm.get('error') or 'unknown error'}"
            )
            return out
        by_period = pm.get("by_period") or {}
    else:
        out["error"] = (
            f"Auto-run only supports annual / monthly balance sheets "
            f"(hist_balance_source={source!r})"
        )
        return out
    if not by_period:
        out["error"] = "No periods extracted from balance file(s)."
        return out
    earliest = sorted(by_period.keys())[0]
    by_pool = by_period[earliest] or {}
    total = sum(float(v or 0) for v in by_pool.values())
    if total <= 0:
        out["error"] = (
            f"Earliest month ({earliest}) has zero total balance "
            f"-- check pool_map labels."
        )
        return out
    out["ok"] = True
    out["period"] = earliest
    out["total"] = total
    out["by_pool"] = dict(by_pool)
    out["pool_distribution"] = {
        k: float(v or 0) / total
        for k, v in by_pool.items()
        if (v or 0) > 0
    }
    return out


def _auto_run_distributed_backfill(
    state: dict[str, Any],
) -> list[str]:
    """If all preconditions are met, run the distributed 5300 backfill
    and return summary messages. Stores result on
    ``state['hist_extracts']['solr_backfill']['last_run']``. Never
    raises -- any failure is captured in the returned messages and on
    ``last_run`` so a scan can never 500 because of this auto-run.

    Preconditions:
      * ``credit_union`` and a numeric ``charter_number`` on Identity
      * ``report_period`` (used as ``target_period``)
      * ``hist_balance_source`` is annual_balance_sheets or
        monthly_balance_sheets
      * ``monthly_bal.pool_map`` has at least one non-Ignore mapping
        so the earliest-month distribution is non-empty
    """
    msgs: list[str] = []
    # Lazy import keeps the rest of auto_setup decoupled from the DB +
    # HTTP transport that the backfill drags in.
    try:
        from cecl_ui.services import solr_5300_backfill  # noqa: F401
    except Exception as exc:  # noqa: BLE001
        return [f"5300 backfill skipped: import failed ({exc})"]

    cu = (state.get("credit_union") or "").strip()
    charter_raw = (state.get("charter_number") or "").strip()
    try:
        charter_int = (
            int(re.sub(r"\D", "", charter_raw)) if charter_raw else 0
        )
    except (TypeError, ValueError):
        charter_int = 0
    report_period = (state.get("report_period") or "").strip()
    source = (state.get("hist_balance_source") or "").strip()

    # Initialise hist_extracts so the result has a place to live and
    # the user lands on Step 4 with the mode already set to distributed.
    he = state.get("hist_extracts")
    if not isinstance(he, dict):
        he = {}
        state["hist_extracts"] = he
    he.setdefault("target_period", report_period)
    he.setdefault("history_months", 84)
    sb = he.get("solr_backfill")
    if not isinstance(sb, dict):
        sb = {}
        he["solr_backfill"] = sb
    sb.setdefault("solr_url", "http://searchserver1.tctrisk.com:8983/solr")
    sb.setdefault("core", "ncua")
    sb["mode"] = "distributed"
    # If a prior wizard pass left target_period blank, fill from the
    # Identity step's report_period.
    if not (he.get("target_period") or "").strip():
        he["target_period"] = report_period
    target_period = (he.get("target_period") or "").strip()
    months = int(he.get("history_months") or 84)

    missing: list[str] = []
    if not cu:
        missing.append("credit_union")
    if not charter_int:
        missing.append("charter_number")
    if not target_period:
        missing.append("report_period")
    if source not in ("annual_balance_sheets", "monthly_balance_sheets"):
        missing.append(f"hist_balance_source ({source!r})")
    if missing:
        return [
            "5300 backfill (distributed) skipped at auto-scan time -- "
            "still needs: " + ", ".join(missing)
        ]

    try:
        dist = _compute_earliest_month_pool_distribution(state)
    except Exception as exc:  # noqa: BLE001
        return [f"5300 backfill skipped: distribution error ({exc})"]
    if not dist.get("ok"):
        return [
            "5300 backfill (distributed) skipped: "
            + (dist.get("error") or "no pool distribution available")
        ]

    # Quarter-ends already covered by the user's uploaded balance file(s)
    # are passed as ``existing_dates`` so the backfill doesn't overwrite
    # them.
    upload_months: set[str] = set()
    try:
        mb_state = state.get("monthly_bal") or {}
        label_to_pool = mb_state.get("pool_map") or {}
        if source == "annual_balance_sheets":
            ann = monthly_bal_parser.pool_balances_for_per_year_files(
                mb_state.get("year_files") or [],
                mb_state.get("per_year_layout") or {},
                label_to_pool,
            )
            upload_months = set((ann.get("by_period") or {}).keys())
        else:  # monthly_balance_sheets
            pm = monthly_bal_parser.pool_balances_for_per_month_files(
                mb_state.get("monthly_files") or [],
                mb_state.get("per_month_layout") or {},
                label_to_pool,
            )
            upload_months = set((pm.get("by_period") or {}).keys())
    except Exception:  # noqa: BLE001
        upload_months = set()

    try:
        from cecl_ui.services import solr_5300_backfill
        result = solr_5300_backfill.backfill_missing_quarters_distributed(
            cu, charter_int, sb["solr_url"], sb["core"],
            target_period, months,
            pool_distribution=dist["pool_distribution"],
            existing_dates=upload_months,
            source_period_iso=dist.get("period") or "",
        )
    except Exception as exc:  # noqa: BLE001
        sb["last_run"] = {
            "ok": False, "mode": "distributed", "error": str(exc),
        }
        return [f"5300 backfill failed: {type(exc).__name__}: {exc}"]

    sb["last_run"] = result
    if result.get("ok"):
        filled = len(result.get("months_filled") or [])
        rows = int(result.get("rows_written") or 0)
        msgs.append(
            f"5300 backfill (distributed): filled {filled} quarter-end(s), "
            f"wrote {rows} row(s) using earliest-month ratios "
            f"from {dist.get('period')}"
        )
    else:
        msgs.append(
            "5300 backfill (distributed) error: "
            + (result.get("error") or "unknown")
        )
    return msgs


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

    # Bulk-stage every analysed loan-data extract into
    # ``sample_uploads.loan_data_files`` so Step 13 (Column Mappings)
    # renders dropdowns immediately without the user having to re-upload
    # each file on Step 12. Mirrors ``_stage_hist_scan_files`` for CO/Recov.
    extract_msgs = _stage_loan_data_extracts(
        state, findings.get("sample_extracts") or []
    )
    if extract_msgs:
        report.setdefault("columns", []).extend(extract_msgs)
        report.setdefault("sample", []).extend(extract_msgs)
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
    elif findings.get("single_hist_bal"):
        shb = findings["single_hist_bal"]
        msgs = _apply_single_hist_bal_to_state(state, shb)
        report.setdefault("monthly_bal", []).extend(msgs)
        state["hist_balance_source"] = "single_workbook"
        report.setdefault("historical", []).append(
            "hist_balance_source set to single_workbook"
        )
        # Seed Step 2 pool_settings + Step 8 pool_map from the workbook's
        # pool labels when no consolidated 'Balance Sheets <CU>'
        # pool_seed is also present (flat hist-bal workbooks ARE the
        # authoritative pool list for CUs that ship them, e.g. TCP CU's
        # ``Historical Balance Sheets-TCP CU 65384.xlsx``).
        if not findings.get("pool_seed"):
            shb_entry = shb.get("entry") or {}
            shb_src = shb_entry.get("name") or "single historical balance workbook"
            flat_report = _apply_flat_pool_labels_to_state(
                state, shb.get("pool_labels") or [], shb_src
            )
            for key, fmsgs in flat_report.items():
                report.setdefault(key, []).extend(fmsgs)
            # Mirror loan_pools msgs to "pools" stepper bucket too.
            if "loan_pools" in flat_report:
                report.setdefault("pools", []).extend(
                    flat_report["loan_pools"]
                )

    # ----- Consolidated pool-grouping balance workbook --------------------
    # (Vizo Financial 'Balance Sheets <CU>.xlsx' carries the authoritative
    # pool list AND the sub-category -> pool mapping; runs INDEPENDENTLY of
    # annual / per-month balance findings so a CU can have both.)
    if findings.get("pool_seed"):
        ps_report = _apply_pool_seed_to_state(state, findings["pool_seed"])
        for key, msgs in ps_report.items():
            report.setdefault(key, []).extend(msgs)
        # Mirror to "pools" stepper bucket as well so the badge shows on
        # Step 2 (no-WARM path) regardless of how the step list keys it.
        if "loan_pools" in ps_report:
            report.setdefault("pools", []).extend(ps_report["loan_pools"])

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
    # Stage ALL monthly CO files into hist_scan.monthly_co_files so Step 5
    # (Historical Charge-Offs) renders the per-file table and the
    # `monthly_co_recov_aggregator` can aggregate them on the next visit.
    bulk_co_msgs = _stage_hist_scan_files(
        state, cls.get("co_files") or [], "monthly_co_files", "co"
    )
    if bulk_co_msgs:
        report.setdefault("co_history", []).extend(bulk_co_msgs)
    if findings.get("recov_file"):
        msgs = _apply_companion_file(
            state, "recovery", findings["recov_file"],
            "recov_files", "recov"
        )
        report.setdefault("co_recov", []).extend(msgs)
        report.setdefault("recov_history", []).extend(msgs)
    bulk_recov_msgs = _stage_hist_scan_files(
        state, cls.get("recov_files") or [], "monthly_recov_files", "recov"
    )
    if bulk_recov_msgs:
        report.setdefault("recov_history", []).extend(bulk_recov_msgs)

    # Flip Step 5/6 source picker to "monthly_files" when per-month
    # files were staged so the user immediately sees them on those
    # steps instead of the empty single-workbook upload field.
    flip_msgs = _auto_flip_hist_source(state)
    for m in flip_msgs:
        if "Charge-off" in m:
            report.setdefault("co_history", []).append(m)
        elif "Recovery" in m:
            report.setdefault("recov_history", []).append(m)

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

    # ----- Distributed 5300 backfill (auto, non-fatal) -----------------
    # Runs after everything else has populated hist_balance_source +
    # monthly_bal.pool_map so the earliest-month distribution can be
    # computed. Skipped gracefully if any precondition is missing.
    try:
        bf_msgs = _auto_run_distributed_backfill(state)
    except Exception as exc:  # noqa: BLE001
        bf_msgs = [f"5300 backfill skipped: {type(exc).__name__}: {exc}"]
    if bf_msgs:
        report.setdefault("historical", []).extend(bf_msgs)

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

    # loan_pools / warm (Step 2) — ensure the wizard lands here after
    # Step 1 auto-scan so the user can confirm or expand the pool list
    # BEFORE mapping loan codes (Step 3) into those pools. Suppressed
    # once the user explicitly clicks Save & Next on Step 2 (recorded
    # in _user_completed_steps via the Phase 9.9 mechanism). Both
    # keys are emitted; first_hil_step_key filters to whichever step
    # belongs to the active step list (loan_pools = NO_WARM path,
    # warm = WARM path).
    user_done = set(state.get("_user_completed_steps") or [])
    pool_settings = state.get("pool_settings") or []
    has_named_pool = any((p.get("name") or "").strip() for p in pool_settings)
    auto_scan_done = bool(state.get("_auto_scan_completed"))
    if "loan_pools" not in user_done and "warm" not in user_done:
        if not has_named_pool:
            add("loan_pools", "required",
                "No loan pools defined yet — add at least one pool before "
                "mapping loan codes")
            add("warm", "required",
                "No loan pools defined yet — upload a WARM file or add "
                "pools manually before mapping loan codes")
        elif auto_scan_done:
            n = len(pool_settings)
            add("loan_pools", "recommended",
                f"{n} loan pool(s) auto-detected — review names, ACL months, "
                "and risk-rated flags before continuing")
            add("warm", "recommended",
                f"{n} loan pool(s) auto-detected — review names, ACL months, "
                "and risk-rated flags before continuing")

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
    mb_source = (mb.get("source") or "").strip()
    if mb_source in ("per_year", "per_month", "single", "single_workbook"):
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
    elif mb_source == "manual":
        # Manual ACL entry mode — no labels to map; treated as
        # complete once the user has picked the mode.
        pass
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


def compute_step_completion(state: dict[str, Any]) -> set[str]:
    """Return wizard step keys whose saved data passes a basic completeness
    check.

    Used by the stepper to display a green ✓ on steps the user has
    filled in -- whether by Step 1 auto-scan, by manual edit on the
    page, or by a mix of both. A step that is "complete" here gets the
    ``auto`` badge in the UI, even if it still has *recommended* HIL
    entries (the HIL banner can flag those as optional review items).

    Steps with *required* HIL entries (e.g. CU name blank) will fail
    their predicate here and continue to render as ``hil``.

    The set covers every step key from any of the three step lists
    (WARM, NO_WARM, SCALE). The caller filters to its active list.
    """
    done: set[str] = set()

    # ----- identity -----
    if (state.get("credit_union") or "").strip() and (state.get("report_period") or "").strip():
        done.add("identity")

    # ----- loan_pools / warm (Step 2 on either path) -----
    pool_settings = state.get("pool_settings") or []
    has_named_pool = any((p.get("name") or "").strip() for p in pool_settings)
    if has_named_pool:
        done.add("loan_pools")
        # WARM path's "warm" step is the upload checkpoint -- it's
        # complete once the WARM file is parsed and pools were
        # extracted. Treat a non-empty pool_settings list as proof.
        warm = state.get("warm") or {}
        if warm.get("saved_path") or warm.get("balance_titles") or has_named_pool:
            done.add("warm")

    # ----- pools (Step 3 -- Loan Code Mapping) -----
    pm = state.get("pool_map") or {}
    if pm:
        # At least one code mapped to a real pool name (anything other
        # than blank or 'Ignore').
        if any(
            (str(v).strip() and str(v).strip().lower() != "ignore")
            for v in pm.values()
        ):
            done.add("pools")
    else:
        # No codes detected -- nothing to do, treat as complete so the
        # user isn't stuck with a perpetual ⚠ on an empty step.
        done.add("pools")

    # ----- balances (WARM path, Step 4 -- Balance Titles) -----
    btm = state.get("balance_title_map") or {}
    if btm:
        done.add("balances")

    # ----- baseline / historical -----
    warm = state.get("warm") or {}
    if warm.get("imported_baseline") or state.get("hist_balance_source"):
        done.add("baseline")
        done.add("historical")
    # Also count an explicit Solr-backfill mode as completion of the
    # historical step on the no-WARM path.
    he = state.get("hist_extracts") or {}
    if he.get("solr_backfill") and isinstance(he["solr_backfill"], dict) and he["solr_backfill"].get("mode"):
        done.add("historical")

    # ----- co_history / recov_history / dq_hist (NO_WARM Steps 5/6/7) -----
    # CRITICAL: Step 5 (Historical Charge-Offs) and Step 6 (Recoveries)
    # render their per-file tables from ``hist_scan.monthly_co_files`` /
    # ``hist_scan.monthly_recov_files`` -- NOT from ``sample_uploads.*``.
    # ``sample_uploads.co_files`` is only the single representative file
    # the auto-scan stages for sample analysis. If the bulk-stage for
    # that section never ran (old draft from before Phase 9.4, or only
    # one side of a combined CO+Recov classification got cross-listed),
    # the step page renders empty even though the breadcrumb would
    # otherwise show green. When auto-scan completed, REQUIRE the
    # corresponding hist_scan list (the real source for the page); when
    # auto-scan never ran (pure manual draft), accept the sample-uploads
    # / state-level path as before.
    sample_uploads = state.get("sample_uploads") or {}
    hist_scan = state.get("hist_scan") or {}
    auto_scan_done = bool(state.get("_auto_scan_completed"))
    has_monthly_co = bool(hist_scan.get("monthly_co_files") or hist_scan.get("co_files"))
    has_monthly_recov = bool(hist_scan.get("monthly_recov_files") or hist_scan.get("recov_files"))
    if auto_scan_done:
        if has_monthly_co:
            done.add("co_history")
        if has_monthly_recov:
            done.add("recov_history")
    else:
        if sample_uploads.get("co_files") or state.get("co_files") or has_monthly_co:
            done.add("co_history")
        if sample_uploads.get("recov_files") or state.get("recov_files") or has_monthly_recov:
            done.add("recov_history")
    dq = state.get("dq_hist") or {}
    if state.get("dq_files") or dq.get("uploaded") or dq.get("source"):
        done.add("dq_hist")

    # ----- monthly_bal -----
    mb = state.get("monthly_bal") or {}
    mb_source = (mb.get("source") or "").strip()
    if mb_source:
        labels = mb.get("parsed_pool_labels") or []
        mapping = mb.get("pool_map") or {}
        if labels and all(mapping.get(lab) for lab in labels):
            done.add("monthly_bal")
        elif not labels and mb_source in ("manual", "single", "single_workbook"):
            # No labels were parsed (e.g. manual ACL entry mode) --
            # treat as complete once the user has picked a source.
            done.add("monthly_bal")

    # ----- grades -----
    if state.get("credit_grades"):
        done.add("grades")

    # ----- credit_pull -----
    cp = state.get("credit_pull") or {}
    if cp.get("uploaded_filename") or cp.get("skipped") or cp.get("use_configured_report"):
        done.add("credit_pull")

    # ----- orig_score -- informational, always considered complete -----
    done.add("orig_score")

    # ----- sample (Loan Data Extract(s)) -----
    extracts = state.get("loan_data_extracts") or []
    loan_files = (
        state.get("loan_data_files")
        or (state.get("sample_uploads") or {}).get("loan_data_files")
        or []
    )
    if (extracts and any(e.get("file_pattern") for e in extracts)) or loan_files:
        done.add("sample")

    # ----- columns (Column Mappings) -----
    cm = state.get("column_mappings") or {}
    if cm.get("member_number") and cm.get("current_balance"):
        done.add("columns")

    # ----- balance_check / co_recov -- informational sanity steps -----
    done.add("balance_check")
    done.add("co_recov")

    # ----- impaired -----
    if sample_uploads.get("impaired_files") or (state.get("impaired") or {}).get("skipped"):
        done.add("impaired")

    # ----- files (File Format) -----
    if (state.get("file_pattern") or "").strip() and (state.get("date_format") or "").strip():
        done.add("files")

    # ----- economic -----
    econ = state.get("economic_data") or {}
    if (econ.get("state") or "").strip() and (econ.get("county") or "").strip():
        done.add("economic")

    # ----- mgmt_adj -----
    ma = state.get("mgmt_adj") or {}
    if ma.get("ltv_baseline") is not None and ma.get("probability_factor") is not None:
        done.add("mgmt_adj")

    # ----- reports -----
    rp = state.get("reports") or {}
    if any(bool(v) for v in rp.values()):
        done.add("reports")

    # ----- review -- never auto-complete; user must click Save -----

    return done
