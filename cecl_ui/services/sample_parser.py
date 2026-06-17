"""Parse a sample loan file and produce wizard-ready suggestions.

Used by Step 2 of the new-CU wizard.  Given an uploaded Excel/CSV file, this
module returns:

* detected header presence
* the list of column headers (or numeric positions if headerless)
* up to N rows of sample values per column (for display)
* suggested ``column_mappings`` (system field -> source column name)
* suggested ``file_pattern`` and ``date_pattern`` derived from the filename
* suggested ``pool_map`` (raw code -> "" placeholder) seeded from distinct
  values found in the suggested loan-pool-code column

The module never writes to session state directly — the route does that.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:  # pragma: no cover
    import pandas as pd


# ---------- Heuristics for column-name matching ----------
#
# Each entry: (system_field, [keyword_substrings_lowercased]).
# Order within each list matters: earlier keywords win over later ones.
# The first column whose name contains any of these substrings is picked.
COLUMN_HEURISTICS: list[tuple[str, list[str]]] = [
    ("member_number",        ["lndacno", "member_number", "member#", "member no",
                              "account_number", "account no", "acct", "member",
                              "account"]),
    ("current_balance",      ["lndcbal", "current_balance", "current bal",
                              "curr_bal", "cur_bal", "balance", "bal"]),
    ("original_fico_score",  ["lndcrsc", "original_fico", "orig_fico",
                              "orig_score", "original_score", "application_creditscore",
                              "application_score"]),
    ("current_fico_score",   ["current_fico", "current_score",
                              "latest_credit_score", "last_credit_score",
                              "primary_score", "current_credit_score",
                              "updated_fico", "refresh_score", "refreshed_score",
                              "fico_score", "fico", "credit_score", "score"]),
    ("loan_pool_code",       ["lndalpc", "loan_pool", "pool_code", "loan_type",
                              "product_code", "loan_code", "pool", "product",
                              "type", "code"]),
    ("interest_rate",        ["lndirte", "interest_rate", "int_rate", "rate",
                              "apr"]),
    ("open_date",            ["lndopen", "open_date", "orig_date",
                              "origination_date", "open"]),
    ("original_loan_amount", ["lndorg", "original_loan_amount", "orig_amt",
                              "original_amount", "loan_amount", "orig"]),
    ("total_available_credit", ["lndcrlim", "total_available_credit",
                                 "available_credit", "credit_limit",
                                 "line_of_credit", "loc_limit", "loc",
                                 "limit"]),
    ("days_delinquent",      ["lnddel", "days_delinquent", "dq_days",
                              "delinquent", "delinq", "dq"]),
    ("business_risk_rating", ["business_risk_rating", "business_risk",
                              "risk_rating", "risk_rate", "rr_grade",
                              "brr"]),
]


def _normalise(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(name).strip().lower()).strip("_")


# ---------- Date format validation ----------
#
# These are the literal values ``import_data.extract_snapshot_date`` knows
# how to translate from the configured ``date_pattern`` regex groups.
# Anything else silently falls through to the YYYY-MM branch and yields
# garbage dates. Single source of truth so the wizard can validate user
# picks before the YAML is written.

ACCEPTED_DATE_FORMATS: tuple[str, ...] = (
    "YYYY-MM",      # default: 2 capture groups (year, month)
    "YYYYMMDD",     # 3 capture groups (year, month, day)
    "MMDDYY",       # 3 capture groups (month, day, year[2 or 4 digits])
    "MMYY",         # 2 capture groups (month, year[2 digits])
    "MMYYYY",       # 2 capture groups (month, year[4 digits])
    "YYYYQ",        # 2 capture groups (year, quarter[1-4])
    "QYYYY",        # 2 capture groups (quarter, year)
)

DATE_FORMAT_LABELS: dict[str, str] = {
    "YYYY-MM":  "YYYY-MM (e.g. 2025-12) - default; matches 2-group regex",
    "YYYYMMDD": "YYYYMMDD (e.g. 20251231) - 3-group, year/month/day",
    "MMDDYY":   "MMDDYY or MM-DD-YYYY (e.g. 12-31-25 or 12-31-2025)",
    "MMYY":     "MMYY (e.g. 12-25)",
    "MMYYYY":   "MMYYYY (e.g. 12-2025)",
    "YYYYQ":    "YYYYQ (e.g. 2025Q4 or 2025-Q4)",
    "QYYYY":    "QYYYY (e.g. Q4-2025)",
}


def validate_date_format(
    date_pattern: str,
    date_format: str,
    sample_filename: str = "",
) -> dict[str, Any]:
    """Validate a (date_pattern, date_format) pair and resolve a sample.

    Returns ``{ok, error, preview, recognized}``:
      * ``recognized`` - True iff ``date_format`` is one ``import_data`` knows.
      * ``ok`` - True iff the regex compiled and (if ``sample_filename``
        given) successfully resolved to a date.
      * ``preview`` - resolved ISO date string (e.g. ``"2025-12-31"``)
        when a sample was supplied and the parse succeeded; else ``""``.
      * ``error`` - human-readable string when something failed.
    """
    out: dict[str, Any] = {
        "ok": False, "error": None, "preview": "",
        "recognized": date_format in ACCEPTED_DATE_FORMATS,
    }
    if not out["recognized"]:
        out["error"] = (
            f"Date format {date_format!r} is not recognized. "
            f"Pick one of: {', '.join(ACCEPTED_DATE_FORMATS)}."
        )
        return out
    try:
        re.compile(date_pattern)
    except re.error as exc:
        out["error"] = f"Invalid date_pattern regex: {exc}"
        return out
    if not sample_filename:
        out["ok"] = True
        return out
    # Defer the import: import_data is a heavy module and we only need
    # one function. This keeps wizard pages snappy.
    try:
        import import_data as _id  # type: ignore[import-not-found]
    except Exception as exc:  # noqa: BLE001
        out["error"] = f"Could not import date resolver: {exc}"
        return out
    iso = _id.extract_snapshot_date(
        sample_filename,
        {"date_pattern": date_pattern, "date_format": date_format},
    )
    if not iso:
        out["error"] = (
            f"Could not extract a date from {sample_filename!r} "
            f"using pattern {date_pattern!r} / format {date_format}."
        )
        return out
    out["ok"] = True
    out["preview"] = iso
    return out


# ---------- Member / Account number layout inference ----------
#
# Looks at the literal account-key cells from the loan extract and
# proposes a ``member_account`` block (mode + delimiter / suffix length)
# that the import pipeline knows how to split. The wizard previously
# defaulted everyone to ``mode=fixed_suffix, suffix_length=3``; CUs
# whose account column is e.g. ``000000470-06`` (delimiter) or
# ``219044`` (no suffix) silently failed credit-pull joins until an
# operator hand-edited the YAML.

_DELIM_INFER_RX = re.compile(r"^\s*(\d+)([\-_LXlx/])(\d+)\s*$")
_DIGITS_RX = re.compile(r"^\s*(\d+)\s*$")


def infer_member_account_from_values(
    values: list[Any],
    sample_size: int = 50,
) -> dict[str, Any] | None:
    """Inspect raw values and propose a ``member_account`` mapping.

    Returns ``None`` when the values are too sparse or ambiguous to
    suggest anything useful (caller should fall back to its own
    default). Otherwise returns a dict shaped like::

        {"mode": "delimiter", "delimiter": "-", "suffix_length": 3,
         "confidence": "high"|"medium"|"low",
         "rationale": "human-readable explanation"}

    The ``suffix_length`` field is always present so the dict can be
    used as a drop-in replacement for the wizard's default; for
    ``mode=delimiter`` it carries the *typical* observed suffix length
    (informational; the importer ignores it). For ``mode=fixed_suffix``
    it is the dominant trailing-segment length actually observed.
    """
    cleaned: list[str] = []
    for v in values:
        if v is None:
            continue
        if isinstance(v, float) and v != v:  # NaN
            continue
        s = str(v).strip()
        if not s or s.lower() in {"nan", "none"}:
            continue
        cleaned.append(s)
        if len(cleaned) >= sample_size:
            break
    if len(cleaned) < 3:
        return None

    delim_hits: dict[str, list[int]] = {}
    digit_only: list[str] = []
    other = 0
    for s in cleaned:
        m = _DELIM_INFER_RX.match(s)
        if m:
            delim_hits.setdefault(m.group(2), []).append(len(m.group(3)))
            continue
        d = _DIGITS_RX.match(s)
        if d:
            digit_only.append(d.group(1))
            continue
        other += 1

    total = len(cleaned)

    # --- delimiter mode ---
    if delim_hits:
        # Pick the most-frequent delimiter. Promote it to delimiter mode
        # only if it covers a clear majority and digit-only rows aren't.
        best_delim, best_lens = max(
            delim_hits.items(), key=lambda kv: len(kv[1])
        )
        share = len(best_lens) / total
        if share >= 0.6:
            # Most-common suffix length (informational).
            suffix_len = max(set(best_lens), key=best_lens.count)
            conf = "high" if share >= 0.85 else "medium"
            return {
                "mode": "delimiter",
                "delimiter": best_delim if best_delim.lower() != "x" else best_delim.upper(),
                "suffix_length": int(suffix_len),
                "confidence": conf,
                "rationale": (
                    f"{len(best_lens)}/{total} sample values match "
                    f"<digits>{best_delim}<digits>; suffix ~"
                    f"{suffix_len} chars."
                ),
            }

    # --- digit-only branch ---
    # If everything is bare digits, decide between fixed_suffix and
    # zero-suffix (member-only). Look at trailing-3 vs trailing-2 vs
    # trailing-4 stability. If the last N chars vary widely AND the
    # length itself varies widely, the column is probably the bare
    # member number with no embedded suffix.
    if digit_only and len(digit_only) / total >= 0.85:
        lens = {len(s) for s in digit_only}
        # If lengths are tightly clustered, prefer fixed_suffix=3 (the
        # historical default); we can't tell from values alone whether
        # the last 3 chars are a suffix or part of the member number,
        # but fixed_suffix=3 has been the heuristic for years and the
        # operator can override on Step 3.
        if len(lens) <= 2 and min(lens) >= 6:
            return {
                "mode": "fixed_suffix",
                "delimiter": "-",
                "suffix_length": 3,
                "confidence": "low",
                "rationale": (
                    f"All {len(digit_only)}/{total} values are digit-only "
                    f"({sorted(lens)} chars); defaulting to "
                    "fixed_suffix=3. Override if the column is the bare "
                    "member number with no embedded suffix."
                ),
            }
        # Highly variable digit-only lengths -> probably bare member
        # numbers with no suffix portion.
        return {
            "mode": "fixed_suffix",
            "delimiter": "-",
            "suffix_length": 0,
            "confidence": "medium",
            "rationale": (
                f"{len(digit_only)}/{total} digit-only values with "
                f"variable lengths {sorted(lens)} - looks like a bare "
                "member number with no suffix."
            ),
        }

    return None


_UNNAMED_RX = re.compile(r"^unnamed:\s*\d+(?:_level_\d+)*$", re.IGNORECASE)


def _is_blank_header(s: str) -> bool:
    """True for cells that should be treated as missing headers — empty,
    pandas' synthetic ``Unnamed: 0`` placeholder, or the literal string
    ``nan`` (which is what ``str(float('nan'))`` produces when a raw
    Excel cell is empty)."""
    if not s:
        return True
    low = s.lower()
    if low == "nan":
        return True
    if _UNNAMED_RX.match(s):
        return True
    return False


def _clean_header(cell: object, idx: int | None = None) -> str:
    """Normalise a spreadsheet header cell: collapse any internal whitespace
    (CR/LF, tabs, multi-space — common with wrap-text Excel cells) into a
    single space and strip ends. Browsers do the same when round-tripping
    `<option value="...">` on form submit, so the dropdown's selected value
    will match the rendered options across saves only if we store the
    normalised shape.

    When the cell is blank/NaN/``Unnamed: N`` and a column index is
    provided, returns ``col_<LETTER>`` (e.g. ``col_H`` for the 8th
    column). This guarantees every dropdown option is unique even when
    the source spreadsheet has multiple unlabelled columns — without it,
    every blank column collapses to the same value and the form would
    silently snap the user's pick to the leftmost blank column on save.
    """
    s = "" if cell is None else re.sub(r"\s+", " ", str(cell)).strip()
    if _is_blank_header(s):
        if idx is not None:
            return f"col_{_excel_col_letter(idx)}"
        return ""
    return s


def _excel_col_letter(idx: int) -> str:
    """0 -> 'A', 25 -> 'Z', 26 -> 'AA', 27 -> 'AB', ..."""
    if idx < 0:
        return ""
    s = ""
    n = idx
    while True:
        s = chr(ord("A") + (n % 26)) + s
        n = n // 26 - 1
        if n < 0:
            break
    return s


def _dedupe_headers(headers: list[str]) -> list[str]:
    """Append a positional suffix to repeated header names.

    Some workbooks (e.g. CUMA MTG Servicing reports) ship with duplicate
    column labels in the header row. Pandas tolerates duplicates but
    ``df[col]`` then returns a DataFrame instead of a Series, which
    breaks downstream ``.tolist()`` / ``.dropna()`` calls. Suffix the
    second+ occurrence with the Excel column letter to keep names unique.
    """
    seen: dict[str, int] = {}
    out: list[str] = []
    for i, h in enumerate(headers):
        if h in seen:
            seen[h] += 1
            out.append(f"{h} ({_excel_col_letter(i)})")
        else:
            seen[h] = 1
            out.append(h)
    return out


def _match_columns(headers: list[str]) -> dict[str, str]:
    """Return {system_field: matched_header} using the heuristics."""
    out: dict[str, str] = {}
    used: set[str] = set()
    norm_headers = [(h, _normalise(h)) for h in headers]
    for field, keywords in COLUMN_HEURISTICS:
        match: str | None = None
        for kw in keywords:
            kw_n = _normalise(kw)
            for original, norm in norm_headers:
                if original in used:
                    continue
                if kw_n in norm or norm in kw_n:
                    match = original
                    break
            if match:
                break
        if match:
            out[field] = match
            used.add(match)
    return out


# ---------- Filename pattern guessing ----------

_DATE_RX_CANDIDATES: list[tuple[str, str]] = [
    # (description, regex). Year-anchored (``20\d{2}``) where we can, so
    # heuristic doesn't grab e.g. "0331" as a year out of "03312026".
    # Order is most-specific first.
    ("YYYY-MM-DD",    r"(20\d{2})-(\d{2})-(\d{2})"),
    ("YYYY_MM_DD",    r"(20\d{2})_(\d{2})_(\d{2})"),
    ("YYYYMMDD",      r"(20\d{2})(\d{2})(\d{2})"),
    ("MM-DD-YYYY",    r"(\d{2})-(\d{2})-(20\d{2})"),
    ("MM_DD_YYYY",    r"(\d{2})_(\d{2})_(20\d{2})"),
    ("MMDDYYYY",      r"(\d{2})(\d{2})(20\d{2})"),
    ("YYYY-MM",       r"(20\d{2})-(\d{2})"),
    ("YYYY_MM",       r"(20\d{2})_(\d{2})"),
    ("YYYYMM",        r"(20\d{2})(\d{2})(?!\d)"),
    ("MM-YYYY",       r"(\d{2})-(20\d{2})"),
    # MM-DD-YY with separators (e.g. "12-31-25", "12_31_25"). Negative
    # look-behind avoids matching the tail half of "2025-12-31" as "25-12-31".
    ("MM-DD-YY",      r"(?<!\d)(\d{2})-(\d{2})-(\d{2})(?!\d)"),
    ("MM_DD_YY",      r"(?<!\d)(\d{2})_(\d{2})_(\d{2})(?!\d)"),
    ("MMDDYY",        r"(\d{2})(\d{2})(\d{2})(?!\d)"),
    # Quarter-end forms: "2025Q4", "2026-Q1", "2026_Q1".
    ("YYYYQ",         r"(20\d{2})[-_ ]?[Qq]([1-4])(?!\d)"),
    ("QYYYY",         r"(?<![A-Za-z])[Qq]([1-4])[-_ ]?(20\d{2})(?!\d)"),
    # Month-name forms — the import engine knows how to translate these.
    ("Mon_YYYY",      r"(?i)(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*[-_ \.]+(20\d{2})"),
    ("YYYY_Mon",      r"(?i)(20\d{2})[-_ \.]+(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*"),
]


# Tokens that should be wildcarded in a filename pattern so the same
# regex matches files dropped in other months / years.
_MONTH_NAME_RX_GROUP = (
    r"(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|"
    r"jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|"
    r"oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
)

# Match candidates inside the original sample filename. Each entry is
# (regex-to-find, regex-to-emit-into-pattern). Order matters: the
# longest / most specific patterns win when matches overlap.
_FNAME_GENERALIZE_RULES: list[tuple[re.Pattern[str], str]] = [
    # Month-name + 4-digit year: "December 2025", "Mar_2026", "Sept-2024"
    (re.compile(
        rf"(?i)\b{_MONTH_NAME_RX_GROUP}[\s_\-\.]+20\d{{2}}\b"),
     rf"{_MONTH_NAME_RX_GROUP}[\s_\-\.]+20\d{{2}}"),
    # Month-name + 2-digit year: "Mar 26", "March_26"
    (re.compile(
        rf"(?i)\b{_MONTH_NAME_RX_GROUP}[\s_\-\.]+\d{{2}}(?!\d)"),
     rf"{_MONTH_NAME_RX_GROUP}[\s_\-\.]+\d{{2}}"),
    # YYYYMMDD (8 digits): "20251130" — match BEFORE YYYYMM so the day
    # part doesn't get truncated.
    (re.compile(
        r"(?<!\d)20\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])(?!\d)"),
     r"20\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])"),
    # MMDDYYYY (8 digits): "03312026"
    (re.compile(
        r"(?<!\d)(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])20\d{2}(?!\d)"),
     r"(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])20\d{2}"),
    # YYYY-MM-DD / YYYY_MM_DD: "2026-03-31"
    (re.compile(
        r"(?<!\d)20\d{2}[-_](?:0[1-9]|1[0-2])[-_](?:0[1-9]|[12]\d|3[01])(?!\d)"),
     r"20\d{2}[-_](?:0[1-9]|1[0-2])[-_](?:0[1-9]|[12]\d|3[01])"),
    # MM-DD-YYYY: "03-31-2026"
    (re.compile(
        r"(?<!\d)(?:0[1-9]|1[0-2])[-_](?:0[1-9]|[12]\d|3[01])[-_]20\d{2}(?!\d)"),
     r"(?:0[1-9]|1[0-2])[-_](?:0[1-9]|[12]\d|3[01])[-_]20\d{2}"),
    # YYYY-MM / YYYY_MM / YYYYMM (any separator, optional)
    (re.compile(r"(?<!\d)20\d{2}[-_]?(?:0[1-9]|1[0-2])(?!\d)"),
     r"20\d{2}[-_]?(?:0[1-9]|1[0-2])"),
    # MM-YYYY / MM_YYYY
    (re.compile(r"(?<!\d)(?:0[1-9]|1[0-2])[-_]20\d{2}(?!\d)"),
     r"(?:0[1-9]|1[0-2])[-_]20\d{2}"),
    # MMYYYY (no separator): "092022"
    (re.compile(r"(?<!\d)(?:0[1-9]|1[0-2])20\d{2}(?!\d)"),
     r"(?:0[1-9]|1[0-2])20\d{2}"),
    # YY-MM (e.g. "25-12 AIRES")
    (re.compile(r"(?<!\d)\d{2}[-_](?:0[1-9]|1[0-2])(?!\d)"),
     r"\d{2}[-_](?:0[1-9]|1[0-2])"),
    # Bare 4-digit year (matched after the combined forms above)
    (re.compile(r"(?<!\d)20\d{2}(?!\d)"), r"20\d{2}"),
    # Bare month name (matched last so combined month+year wins first).
    # Use letter-boundary lookarounds rather than ``\b`` so the rule
    # also fires inside underscore-joined names like
    # ``December_Loan_File_-_Upload.xlsx`` (``\b`` fails to match
    # between a letter and an underscore — both are word chars).
    (re.compile(rf"(?i)(?<![A-Za-z]){_MONTH_NAME_RX_GROUP}(?![A-Za-z])"),
     _MONTH_NAME_RX_GROUP),
    # Version suffix: "V2", "V3", "v10". Wildcards the digit so next
    # quarter's "V3" file matches a pattern derived from a "V2" sample.
    # Restricted to 1-2 digits and an explicit boundary so generic words
    # like "Vault123" don't get rewritten.
    (re.compile(r"(?<![A-Za-z0-9])[Vv]\d{1,2}(?![A-Za-z0-9])"),
     r"[Vv]\d{1,2}"),
    # Trailing numeric suffix like ".01" / ".001" / ".10" that Symitar
    # / Vizo IDLR exports sometimes append to the baseline file
    # (e.g. "airesln V2 DEC 2025.01.xlsx"). Wildcard as an OPTIONAL
    # group so a pattern derived from a ".01"-suffixed baseline also
    # matches the subsequent unsuffixed monthly drops. Anchored to
    # end-of-stem so middle-of-name dotted tokens (e.g. ".v2.") aren't
    # affected; restricted to 1-3 digits to avoid eating obvious
    # non-suffix sequences.
    (re.compile(r"\.\d{1,3}$"), r"(?:\.\d{1,3})?"),
]


def _generalize_filename_pattern(stem: str, ext_part: str) -> str:
    """Build a case-insensitive ``file_pattern`` regex for ``stem`` that
    wildcards every month-name and year token so the same regex still
    matches next month's drop of the same file.

    Tokens stay literal otherwise. Matches anchored ``^…\\.<ext>$``.
    """
    spans: list[tuple[int, int, str]] = []  # (start, end, replacement)
    for rx, repl in _FNAME_GENERALIZE_RULES:
        for m in rx.finditer(stem):
            spans.append((m.start(), m.end(), repl))
    # Resolve overlaps: keep the earliest-starting match; on ties prefer
    # the longer one. Drop anything that overlaps a chosen span.
    spans.sort(key=lambda x: (x[0], -(x[1] - x[0])))
    chosen: list[tuple[int, int, str]] = []
    cursor = 0
    for s, e, r in spans:
        if s < cursor:
            continue
        chosen.append((s, e, r))
        cursor = e
    # Loosen runs of whitespace (escaped as ``\ ``), underscores, and
    # escaped dashes (``\-``) so the pattern still matches across common
    # CU naming variations:
    #   - ``secure_filename`` rewrites spaces to underscores during upload
    #     (sample ``December Loan File - Upload.xlsx`` → staged
    #     ``December_Loan_File_-_Upload.xlsx``).
    #   - Per-quarter drops sometimes lose the dash entirely
    #     (``February Loan File Upload.xlsx``).
    # Collapsing runs of any of {space, underscore, dash} to ``[\s_\-]+``
    # lets one wizard-derived pattern match all three forms.
    #
    # CRITICAL: this collapse must run ONLY on escaped literal stem
    # fragments — never on rule replacement strings. Rule replacements
    # contain ``_`` and ``\-`` literals inside character classes like
    # ``[\s_\-\.]+`` (combined month+year rule) and ``[-_]`` (date-format
    # rules); applying the collapse to the whole body corrupts those
    # classes into malformed nested classes (``[\s[\s_\-]+\.]+``) that
    # silently never match any real filename.
    _SEP_COLLAPSE_RX = re.compile(r"(?:\\ |_|\\-)+")
    _SEP_REPL = r"[\\s_\\-]+"

    def _collapse(s: str) -> str:
        return _SEP_COLLAPSE_RX.sub(_SEP_REPL, s)

    parts: list[str] = []
    last = 0
    for s, e, r in chosen:
        if s > last:
            parts.append(_collapse(re.escape(stem[last:s])))
        parts.append(r)
        last = e
    if last < len(stem):
        parts.append(_collapse(re.escape(stem[last:])))
    body = "".join(parts) if parts else _collapse(re.escape(stem))
    return f"(?i)^{body}\\.{ext_part}$"


def guess_filename_patterns(filename: str) -> dict[str, str]:
    """Derive sensible defaults for file_pattern + date_pattern.

    Returns a dict with ``file_pattern``, ``date_pattern`` and ``date_format``
    (one of ``"YYYY-MM"``, ``"MMDDYY"`` — the formats ``import_data`` knows).

    The emitted ``file_pattern`` wildcards month names, 4-digit years,
    YYYY-MM and MM-YYYY tokens so that next month's drop of the same
    file still matches without regenerating the YAML.
    """
    p = Path(filename)
    stem = p.stem
    ext = p.suffix.lower().lstrip(".")
    if ext in {"xlsx", "xls"}:
        ext_part = r"(xlsx|xls)"
    elif ext == "csv":
        ext_part = "csv"
    else:
        ext_part = ext or r"(xlsx|xls|csv)"
    file_pattern = _generalize_filename_pattern(stem, ext_part)

    # Date pattern: scan filename for the first candidate that matches AND
    # produces a sensible month (1-12). Default = YYYY-MM with YYYY-MM format.
    date_pattern = r"(\d{4})-(\d{2})"
    date_format = "YYYY-MM"
    for desc, rx in _DATE_RX_CANDIDATES:
        m = re.search(rx, stem)
        if not m:
            continue
        # Validate the match implies a real month-of-year, so we don't
        # save a regex that will explode at import time. Month-name forms
        # always produce valid months (the regex restricts to jan..dec).
        if desc.startswith(("YYYY-MM-DD", "YYYY_MM_DD", "YYYYMMDD",
                            "YYYY-MM", "YYYY_MM", "YYYYMM")):
            try:
                if not 1 <= int(m.group(2)) <= 12:
                    continue
            except (ValueError, IndexError):
                continue
        elif desc.startswith(("MM-DD-YYYY", "MM_DD_YYYY", "MMDDYYYY",
                              "MM-DD-YY", "MM_DD_YY",
                              "MM-YYYY", "MMDDYY")):
            try:
                if not 1 <= int(m.group(1)) <= 12:
                    continue
            except (ValueError, IndexError):
                continue
        elif desc == "YYYYQ":
            try:
                if not 1 <= int(m.group(2)) <= 4:
                    continue
            except (ValueError, IndexError):
                continue
        elif desc == "QYYYY":
            try:
                if not 1 <= int(m.group(1)) <= 4:
                    continue
            except (ValueError, IndexError):
                continue
        date_pattern = rx
        # Pick the date_format extract_snapshot_date expects. Month-first
        # 3-group patterns use 'MMDDYY' (it handles 2- or 4-digit years).
        if desc in ("MM-DD-YYYY", "MM_DD_YYYY", "MMDDYYYY",
                    "MM-DD-YY", "MM_DD_YY", "MMDDYY"):
            date_format = "MMDDYY"
        elif desc in ("YYYY-MM-DD", "YYYY_MM_DD", "YYYYMMDD"):
            date_format = "YYYYMMDD"
        elif desc == "MM-YYYY":
            date_format = "MMYYYY"
        elif desc == "YYYYQ":
            date_format = "YYYYQ"
        elif desc == "QYYYY":
            date_format = "QYYYY"
        else:
            date_format = "YYYY-MM"
        break

    return {"file_pattern": file_pattern, "date_pattern": date_pattern,
            "date_format": date_format}


# ---------- Header detection ----------

def _looks_numeric(val: Any) -> bool:
    if val is None:
        return False
    if isinstance(val, bool):
        return False
    if isinstance(val, (int, float)):
        # NaN is a float but not a "real" numeric value for header detection.
        try:
            import math
            if isinstance(val, float) and math.isnan(val):
                return False
        except Exception:  # noqa: BLE001
            pass
        return True
    s = str(val).strip().replace(",", "").replace("$", "").replace("(", "-").replace(")", "")
    if not s:
        return False
    try:
        float(s)
        return True
    except ValueError:
        return False


def _detect_has_header(df_no_header: "pd.DataFrame") -> bool:
    """Heuristic: if the first row is mostly text and the second row is mostly
    numeric/date, treat the first row as a header."""
    if len(df_no_header) < 2:
        return True  # benign default
    row0 = df_no_header.iloc[0].tolist()
    row1 = df_no_header.iloc[1].tolist()
    text_score_0 = sum(1 for v in row0 if v is not None and not _looks_numeric(v))
    num_score_1  = sum(1 for v in row1 if _looks_numeric(v))
    if text_score_0 >= max(2, len(row0) // 2) and num_score_1 >= max(2, len(row1) // 2):
        return True
    # Strong signal: row 0 has many short-string identifier-like cells AND
    # zero numeric cells (typical AIRES / Symitar header rows). Row 1 just
    # needs to have at least a couple of numeric cells anywhere — even text-
    # heavy data rows usually carry balance/score/term numbers somewhere.
    if num_score_1 >= 2:
        row0_numeric = sum(1 for v in row0 if v is not None and _looks_numeric(v))
        row0_strings = sum(1 for v in row0
                           if v is not None
                           and not _looks_numeric(v)
                           and isinstance(v, str)
                           and v.strip())
        if row0_numeric == 0 and row0_strings >= 5:
            return True
    return False


# ---------- Main entry ----------

def _clean_code(val: Any) -> str:
    """Strip standard + non-breaking whitespace, drop empty/'nan' values."""
    if val is None:
        return ""
    s = str(val)
    # str.strip() handles standard whitespace; explicitly strip nbsp + BOM too.
    s = s.strip().strip("\u00a0\ufeff\u200b\t")
    # Collapse internal runs of whitespace to a single space (defensive).
    s = re.sub(r"\s+", " ", s).strip()
    if not s or s.lower() == "nan":
        return ""
    return s


def extract_pool_codes(
    file_path: str | Path,
    column_name: str,
    header_row: int | None = None,
    split_char: str = "/",
    max_codes: int = 40,
) -> list[str]:
    """Re-read the saved sample file and pull distinct codes from one column.

    Used both internally and from the wizard when the user changes the
    loan_pool_code column mapping after the original upload.
    """
    path = Path(file_path)
    suffix = path.suffix.lower()
    try:
        if suffix in {".xlsx", ".xls"}:
            import pandas as pd
            raw = pd.read_excel(path, header=None, dtype=object)
        elif suffix == ".csv":
            import pandas as pd
            raw = pd.read_csv(path, header=None, dtype=object,
                              keep_default_na=False)
        else:
            return []
    except Exception:  # noqa: BLE001
        return []

    if raw.empty:
        return []

    has_header = _detect_has_header(raw) if header_row is None else (header_row >= 1)
    if has_header:
        hdr_idx = max(0, (header_row or 1) - 1)
        headers = [
            _clean_header(h, i)
            for i, h in enumerate(raw.iloc[hdr_idx].tolist())
        ]
        headers = _dedupe_headers(headers)
        body = raw.iloc[hdr_idx + 1:].reset_index(drop=True)
    else:
        headers = [f"col_{_excel_col_letter(i)}" for i in range(raw.shape[1])]
        body = raw.copy()
    body.columns = headers

    if column_name not in body.columns:
        return []

    codes: list[str] = []
    seen: set[str] = set()
    for raw_val in body[column_name].dropna().tolist():
        code = _clean_code(raw_val)
        if not code:
            continue
        # "11 / New Car" -> "11"
        if split_char and split_char in code:
            code = _clean_code(code.split(split_char)[0])
        if not code or code in seen:
            continue
        seen.add(code)
        codes.append(code)
        if len(codes) >= max_codes:
            break
    return codes


# ---------- Pool-map upload ----------

# Header keywords to identify the "code" vs "name" columns in an
# uploaded pool-map file.  Order matters: more specific terms first.
_POOL_CODE_KEYWORDS = (
    "code", "raw", "id", "lndaltc", "loan_type", "loantype",
    "type_code", "product_code", "category_code",
)
_POOL_NAME_KEYWORDS = (
    "pool", "name", "description", "label", "category",
    "product", "loan_type_desc", "type_desc",
)


def _score_header(header: str, keywords: tuple[str, ...]) -> int:
    h = (header or "").lower().strip()
    if not h:
        return 0
    for i, kw in enumerate(keywords):
        if kw in h:
            return len(keywords) - i  # earlier keyword = higher score
    return 0


def _read_row_cells(file_path: str | Path, row_1based: int) -> list[str]:
    """Return the raw cell values of ``row_1based`` (1-based) from
    an .xlsx/.xls/.csv file as cleaned strings. Trailing empties are
    trimmed; embedded blanks are preserved. Returns ``[]`` on any
    parsing error or when the row is out of range."""
    p = Path(file_path)
    suffix = p.suffix.lower()
    try:
        import pandas as pd
        if suffix in {".xlsx", ".xls"}:
            raw = pd.read_excel(p, header=None, dtype=object,
                                nrows=row_1based)
        elif suffix == ".csv":
            raw = pd.read_csv(p, header=None, dtype=object,
                              keep_default_na=False, nrows=row_1based)
        else:
            return []
        idx = row_1based - 1
        if idx < 0 or idx >= len(raw):
            return []
        cells = [
            "" if v is None or str(v).strip().lower() == "nan"
            else str(v).strip()
            for v in raw.iloc[idx].tolist()
        ]
        # Strip trailing blanks
        while cells and cells[-1] == "":
            cells.pop()
        return cells
    except Exception:  # noqa: BLE001
        return []


def parse_pool_map_file(
    file_path: str | Path,
    code_col: str | None = None,
    name_col: str | None = None,
    header_row: int | None = None,
) -> dict[str, Any]:
    """Read an uploaded loan-code -> pool-name map file (xlsx/csv).

    Returns:
        {
          "headers": ["LNDALTC", "Pool Name", ...],
          "code_column": "LNDALTC",
          "name_column": "Pool Name",
          "rows": [{"code": "AUTO_NEW", "name": "New Vehicle"}, ...],
          "row_count": 12,
        }

    Picks the two columns by header keyword scoring unless the caller
    overrides via ``code_col`` / ``name_col``.  If the file has only two
    columns we treat them as code/name without scoring.
    """
    path = Path(file_path)
    suffix = path.suffix.lower()
    # Convert 1-based header_row from caller to pandas' 0-based index;
    # ``None`` means auto-detect.
    hdr_idx: int | None
    if header_row is None or header_row <= 0:
        hdr_idx = None
    else:
        hdr_idx = int(header_row) - 1

    if suffix in {".xlsx", ".xls"}:
        import pandas as pd
        if hdr_idx is None:
            # Read the first ~25 rows raw; pick the row with the most
            # non-empty string cells as the header row.
            raw = pd.read_excel(path, header=None, dtype=object, nrows=25)
            best_row = 0
            best_score = -1
            for i in range(len(raw)):
                row = raw.iloc[i]
                score = sum(
                    1 for v in row
                    if v is not None
                    and str(v).strip() != ""
                    and str(v).strip().lower() != "nan"
                )
                if score > best_score:
                    best_score = score
                    best_row = i
            hdr_idx = best_row
        df = pd.read_excel(path, header=hdr_idx, dtype=object)
    elif suffix == ".csv":
        import pandas as pd
        if hdr_idx is None:
            raw = pd.read_csv(path, header=None, dtype=object,
                              keep_default_na=False, nrows=25)
            best_row = 0
            best_score = -1
            for i in range(len(raw)):
                row = raw.iloc[i]
                score = sum(
                    1 for v in row
                    if v is not None and str(v).strip() != ""
                )
                if score > best_score:
                    best_score = score
                    best_row = i
            hdr_idx = best_row
        df = pd.read_csv(path, header=hdr_idx, dtype=object,
                         keep_default_na=False)
    else:
        raise ValueError(
            f"Unsupported pool-map file type '{suffix}'. "
            "Please upload an .xlsx, .xls, or .csv file."
        )

    if df.shape[1] < 2:
        raise ValueError(
            "Pool-map file must have at least two columns "
            "(loan code and pool name)."
        )

    headers = [_clean_header(h, i) for i, h in enumerate(df.columns.tolist())]
    headers = _dedupe_headers(headers)
    df.columns = headers

    # Pick columns
    if code_col and code_col in headers:
        code_h = code_col
    elif df.shape[1] == 2:
        code_h = headers[0]
    else:
        scored = sorted(
            ((h, _score_header(h, _POOL_CODE_KEYWORDS)) for h in headers),
            key=lambda x: x[1],
            reverse=True,
        )
        code_h = scored[0][0] if scored[0][1] > 0 else headers[0]

    if name_col and name_col in headers and name_col != code_h:
        name_h = name_col
    elif df.shape[1] == 2:
        name_h = headers[1]
    else:
        scored = sorted(
            ((h, _score_header(h, _POOL_NAME_KEYWORDS))
             for h in headers if h != code_h),
            key=lambda x: x[1],
            reverse=True,
        )
        name_h = scored[0][0] if scored and scored[0][1] > 0 else (
            [h for h in headers if h != code_h][0]
        )

    rows: list[dict[str, str]] = []
    seen: set[str] = set()
    for _, r in df.iterrows():
        code = _clean_code(r.get(code_h))
        name = _clean_code(r.get(name_h))
        if not code or code in seen:
            continue
        seen.add(code)
        rows.append({"code": code, "name": name})

    return {
        "headers": headers,
        "code_column": code_h,
        "name_column": name_h,
        "header_row": (hdr_idx + 1) if hdr_idx is not None else 1,
        "rows": rows,
        "row_count": len(rows),
    }


def analyse_sample_file(
    file_path: str | Path,
    original_filename: str,
    max_pool_codes: int = 40,
    sample_rows: int = 5,
    header_row: int | None = None,
    split_char: str | None = "/",
) -> dict[str, Any]:
    """Parse the uploaded sample and return wizard suggestions.

    ``header_row`` is a 1-indexed override:
        None  -> auto-detect via heuristic
        0     -> file has no header row
        N>=1  -> treat row N (1-indexed) as the header; rows above it are
                 dropped (cover sheets, blank rows, etc.)

    Returns a dict with keys:
        ok (bool)
        error (str | None)
        filename
        saved_path (str)                  (path to the file on disk for re-parse)
        file_pattern, date_pattern        (filename-derived guesses)
        has_header (bool)
        header_row (int | None)           (1-indexed; None = no header)
        headers (list[str])               (column names — synthetic if no header)
        sample_rows (list[dict])          (first N rows for preview)
        column_suggestions (dict)         {system_field: header}
        pool_code_suggestions (list[str]) (distinct raw codes)
    """
    path = Path(file_path)
    suffix = path.suffix.lower()
    try:
        if suffix in {".xlsx", ".xls"}:
            import pandas as pd  # deferred: pandas cold-import is slow on network drives
            raw = pd.read_excel(path, header=None, dtype=object)
        elif suffix == ".csv":
            import pandas as pd  # deferred: pandas cold-import is slow on network drives
            # Sniff a few bytes for BOM / common Windows credit-union CSV encodings.
            # AIRES-style exports are sometimes UTF-16 LE (BOM ff fe) or Windows-1252.
            try:
                with open(path, "rb") as fh:
                    head = fh.read(4)
            except Exception:
                head = b""
            encodings_to_try: list[str | None] = []
            if head.startswith(b"\xff\xfe") or head.startswith(b"\xfe\xff"):
                encodings_to_try = ["utf-16", "utf-8-sig", "cp1252", "latin-1"]
            elif head.startswith(b"\xef\xbb\xbf"):
                encodings_to_try = ["utf-8-sig", "cp1252", "latin-1"]
            else:
                encodings_to_try = [None, "utf-8-sig", "cp1252", "latin-1"]
            # Auto-detect separator (handles tab-delimited .csv exports too).
            last_exc: Exception | None = None
            raw = None
            for enc in encodings_to_try:
                for sep in (None, ",", "\t", ";", "|"):
                    try:
                        raw = pd.read_csv(
                            path,
                            header=None,
                            dtype=object,
                            keep_default_na=False,
                            encoding=enc,
                            sep=sep,
                            engine="python" if sep is None else "c",
                        )
                        last_exc = None
                        break
                    except Exception as exc:  # noqa: BLE001
                        last_exc = exc
                        continue
                if raw is not None:
                    break
            if raw is None:
                raise last_exc if last_exc else RuntimeError("Could not parse CSV")
        else:
            return {"ok": False, "error": f"Unsupported file type: {suffix}"}
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "error": f"Could not read file: {exc}"}

    if raw.empty:
        return {"ok": False, "error": "File is empty."}

    has_header = _detect_has_header(raw) if header_row is None else (header_row >= 1)
    header_row_used: int | None = None
    if has_header:
        if header_row is None or header_row < 1:
            hdr_idx = 0  # auto-detect path: header is the first row
        else:
            hdr_idx = min(header_row - 1, len(raw) - 1)
        header_row_used = hdr_idx + 1
        headers = [
            _clean_header(h, i)
            for i, h in enumerate(raw.iloc[hdr_idx].tolist())
        ]
        headers = _dedupe_headers(headers)
        body = raw.iloc[hdr_idx + 1:].reset_index(drop=True)
        body.columns = headers
    else:
        headers = [f"col_{_excel_col_letter(i)}" for i in range(raw.shape[1])]
        body = raw.copy()
        body.columns = headers

    # Column suggestions
    col_sugg = _match_columns(headers)

    # Pool-code suggestions: distinct values from the pool column, if found.
    # Phase 9.26: ``split_char`` (passed by the caller, default ``"/"``) gates
    # the prefix-extraction step. Symitar/Episys-style codes ``11/New Car``
    # need ``"/"`` so the seed becomes ``"11"``; CUMA mortgage codes like
    # ``15/15 ARM`` need ``split_char=""`` (or ``None``) so the full label
    # ``"15/15 ARM"`` survives — otherwise the seeded ``"15"`` collides with
    # any future Symitar loan code 15.
    pool_codes: list[str] = []
    pool_col = col_sugg.get("loan_pool_code")
    if pool_col and pool_col in body.columns:
        seen: set[str] = set()
        for raw_val in body[pool_col].dropna().tolist():
            code = _clean_code(raw_val)
            if not code:
                continue
            if split_char and split_char in code:
                code = _clean_code(code.split(split_char)[0])
            if not code or code in seen:
                continue
            seen.add(code)
            pool_codes.append(code)
            if len(pool_codes) >= max_pool_codes:
                break

    # Sample rows for preview
    preview: list[dict[str, Any]] = []
    for i in range(min(sample_rows, len(body))):
        row = {}
        for h in headers:
            v = body.iloc[i].get(h)
            row[h] = "" if v is None or (isinstance(v, float) and v != v) else str(v)
        preview.append(row)

    fname_guesses = guess_filename_patterns(original_filename)

    # Member/account layout inference: peek at the column heuristically
    # picked as the member-number column. Falls back gracefully when no
    # such column was detected.
    ma_suggestion: dict[str, Any] | None = None
    member_col = col_sugg.get("member_number")
    if member_col and member_col in body.columns:
        try:
            ma_suggestion = infer_member_account_from_values(
                body[member_col].head(60).tolist()
            )
        except Exception:  # noqa: BLE001
            ma_suggestion = None

    return {
        "ok": True,
        "error": None,
        "filename": original_filename,
        "saved_path": str(path),
        "file_pattern": fname_guesses["file_pattern"],
        "date_pattern": fname_guesses["date_pattern"],
        "date_format": fname_guesses.get("date_format", "YYYY-MM"),
        "has_header": has_header,
        "header_row": header_row_used,
        "row_count_total": int(len(raw)),
        "headers": headers,
        "sample_rows": preview,
        "column_suggestions": col_sugg,
        "pool_code_suggestions": pool_codes,
        "member_account_suggestion": ma_suggestion,
    }
