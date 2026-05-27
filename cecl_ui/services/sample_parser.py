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
                              "orig_score", "original_score", "fico_score",
                              "fico", "credit_score", "score"]),
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
]


def _normalise(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(name).strip().lower()).strip("_")


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
    # (description, regex)
    ("YYYY-MM",       r"(\d{4})-(\d{2})"),
    ("YYYY_MM",       r"(\d{4})_(\d{2})"),
    ("YYYYMM",        r"(\d{4})(\d{2})"),
    ("MM-YYYY",       r"(\d{2})-(\d{4})"),
    ("MMDDYY",        r"(\d{2})(\d{2})(\d{2})"),
    # Month-name forms — the import engine knows how to translate these.
    ("Mon_YYYY",      r"(?i)(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*[-_ \.]+(20\d{2})"),
    ("YYYY_Mon",      r"(?i)(20\d{2})[-_ \.]+(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*"),
]


def guess_filename_patterns(filename: str) -> dict[str, str]:
    """Derive sensible defaults for file_pattern and date_pattern."""
    p = Path(filename)
    stem = p.stem
    ext = p.suffix.lower().lstrip(".")
    # File pattern: keep the leading alphabetic prefix, allow .* in the middle,
    # accept either xlsx/xls if Excel, otherwise the file's own extension.
    prefix_match = re.match(r"^([A-Za-z][A-Za-z0-9]*)", stem)
    prefix = prefix_match.group(1) if prefix_match else stem
    if ext in {"xlsx", "xls"}:
        ext_part = r"(xlsx|xls)"
    elif ext == "csv":
        ext_part = "csv"
    else:
        ext_part = ext or r"(xlsx|xls|csv)"
    file_pattern = f"{re.escape(prefix)}.*\\.{ext_part}$"

    # Date pattern: scan filename for the first matching candidate.
    date_pattern = r"(\d{4})-(\d{2})"  # default
    for _desc, rx in _DATE_RX_CANDIDATES:
        if re.search(rx, stem):
            date_pattern = rx
            break

    return {"file_pattern": file_pattern, "date_pattern": date_pattern}


# ---------- Header detection ----------

def _looks_numeric(val: Any) -> bool:
    if val is None:
        return False
    if isinstance(val, (int, float)):
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
    return text_score_0 >= max(2, len(row0) // 2) and num_score_1 >= max(2, len(row1) // 2)


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
            str(h).strip() if h is not None and str(h).strip()
            else f"col_{_excel_col_letter(i)}"
            for i, h in enumerate(raw.iloc[hdr_idx].tolist())
        ]
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


def parse_pool_map_file(
    file_path: str | Path,
    code_col: str | None = None,
    name_col: str | None = None,
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
    if suffix in {".xlsx", ".xls"}:
        import pandas as pd
        df = pd.read_excel(path, dtype=object)
    elif suffix == ".csv":
        import pandas as pd
        df = pd.read_csv(path, dtype=object, keep_default_na=False)
    else:
        raise ValueError(
            f"Unsupported pool-map file type '{suffix}'. "
            "Please upload an .xlsx, .xls, or .csv file."
        )

    if df.empty or df.shape[1] < 2:
        raise ValueError(
            "Pool-map file must have at least two columns "
            "(loan code and pool name)."
        )

    headers = [str(h).strip() for h in df.columns.tolist()]

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
        "rows": rows,
        "row_count": len(rows),
    }


def analyse_sample_file(
    file_path: str | Path,
    original_filename: str,
    max_pool_codes: int = 40,
    sample_rows: int = 5,
    header_row: int | None = None,
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
            raw = pd.read_csv(path, header=None, dtype=object, keep_default_na=False)
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
            str(h).strip() if h is not None and str(h).strip()
            else f"col_{_excel_col_letter(i)}"
            for i, h in enumerate(raw.iloc[hdr_idx].tolist())
        ]
        body = raw.iloc[hdr_idx + 1:].reset_index(drop=True)
        body.columns = headers
    else:
        headers = [f"col_{_excel_col_letter(i)}" for i in range(raw.shape[1])]
        body = raw.copy()
        body.columns = headers

    # Column suggestions
    col_sugg = _match_columns(headers)

    # Pool-code suggestions: distinct values from the pool column, if found
    pool_codes: list[str] = []
    pool_col = col_sugg.get("loan_pool_code")
    if pool_col and pool_col in body.columns:
        seen: set[str] = set()
        for raw_val in body[pool_col].dropna().tolist():
            code = _clean_code(raw_val)
            if not code:
                continue
            if "/" in code:
                code = _clean_code(code.split("/")[0])
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

    return {
        "ok": True,
        "error": None,
        "filename": original_filename,
        "saved_path": str(path),
        "file_pattern": fname_guesses["file_pattern"],
        "date_pattern": fname_guesses["date_pattern"],
        "has_header": has_header,
        "header_row": header_row_used,
        "row_count_total": int(len(raw)),
        "headers": headers,
        "sample_rows": preview,
        "column_suggestions": col_sugg,
        "pool_code_suggestions": pool_codes,
    }
