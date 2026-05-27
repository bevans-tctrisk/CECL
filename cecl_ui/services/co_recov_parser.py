"""Parser helpers for the Charge-Off / Recovery wizard step.

Provides:

* ``inspect_file(path, max_rows=12)`` — read a CO/recov file the same way
  the report engine does (concat all sheets), return column letter labels
  (A, B, C...) plus a small preview so the user can pick which 0-based
  column holds account / code / amount / date.

* ``extract_codes(path, parse_config)`` — apply the saved parse_config
  (has_header / skip_rows / account_col / code_col / amount_col / date_col)
  and return the distinct loan-pool codes that appear in the file.

* ``validate_codes(state, kind)`` — for each saved CO or recov file,
  run ``extract_codes`` and compare against the wizard's loan-code map
  (``state['pool_map']``). Returns per-file totals and a deduped list of
  unmapped codes.

The parse_config shape mirrors what
``generate_report._parse_chargeoff_file`` and ``_parse_recovery_file``
already accept, so what the wizard saves is exactly what the engine reads.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _col_letter(idx: int) -> str:
    """0-based column index -> Excel-style letter (0->A, 25->Z, 26->AA)."""
    n = idx + 1
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def _read_concat(filepath: str | Path):
    """Mirror ``generate_report._read_data_file``: read as DataFrame w/o header.

    For Excel files with multiple sheets concatenates all sheets that share
    the maximum column count.
    """
    import pandas as pd  # deferred — pandas import is slow on network drives

    ext = os.path.splitext(str(filepath))[1].lower()
    if ext == ".csv":
        return pd.read_csv(filepath, header=None, dtype=object,
                           keep_default_na=False)
    xl = pd.ExcelFile(filepath)
    parts = []
    for s in xl.sheet_names:
        d = pd.read_excel(xl, sheet_name=s, header=None, dtype=object)
        if not d.empty:
            parts.append(d)
    if not parts:
        return pd.DataFrame()
    if len(parts) == 1:
        return parts[0]
    target_cols = max(d.shape[1] for d in parts)
    parts = [d for d in parts if d.shape[1] == target_cols]
    return pd.concat(parts, ignore_index=True)


def _looks_like_loan_code(val: Any) -> bool:
    """Borrowed from generate_report — 1-5 chars, has a letter, not all digits."""
    if val is None:
        return False
    s = str(val).strip()
    if not (1 <= len(s) <= 5):
        return False
    if s.isdigit():
        return False
    return any(ch.isalpha() for ch in s)


# ---------------------------------------------------------------------------
# Public
# ---------------------------------------------------------------------------

def inspect_file(filepath: str | Path, max_rows: int = 12) -> dict[str, Any]:
    """Read the file and return a small preview for column-mapping UI.

    Returns::
        {
            "ok": bool, "error": str|None,
            "filename": str,
            "ncols": int,
            "columns": [{"index": 0, "letter": "A"}, ...],
            "rows": [["...", "..."], ...],   # first max_rows rows, str-coerced
        }
    """
    out: dict[str, Any] = {
        "ok": False, "error": None, "filename": Path(str(filepath)).name,
        "ncols": 0, "columns": [], "rows": [],
    }
    p = Path(str(filepath))
    if not p.exists():
        out["error"] = f"File not found: {filepath}"
        return out
    try:
        df = _read_concat(p)
    except Exception as exc:  # noqa: BLE001
        out["error"] = f"Could not read file: {exc}"
        return out
    if df.empty:
        out["error"] = "File is empty."
        return out

    ncols = df.shape[1]
    out["ncols"] = ncols
    # Capture first-row values so the UI can show "A — Account Number" style
    # labels in column-mapping dropdowns when the file has a header row.
    first_row = df.iloc[0] if not df.empty else []
    headers: list[str] = []
    for v in (first_row.tolist() if hasattr(first_row, "tolist") else list(first_row)):
        if v is None:
            headers.append("")
        else:
            try:
                # NaN check without importing pd here
                if isinstance(v, float) and v != v:
                    headers.append("")
                    continue
            except Exception:  # noqa: BLE001
                pass
            headers.append(str(v).strip())
    out["columns"] = [
        {"index": i, "letter": _col_letter(i),
         "header": headers[i] if i < len(headers) else ""}
        for i in range(ncols)
    ]
    rows: list[list[str]] = []
    for _, row in df.head(max_rows).iterrows():
        rows.append([
            "" if v is None or (isinstance(v, float) and v != v)  # NaN check
            else str(v) for v in row.tolist()
        ])
    out["rows"] = rows
    out["ok"] = True
    return out


def _suggest_columns(filepath: str | Path) -> dict[str, Any]:
    """Heuristic column-index suggestions for a CO/recov file.

    Used to seed the wizard mapping when the user first uploads a file.
    Returns a dict with optional keys: ``has_header`` (bool),
    ``account_col`` / ``code_col`` / ``amount_col`` / ``date_col`` (int).
    """
    import pandas as pd  # deferred

    out: dict[str, Any] = {"has_header": False}
    try:
        df = _read_concat(filepath)
    except Exception:  # noqa: BLE001
        return out
    if df.empty:
        return out

    first = df.iloc[0]
    first_a = first.iloc[0] if len(first) > 0 else None
    col0_numeric = isinstance(first_a, (int, float)) and not pd.isna(first_a)
    first_vals = ["" if pd.isna(v) else str(v).lower() for v in first]
    has_header = (not col0_numeric) and any(
        ("account" in v) or ("charge" in v) or ("loan" in v) or ("acct" in v)
        for v in first_vals
    )
    out["has_header"] = bool(has_header)
    out["account_col"] = 0
    if has_header:
        for c, hdr in enumerate(first_vals):
            if "fico" in hdr or "score" in hdr:
                continue
            if "account" in hdr or "acct" in hdr:
                out["account_col"] = c
            elif (("security" in hdr) or ("code" in hdr)) and (
                    "account" not in hdr and "sub" not in hdr):
                out["code_col"] = c
            elif any(k in hdr for k in (
                    "amount", "chg off am", "chargeoff am", "pymt", "principal")):
                out["amount_col"] = c
            elif "date" in hdr or "effective" in hdr:
                out["date_col"] = c
        return out

    # No header — sniff data rows
    body = df.iloc[1:] if has_header else df
    ncols = body.shape[1]
    for c in range(ncols):
        sample = body[c].dropna()
        if sample.empty:
            continue
        first_val = sample.iloc[0]
        if "code_col" not in out and _looks_like_loan_code(first_val):
            out["code_col"] = c
        elif "date_col" not in out and isinstance(first_val, pd.Timestamp):
            out["date_col"] = c
        elif "amount_col" not in out and isinstance(first_val, (int, float)) and c > 1:
            out["amount_col"] = c
    return out


def extract_codes(filepath: str | Path, parse_config: dict[str, Any]) -> dict[str, Any]:
    """Apply ``parse_config`` and return distinct loan-pool codes from the file."""
    import pandas as pd  # deferred

    out: dict[str, Any] = {"ok": False, "error": None, "codes": []}
    p = Path(str(filepath))
    if not p.exists():
        out["error"] = f"File not found: {filepath}"
        return out
    try:
        df = _read_concat(p)
    except Exception as exc:  # noqa: BLE001
        out["error"] = f"Could not read file: {exc}"
        return out
    if df.empty:
        out["error"] = "File is empty."
        return out

    if parse_config.get("has_header"):
        df = df.iloc[1:]
    skip = int(parse_config.get("skip_rows") or 0)
    if skip > 0:
        df = df.iloc[skip:]

    code_col = parse_config.get("code_col")
    account_col = parse_config.get("account_col", 0)
    if code_col is None:
        out["error"] = "No code_col set in parse_config."
        return out
    code_col = int(code_col)
    if code_col < 0 or code_col >= df.shape[1]:
        out["error"] = f"code_col {code_col} is out of range (file has {df.shape[1]} cols)."
        return out

    # Drop rows where account column isn't numeric (filter out totals/comments)
    if account_col is not None and 0 <= int(account_col) < df.shape[1]:
        ac = int(account_col)
        acct_series = df.iloc[:, ac]
        acct_numeric = acct_series.apply(lambda v: pd.notna(v) and (
            str(v).replace("-", "").replace(" ", "").isdigit()
            if isinstance(v, str)
            else isinstance(v, (int, float)) and not pd.isna(v)))
        df = df[acct_numeric]

    seen: set[str] = set()
    codes: list[str] = []
    for raw in df.iloc[:, code_col].dropna().tolist():
        s = str(raw).strip()
        if not s:
            continue
        # Loan codes can contain a "/" suffix in some CUs (e.g. "VA/01") —
        # use the prefix for mapping purposes, matching the engine.
        if "/" in s:
            s = s.split("/", 1)[0].strip()
        if not s or s in seen:
            continue
        seen.add(s)
        codes.append(s)

    out["codes"] = codes
    out["ok"] = True
    return out


def validate_codes(state: dict[str, Any], kind: str) -> dict[str, Any]:
    """Validate codes in saved CO or recov files against state['pool_map'].

    ``kind`` is ``"co"`` or ``"recov"``. Returns::
        {
            "ok": bool, "error": str|None,
            "configured": bool,           # True iff parse_config has code_col
            "files": [
                {"name", "path", "ok", "error",
                 "total_codes", "mapped_codes", "unmapped_codes": [str]},
                ...
            ],
            "unmapped_codes": [str],      # union across files, dedup, ordered
        }
    """
    out: dict[str, Any] = {
        "ok": False, "error": None, "configured": False,
        "files": [], "unmapped_codes": [],
    }
    if kind not in ("co", "recov"):
        out["error"] = f"Unknown kind: {kind}"
        return out

    cfg_key = "co_columns" if kind == "co" else "recov_columns"
    upload_key = "co_files" if kind == "co" else "recov_files"

    cfg = state.get(cfg_key) or {}
    if cfg.get("code_col") in (None, ""):
        out["error"] = "Set the column mapping (Account / Code / Amount) first."
        return out
    out["configured"] = True

    pool_map = state.get("pool_map") or {}
    # Normalize the wizard pool_map keys for case-insensitive lookups.
    mapped_keys = {str(k).strip().lower() for k in pool_map}

    uploads = (state.get("sample_uploads") or {}).get(upload_key) or []
    if not uploads:
        out["error"] = "No file uploaded yet."
        return out

    seen_unmapped: set[str] = set()
    union: list[str] = []
    for entry in uploads:
        path = entry.get("path")
        info: dict[str, Any] = {
            "name": entry.get("name", ""), "path": path,
            "ok": False, "error": None,
            "total_codes": 0, "mapped_codes": 0, "unmapped_codes": [],
        }
        if not path or not Path(str(path)).exists():
            info["error"] = "File no longer exists on disk."
            out["files"].append(info)
            continue
        result = extract_codes(path, cfg)
        if not result["ok"]:
            info["error"] = result["error"]
            out["files"].append(info)
            continue
        codes = result["codes"]
        info["ok"] = True
        info["total_codes"] = len(codes)
        unmapped = []
        for code in codes:
            key = code.strip().lower()
            if key in mapped_keys and (pool_map.get(code) or
                                       next((v for k, v in pool_map.items()
                                             if str(k).strip().lower() == key), "")):
                info["mapped_codes"] += 1
            else:
                unmapped.append(code)
                if key not in seen_unmapped:
                    seen_unmapped.add(key)
                    union.append(code)
        info["unmapped_codes"] = unmapped
        out["files"].append(info)

    out["unmapped_codes"] = union
    out["ok"] = True
    return out
