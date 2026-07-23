"""Structural profiler for client delivery files (Setup Router, Phase 0).

METADATA ONLY. A :class:`FileProfile` captures sheet/column structure,
dtypes, value ranges, and — for *low-cardinality, code-like* columns only —
the distinct values. It never captures high-cardinality columns (names,
member numbers, addresses) or any loan-level row, so a profile is safe to
log, persist in the pattern registry, or hand to an AI escalation stage.

The profiler is deliberately format-agnostic: it profiles whatever files a
CU delivers and tags each with an ``archetype`` (loan_extract,
monthly_balances, participation, impaired, credit_pull) so later phases can
consume the same profiles without re-parsing.
"""

from __future__ import annotations

import os
import re
from dataclasses import asdict, dataclass, field
from typing import Any

import pandas as pd

# Distinct-value ceiling for emitting sample values. Code columns (loan
# types, statuses) fall well under this; identity columns (member #, names)
# blow past it and are therefore never sampled — the privacy guarantee.
_LOW_CARD_MAX = 60
_SAMPLE_CAP = 40
_HEADER_SCAN_ROWS = 12


@dataclass
class ColumnProfile:
    name: str
    index: int                     # 0-based position in the sheet
    dtype: str                     # numeric | text | date | mixed | empty
    non_null: int
    distinct: int
    numeric_frac: float
    min_v: float | None = None
    max_v: float | None = None
    code_values: list[str] = field(default_factory=list)  # only if low-cardinality


@dataclass
class SheetProfile:
    sheet: str
    header_row: int                # 1-based row where the header was detected
    n_rows: int                    # data rows below the header
    n_cols: int
    columns: list[ColumnProfile]


@dataclass
class FileProfile:
    path: str
    filename: str
    archetype: str
    sheets: list[SheetProfile]

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    def column_names(self) -> set[str]:
        return {c.name.lower() for sp in self.sheets for c in sp.columns}


# ---------------------------------------------------------------------------
# Raw reading + header detection
# ---------------------------------------------------------------------------
def _read_raw(path: str, sheet: str | int | None = None,
              nrows: int | None = None) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(path, sheet_name=sheet, header=None,
                             nrows=nrows, dtype=object)
    return pd.read_csv(path, header=None, nrows=nrows, dtype=object,
                       keep_default_na=False, na_values=[""])


def _detect_header_row(raw: pd.DataFrame, scan: int = _HEADER_SCAN_ROWS) -> int:
    """Return the 0-based index of the most likely header row.

    Heuristic: the row (within the first ``scan``) with the most non-empty,
    mostly-textual, distinct cells. Handles title/blank preludes (monthly
    balance files start on row 1 with dates; participation remittance files
    put the header on row 3).
    """
    best_i, best_score = 0, -1.0
    for i in range(min(scan, len(raw))):
        cells = [c for c in raw.iloc[i].tolist()
                 if c is not None and str(c).strip() != ""]
        if not cells:
            continue
        text_like = sum(
            1 for c in cells
            if isinstance(c, str)
            and not str(c).strip().replace(".", "", 1).isdigit()
        )
        distinct = len({str(c).strip() for c in cells})
        score = text_like + 0.5 * distinct
        if score > best_score:
            best_score, best_i = score, i
    return best_i


def _col_dtype(series: pd.Series) -> tuple[str, float]:
    s = series.dropna()
    if s.empty:
        return "empty", 0.0
    num = pd.to_numeric(s, errors="coerce")
    frac = float(num.notna().mean())
    if frac >= 0.95:
        return "numeric", frac
    if frac <= 0.05:
        return "text", frac
    return "mixed", frac


def _norm_code(x: str) -> str:
    """Normalize a code cell: '85.0' -> '85', trim; leave non-numeric as-is."""
    x = str(x).strip()
    if x.replace(".", "", 1).isdigit():
        try:
            return str(int(float(x)))
        except (TypeError, ValueError):
            return x
    return x


# ---------------------------------------------------------------------------
# Profiling
# ---------------------------------------------------------------------------
def profile_sheet(raw: pd.DataFrame, sheet_name: str) -> SheetProfile:
    hdr = _detect_header_row(raw)
    header = [str(c).strip() if c is not None else ""
              for c in raw.iloc[hdr].tolist()]
    body = raw.iloc[hdr + 1:].reset_index(drop=True)
    body.columns = range(body.shape[1])

    cols: list[ColumnProfile] = []
    for j, name in enumerate(header):
        if j >= body.shape[1]:
            break
        col = body[j]
        dtype, numfrac = _col_dtype(col)
        nn = int(col.notna().sum())
        clean = col.dropna().astype(str).str.strip()
        distinct = int(clean.nunique())
        cp = ColumnProfile(
            name=name or f"col_{j}", index=j, dtype=dtype,
            non_null=nn, distinct=distinct, numeric_frac=round(numfrac, 3),
        )
        if dtype in ("numeric", "mixed"):
            num = pd.to_numeric(col, errors="coerce")
            if num.notna().any():
                cp.min_v = float(num.min())
                cp.max_v = float(num.max())
        # Emit distinct values ONLY for low-cardinality columns (codes,
        # statuses). Identity/PII columns exceed the ceiling -> never sampled.
        if 0 < distinct <= _LOW_CARD_MAX:
            cp.code_values = sorted({_norm_code(v) for v in clean.unique()})[:_SAMPLE_CAP]
        cols.append(cp)
    return SheetProfile(sheet=sheet_name, header_row=hdr + 1,
                        n_rows=len(body), n_cols=len(cols), columns=cols)


def classify_archetype(sheets: list[SheetProfile], filename: str) -> str:
    fn = filename.lower()
    names = {c.name.lower() for sp in sheets for c in sp.columns}
    sheet_names = " ".join(sp.sheet.lower() for sp in sheets)

    # Specific deliverables first — many of these filenames also contain the
    # word "loans", so they must be matched BEFORE the structural loan-extract
    # check or they masquerade as the main extract (e.g. "Improved
    # Deteriorated Loans Report").
    if "investor upb" in names or "remittance" in sheet_names \
            or "participation" in fn:
        return "participation"
    if "impaired" in fn or "impaired loans" in sheet_names:
        return "impaired"
    if "credit pull" in fn:
        return "credit_pull"
    if any("alll" in n for n in names) or ("monthly" in fn and "balance" in fn) \
            or "balances by pool" in fn:
        return "monthly_balances"
    if any(k in fn for k in ("charge off", "chargeoff", "recovery",
                             "improved", "deteriorated", "tracking report")):
        return "report_other"

    # Structural loan-extract test: an account/member key + a balance column +
    # enough rows/columns to be a loan-level file (not a small report).
    has_acct = any(re.search(r"account|member|\bacct\b|\bcif\b|borr", n)
                   for n in names)
    has_bal = any("balance" in n for n in names)
    max_rows = max((sp.n_rows for sp in sheets), default=0)
    max_cols = max((sp.n_cols for sp in sheets), default=0)
    if has_acct and has_bal and max_rows >= 50 and max_cols >= 8:
        return "loan_extract"
    return "unknown"


def profile_file(path: str, max_sheets: int = 8) -> FileProfile:
    """Profile every (or the first ``max_sheets``) sheet of ``path``."""
    ext = os.path.splitext(path)[1].lower()
    sheets: list[SheetProfile] = []
    if ext in (".xlsx", ".xls"):
        xl = pd.ExcelFile(path)
        for sheet in xl.sheet_names[:max_sheets]:
            raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)
            if raw is None or raw.empty:
                continue
            sheets.append(profile_sheet(raw, str(sheet)))
    else:
        raw = _read_raw(path)
        sheets.append(profile_sheet(raw, "csv"))

    filename = os.path.basename(path)
    return FileProfile(path=path, filename=filename,
                       archetype=classify_archetype(sheets, filename),
                       sheets=sheets)
