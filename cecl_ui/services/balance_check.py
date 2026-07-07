"""Balance Adjustment helper.

Compares the per-pool balance recorded in the recurring Monthly Balance
file (Step "Monthly Balance File") against the sum of ``current_balance``
in the user's Loan Data Extract(s), grouped by the same pool name. The
wizard renders the result as a side-by-side table so the user can spot
mis-mapped pool codes / balance-format issues before generating reports.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd

from cecl_ui.services import monthly_bal_parser


_UNNAMED_RX = re.compile(r"^unnamed:\s*\d+(?:_level_\d+)*$", re.IGNORECASE)
_WS_RX = re.compile(r"\s+")


def _excel_idx_to_letter(idx: int) -> str:
    """0 -> 'A', 25 -> 'Z', 26 -> 'AA'..."""
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


def _normalize_header_columns(cols) -> list[str]:
    """Mirror ``sample_parser._clean_header`` so wizard-saved mapping
    values (e.g. ``'Current Loan Bal'`` for a wrap-text cell, or
    ``'col_X'`` for a blank header) resolve correctly against the
    DataFrame's columns at lookup time.

    Also dedupes duplicate header names by appending an Excel-letter
    suffix to the second+ occurrence (mirrors
    ``sample_parser._dedupe_headers``). Some workbooks (e.g. CUMA MTG
    Servicing reports) ship with duplicate column labels, which would
    otherwise cause ``df[name]`` to return a DataFrame slice instead of
    a Series and break downstream string ops.
    """
    out: list[str] = []
    for i, c in enumerate(cols):
        s = "" if c is None else _WS_RX.sub(" ", str(c)).strip()
        low = s.lower()
        if (not s) or low == "nan" or _UNNAMED_RX.match(s):
            out.append(f"col_{_excel_idx_to_letter(i)}")
        else:
            out.append(s)
    seen: dict[str, int] = {}
    deduped: list[str] = []
    for i, name in enumerate(out):
        if name in seen:
            seen[name] += 1
            deduped.append(f"{name} ({_excel_idx_to_letter(i)})")
        else:
            seen[name] = 1
            deduped.append(name)
    return deduped


def _report_period_cutoff(state: dict[str, Any]) -> str:
    """Return the global Report Period anchor (Step 1) as an ISO
    month-end date string ``YYYY-MM-DD``, or ``""`` if not set.

    Used to scope period-keyed balance lookups so data dated AFTER the
    report period is ignored (e.g. Apr-Dec columns in a Jan-Dec balance
    sheet when running March).
    """
    rp = (state.get("report_period") or "").strip()
    if not rp:
        return ""
    try:
        y_s, m_s = rp.split("-", 1)[0], rp.split("-", 1)[1][:2]
        y = int(y_s)
        m = int(m_s)
    except (ValueError, IndexError):
        return ""
    if not (2000 <= y <= 2100 and 1 <= m <= 12):
        return ""
    # Compute the last day of the month without importing calendar.
    if m == 12:
        next_first = pd.Timestamp(year=y + 1, month=1, day=1)
    else:
        next_first = pd.Timestamp(year=y, month=m + 1, day=1)
    last = (next_first - pd.Timedelta(days=1)).date()
    return last.isoformat()


def _load_loan_extract(
    path: str | Path,
    has_header: bool,
    header_row: int | None = None,
) -> pd.DataFrame | None:
    """Read a single loan-extract file. Returns None on failure.

    ``header_row`` is a 1-indexed override (e.g. ``2`` when the real column
    names live on the second row because the first row is just position
    numbers). When omitted and ``has_header`` is True, defaults to row 1.
    """
    p = Path(path)
    if not p.exists():
        return None
    suffix = p.suffix.lower()
    # Translate (has_header, header_row) into the pandas `header=` arg.
    if has_header:
        hdr = int(header_row) - 1 if header_row and int(header_row) > 0 else 0
    else:
        hdr = None
    try:
        if suffix == ".csv":
            df = pd.read_csv(p, header=hdr)
        elif suffix in (".xlsx", ".xls"):
            df = pd.read_excel(p, header=hdr)
        else:
            return None
    except Exception:  # noqa: BLE001
        return None
    if has_header:
        df.columns = _normalize_header_columns(df.columns)
    return df


def _get_col(df: pd.DataFrame, ref: Any, has_header: bool) -> pd.Series | None:
    """Resolve a column-mapping reference to a Series.

    ``ref`` is either an actual header name (when ``has_header``) or a
    "col_A"-style letter / 0-based int when not.
    """
    if ref is None or ref == "":
        return None
    if has_header:
        if ref in df.columns:
            return df[ref]
        return None
    # No-header: accept "col_A", "A", or an integer.
    if isinstance(ref, int):
        idx = ref
    else:
        s = str(ref).strip().upper()
        if s.startswith("COL_"):
            s = s[4:]
        if s.isdigit():
            idx = int(s)
        elif s.isalpha():
            idx = 0
            for ch in s:
                idx = idx * 26 + (ord(ch) - ord("A") + 1)
            idx -= 1
        else:
            return None
    if idx < 0 or idx >= df.shape[1]:
        return None
    return df.iloc[:, idx]


def _clean_balance(series: pd.Series, remove_chars, accounting_negatives) -> pd.Series:
    s = series.astype(str)
    for ch in (remove_chars or []):
        s = s.str.replace(ch, "", regex=False)
    if accounting_negatives:
        s = s.str.replace("(", "-", regex=False).str.replace(")", "", regex=False)
    return pd.to_numeric(s, errors="coerce")


def _normalize_raw_code(x: Any) -> str:
    """Coerce a pool-code cell to a clean string.

    Handles NaN/None/float values that survive ``astype(str).str.*``
    chains (pandas preserves NaN through the str accessor) and
    normalises numeric strings like ``"85.0"`` to ``"85"``.
    """
    if x is None:
        return ""
    try:
        if isinstance(x, float) and pd.isna(x):
            return ""
    except Exception:  # noqa: BLE001
        pass
    s = str(x).strip()
    if not s or s.lower() == "nan":
        return ""
    if s.replace(".", "", 1).isdigit():
        try:
            return str(int(float(s)))
        except (ValueError, OverflowError):
            return s
    return s


def _map_pool_codes(series: pd.Series, pool_map: dict, split_char: str,
                    default_pool: str) -> pd.Series:
    if split_char:
        raw = series.astype(str).str.split(split_char).str[0].str.strip()
    else:
        raw = series.astype(str).str.strip()
    # Normalise "85.0" -> "85" for numeric codes (defensive against
    # NaN/float survivors).
    raw = raw.apply(_normalize_raw_code)
    pmap = {str(k): v for k, v in (pool_map or {}).items()}
    return raw.map(pmap).fillna(default_pool or "")


def canonical_pool_order(state: dict[str, Any]) -> list[str]:
    """Return pools in the order they'll appear in the report.

    Mirrors the ordering used by the Mgmt Adjustments step and the report
    engine: ``pool_settings`` (canonical WARM order), then ``warm.pools``,
    then any other pool names referenced by ``pool_map`` values.
    """
    out: list[str] = []
    seen: set[str] = set()

    def _add(name: Any) -> None:
        if not isinstance(name, str):
            return
        s = name.strip()
        if not s:
            return
        key = s.lower()
        if key in seen:
            return
        seen.add(key)
        out.append(s)

    for ps in (state.get("pool_settings") or []):
        _add((ps or {}).get("name"))
    for name in ((state.get("warm") or {}).get("pools") or []):
        _add(name)
    for v in (state.get("pool_map") or {}).values():
        _add(v)
    return out


def _loan_balances_from_db(state: dict[str, Any]) -> dict[str, Any] | None:
    """Fallback loan-side balances sourced from the imported ``monthly_loan_data``
    DB snapshot, for CUs whose wizard draft has no readable sample loan files
    (e.g. a script-built config where each extract entry's ``path`` is absent).

    Resolves the snapshot as the latest on/before the Report Period (or the
    latest available for the CU), sums ``current_balance`` per ``loan_pool``,
    and returns the same shape as :func:`loan_balances_by_pool`. Returns
    ``None`` when the CU/DB/snapshot can't be resolved so the caller keeps its
    own "no extracts" error.
    """
    cu = (state.get("credit_union") or "").strip()
    if not cu:
        return None
    try:
        from sqlalchemy import create_engine, text
        from cecl_credentials import get_database_url
        engine = create_engine(get_database_url())
    except Exception:  # noqa: BLE001
        return None
    cutoff = _report_period_cutoff(state)
    default_pool = (state.get("default_pool") or "Other/Uncategorized").strip()
    try:
        with engine.connect() as conn:
            snap = None
            if cutoff:
                snap = conn.execute(
                    text("SELECT MAX(snapshot_date) FROM monthly_loan_data "
                         "WHERE credit_union=:c AND snapshot_date<=:cut"),
                    {"c": cu, "cut": cutoff},
                ).scalar()
            if not snap:
                snap = conn.execute(
                    text("SELECT MAX(snapshot_date) FROM monthly_loan_data "
                         "WHERE credit_union=:c"),
                    {"c": cu},
                ).scalar()
            if not snap:
                return None
            rows = conn.execute(
                text("SELECT loan_pool, SUM(current_balance) AS t, "
                     "COUNT(*) AS n FROM monthly_loan_data "
                     "WHERE credit_union=:c AND snapshot_date=:s "
                     "GROUP BY loan_pool"),
                {"c": cu, "s": snap},
            ).fetchall()
    except Exception:  # noqa: BLE001
        return None
    if not rows:
        return None
    by_pool: dict[str, float] = {}
    default_total = 0.0
    grand_total = 0.0
    row_count = 0
    for pool, t, n in rows:
        name = str(pool or "").strip()
        if name in ("Ignore", "Exclude"):
            continue
        v = float(t or 0.0)
        row_count += int(n or 0)
        grand_total += v
        if name == default_pool:
            default_total += v
        else:
            by_pool[name] = by_pool.get(name, 0.0) + v
    snap_str = str(snap)[:10]
    return {
        "ok": True,
        "error": None,
        "by_pool": by_pool,
        "by_pool_code": {},
        "default_total": default_total,
        "file_count": 0,
        "row_count": row_count,
        "files": [{
            "name": f"Imported snapshot {snap_str} (monthly_loan_data)",
            "rows": row_count,
            "total": grand_total,
            "current_balance_col": "current_balance (DB)",
            "loan_pool_code_col": "loan_pool (DB)",
            "error": None,
        }],
        "unmapped_codes": [],
        "source": "db",
        "period": snap_str,
    }


def loan_balances_by_pool(state: dict[str, Any]) -> dict[str, Any]:
    """Sum ``current_balance`` per mapped pool across all loan-data extracts.

    Returns ``{ok, error, by_pool, by_pool_code, default_total, file_count,
    row_count, files: [{name, rows, total}], unmapped_codes: [...]}``.

    ``by_pool_code`` is ``{pool_name: [(raw_code, balance), ...]}`` sorted
    by descending balance, so the UI can render an expandable per-loan-code
    breakdown for each pool.
    """
    top_col_map = state.get("column_mappings") or {}
    top_has_header = bool(state.get("has_header", True))
    remove_chars = state.get("balance_remove_chars") or []
    accounting_negatives = bool(state.get("accounting_negatives", True))
    pool_map = state.get("pool_map") or {}
    top_split_char = state.get("pool_code_split") or ""
    default_pool = state.get("default_pool") or "Other/Uncategorized"

    # Collect candidate file entries: prefer the multi-upload list (so we
    # can honor per-file column_mappings / has_header / header_row), but
    # fall back to the single-sample saved_path if that's all the user
    # provided. Each tuple: (display_name, path, col_map, has_header,
    # header_row_or_None, file_split_char).
    #
    # Phase 9.22: ``file_split_char`` lets each extract override the
    # top-level ``pool_code_split``. When the entry has the key
    # ``pool_code_split`` (including ``""`` to mean "no split"), it wins;
    # otherwise the top-level value is inherited. CUMA mortgage extracts
    # with codes like ``15/15 ARM`` need an explicit per-file ``""`` so
    # the importer doesn't truncate to ``15``.
    files: list[tuple[str, str, dict, bool, int | None, str]] = []
    uploads = state.get("sample_uploads") or {}
    for entry in (uploads.get("loan_data_files") or []):
        n = entry.get("name") or Path(entry.get("path", "")).name
        p = entry.get("path") or ""
        if not p:
            continue
        entry_cm = entry.get("column_mappings") or {}
        cm = entry_cm if entry_cm else top_col_map
        hh = bool(entry.get("has_header", top_has_header)) \
            if "has_header" in entry else top_has_header
        # header_row is 1-indexed; some files (e.g. AIRES extracts) have a
        # leading numeric-position row, so the real headers live on row 2.
        hr_raw = entry.get("header_row")
        try:
            hr = int(hr_raw) if hr_raw not in (None, "", 0) else None
        except (TypeError, ValueError):
            hr = None
        if "pool_code_split" in entry:
            file_split = str(entry.get("pool_code_split") or "")
        else:
            file_split = top_split_char
        files.append((n, p, cm, hh, hr, file_split))
    if not files:
        sample = state.get("sample") or {}
        if sample.get("saved_path"):
            files.append((sample.get("filename") or "sample",
                          sample["saved_path"], top_col_map,
                          top_has_header, None, top_split_char))

    # Fail-fast if NO file (including fallback top-level) has the required
    # column mapping pair. Per-file misses degrade gracefully below.
    if files and not any(
        (cm.get("current_balance")
         and (cm.get("loan_pool_code") or cm.get("loan_pool_code_static")))
        for _, _, cm, _, _, _ in files
    ):
        return {"ok": False, "error":
                "Set both 'current_balance' and 'loan_pool_code' on the "
                "Column Mappings step first.",
                "by_pool": {}, "by_pool_code": {}, "default_total": 0.0,
                "file_count": 0, "row_count": 0, "files": [],
                "unmapped_codes": []}

    if not files:
        db = _loan_balances_from_db(state)
        if db is not None:
            return db
        return {"ok": False, "error":
                "No loan-data extracts have been uploaded yet.",
                "by_pool": {}, "by_pool_code": {}, "default_total": 0.0,
                "file_count": 0, "row_count": 0, "files": [],
                "unmapped_codes": []}

    by_pool: dict[str, float] = {}
    # by_pool_code[pool][raw_code] = running balance total.
    by_pool_code: dict[str, dict[str, float]] = {}
    default_total = 0.0
    total_rows = 0
    file_summaries: list[dict[str, Any]] = []
    unmapped: dict[str, float] = {}

    for name, path, col_map, has_header, header_row, file_split in files:
        bal_ref = col_map.get("current_balance")
        pool_ref = col_map.get("loan_pool_code")
        static_pool = (col_map.get("loan_pool_code_static") or "").strip()
        # The column references chosen for this file — surfaced back to
        # the UI so users can confirm at a glance which mapping each
        # file's total is being summed on. ``loan_pool_code_col`` may
        # carry a sentinel ``"(static: <code>)"`` when a fixed code is
        # being applied to every row in lieu of a real column.
        cb_col_display = str(bal_ref) if bal_ref else ""
        if static_pool:
            pool_col_display = f"(static: {static_pool})"
        else:
            pool_col_display = str(pool_ref) if pool_ref else ""
        if not bal_ref or (not pool_ref and not static_pool):
            file_summaries.append({"name": name, "rows": 0, "total": 0.0,
                                   "current_balance_col": cb_col_display,
                                   "loan_pool_code_col": pool_col_display,
                                   "error":
                                   "current_balance / loan_pool_code not "
                                   "mapped for this file"})
            continue
        df = _load_loan_extract(path, has_header, header_row)
        if df is None or df.empty:
            file_summaries.append({"name": name, "rows": 0, "total": 0.0,
                                   "current_balance_col": cb_col_display,
                                   "loan_pool_code_col": pool_col_display,
                                   "error": "Could not read file"})
            continue
        bal_series = _get_col(df, bal_ref, has_header)
        if static_pool:
            # Every row in this file uses the static raw code; it still
            # flows through pool_map so the user can map it to a pool name.
            pool_series = pd.Series([static_pool] * len(df), index=df.index)
        else:
            pool_series = _get_col(df, pool_ref, has_header)
        if bal_series is None or pool_series is None:
            file_summaries.append({"name": name, "rows": 0, "total": 0.0,
                                   "current_balance_col": cb_col_display,
                                   "loan_pool_code_col": pool_col_display,
                                   "error":
                                   "Required columns not found in file"})
            continue

        balances = _clean_balance(bal_series, remove_chars, accounting_negatives)
        mapped = _map_pool_codes(pool_series, pool_map, file_split, default_pool)
        # Treat NaN balance as 0 for grouping.
        balances = balances.fillna(0.0)

        df_calc = pd.DataFrame({"pool": mapped, "bal": balances})
        # Identify raw codes that fell through to default_pool so the user
        # can spot missing mapping rows.
        if file_split:
            raw_codes = pool_series.astype(str).str.split(file_split).str[0].str.strip()
        else:
            raw_codes = pool_series.astype(str).str.strip()
        raw_codes = raw_codes.apply(_normalize_raw_code)
        pmap_keys = {str(k) for k in (pool_map or {})}
        for code, pool_name, bal in zip(raw_codes, mapped, balances):
            b = float(bal or 0.0)
            if code and code not in pmap_keys:
                unmapped[code] = unmapped.get(code, 0.0) + b
            if code:
                bucket = by_pool_code.setdefault(pool_name or default_pool, {})
                bucket[code] = bucket.get(code, 0.0) + b

        sums = df_calc.groupby("pool")["bal"].sum()
        file_total = float(sums.sum())
        for pool_name, val in sums.items():
            v = float(val)
            if pool_name == default_pool:
                default_total += v
            else:
                by_pool[pool_name] = by_pool.get(pool_name, 0.0) + v
        total_rows += len(df_calc)
        file_summaries.append({"name": name, "rows": int(len(df_calc)),
                               "total": file_total,
                               "current_balance_col": cb_col_display,
                               "loan_pool_code_col": pool_col_display,
                               "error": None})

    # If every candidate file was unreadable (e.g. a script-built config
    # whose sample paths are absent, or archived-away uploads), fall back to
    # the imported DB snapshot so the reconciliation still shows real
    # per-pool loan balances instead of a blank column.
    if total_rows == 0 and not by_pool and not default_total:
        db = _loan_balances_from_db(state)
        if db is not None:
            return db

    # Sort each pool's code breakdown by descending balance.
    by_pool_code_sorted: dict[str, list[tuple[str, float]]] = {
        pool: sorted(codes.items(), key=lambda kv: -kv[1])
        for pool, codes in by_pool_code.items()
    }

    return {
        "ok": True,
        "error": None,
        "by_pool": by_pool,
        "by_pool_code": by_pool_code_sorted,
        "default_total": default_total,
        "file_count": len(files),
        "row_count": total_rows,
        "files": file_summaries,
        "unmapped_codes": sorted(unmapped.items(), key=lambda kv: -kv[1]),
    }


def monthly_balances_by_pool(state: dict[str, Any]) -> dict[str, Any]:
    """Pull per-pool balances for the latest period from the Monthly Balance
    file. Dispatches on the chosen ``monthly_bal.source`` mode:

    * ``single``    — one quarterly file with all months in column-band
                      layout (delegates to
                      ``monthly_bal_parser.pool_balances_for_latest_period``).
    * ``per_month`` — one balance-sheet file per month-end (delegates to
                      ``pool_balances_for_per_month_files`` and picks the
                      latest period).
    * ``manual``    — user-entered ``{pool: {YYYY-MM-DD: float}}`` grid.
    """
    mb = state.get("monthly_bal") or {}
    source = (mb.get("source") or "single").strip().lower()

    # ── per_year: one annual workbook per year w/ 12 month-end columns
    if source == "per_year":
        files = [
            e for e in (mb.get("year_files") or [])
            if e.get("saved_path")
        ]
        if not files:
            return {"ok": False, "error":
                    "No annual balance-sheet files have been added on "
                    "the Monthly Balance File step.",
                    "period": "", "by_pool": {}, "raw_rows": []}
        layout = mb.get("per_year_layout") or {}
        result = monthly_bal_parser.pool_balances_for_per_year_files(
            year_files=files,
            layout=layout,
            label_to_pool=mb.get("pool_map") or {},
            exclude_labels=mb.get("exclude_labels") or [],
        )
        by_period = result.get("by_period") or {}
        if not by_period:
            return {"ok": False,
                    "error": (result.get("error")
                              or "Could not extract any pool balances "
                              "from the annual workbook(s)."),
                    "period": "", "by_pool": {}, "raw_rows": []}
        # Skip future-period columns that exist in the layout but have
        # no data yet (e.g. running in May with a Jan-Dec sheet -
        # June-Dec columns are blank).
        populated = [
            p for p, b in by_period.items() if (b or {}).get("by_pool")
        ]
        if not populated:
            return {"ok": False,
                    "error": "Annual workbook(s) had period columns but "
                    "no balances mapped to pools. Check the pool map.",
                    "period": "", "by_pool": {}, "raw_rows": []}
        # Honour the global Report Period anchor: ignore any period
        # AFTER report_period (Step 1). Falls back to "latest with
        # data" when report_period is blank or the anchor is earlier
        # than every populated column.
        anchor = _report_period_cutoff(state)
        if anchor:
            in_range = [p for p in populated if p <= anchor]
            if in_range:
                populated = in_range
        latest = max(populated)
        bucket = by_period[latest] or {}
        return {"ok": True, "error": result.get("error"),
                "period": latest,
                "by_pool": bucket.get("by_pool") or {},
                "raw_rows": bucket.get("raw_rows") or []}

    # ── per_month: one balance-sheet file per month-end ───────────────
    if source == "per_month":
        files = [
            e for e in (mb.get("monthly_files") or [])
            if (e.get("saved_path") and e.get("period"))
        ]
        if not files:
            return {"ok": False, "error":
                    "No per-month balance-sheet files have been added on "
                    "the Monthly Balance File step.",
                    "period": "", "by_pool": {}, "raw_rows": []}
        layout = mb.get("per_month_layout") or {}
        result = monthly_bal_parser.pool_balances_for_per_month_files(
            monthly_files=files,
            layout=layout,
            label_to_pool=mb.get("pool_map") or {},
            exclude_labels=mb.get("exclude_labels") or [],
        )
        by_period = result.get("by_period") or {}
        if not by_period:
            return {"ok": False,
                    "error": (result.get("error")
                              or "Could not extract any pool balances from "
                              "the per-month files."),
                    "period": "", "by_pool": {}, "raw_rows": []}
        periods = list(by_period.keys())
        anchor = _report_period_cutoff(state)
        if anchor:
            in_range = [p for p in periods if p <= anchor]
            if in_range:
                periods = in_range
        latest = max(periods)
        bucket = by_period[latest] or {}
        return {"ok": True, "error": result.get("error"),
                "period": latest,
                "by_pool": bucket.get("by_pool") or {},
                "raw_rows": bucket.get("raw_rows") or []}

    # ── manual: user-entered grid ─────────────────────────────────────
    if source == "manual":
        entries = mb.get("manual_entries") or {}
        months = [m for m in (mb.get("manual_months") or []) if m]
        if not entries or not months:
            return {"ok": False, "error":
                    "No manual monthly balances have been entered on the "
                    "Monthly Balance File step.",
                    "period": "", "by_pool": {}, "raw_rows": []}
        anchor = _report_period_cutoff(state)
        candidate_months = months
        if anchor:
            in_range = [m for m in months if m <= anchor]
            if in_range:
                candidate_months = in_range
        latest = max(candidate_months)
        by_pool: dict[str, float] = {}
        raw_rows: list[dict[str, Any]] = []
        for pool, row in entries.items():
            val = (row or {}).get(latest)
            if val is None:
                continue
            try:
                f = float(val)
            except (TypeError, ValueError):
                continue
            by_pool[pool] = by_pool.get(pool, 0.0) + f
            raw_rows.append({"label": pool, "balance": f,
                             "mapped_pool": pool})
        return {"ok": True, "error": None, "period": latest,
                "by_pool": by_pool, "raw_rows": raw_rows}

    # ── single: legacy column-band quarterly file ─────────────────────
    saved = mb.get("saved_path")
    if not saved:
        return {"ok": False, "error":
                "No Monthly Balance file uploaded on the Monthly Balance "
                "File step.",
                "period": "", "by_pool": {}, "raw_rows": []}
    sheet = mb.get("sheet") or ""
    header_row = int(mb.get("header_row") or 0)
    pool_name_col = mb.get("pool_name_col") or ""
    if not sheet or not header_row or not pool_name_col:
        # Re-derive layout from the file on disk so a draft that was
        # saved before the layout fields were persisted (or whose
        # layout was inadvertently cleared) still works.
        recovered = monthly_bal_parser.analyse_file(saved)
        if recovered.get("ok"):
            sheet = sheet or recovered.get("sheet") or ""
            header_row = header_row or int(recovered.get("header_row") or 0)
            pool_name_col = pool_name_col or recovered.get(
                "pool_name_col") or "A"
    return monthly_bal_parser.pool_balances_for_latest_period(
        saved_path=saved,
        sheet=sheet,
        header_row=header_row,
        pool_name_col=pool_name_col,
        label_to_pool=mb.get("pool_map") or {},
        period=_report_period_cutoff(state) or None,
        exclude_labels=mb.get("exclude_labels") or [],
    )


def compare(state: dict[str, Any]) -> dict[str, Any]:
    """Build the side-by-side table the wizard renders.

    Returns::

        {
          "ok": bool,
          "error": str | None,
          "period": "YYYY-MM-DD" | "",
          "rows": [
            {"pool": str,
             "monthly": float | None,
             "loans": float | None,
             "diff": float | None,
             "pct": float | None},
            ...
          ],
          "totals": {"monthly": float, "loans": float, "diff": float},
          "loan_summary": {file_count, row_count, files, default_total,
                           unmapped_codes},
          "monthly_summary": {raw_rows: [...]},
        }
    """
    monthly = monthly_balances_by_pool(state)
    loans = loan_balances_by_pool(state)

    err_parts: list[str] = []
    if not monthly.get("ok"):
        err_parts.append(monthly.get("error") or "Monthly balance read failed")
    if not loans.get("ok"):
        err_parts.append(loans.get("error") or "Loan extract read failed")

    monthly_pools = monthly.get("by_pool") or {}
    loan_pools = loans.get("by_pool") or {}
    by_pool_code = loans.get("by_pool_code") or {}

    # Build the canonical pool list, then append any extras that showed up
    # only in the data files (case-insensitive de-dupe).
    ordered = canonical_pool_order(state)
    seen = {p.lower() for p in ordered}
    for extra in sorted(set(monthly_pools) | set(loan_pools),
                        key=lambda s: s.lower()):
        if extra and extra.lower() not in seen:
            seen.add(extra.lower())
            ordered.append(extra)

    # Case-insensitive lookups for incoming data.
    monthly_lc = {k.lower(): (k, v) for k, v in monthly_pools.items()}
    loans_lc = {k.lower(): (k, v) for k, v in loan_pools.items()}
    codes_lc = {k.lower(): v for k, v in by_pool_code.items()}

    rows: list[dict[str, Any]] = []
    total_m = 0.0
    total_l = 0.0
    for pool in ordered:
        key = pool.lower()
        m = monthly_lc.get(key, (pool, None))[1]
        l = loans_lc.get(key, (pool, None))[1]
        diff = None
        pct = None
        if m is not None and l is not None:
            diff = l - m
            if m:
                pct = (diff / m) * 100.0
        rows.append({
            "pool": pool,
            "monthly": m,
            "loans": l,
            "diff": diff,
            "pct": pct,
            "loan_codes": codes_lc.get(key, []),
        })
        total_m += float(m or 0.0)
        total_l += float(l or 0.0)

    return {
        "ok": not err_parts,
        "error": "; ".join(err_parts) if err_parts else None,
        "period": monthly.get("period") or "",
        "rows": rows,
        "totals": {
            "monthly": total_m,
            "loans": total_l,
            "diff": total_l - total_m,
        },
        "loan_summary": {
            "file_count": loans.get("file_count", 0),
            "row_count": loans.get("row_count", 0),
            "files": loans.get("files", []),
            "default_total": loans.get("default_total", 0.0),
            "unmapped_codes": loans.get("unmapped_codes", []),
        },
        "monthly_summary": {
            "raw_rows": monthly.get("raw_rows", []),
        },
    }


def compare_run(cfg: dict[str, Any], snapshot_iso: str,
                short_name: str | None = None) -> dict[str, Any]:
    """Per-pool balance comparison for the Run-New-Quarter intercept.

    Reads loan balances from the ``monthly_loan_data`` DB table (the
    freshly-imported snapshot) and monthly balances via the same
    ``generate_report.load_monthly_balances`` loader the report engine
    uses (so the figures shown match what the reports will use).

    Returns the same dict shape as :func:`compare` so the existing
    Step-14 template patterns translate cleanly.
    """
    cu = (cfg.get("credit_union") or "").strip()
    if not cu or not snapshot_iso:
        return {"ok": False,
                "error": "Missing credit_union or snapshot date.",
                "period": snapshot_iso or "",
                "monthly_period": "",
                "rows": [],
                "totals": {"monthly": 0.0, "loans": 0.0, "diff": 0.0},
                "loan_summary": {"file_count": 0, "row_count": 0,
                                 "files": [], "default_total": 0.0,
                                 "unmapped_codes": []},
                "monthly_summary": {"raw_rows": []}}

    # ── Loan side: post-import DB snapshot ───────────────────────────
    try:
        from sqlalchemy import create_engine, text
        from cecl_credentials import get_database_url
        engine = create_engine(get_database_url())
        df = pd.read_sql(
            text("SELECT loan_pool, current_balance "
                 "FROM monthly_loan_data "
                 "WHERE credit_union=:c AND snapshot_date=:s"),
            engine,
            params={"c": cu, "s": snapshot_iso},
        )
    except Exception as exc:  # noqa: BLE001
        return {"ok": False,
                "error": f"Could not read loans from DB: {exc}",
                "period": snapshot_iso,
                "monthly_period": "",
                "rows": [],
                "totals": {"monthly": 0.0, "loans": 0.0, "diff": 0.0},
                "loan_summary": {"file_count": 0, "row_count": 0,
                                 "files": [], "default_total": 0.0,
                                 "unmapped_codes": []},
                "monthly_summary": {"raw_rows": []}}

    if df is None or df.empty:
        return {"ok": False,
                "error": (f"No imported loans found for {cu} at "
                          f"{snapshot_iso}. Import may have failed."),
                "period": snapshot_iso,
                "monthly_period": "",
                "rows": [],
                "totals": {"monthly": 0.0, "loans": 0.0, "diff": 0.0},
                "loan_summary": {"file_count": 0, "row_count": 0,
                                 "files": [], "default_total": 0.0,
                                 "unmapped_codes": []},
                "monthly_summary": {"raw_rows": []}}

    loan_by_pool: dict[str, float] = (
        df.groupby("loan_pool")["current_balance"].sum().to_dict()
    )
    # Drop intentional excludes from the comparison surface.
    for excl in ("Ignore", "Exclude"):
        loan_by_pool.pop(excl, None)
    loan_row_count = int(len(df))

    # ── Monthly side: report-engine loader (handles all 4 modes) ────
    monthly_by_pool: dict[str, float] = {}
    monthly_period = ""
    monthly_err: str | None = None
    # When the CU has no monthly balance by pool/type for the snapshot
    # month, show zeros for the monthly balance instead of comparing the
    # current loan extract against a stale prior month.
    no_monthly_data = False
    try:
        import generate_report  # local import keeps service decoupled
        mb_df, _alll = generate_report.load_monthly_balances(cfg)
        if mb_df is not None and not mb_df.empty \
                and {"pool", "date", "balance"}.issubset(mb_df.columns):
            mb_df = mb_df.copy()
            mb_df["date"] = pd.to_datetime(mb_df["date"], errors="coerce")
            mb_df = mb_df.dropna(subset=["date"])
            snap_ts = pd.to_datetime(snapshot_iso)
            same = mb_df[mb_df["date"].dt.normalize() == snap_ts.normalize()]
            if same.empty:
                # No monthly balance by pool/type for the snapshot month.
                # Do NOT fall back to the most recent prior period —
                # comparing the current loan extract against a stale month
                # (e.g. Dec balances vs Mar loans) is misleading. Show zeros
                # for the monthly balance with no per-pool difference.
                no_monthly_data = True
                monthly_period = ""
                monthly_err = (
                    "No monthly balance by pool/type for "
                    f"{snapshot_iso} — showing zeros for the monthly balance "
                    "(no prior-period comparison)."
                )
            else:
                monthly_period = snapshot_iso
                monthly_by_pool = (
                    same.groupby("pool")["balance"].sum().to_dict()
                )
        else:
            monthly_err = (
                "Monthly balance file did not yield any per-pool data."
            )
    except Exception as exc:  # noqa: BLE001
        monthly_err = f"Monthly balance read failed: {exc}"

    # Restrict the comparison to real loan pools. The Monthly Balance file
    # is often a full GL balance sheet (especially in ``per_month`` mode):
    # its non-loan line items (ACH Clearing, ATM Machine, Christmas Clubs,
    # Accrued Interest, CECL, ...) fall through the balance ``pool_map`` and
    # would otherwise show up as bogus pools here. Mirror the report
    # engine's ``load_historical_data`` filter: keep only pools this CU has
    # configured (``pools`` / ``pool_order`` / ``not_risk_rated`` / the
    # Monthly-Balance ``pool_map`` targets) or that actually carry imported
    # loans this snapshot.
    configured_lc: set[str] = set()
    for _p in (cfg.get("pools") or []):
        _n = _p.get("name") if isinstance(_p, dict) else _p
        if _n:
            configured_lc.add(str(_n).strip().lower())
    for _n in (cfg.get("pool_order") or []):
        if _n:
            configured_lc.add(str(_n).strip().lower())
    for _n in (cfg.get("not_risk_rated") or []):
        if _n:
            configured_lc.add(str(_n).strip().lower())
    for _v in ((cfg.get("monthly_balance") or {}).get("pool_map") or {}).values():
        if _v:
            configured_lc.add(str(_v).strip().lower())
    if configured_lc:
        loan_pools_lc = {str(k).strip().lower() for k in loan_by_pool}
        dropped = [
            k for k in monthly_by_pool
            if str(k).strip().lower() not in configured_lc
            and str(k).strip().lower() not in loan_pools_lc
        ]
        if dropped:
            monthly_by_pool = {
                k: v for k, v in monthly_by_pool.items()
                if k not in set(dropped)
            }
            print(f"    Balance check: dropped {len(dropped)} non-loan "
                  f"balance-sheet line item(s): "
                  f"{', '.join(sorted(dropped)[:6])}"
                  f"{'…' if len(dropped) > 6 else ''}")

    # ── Build rows ──────────────────────────────────────────────────
    # Order pools to match the reports rather than alphabetically. Mirror
    # generate_report's canonical ordering: pools listed in
    # ``config['pool_order']`` come first in their declared order, then any
    # extras alphabetically; within that, risk-rated pools precede
    # not-risk-rated pools (same as the report sheets).
    cfg_order = cfg.get("pool_order", []) or []
    nrr = {str(n).strip() for n in (cfg.get("not_risk_rated", []) or [])}
    order_idx = {name: i for i, name in enumerate(cfg_order)}
    fallback = len(cfg_order)
    pool_names = set(loan_by_pool.keys()) | set(monthly_by_pool.keys())
    rr_pools = [p for p in pool_names if str(p).strip() not in nrr]
    nrr_pools = [p for p in pool_names if str(p).strip() in nrr]
    rr_pools.sort(key=lambda p: (order_idx.get(p, fallback), str(p).lower()))
    nrr_pools.sort(key=lambda p: (order_idx.get(p, fallback), str(p).lower()))
    all_pools = rr_pools + nrr_pools
    rows: list[dict[str, Any]] = []
    total_m = 0.0
    total_l = 0.0
    for pool in all_pools:
        m_val = monthly_by_pool.get(pool)
        l_val = loan_by_pool.get(pool)
        m_f = float(m_val) if m_val is not None else None
        l_f = float(l_val) if l_val is not None else None
        if no_monthly_data:
            # Zeros on the monthly balance; suppress the difference so we do
            # not present a comparison against a period we do not have.
            m_f = 0.0
        diff = None
        pct = None
        if not no_monthly_data and m_f is not None and l_f is not None:
            diff = l_f - m_f
            if m_f:
                pct = (diff / m_f) * 100.0
        rows.append({
            "pool": pool,
            "monthly": m_f,
            "loans": l_f,
            "diff": diff,
            "pct": pct,
            "loan_codes": [],
        })
        total_m += float(m_f or 0.0)
        total_l += float(l_f or 0.0)

    # Per-pool per-code breakdown — derived by re-reading the loan
    # extracts in Raw_Uploads via the same engine used in wizard Step 14.
    # Used by the run-flow inline pool-mapping editor so users can see
    # which raw loan codes currently sit in each pool and reassign them
    # without leaving the Balance Adjustment page.
    try:
        adapter_state = _build_state_for_run(cfg, short_name,
                                             snapshot_iso=snapshot_iso)
        by_code = loan_balances_by_pool(adapter_state)
        if by_code.get("ok"):
            code_map = by_code.get("by_pool_code") or {}
            for r in rows:
                r["loan_codes"] = list(code_map.get(r["pool"]) or [])
    except Exception:  # noqa: BLE001
        pass

    return {
        "ok": True,
        "error": monthly_err,
        "period": snapshot_iso,
        "monthly_period": monthly_period,
        "rows": rows,
        "totals": {
            "monthly": total_m,
            "loans": total_l,
            "diff": 0.0 if no_monthly_data else (total_l - total_m),
        },
        "loan_summary": {
            "file_count": 0,
            "row_count": loan_row_count,
            "files": [],
            "default_total": 0.0,
            "unmapped_codes": [],
        },
        "monthly_summary": {"raw_rows": []},
    }


def _build_state_for_run(cfg: dict[str, Any],
                         short_name: str | None,
                         snapshot_iso: str | None = None) -> dict[str, Any]:
    """Construct a minimal state-like dict suitable for
    :func:`loan_balances_by_pool` from a runtime YAML ``cfg`` plus the
    staged ``Raw_Uploads/<short>/`` folder.

    Walks the workspace's ``Raw_Uploads/<short_name>/`` directory and
    matches each spreadsheet against the configured ``file_pattern``
    (per-extract first, then the top-level fallback). The resulting
    ``sample_uploads.loan_data_files`` list mirrors the wizard's shape
    so the existing per-code aggregation works without modification.

    When ``snapshot_iso`` is provided, each candidate file's filename is
    further filtered via :func:`import_data.extract_snapshot_date` so
    only files belonging to the requested snapshot contribute to the
    per-code breakdown. This prevents per-code totals from being inflated
    by historical snapshots that linger in ``Archive/<short>/`` after
    prior quarters' imports. Files whose date can't be parsed are
    INCLUDED (fallback: single-period CUs or undated filenames).
    """
    from cecl_ui.services import config_service  # local: avoid cycles
    raw_dir = None
    archive_dir = None
    if short_name:
        try:
            from flask import current_app
            ws = current_app.config["WORKSPACE_ROOT"]
            raw_dir = config_service.raw_uploads_dir(ws) / short_name
            # Phase 9.35f: also scan Archive/<short>/ so the per-code
            # breakdown still works AFTER import has archived the files.
            from pathlib import Path
            archive_dir = Path(ws) / "Archive" / short_name
        except Exception:  # noqa: BLE001 — best-effort outside Flask
            raw_dir = None
            archive_dir = None

    loan_files: list[dict[str, Any]] = []
    # Compile patterns once.
    per_extract: list[tuple[re.Pattern, dict]] = []
    for ex in (cfg.get("loan_data_extracts") or []):
        pat = (ex or {}).get("file_pattern") or ""
        if not pat:
            continue
        try:
            per_extract.append((re.compile(pat), ex or {}))
        except re.error:
            continue
    top_pat = cfg.get("file_pattern") or ""
    try:
        top_rx = re.compile(top_pat) if top_pat else None
    except re.error:
        top_rx = None

    allowed = {".xlsx", ".xlsm", ".xls", ".csv"}
    seen_names: set[str] = set()

    # Snapshot filter setup. ``extract_snapshot_date`` needs ``date_pattern``
    # in the config; absence falls back to filename hints inside the func.
    snap_filter_active = bool(snapshot_iso)
    _extract_date = None
    if snap_filter_active:
        try:
            from import_data import extract_snapshot_date as _extract_date
        except Exception:  # noqa: BLE001 — outside Flask / missing deps
            _extract_date = None
            snap_filter_active = False
    # extract_snapshot_date requires a `date_pattern`; synthesize a
    # liberal default when the YAML doesn't carry one so the function's
    # name-based fallbacks (month names, MMDDYYYY, etc.) still fire.
    snap_cfg = dict(cfg)
    snap_cfg.setdefault("date_pattern", r"(\d{4})[-_](\d{2})")

    def _norm(name: str) -> str:
        # Normalize underscore/space + case so the Phase 9.35 auto-restore
        # copies (which secure_filename'd spaces into underscores) don't
        # double-count next to the original filenames.
        return re.sub(r"[\s_]+", " ", name).strip().lower()

    def _collect_from(dir_path) -> None:
        if not dir_path or not dir_path.is_dir():
            return
        # rglob handles Archive's nested-per-extract subdirs cleanly
        # and is a no-op extra cost on flat Raw_Uploads layouts.
        for entry in dir_path.rglob("*"):
            if not entry.is_file() or entry.name.startswith("~$"):
                continue
            if entry.suffix.lower() not in allowed:
                continue
            key = _norm(entry.name)
            if key in seen_names:
                continue
            matched_ex: dict | None = None
            for rx, ex in per_extract:
                if rx.search(entry.name):
                    matched_ex = ex
                    break
            if matched_ex is None and top_rx is not None and \
                    not top_rx.search(entry.name):
                continue
            # Phase 9.40: skip files that belong to a DIFFERENT snapshot
            # so per-code breakdowns reflect only the requested period.
            # Undated filenames (extract returns None) are kept so
            # single-period CUs still get per-code data.
            if snap_filter_active and _extract_date is not None:
                try:
                    file_snap = _extract_date(entry.name, snap_cfg)
                except Exception:  # noqa: BLE001
                    file_snap = None
                if file_snap and file_snap != snapshot_iso:
                    continue
            ex_src = matched_ex or {}
            file_entry: dict[str, Any] = {
                "name": entry.name,
                "path": str(entry),
            }
            if "column_mappings" in ex_src:
                file_entry["column_mappings"] = ex_src.get(
                    "column_mappings") or {}
            if "has_header" in ex_src:
                file_entry["has_header"] = ex_src.get("has_header")
            if "header_row" in ex_src:
                file_entry["header_row"] = ex_src.get("header_row")
            if "pool_code_split" in ex_src:
                file_entry["pool_code_split"] = ex_src.get(
                    "pool_code_split") or ""
            # Phase 9.24c parity — per-extract member_account override
            # (e.g. CUMA Mortgage vs Symitar ceclXX shape). Needed so
            # impaired_parser._build_loan_index can build the correct
            # member-suffix key per file.
            if "member_account" in ex_src:
                file_entry["member_account"] = ex_src.get(
                    "member_account") or {}
            seen_names.add(key)
            loan_files.append(file_entry)

    # Raw_Uploads first so any duplicate name in Archive is skipped.
    _collect_from(raw_dir)
    _collect_from(archive_dir)

    return {
        "column_mappings": cfg.get("column_mappings") or {},
        "has_header": bool(cfg.get("has_header", True)),
        "pool_map": cfg.get("pool_map") or {},
        "pool_code_split": cfg.get("pool_code_split") or "",
        "default_pool": cfg.get("default_pool") or "Other/Uncategorized",
        "balance_remove_chars": cfg.get("balance_remove_chars") or [],
        "accounting_negatives": bool(cfg.get("accounting_negatives", True)),
        # Phase 9.37b — the run-time impaired verification page relies on
        # these to (a) classify credit_grade correctly instead of falling
        # back to "Not Reported" for every row, and (b) build the correct
        # member-suffix key so loan-extract enrichment actually matches.
        # Without them, impaired_parser._build_loan_index / _grade_for
        # silently degrade to the empty-grades and default fixed_suffix
        # paths, producing 100% unmatched rows with credit_grade blank.
        "credit_grades": cfg.get("credit_grades") or [],
        "no_score_label": cfg.get("no_score_label") or "Not Reported",
        "member_account": cfg.get("member_account") or {},
        "sample_uploads": {"loan_data_files": loan_files},
    }
