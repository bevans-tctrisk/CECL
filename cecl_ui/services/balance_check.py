"""Balance Adjustment helper.

Compares the per-pool balance recorded in the recurring Monthly Balance
file (Step "Monthly Balance File") against the sum of ``current_balance``
in the user's Loan Data Extract(s), grouped by the same pool name. The
wizard renders the result as a side-by-side table so the user can spot
mis-mapped pool codes / balance-format issues before generating reports.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from cecl_ui.services import monthly_bal_parser


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
        df.columns = [str(c).strip() for c in df.columns]
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


def _map_pool_codes(series: pd.Series, pool_map: dict, split_char: str,
                    default_pool: str) -> pd.Series:
    if split_char:
        raw = series.astype(str).str.split(split_char).str[0].str.strip()
    else:
        raw = series.astype(str).str.strip()
    # Normalise "85.0" -> "85" for numeric codes.
    raw = raw.apply(
        lambda x: str(int(float(x))) if x.replace(".", "", 1).isdigit() else x
    )
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
    split_char = state.get("pool_code_split") or ""
    default_pool = state.get("default_pool") or "Other/Uncategorized"

    # Collect candidate file entries: prefer the multi-upload list (so we
    # can honor per-file column_mappings / has_header / header_row), but
    # fall back to the single-sample saved_path if that's all the user
    # provided. Each tuple: (display_name, path, col_map, has_header,
    # header_row_or_None).
    files: list[tuple[str, str, dict, bool, int | None]] = []
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
        files.append((n, p, cm, hh, hr))
    if not files:
        sample = state.get("sample") or {}
        if sample.get("saved_path"):
            files.append((sample.get("filename") or "sample",
                          sample["saved_path"], top_col_map,
                          top_has_header, None))

    # Fail-fast if NO file (including fallback top-level) has the required
    # column mapping pair. Per-file misses degrade gracefully below.
    if files and not any(
        (cm.get("current_balance")
         and (cm.get("loan_pool_code") or cm.get("loan_pool_code_static")))
        for _, _, cm, _, _ in files
    ):
        return {"ok": False, "error":
                "Set both 'current_balance' and 'loan_pool_code' on the "
                "Column Mappings step first.",
                "by_pool": {}, "by_pool_code": {}, "default_total": 0.0,
                "file_count": 0, "row_count": 0, "files": [],
                "unmapped_codes": []}

    if not files:
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

    for name, path, col_map, has_header, header_row in files:
        bal_ref = col_map.get("current_balance")
        pool_ref = col_map.get("loan_pool_code")
        static_pool = (col_map.get("loan_pool_code_static") or "").strip()
        if not bal_ref or (not pool_ref and not static_pool):
            file_summaries.append({"name": name, "rows": 0, "total": 0.0,
                                   "error":
                                   "current_balance / loan_pool_code not "
                                   "mapped for this file"})
            continue
        df = _load_loan_extract(path, has_header, header_row)
        if df is None or df.empty:
            file_summaries.append({"name": name, "rows": 0, "total": 0.0,
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
                                   "error":
                                   "Required columns not found in file"})
            continue

        balances = _clean_balance(bal_series, remove_chars, accounting_negatives)
        mapped = _map_pool_codes(pool_series, pool_map, split_char, default_pool)
        # Treat NaN balance as 0 for grouping.
        balances = balances.fillna(0.0)

        df_calc = pd.DataFrame({"pool": mapped, "bal": balances})
        # Identify raw codes that fell through to default_pool so the user
        # can spot missing mapping rows.
        if split_char:
            raw_codes = pool_series.astype(str).str.split(split_char).str[0].str.strip()
        else:
            raw_codes = pool_series.astype(str).str.strip()
        raw_codes = raw_codes.apply(
            lambda x: str(int(float(x))) if x.replace(".", "", 1).isdigit() else x
        )
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
                               "total": file_total, "error": None})

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
        )
        by_period = result.get("by_period") or {}
        if not by_period:
            return {"ok": False,
                    "error": (result.get("error")
                              or "Could not extract any pool balances from "
                              "the per-month files."),
                    "period": "", "by_pool": {}, "raw_rows": []}
        latest = max(by_period.keys())
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
        latest = max(months)
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
