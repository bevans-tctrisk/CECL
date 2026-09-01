"""Derive the charge-off split by credit-grade migration status.

Background
----------
The "Charge off by Credit Grade Migration" bar on every ``Risk Change`` /
``Risk Chg <pool>`` tab is fed from ``hist['impaired']['co_by_status']`` and
``hist['impaired']['co_by_pool']``.  Historically those two keys had exactly
one producer -- ``generate_report.load_impaired_data``, which reads the
``CO Data Entry`` tab out of a legacy ~80-tab CECL-Migration-WARM workbook.
Credit unions onboarded through the wizard never had such a workbook, so the
dict stayed empty and ``_sheet_risk_change`` wrote literal zeros
(see ``docs/pdf_migration/04_blank_charts.md``).

The split is a comparison of two credit scores for the same loan: the score
at ORIGINATION and the score AT / NEAR THE CHARGE-OFF EVENT.  Neither is in
the charge-off feed for most credit unions -- the CO upload carries account,
loan code, amount and date and nothing else.  ``07_chargeoff_feasibility.md``
concluded from that the split was "not derivable".

It is derivable, from a source that assessment did not consider: the credit
union's own **loan-snapshot history**.  ``monthly_loan_data`` keeps one row
per loan per month with ``current_fico_score`` / ``original_fico_score``,
keyed on the same full-account string the charge-off file carries.  A loan
that charges off in month Q was open, scored and in the extract in the months
before Q.  So:

* **score at charge-off** = ``current_fico_score`` from the LATEST snapshot at
  or before the charge-off date in which the loan still appears;
* **original score**      = ``original_fico_score`` from the EARLIEST snapshot
  in which the loan appears.  Where the credit union supplies a genuine
  origination score this is that score; where ``import_data`` gap-filled
  ``original = current``, it is the loan's earliest OBSERVED score, which is
  the best available proxy and is exactly what accrues in value as more
  quarters are imported.

Three sources are tried in order, best first:

1. ``CO Data Entry`` from the WARM -- untouched; this module never runs when
   the WARM supplied the keys.
2. **Score columns on the charge-off file itself.**  Some cores export them
   (Utah Community FCU's ``2026Q2_tct_chargeoff.csv`` carries
   ``APPLICATION_CREDITSCORE`` / ``CREDITSCORE_MOSTRECENT``).  When
   ``historical_file_formats.chargeoff`` names ``orig_score_col`` /
   ``curr_score_col`` those win -- they are the same numbers the WARM analyst
   pastes into ``CO Data``.
3. **The snapshot lookback join** described above.

Nothing here writes to the database and nothing here overrides a
WARM-supplied value.

Everything below is pure: no schema change, no persistence.  The history it
reads is already accumulating in ``monthly_loan_data`` on every import, so
the same charge-off resolves better each quarter without a backfill.
"""

from __future__ import annotations

import os
import re

STATUS_LABELS = ("Improved", "Deteriorated", "Unchanged", "Not Reported")

# A charge-off whose score we cannot recover lands in ``Not Reported`` -- the
# WARM's own bucket for "this loan had no credit score".  When that bucket
# swallows more than this share of the charge-off dollars the chart carries no
# information, so the derivation refuses and the caller leaves the block empty
# for the report's empty-state to handle.
MAX_NOT_REPORTED_SHARE = 0.75

# A charge-off whose original score was never supplied is classified
# ``Unchanged`` under the WARM's ``original = current`` gap-fill convention --
# the same convention ``import_data`` applies to live loans and the Risk Change
# matrix displays.  That is defensible per loan, but a chart whose bars are
# mostly that assumption asserts "credit did not deteriorate before these loans
# charged off", which is a specific claim nobody measured.  At least this share
# of the charge-off dollars must carry a genuinely measured pair of scores.
MIN_MEASURED_SHARE = 0.25

# Provenances that represent two independently sourced scores.  ``co_file``
# is both scores off the charge-off feed; ``row_pair`` is the origination /
# current pair the loan carried on its last snapshot; ``lookback_span`` is two
# genuinely different observations of the same loan over time.  Deliberately
# NOT included: ``co_file_current`` (one score only, so the loan lands in
# ``Not Reported`` anyway) and ``gap_fill`` (an assumption).
MEASURED_PROVENANCE = ("co_file", "row_pair", "lookback_span",
                       "co_file_orig+gap_fill", "co_file_orig+row_pair",
                       "co_file_orig+lookback_span",
                       "gap_fill+co_file_current", "row_pair+co_file_current",
                       "lookback_span+co_file_current")

# The charge-off population the chart covers: a trailing window ending on the
# report date.  Three years is not a guess -- Utah Community FCU's WARM
# ``CO Data Entry`` matrix reproduces to the penny, in total and in all six
# live pools, as "every charge-off on the ``CO Data`` tab dated 2023-07 or
# later, excluding the Exclude and HIDE-* pools", against a 2026-06 report
# date.  That is exactly 36 months.  (``07_chargeoff_feasibility.md`` read the
# $28.95M as an ~11-year cumulative figure; it is not.)
DEFAULT_LOOKBACK_YEARS = 3


# ---------------------------------------------------------------------------
# Charge-off file reading -- per LOAN, not per loan code
# ---------------------------------------------------------------------------

def _clean_id_series(s):
    import pandas as pd  # noqa: F401  (kept for parity with import_data)
    out = s.astype(str).str.strip()
    out = out.str.replace(r"\.0$", "", regex=True)
    return out


def _derive_full_account(cfg_df, parse_cfg):
    """Mirror ``import_data.derive_member_account`` for a CO parse-config block.

    The charge-off block carries the same ``member_account`` shape the loan
    extract does (``mode`` = fixed_suffix / delimiter / split), with
    ``member_col`` standing in for ``member_number`` and ``account_col`` for
    ``loan_suffix``.  Producing the string the same way is what makes the
    charge-off row joinable to ``monthly_loan_data.member_number``.
    """
    import pandas as pd

    ma = parse_cfg.get("member_account") or {}
    mode = ma.get("mode") or "fixed_suffix"
    acc_i = parse_cfg.get("account_col")
    mem_i = parse_cfg.get("member_col")
    ncols = cfg_df.shape[1]

    def _col(i):
        return _clean_id_series(cfg_df.iloc[:, int(i)])

    def _ok(i):
        return i is not None and 0 <= int(i) < ncols

    if mode == "split" and _ok(mem_i) and _ok(acc_i) and int(mem_i) != int(acc_i):
        member = _col(mem_i)
        suffix = _col(acc_i)
        raw_len = ma.get("suffix_length")
        try:
            pad = 3 if raw_len is None else int(raw_len)
        except (TypeError, ValueError):
            pad = 3
        if pad > 0:
            suffix = suffix.str.zfill(pad)
        return member + suffix

    base = acc_i if _ok(acc_i) else mem_i
    if not _ok(base):
        return pd.Series([""] * len(cfg_df), index=cfg_df.index)
    raw = _col(base)
    if mode == "delimiter":
        delim = ma.get("delimiter") or "-"
        parts = raw.str.split(re.escape(delim), n=1, expand=True, regex=True)
        left = parts[0].fillna(raw)
        if parts.shape[1] > 1:
            return left + parts[1].fillna("")
        return left
    return raw


def _parse_co_file_loans(filepath, parse_cfg):
    """One charge-off file -> ``DataFrame[account, code, amount, date,
    orig_score, curr_score]``.

    Column resolution (header-text override, account-row filter, strict
    columns) is copied from ``generate_report._parse_chargeoff_file`` so the
    same rows are selected; the only difference is that the account column is
    *kept* instead of being used purely as a row filter.
    """
    import pandas as pd
    import generate_report as gr

    empty = pd.DataFrame(columns=["account", "code", "amount", "date",
                                  "orig_score", "curr_score"])
    df = gr._read_data_file(filepath)
    if df is None or df.empty:
        return empty

    cfg_df = df.copy()
    if parse_cfg.get("has_header", False):
        cfg_df = cfg_df.iloc[1:]
    skip = int(parse_cfg.get("skip_rows", 0) or 0)
    if skip > 0:
        cfg_df = cfg_df.iloc[skip:]
    if cfg_df.empty:
        return empty

    account_col = parse_cfg.get("account_col", 0)
    code_col = parse_cfg.get("code_col")
    amount_col = parse_cfg.get("amount_col")
    date_col = parse_cfg.get("date_col")
    if parse_cfg.get("has_header", False) and not parse_cfg.get("strict_columns"):
        hdr = gr._resolve_hist_cols_by_header(df.iloc[0], "co")
        if "amount" in hdr and "code" in hdr:
            account_col = hdr.get("account", account_col)
            code_col = hdr["code"]
            amount_col = hdr["amount"]
            if "date" in hdr:
                date_col = hdr["date"]

    ncols = cfg_df.shape[1]

    def _ok(i):
        return i is not None and str(i) != "" and 0 <= int(i) < ncols

    if not _ok(amount_col):
        return empty
    if _ok(account_col):
        cfg_df = cfg_df[cfg_df.iloc[:, int(account_col)].apply(gr._is_numeric_or_date)]
    if cfg_df.empty:
        return empty

    local = dict(parse_cfg)
    local["account_col"] = account_col
    out = pd.DataFrame({
        "account": _derive_full_account(cfg_df, local).values,
        "code": (cfg_df.iloc[:, int(code_col)].values if _ok(code_col)
                 else [parse_cfg.get("code_static")] * len(cfg_df)),
        "amount": pd.to_numeric(cfg_df.iloc[:, int(amount_col)],
                                errors="coerce").values,
    })
    if _ok(date_col):
        out["date"] = gr._coerce_mixed_dates(cfg_df.iloc[:, int(date_col)].values).values
    else:
        out["date"] = pd.NaT

    # Optional: the charge-off feed carries the scores itself (Utah-style).
    orig_col = parse_cfg.get("orig_score_col")
    curr_col = parse_cfg.get("curr_score_col")
    if (orig_col is None and curr_col is None
            and parse_cfg.get("has_header", False)
            and parse_cfg.get("auto_detect_score_columns")):
        orig_col, curr_col = detect_score_columns(df.iloc[0].tolist())
        if orig_col is not None or curr_col is not None:
            print(f"    CO migration split: detected score column(s) on "
                  f"{os.path.basename(str(filepath))} -- "
                  f"original={orig_col}, at-charge-off={curr_col}.")
    for field, col in (("orig_score", orig_col), ("curr_score", curr_col)):
        if _ok(col):
            out[field] = pd.to_numeric(cfg_df.iloc[:, int(col)],
                                       errors="coerce").values
        else:
            out[field] = float("nan")

    out = out.dropna(subset=["amount"])
    return out[out["amount"] != 0]


# Header labels that identify a score column on a charge-off file.  Ordered:
# the first pattern that matches a header claims that column.
_ORIG_SCORE_PAT = re.compile(
    r"(orig|application|at\s*loan\s*orig|app[_ ]?cr)", re.IGNORECASE)
_CURR_SCORE_PAT = re.compile(
    r"(most[\s_]*recent|current|at\s*charge|hard)", re.IGNORECASE)
_ANY_SCORE_PAT = re.compile(r"(fico|credit[\s_]*score|creditscore|score)",
                            re.IGNORECASE)
_NOT_SCORE_PAT = re.compile(r"(date|amount|amt|balance|rate|count)",
                            re.IGNORECASE)


def detect_score_columns(header_row):
    """Suggest ``(orig_score_col, curr_score_col)`` from a header row.

    Charge-off feeds routinely carry the scores the chart needs under headers
    nobody has mapped -- ``APPLICATION_CREDITSCORE`` /
    ``CREDITSCORE_MOSTRECENT`` (Utah Community FCU), ``FICO Score at Loan
    Orig`` (WNC, NOVA, Jackson River), ``CreditScore`` (Shuford), ``FICO``
    (SCI, United Community).  This is advisory: it feeds the wizard's
    suggestion and this module's diagnostics.  It is only *used* when a
    credit union opts in with ``chargeoff.auto_detect_score_columns: true``,
    because guessing a score column wrong would replace a blank chart with a
    wrong one.

    Returns ``(orig_idx_or_None, curr_idx_or_None)``.
    """
    labels = []
    for i, val in enumerate(header_row):
        lab = "" if val is None else str(val).strip()
        labels.append((i, lab))
    scoreish = [(i, lab) for i, lab in labels
                if lab and _ANY_SCORE_PAT.search(lab)
                and not _NOT_SCORE_PAT.search(lab)]
    if not scoreish:
        return None, None
    orig = next((i for i, lab in scoreish if _ORIG_SCORE_PAT.search(lab)), None)
    curr = next((i for i, lab in scoreish
                 if _CURR_SCORE_PAT.search(lab) and i != orig), None)
    if orig is None and curr is None and len(scoreish) == 1:
        # A single unqualified score column on a charge-off file is the score
        # as of the charge-off, not the origination score.
        curr = scoreish[0][0]
    return orig, curr


def _shared_locator(co_cfg, rc_cfg):
    """``generate_report``'s combined CO+recovery file test, verbatim."""
    if not (co_cfg and rc_cfg):
        return False
    for k in ("account_col", "code_col", "date_col", "has_header", "skip_rows"):
        if co_cfg.get(k) != rc_cfg.get(k):
            return False
    return co_cfg.get("amount_col") != rc_cfg.get("amount_col")


_MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
           "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}


def _file_period(fname, config):
    """(year, month) parsed from a filename, mirroring ``generate_report``."""
    iso = None
    try:
        from import_data import _try_common_date_layouts
        iso = _try_common_date_layouts(fname)
    except Exception:  # noqa: BLE001
        iso = None
    if not iso or len(iso) < 7:
        try:
            from import_data import extract_snapshot_date
            iso = extract_snapshot_date(fname, config)
        except Exception:  # noqa: BLE001
            iso = None
    if not iso or len(iso) < 7:
        m = re.search(r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*"
                      r"[\s_\-\.]*((?:19|20)?\d{2})(?!\d)", fname, re.IGNORECASE)
        if m:
            yr = int(m.group(2))
            yr += 2000 if yr < 100 else 0
            if 2000 <= yr <= 2099:
                return yr, _MONTHS[m.group(1)[:3].lower()]
        m = re.search(r"\b(0[1-9]|1[0-2])[-_\. ]?((?:19|20)?\d{2})\b", fname)
        if m:
            yr = int(m.group(2))
            yr += 2000 if yr < 100 else 0
            if 2000 <= yr <= 2099:
                return yr, int(m.group(1))
        return None
    try:
        return int(iso[0:4]), int(iso[5:7])
    except (TypeError, ValueError):
        return None


def load_chargeoff_loans(config):
    """Every charge-off the pipeline already sees, one row per LOAN.

    Returns ``(DataFrame, [file paths])``.  Columns: ``account``, ``code``,
    ``pool``, ``amount``, ``co_date``, ``date_source``, ``orig_score``,
    ``curr_score``, ``file``.

    File discovery, format routing and the filename-date fallback replicate
    ``generate_report.load_chargeoff_recovery_history`` exactly, so this sees
    the same population of charge-off rows that already feeds the by-year and
    by-pool charge-off numbers -- and therefore sums to the same dollars.
    """
    import pandas as pd
    import generate_report as gr

    cols = ["account", "code", "pool", "amount", "co_date", "date_source",
            "orig_score", "curr_score", "file"]
    hff = config.get("historical_file_formats") or {}
    co_cfg = hff.get("chargeoff")
    rc_cfg = hff.get("recovery")
    formats = hff.get("formats") or []
    multi_format = bool(formats)
    combined_mode = _shared_locator(co_cfg, rc_cfg)

    compiled = []
    for fmt in formats:
        pat = fmt.get("file_pattern") or ""
        try:
            compiled.append((re.compile(pat, re.IGNORECASE) if pat else None, fmt))
        except re.error:
            compiled.append((None, fmt))

    def _resolve(fname):
        if multi_format:
            for rx, fmt in compiled:
                if rx is not None and rx.search(fname):
                    c, r = fmt.get("chargeoff"), fmt.get("recovery")
                    return c, r, _shared_locator(c, r)
            return None, None, False
        return co_cfg, rc_cfg, combined_mode

    data_dir = gr.resolve_path(config.get("data_directory", ""))
    if not data_dir or not os.path.isdir(data_dir):
        return pd.DataFrame(columns=cols), []
    quarters = gr._find_quarter_folders(data_dir)
    if not quarters:
        quarters = [(data_dir, "2026-12")]

    frames, files, seen = [], [], set()
    for folder, qlabel in quarters:
        try:
            names = sorted(os.listdir(folder))
        except OSError:
            continue
        for fname in names:
            low = fname.lower()
            path = os.path.join(folder, fname)
            if path in seen or low.startswith("~$"):
                continue
            if not low.endswith((".xlsx", ".xls", ".csv")):
                continue
            pc, rc, comb = _resolve(fname)
            if pc is None:
                continue
            if comb:
                want = not ("proposed" in low or "3yr" in low)
                if not multi_format and not (("charge" in low and "off" in low)
                                             or "recov" in low):
                    want = False
            else:
                want = (("charge" in low and "off" in low)
                        and "proposed" not in low and "3yr" not in low
                        and "recov" not in low)
                if multi_format and not want and not rc:
                    want = True
            if not want:
                continue
            seen.add(path)
            try:
                part = _parse_co_file_loans(path, pc)
            except Exception as exc:  # noqa: BLE001
                print(f"    CO migration split: could not read {fname}: {exc}")
                continue
            if part.empty:
                continue
            period = _file_period(fname, config)
            part["file"] = fname
            part["_fy"] = period[0] if period else int(qlabel[:4])
            part["_fm"] = period[1] if period else (
                int(qlabel[5:7]) if len(qlabel) >= 7 else 12)
            frames.append(part)
            files.append(path)

    if not frames:
        return pd.DataFrame(columns=cols), files

    out = pd.concat(frames, ignore_index=True)
    pool_map = {str(k).strip(): v for k, v in (config.get("pool_map") or {}).items()}

    def _pool(code):
        s = str(code).strip()
        for cand in (s, s.upper(), s.lower()):
            if cand in pool_map:
                return pool_map[cand]
        try:
            i = str(int(float(s)))
            if i in pool_map:
                return pool_map[i]
        except (ValueError, TypeError):
            pass
        upper = s.upper()
        for v in set(pool_map.values()):
            if str(v).upper() == upper:
                return v
        return None

    out["pool"] = out["code"].map(_pool)
    out["account"] = out["account"].astype(str).str.strip()
    row_date = pd.to_datetime(out["date"], errors="coerce")
    # A charge-off date outside a sane window is a mis-parsed cell, not a date.
    row_date = row_date.where((row_date >= pd.Timestamp("2000-01-01"))
                              & (row_date <= pd.Timestamp("2099-12-31")))
    fallback = pd.to_datetime(
        out["_fy"].astype(int).astype(str) + "-"
        + out["_fm"].astype(int).astype(str).str.zfill(2) + "-01",
        errors="coerce") + pd.offsets.MonthEnd(0)
    out["co_date"] = row_date.fillna(fallback)
    out["date_source"] = ["row" if pd.notna(d) else "filename" for d in row_date]
    return out[cols], files


# ---------------------------------------------------------------------------
# Score history
# ---------------------------------------------------------------------------

def normalize_account(value):
    """Loose form of an account string for a fallback join.

    Charge-off files and loan extracts are mapped independently -- each has
    its own ``member_account`` block -- so the same loan can come out
    ``663101`` on one side and ``00000663101`` on the other.  Stripping
    non-digits and leading zeros makes those meet.  Only ever used when the
    literal string misses AND the normalized form is unambiguous within the
    credit union, so it can never merge two real accounts.
    """
    digits = re.sub(r"\D", "", str(value or ""))
    return digits.lstrip("0")


def load_score_history(credit_union, engine=None):
    """``{account: [(snapshot_date, current_fico, original_fico), ...]}``.

    One entry per (account, snapshot) ordered oldest-first.  Where a snapshot
    holds several rows for the same account string (a double-imported month,
    or a member-level extract), the highest-scored row wins so a blank
    duplicate cannot mask a real score.
    """
    import pandas as pd
    from sqlalchemy import text

    if engine is None:
        os.environ.setdefault("DATABASE_URL", "")
        from cecl_credentials import get_database_url
        from sqlalchemy import create_engine
        engine = create_engine(get_database_url())

    df = pd.read_sql(
        text("SELECT snapshot_date, member_number, current_fico_score, "
             "       original_fico_score "
             "FROM monthly_loan_data WHERE credit_union = :cu "
             "ORDER BY snapshot_date"),
        engine, params={"cu": credit_union})
    hist = {}
    if df.empty:
        return hist
    df["member_number"] = df["member_number"].astype(str).str.strip()
    df["snapshot_date"] = pd.to_datetime(df["snapshot_date"])
    for acct, snap, cur, orig in zip(df["member_number"], df["snapshot_date"],
                                     df["current_fico_score"],
                                     df["original_fico_score"]):
        try:
            cur = int(cur) if cur is not None and cur == cur else 0
        except (TypeError, ValueError):
            cur = 0
        try:
            orig = int(orig) if orig is not None and orig == orig else 0
        except (TypeError, ValueError):
            orig = 0
        rows = hist.setdefault(acct, [])
        if rows and rows[-1][0] == snap:
            if cur > rows[-1][1]:
                rows[-1] = (snap, cur, orig)
        else:
            rows.append((snap, cur, orig))
    return hist


def build_alias_index(history):
    """``{normalized_account: literal_account}`` for unambiguous keys only.

    A normalized form that collapses two or more distinct account strings is
    dropped: matching it would silently merge two loans.
    """
    counts = {}
    for acct in history:
        key = normalize_account(acct)
        if not key:
            continue
        counts.setdefault(key, []).append(acct)
    return {k: v[0] for k, v in counts.items() if len(v) == 1}


def recover_scores(account, co_date, history, alias=None):
    """Recover ``(original_score, score_at_chargeoff, provenance)`` for a loan.

    The loan's row in ``monthly_loan_data`` disappears the month the core stops
    exporting it, taking both of its scores with it.  This walks back to the
    LAST snapshot at or before the charge-off in which the loan was still
    there and scored, and reads the pair off that row.

    A snapshot *after* the charge-off is never used.  For credit unions whose
    core keeps the charged-off back-book in the extract (``Exclude`` pool), a
    later snapshot still carries the loan, but its two scores are frozen and
    equal -- pairing them manufactures a spurious ``Unchanged``.  That is the
    trap ``07_chargeoff_feasibility.md`` measured on Mountain CU.

    Original score, best source first:

    ``row_pair``
        the recovered row's own ``original_fico_score``, when it differs from
        that row's current score.  A difference proves the credit union (or
        its dated credit pull) really supplied an origination score for this
        loan rather than ``import_data``'s ``original = current`` gap-fill.
        This is the same pair the Risk Change matrix would have shown for the
        loan while it was still live.
    ``lookback_span``
        the loan's earliest observed ``current_fico_score``, when the credit
        union refreshes scores between snapshots and that earliest score
        differs from the score at charge-off.  Only useful where the extract
        carries a re-pulled bureau score; most credit unions apply a single
        dated credit pull, so their per-loan score is constant and this never
        fires.
    ``gap_fill``
        neither of the above: original is unknown and ``original = current``
        is assumed, the documented WARM convention already applied by
        ``import_data`` and by the Risk Change matrix.  Classifies as
        ``Unchanged``, but is counted separately so a chart built mostly out
        of assumptions can be refused.

    Failure provenances -- ``no_history`` / ``unscored`` / ``no_prior`` --
    return zeros and belong in ``Not Reported``.
    """
    rows = history.get(account)
    if rows is None and alias:
        literal = alias.get(normalize_account(account))
        if literal is not None:
            rows = history.get(literal)
    if not rows:
        return 0, 0, "no_history"

    prior = rows if co_date is None else [r for r in rows if r[0] <= co_date]
    if not prior:
        return 0, 0, "no_prior"
    scored = [r for r in prior if r[1] > 0]
    if not scored:
        return 0, 0, "unscored"

    at_co = scored[-1][1]

    for _snap, cur, og in reversed(scored):
        if og > 0 and og != cur:
            return og, at_co, "row_pair"

    if scored[0][1] != at_co:
        return scored[0][1], at_co, "lookback_span"

    return at_co, at_co, "gap_fill"


# ---------------------------------------------------------------------------
# Derivation
# ---------------------------------------------------------------------------

def derive_co_by_migration(config, snapshot_date, grades, no_score=None,
                           engine=None, lookback_years=None,
                           co_rows=None, history=None, verbose=True):
    """Return ``(co_by_status, co_by_pool, diagnostics)``.

    Shapes match ``load_impaired_data._read_migration_blocks``::

        co_by_status = {'Improved': {'balance': float, 'pct': float}, ...}
        co_by_pool   = {'<pool>': {'Improved': {...}, ...}, ...}

    ``({}, {}, diag)`` whenever the split cannot be produced honestly, so the
    caller leaves today's behaviour alone.
    """
    import pandas as pd
    from cecl_engine import assign_credit_grade
    from dq_migration_split import classify_migration, _label_sets

    diag = {"status": "ok"}
    co_cfg = config.get("chargeoff") or {}
    if co_cfg.get("derive_migration_split") is False:
        diag["status"] = "disabled"
        return {}, {}, diag

    if no_score is None:
        no_score = config.get("no_score_label", "Not Reported")
    try:
        n_top = int(config.get("top_grades_double_drop", 3))
    except (TypeError, ValueError):
        n_top = 3
    if lookback_years is None:
        lookback_years = co_cfg.get("migration_lookback_years",
                                    DEFAULT_LOOKBACK_YEARS)
    try:
        lookback_years = float(lookback_years)
    except (TypeError, ValueError):
        lookback_years = DEFAULT_LOOKBACK_YEARS

    if co_rows is None:
        co_rows, _files = load_chargeoff_loans(config)
    if co_rows is None or co_rows.empty:
        diag["status"] = "no_chargeoff_rows"
        return {}, {}, diag

    co = co_rows.copy()
    co["co_date"] = pd.to_datetime(co["co_date"], errors="coerce")
    asof = pd.to_datetime(snapshot_date, errors="coerce")
    diag["rows_all"] = int(len(co))
    diag["amount_all"] = float(co["amount"].sum())
    if pd.notna(asof):
        floor = asof - pd.DateOffset(years=int(lookback_years))
        co = co[(co["co_date"].isna()) | ((co["co_date"] >= floor)
                                          & (co["co_date"] <= asof))]
    co = co[co["pool"].notna()]
    diag["rows_in_window"] = int(len(co))
    diag["amount_in_window"] = float(co["amount"].sum()) if len(co) else 0.0
    if co.empty:
        diag["status"] = "no_rows_in_window"
        return {}, {}, diag

    if history is None:
        history = load_score_history(config["credit_union"], engine=engine)
    diag["history_accounts"] = len(history)
    alias = build_alias_index(history)

    fico_labels, brr_labels, brr_pool_lcs = _label_sets(config, grades, no_score)
    fico_idx = {g: i for i, g in enumerate(fico_labels)}
    brr_idx = {g: i for i, g in enumerate(brr_labels)} if brr_labels else None

    def _blank():
        return {lbl: 0.0 for lbl in STATUS_LABELS}

    grand, by_pool = _blank(), {}
    prov_counts, prov_amount = {}, {}
    n_both = 0
    amt_both = 0.0
    amt_measured = 0.0

    for acct, pool, amount, co_date, o_file, c_file in zip(
            co["account"], co["pool"], co["amount"], co["co_date"],
            co["orig_score"], co["curr_score"]):
        try:
            bal = float(amount)
        except (TypeError, ValueError):
            continue
        if bal != bal:
            continue

        # The two scores are sourced independently: whichever side the
        # charge-off feed supplies wins, the other comes from the snapshot
        # lookback.  Several credit unions carry exactly one of the two --
        # Curis CU's charge-off file has an origination FICO (its ``FICO
        # Date`` sits years before the charge-off), Shuford's has the score
        # at write-off -- so pairing per side recovers loans that neither
        # source could resolve alone.
        o_ok = o_file is not None and o_file == o_file and float(o_file) > 0
        c_ok = c_file is not None and c_file == c_file and float(c_file) > 0
        orig_l, co_l, lb = recover_scores(
            acct, None if pd.isna(co_date) else co_date, history, alias)

        if c_ok:
            co_score, src_c = int(c_file), "co_file"
        else:
            co_score, src_c = co_l, lb
        if o_ok:
            orig_score, src_o = int(o_file), "co_file"
        else:
            orig_score, src_o = orig_l, lb

        if co_score <= 0 or orig_score <= 0:
            prov = lb if lb in ("no_history", "no_prior", "unscored") else "partial"
        elif src_o == "gap_fill" and src_c == "gap_fill":
            prov = "gap_fill"
        elif src_o == "co_file" and src_c == "co_file":
            prov = "co_file"
        elif src_o == "co_file":
            prov = "co_file_orig+" + src_c
        elif src_c == "co_file":
            prov = src_o + "+co_file_current"
        else:
            prov = src_o

        prov_counts[prov] = prov_counts.get(prov, 0) + 1
        prov_amount[prov] = prov_amount.get(prov, 0.0) + bal
        if prov in MEASURED_PROVENANCE:
            amt_measured += bal

        pool_key = str(pool).strip()
        by_pool.setdefault(pool_key, _blank())

        if co_score > 0 and orig_score > 0:
            n_both += 1
            amt_both += bal
            cg = assign_credit_grade(co_score, grades, no_score)
            og = assign_credit_grade(orig_score, grades, no_score)
            index = (brr_idx if (brr_idx and pool_key.lower() in brr_pool_lcs)
                     else fico_idx)
            status = classify_migration(cg, og, index, n_top, no_score)
            if status is None:
                status = "Not Reported"
        else:
            status = "Not Reported"

        grand[status] += bal
        by_pool[pool_key][status] += bal

    total = sum(grand.values())
    diag.update({
        "rows_used": int(sum(prov_counts.values())),
        "provenance": prov_counts,
        "provenance_amount": prov_amount,
        "both_scores_rows": n_both,
        "both_scores_amount": amt_both,
        "measured_amount": amt_measured,
        "total": total,
        "not_reported": grand["Not Reported"],
    })
    if total <= 0:
        diag["status"] = "zero_total"
        return {}, {}, diag

    nr_share = grand["Not Reported"] / total
    measured_share = amt_measured / total
    diag["not_reported_share"] = nr_share
    diag["measured_share"] = measured_share
    if nr_share > MAX_NOT_REPORTED_SHARE:
        diag["status"] = "refused_low_coverage"
        if verbose:
            print(f"    *** CO migration split REFUSED: only "
                  f"{(1 - nr_share) * 100:.1f}% of ${total:,.2f} in charge-offs "
                  f"could be scored from loan history "
                  f"({n_both:,} of {diag['rows_used']:,} loans recovered both "
                  f"an origination and a charge-off score). Leaving the Risk "
                  f"Change charge-off charts unpopulated -- more monthly loan "
                  f"snapshots are needed before this chart can be honest.")
        return {}, {}, diag

    if measured_share < MIN_MEASURED_SHARE:
        diag["status"] = "refused_mostly_assumed"
        if verbose:
            print(f"    *** CO migration split REFUSED: only "
                  f"{measured_share * 100:.1f}% of ${total:,.2f} in charge-offs "
                  f"carries a genuinely measured pair of scores. The rest would "
                  f"plot as 'Unchanged' purely by the original=current gap-fill "
                  f"convention -- an assumption, not a measurement. Leaving the "
                  f"Risk Change charge-off charts unpopulated.")
        return {}, {}, diag

    if verbose:
        print(f"    CO migration split derived from loan-snapshot history: "
              f"{n_both:,} of {diag['rows_used']:,} charged-off loan(s) "
              f"recovered both scores; ${total:,.2f} charged off, "
              f"{(1 - nr_share) * 100:.1f}% scored, "
              f"{measured_share * 100:.1f}% measured "
              f"({int(lookback_years)}-year window to {snapshot_date}).")

    def _shape(counts):
        tot = sum(counts.values())
        return {lbl: {"balance": round(counts[lbl], 2),
                      "pct": (counts[lbl] / tot) if tot else 0.0}
                for lbl in STATUS_LABELS}

    return _shape(grand), {p: _shape(c) for p, c in by_pool.items()}, diag


def fill_missing_co_migration(hist, config, snapshot_date, grades,
                              no_score=None, engine=None):
    """Populate ``hist['impaired']['co_by_status'|'co_by_pool']`` if absent.

    A WARM-supplied value always wins: when both keys already hold a non-empty
    dict this is a no-op and no charge-off file is even opened, so credit
    unions with a working ``CO Data Entry`` tab are byte-for-byte unaffected.

    Returns ``True`` when it filled something.
    """
    if hist is None:
        return False
    imp = hist.get("impaired") or {}
    have_status = bool(imp.get("co_by_status"))
    have_pool = bool(imp.get("co_by_pool"))
    if have_status and have_pool:
        return False

    status, by_pool, diag = derive_co_by_migration(
        config, snapshot_date, grades, no_score=no_score, engine=engine)
    if not status and not by_pool:
        imp.setdefault("co_unavailable_reason", diag.get("status"))
        hist["impaired"] = imp
        return False

    filled = []
    if not have_status and status:
        imp["co_by_status"] = status
        filled.append("co_by_status")
    if not have_pool and by_pool:
        imp["co_by_pool"] = by_pool
        filled.append(f"co_by_pool ({len(by_pool)} pools)")
    if not filled:
        return False
    imp["co_source"] = "derived_from_snapshot_history"
    imp["co_diagnostics"] = diag
    hist["impaired"] = imp
    print(f"    CO migration split: filled {', '.join(filled)} "
          f"(no WARM 'CO Data Entry' source for this snapshot).")
    return True
