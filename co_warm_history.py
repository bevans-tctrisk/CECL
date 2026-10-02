"""Path A -- score history recovered from the WARM workbooks' ``Mmm-YY`` tabs.

Two families of ``Mmm-YY`` tab live in a CECL-Migration-WARM workbook, and
neither is fully read by anything in the pipeline today:

``Mmm-YY Credit Pull``
    The credit union's dated bureau pull: member number + FICO, two or three
    columns (``Member Number`` / ``FICO``, or ``Member #`` /
    ``Current Credit Score`` / ``Open Date``).
    ``import_data._load_dated_credit_pulls`` reads these, but only from the
    single newest WARM in the folder (``max(warms, key=os.path.getmtime)``)
    and only for the opt-in ``credit_pull.origination_aware`` path.
    **MEMBER level.**

``Mmm-YY Data``
    The full quarterly loan extract, archived inside the workbook -- 60+
    columns including ``Member #-Suffix``, ``Original Credit Score``,
    ``Original Credit Grade``, ``Current Credit Score``,
    ``Current Credit Grade``, ``Loan Pool``, ``Current Balance``,
    ``Charge Off Amount``.  **Nothing reads these at all**, and they reach
    years further back than ``monthly_loan_data`` does.  **LOAN level.**

Because the standalone credit-pull files are deleted after six months for
security, the WARM tab is the surviving copy of both.

Each quarter's WARM carries every earlier quarter's tabs, so the newest
workbook usually suffices; the loader still unions across every WARM in the
folder because a tab occasionally gets dropped on a re-issue.

See ``docs/pdf_migration/09_chargeoff_score_history.md`` §A.
"""

from __future__ import annotations

import calendar
import glob
import os
import re
import zipfile

_SHEET_RE = re.compile(r'<sheet[^>]*name="([^"]*)"')
_MMMYY_RE = re.compile(r"^([A-Za-z]{3})-(\d{2})\b")
_MON = {m.lower(): i for i, m in enumerate(calendar.month_abbr) if m}

# Column-header candidates, lowercased. Order is preference order.
_ACCT_HDRS = ("member #-suffix", "member-sfx", "member #-sfx",
              "member-suffix", "account number")
_MEMBER_HDRS = ("member #", "member number", "member no", "member")
_CUR_HDRS = ("current credit score", "curr score from pull", "curr scr hard",
             "fico", "credit score", "score")
_ORG_HDRS = ("original credit score", "org score")


def sheet_date(name):
    """``'Dec-24 Credit Pull'`` -> ``Timestamp('2024-12-31')``; else ``None``."""
    import pandas as pd
    m = _MMMYY_RE.match(str(name).strip())
    if not m:
        return None
    mon = _MON.get(m.group(1).lower())
    if not mon:
        return None
    yr = 2000 + int(m.group(2))
    return pd.Timestamp(yr, mon, calendar.monthrange(yr, mon)[1])


def classify_tab(name):
    """``'pull'`` / ``'data'`` / ``None`` for a sheet name."""
    if sheet_date(name) is None:
        return None
    low = str(name).strip().lower()
    if "credit pull" in low:
        return "pull"
    if low.endswith("data") and "cc data" not in low:
        return "data"
    return None


def list_tabs(path):
    """Sheet names straight out of the zip -- no workbook parsing, no pandas.

    Cheap enough to run over every WARM on a network share; a 90-tab, 40MB
    workbook answers in milliseconds because only ``xl/workbook.xml`` is read.
    """
    try:
        with zipfile.ZipFile(path) as z:
            return _SHEET_RE.findall(
                z.read("xl/workbook.xml").decode("utf-8", "replace"))
    except Exception:  # noqa: BLE001
        return []


def warm_files(config, resolve=None):
    """Every CECL-Migration-WARM workbook reachable for this credit union."""
    cp = config.get("credit_pull") or {}
    folders = [cp.get("fallback_report_folder"), cp.get("source_folder"),
               config.get("warm_folder"), config.get("impaired_folder")]
    if resolve is None:
        from generate_report import resolve_path as resolve
    out, seen = [], set()
    for f in folders:
        if not f:
            continue
        try:
            d = resolve(f)
        except Exception:  # noqa: BLE001
            d = f
        if not d or not os.path.isdir(d) or d in seen:
            continue
        seen.add(d)
        for p in sorted(glob.glob(os.path.join(d, "*.xls*"))):
            b = os.path.basename(p)
            low = b.lower()
            if b.startswith("~$") or "warm" not in low or "dnu" in low:
                continue
            out.append(p)
    return out


def _pick(cols_lc, candidates):
    for c in candidates:
        if c in cols_lc:
            return cols_lc[c]
    return None


def load_warm_tabs(config, resolve=None, verbose=True, want_data=True,
                   want_pulls=True):
    """Read every ``Mmm-YY Data`` and ``Mmm-YY Credit Pull`` tab.

    Returns ``(loan_hist, member_hist, meta)``::

        loan_hist   {full_account: [(date, current, original), ...]}  oldest first
        member_hist {member_key:   [(date, score), ...]}              oldest first

    ``full_account`` is the ``Member #-Suffix`` string exactly as the WARM
    writes it; the caller normalises before joining, because the WARM tab and
    the loan extract are mapped independently.
    """
    import pandas as pd

    files = warm_files(config, resolve)
    loan_hist, member_hist = {}, {}
    data_tabs, pull_tabs, skipped = {}, {}, []

    # Newest workbook first, so the freshest copy of a repeated tab wins and
    # older workbooks only contribute tabs the newer ones dropped.
    files = sorted(files, key=lambda p: os.path.getmtime(p), reverse=True)
    for path in files:
        want = []
        for t in list_tabs(path):
            kind = classify_tab(t)
            if kind == "data" and want_data and t not in data_tabs:
                want.append((kind, t, sheet_date(t)))
            elif kind == "pull" and want_pulls and t not in pull_tabs:
                want.append((kind, t, sheet_date(t)))
        if not want:
            continue
        try:
            xl = pd.ExcelFile(path)
        except Exception as exc:  # noqa: BLE001
            skipped.append((os.path.basename(path), str(exc)[:70]))
            continue
        for kind, tab, when in want:
            try:
                df = pd.read_excel(xl, tab)
            except Exception as exc:  # noqa: BLE001
                skipped.append((tab, str(exc)[:70]))
                continue
            cols_lc = {str(c).strip().lower(): c for c in df.columns}
            if kind == "data":
                acct = _pick(cols_lc, _ACCT_HDRS)
                cur = _pick(cols_lc, _CUR_HDRS)
                if acct is None or cur is None:
                    skipped.append((tab, "no account / score column"))
                    continue
                org = _pick(cols_lc, _ORG_HDRS)
                a = df[acct].astype(str).str.strip()
                c = pd.to_numeric(df[cur], errors="coerce").fillna(0)
                o = (pd.to_numeric(df[org], errors="coerce").fillna(0)
                     if org is not None else c * 0)
                n = 0
                for ak, cv, ov in zip(a, c, o):
                    if not ak or ak in ("nan", "None", "NaT"):
                        continue
                    try:
                        cv, ov = int(cv), int(ov)
                    except (TypeError, ValueError):
                        continue
                    loan_hist.setdefault(ak, []).append((when, cv, ov))
                    n += 1
                data_tabs[tab] = (when, n, os.path.basename(path))
            else:
                mem = _pick(cols_lc, _MEMBER_HDRS)
                sc = _pick(cols_lc, _CUR_HDRS)
                if mem is None or sc is None:
                    # Brian's two-column convention with unrecognised headers:
                    # column 0 is the member, column 1 the score.
                    if len(df.columns) >= 2:
                        mem, sc = df.columns[0], df.columns[1]
                    else:
                        skipped.append((tab, "no member / score column"))
                        continue
                s = pd.to_numeric(df[sc], errors="coerce").fillna(0)
                n = 0
                for mv, sv in zip(df[mem], s):
                    try:
                        mk = str(int(float(mv)))
                        sv = int(sv)
                    except (TypeError, ValueError):
                        continue
                    if sv <= 0:
                        continue
                    member_hist.setdefault(mk, []).append((when, sv))
                    n += 1
                pull_tabs[tab] = (when, n, os.path.basename(path))

    for v in loan_hist.values():
        v.sort(key=lambda r: r[0])
    for v in member_hist.values():
        v.sort(key=lambda r: r[0])

    meta = {
        "files": len(files),
        "data_tabs": sorted(data_tabs.items(), key=lambda kv: kv[1][0]),
        "pull_tabs": sorted(pull_tabs.items(), key=lambda kv: kv[1][0]),
        "loan_accounts": len(loan_hist),
        "members": len(member_hist),
        "skipped": skipped,
    }
    if verbose:
        dts = [str(v[0].date()) for _k, v in meta["data_tabs"]]
        pts = [str(v[0].date()) for _k, v in meta["pull_tabs"]]
        print(f"    Path A: {len(files)} WARM file(s); "
              f"{len(data_tabs)} 'Mmm-YY Data' tab(s) "
              f"[{dts[0] if dts else '-'} .. {dts[-1] if dts else '-'}], "
              f"{len(pull_tabs)} 'Mmm-YY Credit Pull' tab(s) "
              f"[{pts[0] if pts else '-'} .. {pts[-1] if pts else '-'}]; "
              f"{len(loan_hist):,} loan accounts, {len(member_hist):,} members.")
    return loan_hist, member_hist, meta


def score_movement(hist, score_index=1):
    """``(accounts, in_2plus, ever_changed, mean_range)`` for a history dict.

    The same measurement ``09 §3.2`` ran against ``monthly_loan_data``, so
    Path A and Path B numbers are directly comparable.
    """
    n_multi = n_moved = 0
    total_range = 0
    for rows in hist.values():
        scored = [r for r in rows if r[score_index] > 0]
        if len(scored) < 2:
            continue
        n_multi += 1
        vals = [r[score_index] for r in scored]
        if len(set(vals)) >= 2:
            n_moved += 1
        total_range += max(vals) - min(vals)
    return (len(hist), n_multi, n_moved,
            (total_range / n_multi) if n_multi else 0.0)
