"""Pool-code-column detector (Setup Router, Phase 1).

Given a loan extract that may carry more than one plausible "loan type / pool
code" column, pick the one that makes the per-pool balances **tie to the
monthly balance file**. This is the U-vs-V decision from the Cottonwood
onboarding, resolved by ground truth rather than by which column *looks* like
a code.

Design principle (see docs/setup_router_design.md): the winner is the column
whose mapped per-pool sums best match the book. Confidence is the tie result,
not a heuristic score. When no candidate ties, the detector returns
``winner=None`` (a Red decision -> surface for human review), never a guess.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd


@dataclass
class CandidateResult:
    column: str
    key_coverage: float        # fraction of rows whose code is a known pool key
    pools_tied: int            # pools matching the book within tolerance
    pools_total: int
    total_abs_diff: float      # sum |extract_pool - book_pool| across pools
    mapped_total: float        # balance that mapped to a real pool
    excluded_total: float      # balance that fell to the default/Exclude pool


@dataclass
class DetectionResult:
    winner: str | None
    tolerance: float
    candidates: list[CandidateResult] = field(default_factory=list)
    evidence: str = ""

    @property
    def decision_tier(self) -> str:
        """Green/Yellow/Red per the design's validation-gated confidence."""
        if not self.winner or not self.candidates:
            return "red"
        top = self.candidates[0]
        if top.pools_tied == top.pools_total and top.pools_total > 0:
            return "green"
        if top.pools_tied >= max(1, top.pools_total - 1):
            return "yellow"
        return "red"


def _norm_code(x: str) -> str:
    x = str(x).strip()
    if x.replace(".", "", 1).isdigit():
        try:
            return str(int(float(x)))
        except (TypeError, ValueError):
            return x
    return x


def _map_codes(series: pd.Series, pool_map: dict, default: str) -> pd.Series:
    keys = {str(k): v for k, v in pool_map.items()}
    raw = series.astype(str).str.strip().map(_norm_code)
    return raw.map(keys).fillna(default)


def find_candidate_columns(df: pd.DataFrame, pool_map: dict, *,
                           max_distinct: int = 80,
                           min_coverage: float = 0.30) -> list[tuple[str, float]]:
    """Return ``[(column, key_coverage)]`` for columns that look like pool
    codes: low-cardinality and where a meaningful fraction of rows carry a
    value that is a known ``pool_map`` key. Sorted best-coverage-first."""
    keys = {str(k) for k in pool_map}
    out: list[tuple[str, float]] = []
    for col in df.columns:
        s = df[col].dropna().astype(str).str.strip()
        if s.empty:
            continue
        distinct = s.nunique()
        if distinct == 0 or distinct > max_distinct:
            continue
        norm = s.map(_norm_code)
        coverage = float(norm.isin(keys).mean())
        if coverage >= min_coverage:
            out.append((col, round(coverage, 3)))
    out.sort(key=lambda t: -t[1])
    return out


def score_column(df: pd.DataFrame, col: str, balance_col: str,
                 pool_map: dict, monthly_by_pool: dict, *,
                 default: str = "Exclude", tol: float = 1.0,
                 coverage: float = 0.0) -> CandidateResult:
    pools = _map_codes(df[col], pool_map, default)
    bal = pd.to_numeric(df[balance_col], errors="coerce").fillna(0.0)
    by_pool = bal.groupby(pools).sum()

    tied = 0
    total_abs = 0.0
    for pool, book in monthly_by_pool.items():
        got = float(by_pool.get(pool, 0.0))
        total_abs += abs(got - float(book))
        if abs(got - float(book)) <= tol:
            tied += 1
    excluded = float(by_pool.get(default, 0.0))
    return CandidateResult(
        column=col, key_coverage=round(coverage, 3),
        pools_tied=tied, pools_total=len(monthly_by_pool),
        total_abs_diff=round(total_abs, 2),
        mapped_total=round(float(bal.sum() - excluded), 2),
        excluded_total=round(excluded, 2),
    )


def detect(df: pd.DataFrame, balance_col: str, pool_map: dict,
           monthly_by_pool: dict, *, default: str = "Exclude",
           tol: float = 1.0) -> DetectionResult:
    """Rank candidate code columns by how well their per-pool sums tie to the
    monthly book; return the winner (or None) with evidence."""
    candidates = find_candidate_columns(df, pool_map)
    results: list[CandidateResult] = []
    for col, cov in candidates:
        results.append(score_column(df, col, balance_col, pool_map,
                                    monthly_by_pool, default=default,
                                    tol=tol, coverage=cov))
    # Best = most pools tied, then smallest total dollar difference.
    results.sort(key=lambda r: (-r.pools_tied, r.total_abs_diff))
    winner = results[0].column if results and results[0].pools_tied > 0 else None

    if winner:
        top = results[0]
        runner = (f"; runner-up '{results[1].column}' tied "
                  f"{results[1].pools_tied}/{results[1].pools_total}"
                  if len(results) > 1 else "")
        evidence = (f"'{winner}' ties {top.pools_tied}/{top.pools_total} pools "
                    f"to the monthly book (total diff ${top.total_abs_diff:,.2f})"
                    f"{runner}")
    else:
        evidence = ("No candidate column ties any pool to the monthly book "
                    "— surface for human review (Red).")
    return DetectionResult(winner=winner, tolerance=tol,
                           candidates=results, evidence=evidence)
