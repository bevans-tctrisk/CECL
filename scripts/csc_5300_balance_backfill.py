"""5300 distributed balance backfill for Central Susquehanna Comm FCU (14180).

WHY: CSC uses monthly_balance.source == 'single', which the wizard's
auto-distributed-backfill does NOT handle (it only covers annual /
monthly balance-sheet sources). So loan_code_history was never seeded and
the report couldn't extend historical balances past the monthly file
(~2023-01). 84-month Real Estate pools showed blank 2019-2022 columns.

WHAT: Replicate the house 'distributed' method: distribute each 5300
quarter's TOTAL loan balance across the configured pools using the
earliest monthly-balance-file month's pool mix, writing one
loan_code_history row per pool (source '5300-distributed:<qe>'). The
uploaded monthly-balance months are passed as existing_dates so only
pre-file quarters (2018Q4-2022) get filled; the file stays authoritative
for 2023+.
"""
import os, sys
sys.path.insert(0, r'C:\dev\CECL')
os.environ['CECL_WORKSPACE_ROOT'] = r'Z:\Shared\TCT Files\CECL - CM Files'
import cecl_credentials
os.environ['DATABASE_URL'] = cecl_credentials.get_database_url()

import yaml
import pandas as pd
import sqlalchemy as sa
import generate_report
from cecl_ui.services import solr_5300_backfill as bf

CU = 'Central Susquehanna Comm FCU'
CHARTER = 14180
SOLR = 'http://searchserver1.tctrisk.com:8983/solr'
CORE = 'ncua'
TARGET = '2026-06-30'
MONTHS = 90  # covers the 84-month June-2026 Real Estate window (back to ~2019) + buffer

CFG = r'Z:\Shared\TCT Files\CECL - CM Files\client_configs\central_susquehanna_comm_fcu.yaml'
with open(CFG, encoding='utf-8') as fh:
    config = yaml.safe_load(fh)

# --- earliest-month pool distribution from the monthly balance file ---
mb_df, _ = generate_report.load_monthly_balances(config)
mb_df = mb_df.copy()
mb_df['date'] = pd.to_datetime(mb_df['date'], errors='coerce')
mb_df = mb_df.dropna(subset=['date'])
earliest = mb_df['date'].min()
earliest_iso = earliest.date().isoformat()
first = mb_df[mb_df['date'] == earliest]
by_pool = first.groupby('pool')['balance'].sum()
by_pool = by_pool[by_pool > 0]
total = float(by_pool.sum())
pool_distribution = {p: float(v) / total for p, v in by_pool.items()}

# Protect the ENTIRE monthly-file era with TRUE calendar month-ends so the
# 5300 distributed backfill fills ONLY the pre-file gap (2018-2022). The
# file stores day-23 dates, so we must normalise to month-end to match the
# backfill's quarter_fill_dates (which are true month-ends) -- otherwise
# 5300 rows land alongside file rows in 2023-2026 and distort the averages.
import calendar as _cal
file_lo = mb_df['date'].min()
protect_hi = pd.Timestamp('2026-12-31')  # cover the whole file era + buffer
existing_dates = set()
_cur = pd.Timestamp(file_lo.year, file_lo.month, 1)
while _cur <= protect_hi:
    last = _cal.monthrange(_cur.year, _cur.month)[1]
    existing_dates.add(_cur.replace(day=last).date().isoformat())
    _cur = (_cur + pd.offsets.MonthBegin(1))

# Clean slate: remove any distributed rows from a prior run so the
# protected-range fix takes full effect (idempotent).
eng0 = sa.create_engine(cecl_credentials.get_database_url())
with eng0.begin() as cx:
    d0 = cx.execute(sa.text(
        "DELETE FROM loan_code_history WHERE cu=:c AND source LIKE '5300-distributed:%'"),
        {'c': CU})
    print(f"Cleared {d0.rowcount} prior distributed row(s)")

print(f"Earliest monthly-file month: {earliest_iso}  total ${total:,.2f}")
for p, r in sorted(pool_distribution.items(), key=lambda x: -x[1]):
    print(f"   {p:<26} {r:6.3%}   (${by_pool[p]:,.2f})")
print(f"Protected (file-era) month-ends: {min(existing_dates)}..{max(existing_dates)} "
      f"({len(existing_dates)} months)")

print("\n=== running distributed 5300 balance backfill ===")
res = bf.backfill_missing_quarters_distributed(
    CU, CHARTER, SOLR, CORE, TARGET, MONTHS,
    pool_distribution=pool_distribution,
    existing_dates=existing_dates,
    source_period_iso=earliest_iso,
)
print("ok=%s rows_written=%s filled=%s no_data=%s stale_removed=%s err=%s" % (
    res.get('ok'), res.get('rows_written'),
    len(res.get('months_filled') or []),
    len(res.get('months_no_data') or []),
    res.get('stale_rows_removed'), res.get('error')))
mf = res.get('months_filled') or []
if mf:
    ds = sorted(m['period'] for m in mf)
    print("  filled months:", ds[0], "..", ds[-1])
nd = res.get('months_no_data') or []
if nd:
    print("  no-data quarters:", sorted(nd))

print("\n=== post-backfill loan_code_history coverage ===")
eng = sa.create_engine(cecl_credentials.get_database_url())
with eng.connect() as cx:
    row = cx.execute(sa.text(
        "SELECT MIN(as_of_date) mn, MAX(as_of_date) mx, "
        "COUNT(DISTINCT as_of_date) q, COUNT(*) n FROM loan_code_history WHERE cu=:c"),
        {'c': CU}).fetchone()
    print(f"  {row.mn}..{row.mx}  {row.q} dates  {row.n} rows")
    print("\n  Real-estate pool history by year (avg across quarter-ends):")
    for pool in ['1st/2nd Lien Mortgage', 'HELOC', 'RETAINED MTG']:
        rows = cx.execute(sa.text(
            "SELECT EXTRACT(YEAR FROM as_of_date)::int yr, AVG(total_balance) bal "
            "FROM loan_code_history WHERE cu=:c AND loan_code=:p "
            "GROUP BY 1 ORDER BY 1"), {'c': CU, 'p': pool}).fetchall()
        yrs = " ".join(f"{int(r.yr)}=${r.bal:,.0f}" for r in rows)
        print(f"   {pool:<24} {yrs}")
