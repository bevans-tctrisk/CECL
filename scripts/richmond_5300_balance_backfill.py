"""5300 distributed BALANCE backfill for Credit Union of Richmond.

Fills the pre-import quarters (2018-2023) of pool balance history from the NCUA
5300 total loan balance, apportioned across pools using the earliest imported
month's pool mix. Quarters already covered by imported loan data (2024-01..
2026-03) are passed as existing_dates so the accurate imported balances are
never overwritten with the apportioned estimate.
"""
import os, datetime as dt
os.environ['CECL_WORKSPACE_ROOT'] = r'Z:\Shared\TCT Files\CECL - CM Files'
import cecl_credentials
os.environ['DATABASE_URL'] = cecl_credentials.get_database_url()
import sqlalchemy as sa
from cecl_ui.services import solr_5300_backfill as bf

CU = 'Credit Union of Richmond'
CHARTER = 66929
SOLR = 'http://searchserver1.tctrisk.com:8983/solr'
CORE = 'ncua'
TARGET = '2026-03-31'
MONTHS = 90
REF = '2024-01-31'   # earliest imported month -> pool distribution

eng = sa.create_engine(cecl_credentials.get_database_url())
with eng.connect() as cx:
    rows = cx.execute(sa.text(
        "SELECT loan_pool, SUM(current_balance) bal FROM monthly_loan_data "
        "WHERE credit_union=:c AND snapshot_date=:s AND loan_pool NOT IN ('Ignore','Exclude') "
        "GROUP BY loan_pool"), {'c': CU, 's': REF}).fetchall()
    by_pool = {r.loan_pool: float(r.bal or 0) for r in rows if (r.bal or 0) > 0}
    total = sum(by_pool.values())
    pool_distribution = {p: v / total for p, v in by_pool.items()}
    # Quarter-ends already covered by imported loan data -> skip (don't overwrite)
    snaps = cx.execute(sa.text(
        "SELECT DISTINCT snapshot_date FROM monthly_loan_data WHERE credit_union=:c"),
        {'c': CU}).fetchall()
    imported_months = sorted(str(s.snapshot_date) for s in snaps)

# existing = ALL imported month-ends (2024-01..2026-03) so the accurate imported
# balances are never overwritten; backfill then fills only the pre-import range.
existing = set(imported_months)

print('pool_distribution:')
for p, r in sorted(pool_distribution.items(), key=lambda x: -x[1]):
    print('   %-28s %.4f' % (p, r))
print('total ref balance %.2f; existing (skip) quarters: %s' % (total, sorted(existing)))

res = bf.backfill_missing_quarters_distributed(
    CU, CHARTER, SOLR, CORE, TARGET, MONTHS,
    pool_distribution=pool_distribution,
    existing_dates=existing,
    source_period_iso=REF,
)
print('\nok=%s rows=%s filled=%s no_data=%s err=%s' % (
    res.get('ok'), res.get('rows_written'), len(res.get('months_filled') or []),
    len(res.get('months_no_data') or []), res.get('error')))

# Idempotence guard: the service's cleanup only drops rows for pools no longer
# in the list, so a prior run with different existing_dates can leave stray
# distributed rows inside the imported range. Purge any '5300-distributed:%'
# rows at/after the earliest imported month so imported balances always win.
with eng.begin() as cx:
    d = cx.execute(sa.text(
        "DELETE FROM loan_code_history WHERE cu=:c AND source LIKE '5300-distributed:%' "
        "AND as_of_date >= :lo"), {'c': CU, 'lo': imported_months[0]})
    print('purged stray in-range distributed rows:', d.rowcount)

