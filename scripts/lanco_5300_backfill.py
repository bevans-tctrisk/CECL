"""5300 distributed backfill for Lanco FCU (16657).

Fills pre-import history from the NCUA 5300 call report:
  * BALANCE  — distributed across pools by the Dec-2025 imported pool mix,
               written to loan_code_history (loan_code = pool name). The
               report's GL monthly_balance file (2023-05..2026-04) wins for
               the months it covers; the 5300 fills the earlier quarters.
  * CO/RECOV — written to loan_code_chargeoff_history / _recovery_history.
               The CU's own transaction files win at the YEAR level; the
               5300 fills only the years the files don't cover.
DQ is already backfilled (loan_code_delinquency_history, 2018-12..2025-12),
so it is left untouched.
"""
import os
os.environ['CECL_WORKSPACE_ROOT'] = r'Z:\Shared\TCT Files\CECL - CM Files'
os.environ.setdefault('PYTHONPATH', r'C:\dev\CECL')
import cecl_credentials
os.environ['DATABASE_URL'] = cecl_credentials.get_database_url()

import sqlalchemy as sa
from cecl_ui.services import (
    solr_5300_backfill as bf,
    solr_5300_co_backfill as co,
    solr_5300_recov_backfill as rec,
)

CU = 'Lanco FCU'
CHARTER = 16657
SOLR = 'http://searchserver1.tctrisk.com:8983/solr'
CORE = 'ncua'
TARGET = '2025-12-31'   # report period (full month-end date)
MONTHS = 90             # ~7.5 yrs -> back to ~2018-06 (DQ hist starts 2018-12)
REF = '2025-12-31'      # imported month -> pool distribution

eng = sa.create_engine(cecl_credentials.get_database_url())
with eng.connect() as cx:
    rows = cx.execute(sa.text(
        "SELECT loan_pool, SUM(current_balance) bal FROM monthly_loan_data "
        "WHERE credit_union=:c AND snapshot_date=:s "
        "AND loan_pool NOT IN ('Ignore','Exclude') "
        "GROUP BY loan_pool"), {'c': CU, 's': REF}).fetchall()
    by_pool = {r.loan_pool: float(r.bal or 0) for r in rows if (r.bal or 0) > 0}
    total = sum(by_pool.values())
    pool_distribution = {p: v / total for p, v in by_pool.items()}
    snaps = cx.execute(sa.text(
        "SELECT DISTINCT snapshot_date FROM monthly_loan_data WHERE credit_union=:c"),
        {'c': CU}).fetchall()
    imported_months = sorted(str(s.snapshot_date) for s in snaps)

existing = set(imported_months)

print('pool_distribution (from %s, total $%.2f):' % (REF, total))
for p, r in sorted(pool_distribution.items(), key=lambda x: -x[1]):
    print('   %-28s %.4f' % (p, r))
print('existing (skip) imported quarters:', sorted(existing))

print('\n=== BALANCE backfill (distributed) ===')
r1 = bf.backfill_missing_quarters_distributed(
    CU, CHARTER, SOLR, CORE, TARGET, MONTHS,
    pool_distribution=pool_distribution,
    existing_dates=existing,
    source_period_iso=REF,
)
print('ok=%s rows=%s filled=%s no_data=%s stale_removed=%s err=%s' % (
    r1.get('ok'), r1.get('rows_written'), len(r1.get('months_filled') or []),
    len(r1.get('months_no_data') or []), r1.get('stale_rows_removed'), r1.get('error')))

print('\n=== CHARGE-OFF backfill ===')
r2 = co.backfill_missing_chargeoff_quarters(CU, CHARTER, SOLR, CORE, TARGET, MONTHS, overwrite=True)
print('ok=%s rows=%s filled=%s no_data=%s err=%s' % (
    r2.get('ok'), r2.get('rows_written'), len(r2.get('months_filled') or []),
    len(r2.get('months_no_data') or []), r2.get('error')))

print('\n=== RECOVERY backfill ===')
r3 = rec.backfill_missing_recovery_quarters(CU, CHARTER, SOLR, CORE, TARGET, MONTHS, overwrite=True)
print('ok=%s rows=%s filled=%s no_data=%s err=%s' % (
    r3.get('ok'), r3.get('rows_written'), len(r3.get('months_filled') or []),
    len(r3.get('months_no_data') or []), r3.get('error')))

print('\n=== post-backfill coverage ===')
with eng.connect() as cx:
    for t, amt in [('loan_code_history', 'total_balance'),
                   ('loan_code_chargeoff_history', 'chargeoff_amount'),
                   ('loan_code_recovery_history', 'recovery_amount'),
                   ('loan_code_delinquency_history', 'dq_amount')]:
        row = cx.execute(sa.text(
            f"SELECT MIN(as_of_date) mn, MAX(as_of_date) mx, "
            f"COUNT(DISTINCT as_of_date) q, COUNT(*) n FROM {t} WHERE cu=:c"),
            {'c': CU}).fetchone()
        print('   %-32s %s..%s  %s qtrs  %s rows' % (t, row.mn, row.mx, row.q, row.n))
