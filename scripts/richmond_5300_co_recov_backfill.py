import os
os.environ['CECL_WORKSPACE_ROOT'] = r'Z:\Shared\TCT Files\CECL - CM Files'
os.environ.setdefault('PYTHONPATH', r'C:\dev\CECL')
import cecl_credentials
os.environ['DATABASE_URL'] = cecl_credentials.get_database_url()
from cecl_ui.services import solr_5300_co_backfill as co, solr_5300_recov_backfill as rec

CU = 'Credit Union of Richmond'
CHARTER = 66929
SOLR = 'http://searchserver1.tctrisk.com:8983/solr'
CORE = 'ncua'
TARGET = '2026-03-31'   # latest report period (full month-end date required)
MONTHS = 90             # covers both Dec-2025 and Mar-2026 84-month RE windows (back to ~2018-09)

print('=== Charge-off backfill ===')
r1 = co.backfill_missing_chargeoff_quarters(CU, CHARTER, SOLR, CORE, TARGET, MONTHS, overwrite=True)
print('ok=%s rows=%s filled=%s no_data=%s err=%s' % (
    r1.get('ok'), r1.get('rows_written'), len(r1.get('months_filled') or []),
    len(r1.get('months_no_data') or []), r1.get('error')))

print('=== Recovery backfill ===')
r2 = rec.backfill_missing_recovery_quarters(CU, CHARTER, SOLR, CORE, TARGET, MONTHS, overwrite=True)
print('ok=%s rows=%s filled=%s no_data=%s err=%s' % (
    r2.get('ok'), r2.get('rows_written'), len(r2.get('months_filled') or []),
    len(r2.get('months_no_data') or []), r2.get('error')))
