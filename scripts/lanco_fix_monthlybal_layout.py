"""Fix Lanco monthly-balance loading after a rerun stripped the
monthly_balance layout hints (sheet/header_row/pool_name_col/first_date_col)
from the YAML, which made the report engine's auto-detect fail (the workbook
date headers were stored as STRINGS, not datetimes) -> 0 monthly balances ->
Balance Adjustment showed $0.

Two-part robust fix:
 1) Convert the workbook's row-1 date headers from 'YYYY-MM-DD' strings to
    real datetime cells so auto-detect works WITH or WITHOUT the config hints.
 2) Restore the explicit monthly_balance layout fields in the YAML.
"""
import os, sys, datetime
sys.path.insert(0, r'C:\dev\CECL')
os.environ['CECL_WORKSPACE_ROOT'] = r'Z:\Shared\TCT Files\CECL - CM Files'
import cecl_credentials
os.environ['DATABASE_URL'] = cecl_credentials.get_database_url()
import shutil, openpyxl
from cecl_ui.services import config_service

WS = r'Z:\Shared\TCT Files\CECL - CM Files'
WB = r'Z:\Shared\TCT Files\CECL - CM Files\Raw_Uploads\lanco_fcu\Monthly Loan Balances by Type.xlsx'

# ---- 1) convert header-row date strings to real datetimes ----
shutil.copyfile(WB, WB + '.bak.phase_datehdr_' + datetime.datetime.now().strftime('%Y%m%d%H%M%S'))
wb = openpyxl.load_workbook(WB)
ws = wb['Sheet1']
conv = 0
for c in range(2, ws.max_column + 1):
    cell = ws.cell(1, c)
    v = cell.value
    if isinstance(v, str) and len(v) >= 8:
        try:
            d = datetime.datetime.strptime(v[:10], '%Y-%m-%d')
        except ValueError:
            continue
        cell.value = d
        cell.number_format = 'yyyy-mm-dd'
        conv += 1
wb.save(WB)
print(f"1) converted {conv} header date string(s) -> datetime")

# ---- 2) restore monthly_balance layout fields in the YAML ----
cfg = config_service.load_client_config(WS, 'lanco_fcu')
mb = cfg.get('monthly_balance') or {}
# rebuild with layout fields in canonical order, preserving existing values
new_mb = {
    'source': mb.get('source', 'single'),
    'sheet': 'Sheet1',
    'header_row': 1,
    'pool_name_col': 'A',
    'first_date_col': 'B',
    'filename': mb.get('filename', 'Monthly Loan Balances by Type.xlsx'),
    'saved_path': mb.get('saved_path', WB),
    'pool_map': mb.get('pool_map') or {},
}
cfg['monthly_balance'] = new_mb
config_service.save_client_config(WS, 'lanco_fcu', cfg, overwrite=True)
print("2) restored monthly_balance layout fields:",
      {k: new_mb[k] for k in ('sheet','header_row','pool_name_col','first_date_col')})
