import os
os.environ['CECL_WORKSPACE_ROOT'] = r'Z:\Shared\TCT Files\CECL - CM Files'
os.environ.setdefault('PYTHONPATH', r'C:\dev\CECL')
import cecl_credentials
os.environ['DATABASE_URL'] = cecl_credentials.get_database_url()
from cecl_ui.services import pipeline_service as ps

SHORT = 'mcdowell_cornerstone_cu'
for snap in ['2025-12-31', '2026-03-31']:
    print('\n' + '=' * 70)
    print('REPORTS for', snap)
    paths, out = ps.run_reports(SHORT, snapshot_date=snap, reports=['vizo', 'vizo_supp'])
    for p in paths:
        print('  ', p)
    if not paths:
        print('  NO REPORTS. stdout tail:')
        print('\n'.join(out.splitlines()[-25:]))
    try:
        ip = ps.run_impdet_report(SHORT, snapshot_date=snap)
        print('   impdet:', ip)
    except Exception as e:
        print('   impdet ERROR:', e)
