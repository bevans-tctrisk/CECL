"""Verify Phase 9.6 source-flip on live Destinations CU draft."""
import os
import sys
import json
import re

os.environ.setdefault('CECL_WORKSPACE_ROOT', r'Z:\Shared\TCT Files\CECL - CM Files')
sys.path.insert(0, r'C:\Dev\CECL')

from cecl_ui.app import app

DRAFT = r'Z:\Shared\TCT Files\CECL - CM Files\wizard_drafts\destinations_cu__migration.json'
draft = json.load(open(DRAFT, encoding='utf-8'))

print('Pre  hist_co_source=', draft.get('hist_co_source'),
      '  hist_recov_source=', draft.get('hist_recov_source'))
hs = draft.get('hist_scan') or {}
print('Pre  monthly_co=', len(hs.get('monthly_co_files') or []),
      'monthly_recov=', len(hs.get('monthly_recov_files') or []))

with app.test_client() as c:
    with c.session_transaction() as sess:
        sess['setup_state'] = draft
    r = c.get('/setup/step/identity')
    print('GET identity status=', r.status_code)
    with c.session_transaction() as sess:
        st = sess.get('setup_state', {})
        print('Post hist_co_source=', st.get('hist_co_source'),
              '  hist_recov_source=', st.get('hist_recov_source'))
        hs = st.get('hist_scan') or {}
        print('Post monthly_co=', len(hs.get('monthly_co_files') or []),
              'monthly_recov=', len(hs.get('monthly_recov_files') or []))
    r2 = c.get('/setup/step/historical')
    print('Step 5 GET status=', r2.status_code)
    body = r2.data.decode('utf-8', errors='replace')
    co_chosen = re.search(r'name="hist_co_source" value="([^"]+)"[^>]*checked', body)
    print('Step 5 selected co source =', co_chosen.group(1) if co_chosen else None)
    recov_chosen = re.search(r'name="hist_recov_source" value="([^"]+)"[^>]*checked', body)
    print('Step 5 selected recov source =', recov_chosen.group(1) if recov_chosen else None)
