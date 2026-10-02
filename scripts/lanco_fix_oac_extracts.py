"""Fix Lanco OAC (unfunded commitments + negative share) + restore the
loan_data_extracts that a prior save round-trip dropped.

Issues found:
 1) config lost `loan_data_extracts` (only top-level single-extract remained)
    -> multi-file import (credit cards / student loans / negative shares) and
    the unfunded-commitment undrawn calc lose their per-file column maps.
 2) top-level Aires `file_pattern` lacked `\\s*` before `\\.` so it did NOT
    match the real file 'Aires Loan Data 06-30-2026 WO Charge-offs .xlsx'
    (trailing space) -> unfunded found no extract -> $0.
 3) negative-share `balance_pattern`='Negative Share File' didn't match
    'Negative Shares 06-30-2026.xlsx'; `co_summary_pattern` didn't match
    'Charge off and Recovery Neg Shares 06-30-2026.xlsx' -> $0 balance/rate.

Rebuilds loan_data_extracts with INDEPENDENT dicts (no YAML anchors, which are
fragile across the app's safe_load/safe_dump round-trips).
"""
import os, sys, copy
sys.path.insert(0, r'C:\dev\CECL')
os.environ['CECL_WORKSPACE_ROOT'] = r'Z:\Shared\TCT Files\CECL - CM Files'
import cecl_credentials
os.environ['DATABASE_URL'] = cecl_credentials.get_database_url()
from cecl_ui.services import config_service

WS = r'Z:\Shared\TCT Files\CECL - CM Files'
cfg = config_service.load_client_config(WS, 'lanco_fcu')

AIRES_PAT = r'(?i)^Aires Loan Data \d{2}-\d{2}-\d{4} WO Charge-offs\s*\.(xlsx|xls)$'
aires_cm = copy.deepcopy(cfg.get('column_mappings') or {})           # == id001 (Aires)
aires_ma = copy.deepcopy(cfg.get('member_account') or {'mode': 'fixed_suffix', 'suffix_length': 4})

lendkey_cm = {'member_number': 'Borr_ID', 'current_balance': 'Loan_Bal_-_Month_End_Lender',
              'loan_pool_code': 'Loan Type Code', 'original_fico_score': 'Borr_FICO',
              'current_fico_score': 'Borr_FICO', 'open_date': 'First_Disbursement_Date',
              'days_delinquent': 'Days_Delinquent'}
fixed0 = {'mode': 'fixed_suffix', 'suffix_length': 0}

cfg['loan_data_extracts'] = [
    {'label': 'Aires Loan Data', 'file_pattern': AIRES_PAT,
     'column_mappings': aires_cm, 'member_account': aires_ma, 'has_header': True},
    {'label': 'Credit Cardholder',
     'file_pattern': r'(?i)^Credit Cardholder - AIRES \d{1,2}-\d{1,2}-\d{2,4}( V\d+)?\.(xlsx|xls)$',
     'column_mappings': {'member_number': 'Account Number', 'current_balance': 'Current Loan Balance',
                         'loan_pool_code': 'Loan Type Code', 'original_fico_score': 'Credit Score',
                         'current_fico_score': 'Credit Score', 'open_date': 'Date of Loan',
                         'original_loan_amount': 'Original Loan Amount', 'interest_rate': 'Interest Rate (APR)',
                         'total_available_credit': 'Credit Limit', 'days_delinquent': 'Days Delinquent'},
     'member_account': dict(fixed0), 'has_header': True},
    {'label': 'Lendkey Student Loans',
     'file_pattern': r'(?i)^Lendkey Student Loans \d{1,2}-\d{1,2}-\d{4}\.(xlsx|xls)$',
     'column_mappings': dict(lendkey_cm), 'member_account': dict(fixed0), 'has_header': True},
    {'label': 'Lendkey Consolidation',
     'file_pattern': r'(?i)^Lendkey Consolidation \d{1,2}-\d{1,2}-\d{4}\.(xlsx|xls)$',
     'column_mappings': dict(lendkey_cm), 'member_account': dict(fixed0), 'has_header': True},
    {'label': 'Negative Shares',
     'file_pattern': r'(?i)^Negative Shares \d{1,2}-\d{1,2}-\d{2,4}( V\d+)?\.(xlsx|xls)$',
     'column_mappings': {'member_number': 'Account Number', 'current_balance': 'Current Balance',
                         'loan_pool_code': 'Loan Type'},
     'member_account': {'mode': 'delimiter', 'suffix_length': 0, 'delimiter': '-'}, 'has_header': True},
]

# fix top-level Aires file_pattern (import fallback) to tolerate the trailing space
cfg['file_pattern'] = AIRES_PAT

# fix negative-share OAC patterns
for o in cfg.get('other_allowance_considerations') or []:
    if o.get('source') == 'negative_share':
        o['balance_pattern'] = r'(?i)^Negative Shares \d'
        o['co_summary_pattern'] = r'(?i)^Charge off and Recovery Neg Shares'
        # leave co_quarterly_pattern (no quarterly file for Lanco)

config_service.save_client_config(WS, 'lanco_fcu', cfg, overwrite=True)
print("saved. loan_data_extracts:", [e['label'] for e in cfg['loan_data_extracts']])
print("top file_pattern:", cfg['file_pattern'])
ns = next(o for o in cfg['other_allowance_considerations'] if o.get('source')=='negative_share')
print("NS balance_pattern:", ns['balance_pattern'], "| co_summary_pattern:", ns['co_summary_pattern'])
