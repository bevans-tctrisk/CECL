"""Build client_configs/lanco_fcu.yaml for Lanco FCU (16657).

Grade-rated, business-heavy, multi-extract CU. This first pass wires the
main Aires Loan Data extract + Negative Shares and reads the pool map
from the CU's own "Loan Type Codes with Pools.xlsx". Other extracts
(Credit Cards, Lendkey Student/Consolidation, Participations) and the
monthly-balance build are layered on next once the Aires import ties out.
"""
import os
import pandas as pd
import yaml

os.environ.setdefault('CECL_WORKSPACE_ROOT', r'Z:\Shared\TCT Files\CECL - CM Files')
WS = os.environ['CECL_WORKSPACE_ROOT']
SHORT = 'lanco_fcu'
CLIENT = (r'Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients'
          r'\Lanco FCU 16657\Portfolio Management\CECL Migration IDLR')

# ── Pool map: read the CU's code->pool file (col 'Pools', fallback
#    'Pool after Dec 2018', else Ignore). Negative Shares -> Ignore. ──
pm = pd.read_excel(os.path.join(CLIENT, 'Lanco FCU Loan Type Codes with Pools.xlsx'),
                   sheet_name='Assets', header=0, dtype=str)
pm.columns = [str(c).strip() for c in pm.columns]
POOL_MAP = {}
for _, r in pm.iterrows():
    code = str(r.get('Loan Type Code') or '').strip()
    if not code or code.lower() == 'nan':
        continue
    pool = str(r.get('Pools') or '').strip()
    if not pool or pool.lower() == 'nan':
        pool = str(r.get('Pool after Dec 2018') or '').strip()
    if not pool or pool.lower() == 'nan':
        pool = 'Ignore'
    if pool in ('Negative Shares', 'Neg Shares'):
        pool = 'Ignore'
    POOL_MAP[code] = pool
# Extra text codes from other extracts.
POOL_MAP.update({'Visa': 'Credit Cards', 'Student Loan': 'Student Loans',
                 'Neg Shares': 'Ignore'})

POOLS = ['Real Estate', 'Consumer Secured', 'Consumer Indirect',
         'Consumer Unsecured', 'Credit Cards', 'Student Loans', 'Business Loans']
# Business Loans have no consumer FICO -> single-line (no grade breakout);
# everything else is grade-rated.
SINGLE_LINE = {'Business Loans'}

def pool_entry(name):
    return {'name': name, 'risk_rated': name not in SINGLE_LINE, 'acl_months': 36,
            'excluded': False, 'use_default_mgmt_adj': False, 'brr': False}

pools = [pool_entry(p) for p in POOLS] + [
    {'name': 'Ignore', 'risk_rated': False, 'acl_months': None,
     'excluded': False, 'use_default_mgmt_adj': False, 'brr': False}]

# ── Grades A-E, current (post 2025-10-27) Lanco bands ──
grades = [
    {'label': 'A', 'min_score': 740, 'max_score': 900, 'reserve_rate': 0.0011},
    {'label': 'B', 'min_score': 700, 'max_score': 739, 'reserve_rate': 0.0025},
    {'label': 'C', 'min_score': 650, 'max_score': 699, 'reserve_rate': 0.0050},
    {'label': 'D', 'min_score': 610, 'max_score': 649, 'reserve_rate': 0.0116},
    {'label': 'E', 'min_score': 350, 'max_score': 609, 'reserve_rate': 0.0250},
]

# ── Aires main extract ──
aires_cm = {
    'member_number': 'Account Number', 'current_balance': 'CURRENT BAL',
    'loan_pool_code': 'Loan Type Code',
    'original_fico_score': 'CREDIT SCORE', 'current_fico_score': 'CREDIT SCORE',
    'open_date': 'LOAN DATE', 'original_loan_amount': 'ORIGINAL AMT',
    'interest_rate': 'INT RATE', 'total_available_credit': 'CREDIT LIMIT',
    'days_delinquent': 'DAYS DQ',
}
aires_ma = {'mode': 'fixed_suffix', 'suffix_length': 0}
AIRES_PAT = r'(?i)^Aires Loan Data \d{2}-\d{2}-\d{4} WO Charge-offs\.(xlsx|xls)$'
NEGSHARE_PAT = r'(?i)^Negative Shares \d{1,2}-\d{1,2}-\d{2,4}( V\d+)?\.(xlsx|xls)$'
negshare_cm = {'member_number': 'Account Number', 'current_balance': 'Current Balance',
               'loan_pool_code': 'Loan Type'}

# ── Credit Card extract (AIRES format) ──
card_cm = {
    'member_number': 'Account Number', 'current_balance': 'Current Loan Balance',
    'loan_pool_code': 'Loan Type Code',
    'original_fico_score': 'Credit Score', 'current_fico_score': 'Credit Score',
    'open_date': 'Date of Loan', 'original_loan_amount': 'Original Loan Amount',
    'interest_rate': 'Interest Rate (APR)', 'total_available_credit': 'Credit Limit',
    'days_delinquent': 'Days Delinquent',
}
CARD_PAT = r'(?i)^Credit Cardholder - AIRES \d{1,2}-\d{1,2}-\d{2,4}( V\d+)?\.(xlsx|xls)$'

# ── Lendkey student-loan extracts (held/lender balance) ──
student_cm = {
    'member_number': 'Borr_ID', 'current_balance': 'Loan_Bal_-_Month_End_Lender',
    'loan_pool_code': 'Loan Type Code',
    'original_fico_score': 'Borr_FICO', 'current_fico_score': 'Borr_FICO',
    'open_date': 'First_Disbursement_Date',
    'days_delinquent': 'Days_Delinquent',
}
STUDENT_PAT = r'(?i)^Lendkey Student Loans \d{1,2}-\d{1,2}-\d{4}\.(xlsx|xls)$'
CONSOL_PAT = r'(?i)^Lendkey Consolidation \d{1,2}-\d{1,2}-\d{4}\.(xlsx|xls)$'

config = {
    'credit_union': 'Lanco FCU',
    'charter_number': '16657',
    'report_period': '2025-12',
    'file_pattern': AIRES_PAT,
    'date_pattern': r'(\d{2})-(\d{2})-(\d{4})',
    'date_format': 'MMDDYYYY',
    'account_suffix_length': 0,
    'member_account': aires_ma,
    'has_header': True,
    'column_mappings': aires_cm,
    'balance_format': {'remove_chars': ['$'], 'accounting_negatives': True},
    'pool_code_split': '',
    'pool_map': POOL_MAP,
    'default_pool': 'Ignore',
    'credit_grades': grades,
    'no_score_label': 'Not Reported',
    'reports': {'tct': False, 'vizo': True, 'vizo_supp': True, 'impdet': True,
                'mgmt_adj_napkin': True},
    'economic_data': {'state': 'Pennsylvania', 'county': 'Lancaster County',
                      'unemployment_rate': 0.038, 'foreclosures': 1800,
                      'bankruptcies': 4200, 'population': 552984},
    'mgmt_adj': {'ltv_baseline': 0.9, 'probability_factor': 0.35},
    'loan_data_extracts': [
        {'label': 'Aires Loan Data', 'file_pattern': AIRES_PAT,
         'column_mappings': aires_cm, 'member_account': aires_ma, 'has_header': True},
        {'label': 'Credit Cardholder', 'file_pattern': CARD_PAT,
         'column_mappings': card_cm, 'member_account': aires_ma, 'has_header': True},
        {'label': 'Lendkey Student Loans', 'file_pattern': STUDENT_PAT,
         'column_mappings': student_cm, 'member_account': aires_ma, 'has_header': True},
        {'label': 'Lendkey Consolidation', 'file_pattern': CONSOL_PAT,
         'column_mappings': student_cm, 'member_account': aires_ma, 'has_header': True},
        {'label': 'Negative Shares', 'file_pattern': NEGSHARE_PAT,
         'column_mappings': negshare_cm,
         'member_account': {'mode': 'delimiter', 'suffix_length': 0, 'delimiter': '-'},
         'has_header': True},
    ],
    'pools': pools,
    'pool_order': POOLS + ['Ignore'],
    'not_risk_rated': sorted(SINGLE_LINE),
    'acl_months_by_pool': {p: 36 for p in POOLS},
    'historical_file_formats': {
        # Transaction-level "Charge off and Recovery Template MM-DD-YY.xlsx":
        # date(0) | account(1) | trailer(2) | name(3) | Recovery Amount(4) |
        # Charge off Amount(5) | Loan Type Code(6) | ...  Combined mode.
        # strict_columns=True: two amount cols (recovery vs charge-off) mean
        # the header-text heuristic can't tell them apart, so force indices.
        'chargeoff': {'has_header': True, 'skip_rows': 0, 'account_col': 1,
                      'code_col': 6, 'amount_col': 5, 'date_col': 0, 'member_col': 1,
                      'strict_columns': True,
                      'member_account': {'mode': 'fixed_suffix', 'suffix_length': 0}},
        'recovery': {'has_header': True, 'skip_rows': 0, 'account_col': 1,
                     'code_col': 6, 'amount_col': 4, 'date_col': 0, 'member_col': 1,
                     'strict_columns': True,
                     'member_account': {'mode': 'fixed_suffix', 'suffix_length': 0}},
    },
    'data_directory': os.path.join(WS, 'Raw_Uploads', SHORT),
    'loan_source_folder': CLIENT,
    'monthly_balance': {'source': 'single'},
    'acl': {'use_5300_fallback': True},
}

os.makedirs(config['data_directory'], exist_ok=True)
out = os.path.join(WS, 'client_configs', SHORT + '.yaml')
with open(out, 'w', encoding='utf-8') as f:
    yaml.safe_dump(config, f, sort_keys=False, allow_unicode=True)
print('WROTE', out)
print('pools:', POOLS)
print('pool_map codes:', len(POOL_MAP), '| distinct pools:', sorted(set(POOL_MAP.values())))
