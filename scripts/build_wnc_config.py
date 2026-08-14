import os, yaml
os.environ.setdefault('CECL_WORKSPACE_ROOT', r'Z:\Shared\TCT Files\CECL - CM Files')
WS = os.environ['CECL_WORKSPACE_ROOT']
SHORT = 'wnc_community_cu'
CLIENT = r'Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients\WNC Community CU 66854\Portfolio Management\CECL Migration IDLR'

# alpha loan-type code -> pool (from "Loan Type Codes with Pools.xlsx")
POOL_MAP = {
    'CP': 'Consumer Secured', 'NC': 'Consumer Secured', 'PV': 'Consumer Secured',
    'RS': 'Consumer Secured', 'SS': 'Consumer Secured', 'UC': 'Consumer Secured',
    'OD': 'Consumer Unsecured', 'SC': 'Consumer Unsecured', 'SI': 'Consumer Unsecured',
    'EA': 'Real Estate', 'FM': 'Real Estate', 'HE': 'Real Estate', 'HF': 'Real Estate',
    'HO': 'Real Estate', 'MF': 'Real Estate', 'PF': 'Real Estate', 'V3': 'Real Estate',
    'V5': 'Real Estate', 'VM': 'Real Estate',
}
# numeric share-type codes (Aries Shares / negative shares) -> Ignore (not in loan GL)
for c in ['01', '1', '16', '18', '19', '25', '51', '75', '80', '81']:
    POOL_MAP[c] = 'Ignore'

RE_POOLS = {'Real Estate'}
POOLS = ['Real Estate', 'Consumer Secured', 'Consumer Unsecured']

def pool_entry(name):
    return {'name': name, 'risk_rated': False,   # single-line: no credit scores yet
            'acl_months': 84 if name in RE_POOLS else 36,
            'excluded': False, 'use_default_mgmt_adj': name in RE_POOLS, 'brr': False}

pools = [pool_entry(p) for p in POOLS] + [
    {'name': 'Ignore', 'risk_rated': False, 'acl_months': None, 'excluded': False,
     'use_default_mgmt_adj': False, 'brr': False}]

# FICO grades (defined for when scores arrive next quarter; pools single-line now)
grades = [
    {'label': '760+',    'min_score': 760, 'max_score': 900, 'reserve_rate': 0.0011},
    {'label': '700-759', 'min_score': 700, 'max_score': 759, 'reserve_rate': 0.0025},
    {'label': '680-699', 'min_score': 680, 'max_score': 699, 'reserve_rate': 0.0050},
    {'label': '660-679', 'min_score': 660, 'max_score': 679, 'reserve_rate': 0.0075},
    {'label': '640-659', 'min_score': 640, 'max_score': 659, 'reserve_rate': 0.0110},
    {'label': '620-639', 'min_score': 620, 'max_score': 639, 'reserve_rate': 0.0180},
    {'label': '600-619', 'min_score': 600, 'max_score': 619, 'reserve_rate': 0.0250},
    {'label': '<600',    'min_score': 350, 'max_score': 599, 'reserve_rate': 0.0500},
]

loan_cm = {
    'member_number': 'Account Number', 'current_balance': 'LOAN BALANCE',
    'loan_pool_code': 'LOAN TYPE', 'original_fico_score': 'CREDIT SCORE',
    'current_fico_score': 'CREDIT SCORE', 'open_date': 'LOAN ORIGINAL DATE',
    'original_loan_amount': 'LOAN ORIGINAL AMNT', 'interest_rate': 'Annual Rate 1',
    'total_available_credit': 'CREDIT LIMIT',
}
loan_ma = {'mode': 'delimiter', 'suffix_length': 0, 'delimiter': 'L'}
LOAN_PAT = r'(?i)^\d{8}[\s_]+Ariess?[\s_]+Loans?\b.*\.(xlsx|xls)$'

config = {
    'credit_union': 'WNC Community CU',
    'charter_number': '66854',
    'report_period': '2025-12',
    'file_pattern': LOAN_PAT,
    'date_pattern': r'(20\d{2})(\d{2})(\d{2})',
    'date_format': 'YYYYMMDD',
    'account_suffix_length': 0,
    'member_account': loan_ma,
    'has_header': True,
    'column_mappings': loan_cm,
    'balance_format': {'remove_chars': ['$'], 'accounting_negatives': True},
    'pool_code_split': '',
    'pool_map': POOL_MAP,
    'default_pool': 'Ignore',
    'credit_grades': grades,
    'no_score_label': 'Not Reported',
    'reports': {'tct': False, 'vizo': True, 'vizo_supp': True, 'impdet': True},
    'economic_data': {'state': 'North Carolina', 'county': 'Haywood County',
                      'unemployment_rate': 0.038, 'foreclosures': 2479,
                      'bankruptcies': 8603, 'population': 62000},
    'mgmt_adj': {'ltv_baseline': 0.9, 'probability_factor': 0.35},
    'loan_data_extracts': [
        {'label': 'Aries Loans', 'file_pattern': LOAN_PAT,
         'column_mappings': loan_cm, 'member_account': loan_ma, 'has_header': True},
    ],
    'pools': pools,
    'pool_order': POOLS + ['Ignore'],
    'not_risk_rated': list(POOLS),   # all single-line for now (no credit scores)
    'acl_months_by_pool': {p: (84 if p in RE_POOLS else 36) for p in POOLS},
    'historical_file_formats': {
        'chargeoff': {'has_header': True, 'skip_rows': 0, 'account_col': 0,
                      'code_col': 13, 'amount_col': 10, 'date_col': 8, 'member_col': 0,
                      'member_account': {'mode': 'delimiter', 'suffix_length': 0, 'delimiter': 'L'}},
        'recovery': {'has_header': True, 'skip_rows': 0, 'account_col': 0,
                     'code_col': 13, 'amount_col': 12, 'date_col': 8, 'member_col': 0,
                     'member_account': {'mode': 'delimiter', 'suffix_length': 0, 'delimiter': 'L'}},
    },
    'data_directory': os.path.join(WS, 'Raw_Uploads', SHORT),
    'loan_source_folder': CLIENT,
    'monthly_balance': {'source': 'single'},
    # Booked ACL: WNC's monthly balance file carries no ALLL/ACL row, so
    # pull the booked "Allowance for Credit Losses on Loans" (AAS0048)
    # straight from NCUA Form 5300 for any quarter present in the core.
    'acl': {'use_5300_fallback': True},
}

os.makedirs(config['data_directory'], exist_ok=True)
out = os.path.join(WS, 'client_configs', SHORT + '.yaml')
with open(out, 'w', encoding='utf-8') as f:
    yaml.safe_dump(config, f, sort_keys=False, allow_unicode=True)
print('WROTE', out)
print('pools (all single-line):', POOLS, '| pool_map codes:', len(POOL_MAP))
c2 = yaml.safe_load(open(out, encoding='utf-8'))
print('reparse OK; extracts:', [e['label'] for e in c2['loan_data_extracts']])
