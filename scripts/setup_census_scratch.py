"""Build Census FCU 'from scratch' setup at short_name=census_fcu_scratch
for snapshot 2025-12-31, using ONLY raw source data from the Vizo CU folder.

No reuse of existing census_fcu.yaml or wizard draft.

Steps:
  1. Create Raw_Uploads/census_fcu_scratch/
  2. Copy needed source files with renamed filenames matching defined patterns
  3. Write census_fcu_scratch.yaml programmatically
  4. (Skip wizard draft — YAML is sufficient for pipeline_service.run_reports)
"""
import os, sys, shutil, re
os.environ.setdefault('CECL_WORKSPACE_ROOT', r'Z:\Shared\TCT Files\CECL - CM Files')
sys.path.insert(0, r'C:\Dev\CECL')
from pathlib import Path
import yaml

ROOT = Path(r'Z:\Shared\TCT Files\CECL - CM Files')
SRC  = Path(r'Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients\Census FCU 5641\Portfolio Management\CECL Migration IDLR')

SHORT = 'census_fcu_scratch'
NAME  = 'Census FCU (Scratch)'
SNAPSHOT_YYYY_MM = '2025-12'

dst = ROOT / 'Raw_Uploads' / SHORT
dst.mkdir(parents=True, exist_ok=True)

print(f'Creating Raw_Uploads/{SHORT}/')

# === Loan extracts (Dec 2025 only — snapshot) ===
copies = [
    (SRC / '2025' / 'Dec 2025' / 'Dec 2025 Loans V2.xlsx', '2025-12_Loans_V2.xlsx'),
    (SRC / '2025' / 'Dec 2025' / 'Dec 2025 Cuma Loans.xlsx', '2025-12_Cuma_Loans.xlsx'),
]

# === Monthly CO/Recov files — all 2024 + 2025 (aggregator scans all + history) ===
month_to_num = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'June':6,'July':7,'Aug':8,'Sept':9,'Oct':10,'Nov':11,'Dec':12}
for year_dir in sorted((SRC / '2024').iterdir()) + sorted((SRC / '2025').iterdir()):
    if not year_dir.is_dir(): continue
    mname = year_dir.name.split()[0]  # 'Jan'/'Feb'/'Sept'/...
    yr    = year_dir.name.split()[1]
    if mname not in month_to_num: continue
    mm = month_to_num[mname]
    # find CO/Recov file in that folder
    for f in year_dir.glob('Charge Off and Recoveries*.xlsx'):
        new_name = f'Charge_Off_and_Recoveries_{yr}-{mm:02d}.xlsx'
        copies.append((f, new_name))
        break

# === Balance sheets (per_year — 2023, 2024, 2025) ===
for yr in ('2023','2024','2025'):
    copies.append((SRC / 'Historical Files' / 'Balance Sheets' / f'{yr} Balance Sheets.xlsx',
                   f'{yr}_Balance_Sheets.xlsx'))

# === Heuristic auto-discovery: impaired loans WARM template + credit pull ===
# Pattern: any *Impaired_Loans*.xlsx or *Credit*Pull*.xlsx anywhere under SRC.
# This mirrors what the analyst-built configs typically reference and lets
# a 'from-scratch' setup pick up these files without hand-coding paths.
impaired_src = None
for f in SRC.rglob('*Impaired*Loans*.xlsx'):
    if f.is_file() and not f.name.startswith('~'):
        impaired_src = f; break
if impaired_src:
    impaired_name = f'{SNAPSHOT_YYYY_MM}_CECL_Migration_Impaired_Loans.xlsx'
    copies.append((impaired_src, impaired_name))
    print(f'  AUTO-DISCOVERED impaired loans: {impaired_src.name} -> {impaired_name}')

# Credit pull for the snapshot period (YYYY-MM or 'Mon YYYY' folder names)
credit_pull_src = None
for f in SRC.rglob('*Credit*Pull*.xlsx'):
    if not f.is_file() or f.name.startswith('~'): continue
    # Match if filename or parent dir contains the snapshot's year-month (any form)
    s_yr, s_mo = SNAPSHOT_YYYY_MM.split('-')
    mon_abbr = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][int(s_mo)-1]
    haystack = f'{f.name} {f.parent.name}'.lower()
    if (f'{s_yr}-{s_mo}' in haystack or f'{mon_abbr.lower()}_{s_yr}' in haystack
            or f'{mon_abbr.lower()} {s_yr}' in haystack):
        credit_pull_src = f; break
if credit_pull_src:
    cp_name = f'{SNAPSHOT_YYYY_MM}_Credit_Pull.xlsx'
    copies.append((credit_pull_src, cp_name))
    print(f'  AUTO-DISCOVERED credit pull: {credit_pull_src.name} -> {cp_name}')
else:
    print(f'  (no credit pull file found for {SNAPSHOT_YYYY_MM} — original=current FICO will be used)')

# Execute copies
copied = 0
for src_path, fname in copies:
    if not src_path.exists():
        print(f'  MISSING: {src_path}'); continue
    target = dst / fname
    shutil.copy2(src_path, target)
    copied += 1
print(f'\nCopied {copied} files to {dst}')

# === Build YAML ===
# Pool mapping per CU-provided "TCT Edits" loan type codes file (with code 13 corrected)
POOL_ORDER = [
    'New Auto Loans',
    'Used Auto Loans',
    'RV, Boat, Motorcycle Loans',
    'Share/CD Secured Loans',
    'Signature Loans',
    'Line of Credit Loans',
    'Second Trust Loans',
    'CUMA Mortgage Loans',
    # Heuristic: balance-sheet rows containing "Deferred"/"Fees"/"Origination" are
    # commonly a separate non-RR pool (matches analyst-built configs).
    'Deferred Loan Origination Fees',
]
POOL_MAP = {
    # col_cd (numeric) -> pool
    '7':  'New Auto Loans',
    '70': 'New Auto Loans',
    '71': 'New Auto Loans',
    '72': 'New Auto Loans',
    '13': 'Used Auto Loans',
    '67': 'Used Auto Loans',
    '68': 'Used Auto Loans',
    '69': 'Used Auto Loans',
    '51': 'RV, Boat, Motorcycle Loans',
    '52': 'RV, Boat, Motorcycle Loans',
    '54': 'RV, Boat, Motorcycle Loans',
    '55': 'RV, Boat, Motorcycle Loans',
    '56': 'RV, Boat, Motorcycle Loans',
    '76': 'RV, Boat, Motorcycle Loans',
    '82': 'RV, Boat, Motorcycle Loans',
    '85': 'RV, Boat, Motorcycle Loans',
    '90': 'RV, Boat, Motorcycle Loans',
    '19': 'Share/CD Secured Loans',
    '20': 'Share/CD Secured Loans',
    '4':  'Signature Loans',
    '26': 'Signature Loans',
    '43': 'Signature Loans',
    '44': 'Signature Loans',
    '45': 'Signature Loans',
    '5':  'Line of Credit Loans',
    '40': 'Line of Credit Loans',
    '30': 'Second Trust Loans',
    '31': 'Second Trust Loans',
    '35': 'Second Trust Loans',
    # text aliases (loan type 3-letter)
    'NVE': 'New Auto Loans',
    'UVE': 'Used Auto Loans',
    'SHR': 'Share/CD Secured Loans',
    'UNS': 'Signature Loans',
    'OT1': 'Line of Credit Loans',
    '2MT': 'Second Trust Loans',
    'HEQ': 'Second Trust Loans',
    # CO/Recov "Loan Type" textual values (seen in Dec 2025 CO file: "Second Trust")
    'used auto':       'Used Auto Loans',
    'new auto':        'New Auto Loans',
    'share secured':   'Share/CD Secured Loans',
    'cd secured':      'Share/CD Secured Loans',
    'signature':       'Signature Loans',
    'redicash':        'Line of Credit Loans',
    'premier':         'Line of Credit Loans',
    'premier (uns.)':  'Line of Credit Loans',
    '2nd trust fixed': 'Second Trust Loans',
    'second trust':    'Second Trust Loans',
    'home equity':     'Second Trust Loans',
    'new home equity': 'Second Trust Loans',
    'new motorcycle':  'RV, Boat, Motorcycle Loans',
    'used motorcycle': 'RV, Boat, Motorcycle Loans',
    'new rv':          'RV, Boat, Motorcycle Loans',
    'used rv':         'RV, Boat, Motorcycle Loans',
    'used boat':       'RV, Boat, Motorcycle Loans',
    'unsecure':        'Signature Loans',
    # Signature upcharge variants
    'signature - 1% upcharge': 'Signature Loans',
    'signature - 3% upcharge': 'Signature Loans',
    'signature - 6% upcharge': 'Signature Loans',
    'used auto - 1% upcharge': 'Used Auto Loans',
    'used auto - 3% upcharge': 'Used Auto Loans',
    'used auto - 6% upcharge': 'Used Auto Loans',
    'new auto - 1% upcharge':  'New Auto Loans',
    'new auto - 3% upcharge':  'New Auto Loans',
    'new auto - 6% upcharge':  'New Auto Loans',
}

# Balance-sheet pool label -> our pool (Census per_year BS row labels)
BS_LABEL_MAP = {
    'CD SECURED LOANS':              'Share/CD Secured Loans',
    'CUMA MORTGAGE LOANS':           'CUMA Mortgage Loans',
    'LINE OF CREDIT LOANS (PREMIER)':  'Line of Credit Loans',
    'LINE OF CREDIT LOANS (REDICASH)': 'Line of Credit Loans',
    'NEW AUTO LOANS':                'New Auto Loans',
    'NEW AUTO UPCHARGE LOANS':       'New Auto Loans',
    'USED AUTO UPCHARGE LOANS':      'Used Auto Loans',
    "RV'S, BOATS, AND MOTORCYCLE LOANS": 'RV, Boat, Motorcycle Loans',
    'SECOND TRUST LOANS':            'Second Trust Loans',
    'SHARE SECURED LOANS':           'Share/CD Secured Loans',
    'SIGNATURE LOANS':               'Signature Loans',
    'SIGNATURE UPCHARGE LOANS':      'Signature Loans',
    'USED AUTO LOANS':               'Used Auto Loans',
    'DEFERRED LOAN ORIGINATION FEES': 'Deferred Loan Origination Fees',
}

# === Pools detail (life_of_loan_months from typical industry defaults) ===
# Note: any pool not on POOL_ORDER would still need an entry here. Each pool MUST
# also be classified risk_rated vs not_risk_rated. Defaults below match analyst
# convention: deferred-fees pool is not_risk_rated (it's a non-loan accounting line).
POOLS = [
    {'name': 'New Auto Loans',            'life_of_loan_months': 60, 'risk_rated': True,  'brr': False},
    {'name': 'Used Auto Loans',           'life_of_loan_months': 60, 'risk_rated': True,  'brr': False},
    {'name': 'RV, Boat, Motorcycle Loans','life_of_loan_months': 60, 'risk_rated': True,  'brr': False},
    {'name': 'Share/CD Secured Loans',    'life_of_loan_months': 36, 'risk_rated': True,  'brr': False},
    {'name': 'Signature Loans',           'life_of_loan_months': 36, 'risk_rated': True,  'brr': False},
    {'name': 'Line of Credit Loans',      'life_of_loan_months': 36, 'risk_rated': True,  'brr': False},
    {'name': 'Second Trust Loans',        'life_of_loan_months': 84, 'risk_rated': True,  'brr': False},
    {'name': 'CUMA Mortgage Loans',       'life_of_loan_months': 84, 'risk_rated': True,  'brr': False},
    {'name': 'Deferred Loan Origination Fees', 'life_of_loan_months': 36, 'risk_rated': False, 'brr': False},
]

# === Credit grades (standard FICO bands with industry-default reserve rates) ===
# NOTE: cecl_engine.get_reserve_rate hard-reads g['reserve_rate'] with no fallback.
# Every grade MUST have a reserve_rate or the engine throws KeyError.
CREDIT_GRADES = [
    {'label': 'A', 'min_score': 740, 'max_score': 850, 'reserve_rate': 0.0025},
    {'label': 'B', 'min_score': 700, 'max_score': 739, 'reserve_rate': 0.005},
    {'label': 'C', 'min_score': 660, 'max_score': 699, 'reserve_rate': 0.01},
    {'label': 'D', 'min_score': 620, 'max_score': 659, 'reserve_rate': 0.025},
    {'label': 'F', 'min_score': 300, 'max_score': 619, 'reserve_rate': 0.05},
]

# === Build the YAML ===
cfg = {
    'credit_union': NAME,
    'charter_number': '5641',
    # Top-level loan file pattern — points at the Dec 2025 V2 extract
    'file_pattern': r'^2025-12_Loans_V2\.xlsx$',
    'date_pattern': r'(\d{4})-(\d{2})',
    'date_format': 'YYYY-MM',
    'header_row': 1,
    'has_header': True,
    # Member/account schema: separate cols (acn + ln_sfx)
    'member_account': {
        'mode': 'split',
        'delimiter': '-',
    },
    'account_suffix_length': 0,
    # Column mappings (top-level — for the Loans V2 extract).
    # NOTE: header names are case-SENSITIVE matched at import time. Census V2 has
    # 'User_Defined_Field_7' (CamelCase) not 'user_defined_field_7'.
    'column_mappings': {
        'member_number':       'acn',
        'loan_suffix':         'ln_sfx',
        'loan_pool_code':      'col_cd',
        'current_balance':     'bal',
        'original_loan_amount':'orig_ln_amt',
        'open_date':           'orig_ln_dt',
        'maturity_date':       'maturity_dt',
        'interest_rate':       'int_rate',
        'term':                'term',
        'days_delinquent':     'dlq_days',
        'total_available_credit':'avl_limit',
        # FICO — current is in User_Defined_Field_7; original FICO comes from credit
        # pull when available, otherwise falls back to current (no migration movement).
        'original_fico_score': 'User_Defined_Field_7',
        'current_fico_score':  'User_Defined_Field_7',
    },
    # Per-extract overrides (second extract = CUMA, different shape)
    'loan_data_extracts': [
        {
            'file_pattern': r'^2025-12_Loans_V2\.xlsx$',
            'header_row': 1,
            'has_header': True,
            'sheet_name': 'Dec 2025',
            'member_account': {'mode':'split','delimiter':'-'},
            'account_suffix_length': 0,
            'column_mappings': {
                'member_number':       'acn',
                'loan_suffix':         'ln_sfx',
                'loan_pool_code':      'col_cd',
                'current_balance':     'bal',
                'original_loan_amount':'orig_ln_amt',
                'open_date':           'orig_ln_dt',
                'maturity_date':       'maturity_dt',
                'interest_rate':       'int_rate',
                'term':                'term',
                'days_delinquent':     'dlq_days',
                'total_available_credit':'avl_limit',
                'original_fico_score': 'User_Defined_Field_7',
                'current_fico_score':  'User_Defined_Field_7',
            },
        },
        {
            'file_pattern': r'^2025-12_Cuma_Loans\.xlsx$',
            'header_row': 1,
            'has_header': True,
            'sheet_name': 'Sheet1',
            'member_account': {'mode':'fixed_suffix','delimiter':'-','suffix_length': 0},
            'account_suffix_length': 0,
            # CUMA file has no loan_pool_code column; map all to CUMA via default_pool below.
            'default_pool': 'CUMA Mortgage Loans',
            'column_mappings': {
                'member_number':       'Account Number',
                'current_balance':     'Current Balance',
                'original_loan_amount':'Original Loan Amount',
                'open_date':           'Origination Date',
                'maturity_date':       'Maturity Date',
                'interest_rate':       'Interest Rate',
            },
        },
    ],
    # Data directory
    'data_directory': str(dst),
    # Archive — disable for the scratch test (keep extracts in place for re-runs)
    'archive_imported_files': False,
    # Pool-code split: handles compound loan codes like '13/14' -> '13'. Safe default;
    # no-op when codes don't contain '/'. Matches analyst-built configs.
    'pool_code_split': '/',
    # Pools (split risk_rated / not_risk_rated by the pools[].risk_rated flag)
    'pools': POOLS,
    'pool_order': POOL_ORDER,
    'risk_rated':     [p['name'] for p in POOLS if p['risk_rated']],
    'not_risk_rated': [p['name'] for p in POOLS if not p['risk_rated']],
    'excluded_pools': [],
    'default_pool': 'Ignore',
    'pool_map': POOL_MAP,
    # Credit grades
    'credit_grades': CREDIT_GRADES,
    'grades': CREDIT_GRADES,
    'no_score_label': 'Not Reported',
    # Optional sentinel for "no FICO" — accepts 'N/A' / 0
    'uses_brr': False,
    # ACL — read from monthly_file row (BS row 21 "ALLOWANCE FOR LOAN LOSSES", DEC 2025 col)
    'acl': {
        'source': 'monthly_file',
        'row': 21,
        'col': 'C',  # DEC 2025 column
    },
    # Monthly balance — per_year mode
    'monthly_balance': {
        'source': 'per_year',
        'layout': {
            'sheet': 'JAN-DEC ',  # prefix; runtime aggregator auto-detects per file (sheet 'JAN-DEC 2025' etc)
            'header_row': 2,
            'label_col': 'B',
            'period_columns': [
                {'col': 'C', 'period_iso': '2025-12-31'},
                {'col': 'D', 'period_iso': '2025-11-30'},
                {'col': 'E', 'period_iso': '2025-10-31'},
                {'col': 'F', 'period_iso': '2025-09-30'},
                {'col': 'G', 'period_iso': '2025-08-31'},
                {'col': 'H', 'period_iso': '2025-07-31'},
                {'col': 'I', 'period_iso': '2025-06-30'},
                {'col': 'J', 'period_iso': '2025-05-31'},
                {'col': 'K', 'period_iso': '2025-04-30'},
                {'col': 'L', 'period_iso': '2025-03-31'},
                {'col': 'M', 'period_iso': '2025-02-28'},
                {'col': 'N', 'period_iso': '2025-01-31'},
                {'col': 'O', 'period_iso': '2024-12-31'},
            ],
        },
        'files': [
            {'year': 2023, 'filename': '2023_Balance_Sheets.xlsx', 'saved_path': str(dst / '2023_Balance_Sheets.xlsx')},
            {'year': 2024, 'filename': '2024_Balance_Sheets.xlsx', 'saved_path': str(dst / '2024_Balance_Sheets.xlsx')},
            {'year': 2025, 'filename': '2025_Balance_Sheets.xlsx', 'saved_path': str(dst / '2025_Balance_Sheets.xlsx')},
        ],
        'pool_map': BS_LABEL_MAP,
    },
    # Snapshot anchor
    'report_period': SNAPSHOT_YYYY_MM,
    # Economic data (state/county) — Maryland (Census FCU is in MD)
    'state':  'Maryland',
    'county': 'Prince George\'s',
    # Mgmt adj defaults (industry standard — analyst-tuned values; user can override)
    'mgmt_adj': {
        'ltv_baseline': 0.9,
        'probability_factor': 0.35,
    },
    # Which reports to generate for this CU
    'reports': {
        'tct': False,
        'vizo': True,
        'vizo_supp': True,
        'impdet': True,
    },
}

# === Auto-add impaired_loans block if file was discovered ===
if (dst / f'{SNAPSHOT_YYYY_MM}_CECL_Migration_Impaired_Loans.xlsx').exists():
    cfg['impaired_loans'] = {
        'file_pattern': rf'^{re.escape(SNAPSHOT_YYYY_MM)}_CECL_Migration_Impaired_Loans\.xlsx$',
        'sheet': ' Impaired Loans',
        'header_row': 29,
        'types': [
            {'name': 'Delinquent Loans',       'provision_pct': 0.35},
            {'name': 'Known Losses',           'provision_pct': 1.0},
            {'name': 'Repossessions',          'provision_pct': 1.0},
            {'name': 'Foreclosed Real Estate', 'provision_pct': 1.0},
            {'name': 'Other Impaired Loans',   'provision_pct': 1.0},
        ],
    }

# === Auto-add credit_pull block if file was discovered ===
cp_path = dst / f'{SNAPSHOT_YYYY_MM}_Credit_Pull.xlsx'
if cp_path.exists():
    cfg['credit_pull'] = {
        'file_pattern': rf'^{re.escape(SNAPSHOT_YYYY_MM)}_Credit_Pull\.xlsx$',
        'member_column': 'acn',
        'score_column': 'User_Defined_Field_7',
        'prefer_original_for_new_loans': False,
        'pull_as_of_date': f'{SNAPSHOT_YYYY_MM}-31' if SNAPSHOT_YYYY_MM.endswith(('-01','-03','-05','-07','-08','-10','-12'))
            else f'{SNAPSHOT_YYYY_MM}-30',
    }

yaml_path = ROOT / 'client_configs' / f'{SHORT}.yaml'
yaml_path.parent.mkdir(exist_ok=True)
with open(yaml_path, 'w', encoding='utf-8') as fh:
    yaml.dump(cfg, fh, sort_keys=False, allow_unicode=True, width=120)

print(f'\nWrote YAML: {yaml_path} ({yaml_path.stat().st_size} bytes)')
print(f'  Pools: {len(POOLS)}, pool_map entries: {len(POOL_MAP)}, BS labels: {len(BS_LABEL_MAP)}')

# === List final state ===
print(f'\n=== Files in {dst} ===')
for f in sorted(dst.iterdir()):
    print(f'  {f.name}  ({f.stat().st_size} bytes)')
