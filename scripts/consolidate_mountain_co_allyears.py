"""Build an ALL-YEARS (2018-2026) Charge-Off / Recovery consolidation for
Mountain CU: individual loans BY POOL, reading every monthly CO/Recovery file
from the client source folder. Handles the varied file formats across years and
flags GL-level aggregate rows (2018-19 Call Report + Negative Shares) that have
no individual-loan detail. Saves to a NEW file (does not overwrite the
2025/2026 consolidated file)."""
import os, re
import yaml
import pandas as pd

ROOT = r"Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients\Mountain CU 68531"
CFG = r"Z:\Shared\TCT Files\CECL - CM Files\client_configs\mountain_cu.yaml"
OUT = r"Z:\Shared\TCT Files\CECL - CM Files\Reports\Mountain_CU_ChargeOff_Recovery_AllYears_ByPool.xlsx"

MON = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,'jul':7,'aug':8,
       'sep':9,'oct':10,'nov':11,'dec':12}

cfg = yaml.safe_load(open(CFG, encoding='utf-8'))
POOL_MAP = cfg.get('pool_map', {}) or {}
DEFAULT_POOL = cfg.get('default_pool', 'Ignore')
# case-insensitive text lookup for category-name codes (Format B)
PM_LOWER = {str(k).strip().lower(): v for k, v in POOL_MAP.items()}

def code_to_pool(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None  # blank -> subtotal row
    s = str(v).strip()
    if not s:
        return None
    cands = [s]
    try:
        i = int(float(s)); cands += [str(i), str(i).zfill(4)]
    except (ValueError, TypeError):
        pass
    for c in cands:
        if c in POOL_MAP:
            return POOL_MAP[c]
    if s.lower() in PM_LOWER:
        return PM_LOWER[s.lower()]
    return DEFAULT_POOL

# header alias -> canonical field
ALIAS = {
    'charge off date': 'date', 'date': 'date', 'co date': 'date',
    'loan suffix': 'suffix', 'suffix': 'suffix',
    'account': 'account',
    'account number': 'account_number', 'account #': 'account_number',
    'loan type': 'loan_type', 'loan type code': 'loan_type', 'type': 'loan_type',
    'charge off amount': 'charge_off', 'charge off': 'charge_off',
    'chargeoff': 'charge_off', 'charge offs': 'charge_off', 'charge-off': 'charge_off',
    'recovery amount': 'recovery', 'recoveries': 'recovery', 'recovery': 'recovery',
}

def norm(h):
    return re.sub(r'\s+', ' ', str(h)).strip().lower()

def is_co_recov_file(name):
    n = name.lower()
    if not n.endswith(('.xlsx', '.xls')) or n.startswith('~$'):
        return False
    return ('charge off' in n or 'charge-off' in n or 'chargeoff' in n
            or 'charge offs' in n) and ('recover' in n or 'recovery' in n)

def path_year(p):
    m = re.search(r'[\\/](20\d{2})[\\/]', p)
    return m.group(1) if m else ''

def month_from(text):
    m = re.search(r'\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*', text.lower())
    return MON.get(m.group(1)) if m else None

# ---- collect files (exclude archive / parallel-testing duplicates) ----
cand = []
for dp, _dirs, fs in os.walk(ROOT):
    low = dp.lower()
    if 'archiv' in low or 'parallel' in low or 'testing' in low or 'do not use' in low:
        continue
    for f in fs:
        if is_co_recov_file(f):
            cand.append(os.path.join(dp, f))

# Dedup by BASENAME *or* CONTENT-HASH -- the two duplicate patterns need both:
#  * SAME name, DIFFERENT bytes: an "APR26" charge-off file re-exported (with an
#    added totals row) into BOTH "Apr 2026" and "Jun 2026". Same basename ->
#    caught by the name set (a content hash would miss it: the totals row
#    changes the bytes).
#  * DIFFERENT name, SAME bytes: the 2018-19 "Call Report" annual summary copied
#    verbatim into all 12 month folders as "... Monthly.xlsx", "... (1).xlsx",
#    "... (2).xlsx". Distinct basenames -> caught by the content hash.
# Skip a file if EITHER its basename or its data-hash was already kept. Files
# that are genuinely distinct in BOTH name and content (every real monthly file,
# and the per-month "Negative Shares" summaries) are all kept -> no data lost.
import hashlib
seen_name = set()
seen_hash = set()
files = []
dup_count = 0
for full in sorted(cand):
    bn = os.path.basename(full).strip().lower()
    try:
        _df = pd.read_excel(full, sheet_name=0)
        sig = hashlib.md5(
            pd.util.hash_pandas_object(_df.fillna(''), index=False).values.tobytes()
        ).hexdigest()
    except Exception:
        sig = None
    if bn in seen_name or (sig is not None and sig in seen_hash):
        dup_count += 1
        continue
    seen_name.add(bn)
    if sig is not None:
        seen_hash.add(sig)
    files.append(full)

print(f"matched {len(files)} CO/Recovery files ({dup_count} duplicate-name/copy files skipped)")

rows = []
skipped = []
for full in files:
    fn = os.path.basename(full)
    yr = path_year(full)
    # month: prefer the FILENAME (authoritative for the period) over the
    # containing subfolder -- the CU drops May's file into the "Jun 2026"
    # folder, so the folder name would mislabel it as June.
    subfolder = os.path.basename(os.path.dirname(full))
    mo = month_from(fn) or month_from(subfolder)
    period = f"{yr}-{mo:02d}" if (yr and mo) else (yr or '')
    try:
        df = pd.read_excel(full, sheet_name=0)
    except Exception as e:
        skipped.append((fn, str(e))); continue
    if df.empty:
        continue
    cmap = {}
    for c in df.columns:
        key = ALIAS.get(norm(c))
        if key and key not in cmap:
            cmap[key] = c
    if 'charge_off' not in cmap and 'recovery' not in cmap:
        skipped.append((fn, 'no CO/recovery columns')); continue
    for _, r in df.iterrows():
        lt = r[cmap['loan_type']] if 'loan_type' in cmap else None
        co = pd.to_numeric(r[cmap['charge_off']], errors='coerce') if 'charge_off' in cmap else None
        rc = pd.to_numeric(r[cmap['recovery']], errors='coerce') if 'recovery' in cmap else None
        acctnum = r[cmap['account_number']] if 'account_number' in cmap else None
        acct = r[cmap['account']] if 'account' in cmap else None
        suf = r[cmap['suffix']] if 'suffix' in cmap else None
        codate = r[cmap['date']] if 'date' in cmap else None
        co_v = 0.0 if (co is None or pd.isna(co)) else float(co)
        rc_v = 0.0 if (rc is None or pd.isna(rc)) else float(rc)
        if co_v == 0 and rc_v == 0 and (lt is None or (isinstance(lt, float) and pd.isna(lt))):
            continue  # empty row
        pool = code_to_pool(lt)
        # classify: aggregate if GL-level account or a category-name loan type
        acct_str = '' if acctnum is None else str(acctnum).strip()
        lt_str = '' if lt is None else str(lt).strip()
        is_agg = acct_str.upper().startswith('GL') or (lt_str and not any(ch.isdigit() for ch in lt_str))
        if pool is None:
            row_type = 'FILE SUBTOTAL (excluded)'
            pool_disp = 'FILE SUBTOTAL (excluded)'
        elif is_agg:
            row_type = 'AGGREGATE (GL summary - no loan detail)'
            pool_disp = pool
        else:
            row_type = 'DETAIL'
            pool_disp = pool
        rows.append({
            'Year': yr, 'Period': period, 'Source File': fn, 'Row Type': row_type,
            'Charge Off Date': codate, 'Loan Suffix': suf, 'Account': acct,
            'Account Number': acctnum, 'Loan Type': lt, 'Pool': pool_disp,
            'Charge Off Amount': co_v, 'Recovery Amount': rc_v,
        })

allrows = pd.DataFrame(rows)
allrows['CO'] = allrows['Charge Off Amount']
allrows['RC'] = allrows['Recovery Amount']
# report/reserve figure excludes per-file subtotal rows
incl = allrows[allrows['Row Type'] != 'FILE SUBTOTAL (excluded)'].copy()

detail_cols = ['Year','Period','Source File','Row Type','Charge Off Date','Loan Suffix',
               'Account','Account Number','Loan Type','Pool','Charge Off Amount','Recovery Amount']
detail = allrows[detail_cols].sort_values(['Year','Period','Row Type','Pool','Source File']).reset_index(drop=True)

def pivot(val):
    p = incl.pivot_table(index='Pool', columns='Year', values=val, aggfunc='sum',
                         fill_value=0.0, margins=True, margins_name='Total')
    return p.round(2)

co_pool = pivot('CO'); rc_pool = pivot('RC')
by_year = incl.groupby('Year').agg(Charge_Off=('CO','sum'), Recovery=('RC','sum'),
                                   Rows=('CO','size')).round(2).reset_index()

with pd.ExcelWriter(OUT, engine='openpyxl') as xw:
    detail.to_excel(xw, sheet_name='CO_Recovery (individual)', index=False)
    co_pool.to_excel(xw, sheet_name='Charge Offs by Pool-Year')
    rc_pool.to_excel(xw, sheet_name='Recoveries by Pool-Year')
    by_year.to_excel(xw, sheet_name='Totals by Year', index=False)

print('WROTE', OUT)
print('rows total:', len(detail), '| DETAIL:', int((allrows["Row Type"]=="DETAIL").sum()),
      '| AGGREGATE:', int((allrows["Row Type"].str.startswith("AGGREGATE")).sum()),
      '| SUBTOTAL excl:', int((allrows["Row Type"].str.startswith("FILE SUBTOTAL")).sum()))
print('\n=== Totals by Year (file-based, excl subtotals) ===')
print(by_year.to_string(index=False))
if skipped:
    print('\nskipped files:', len(skipped))
    for s in skipped[:10]: print('  ', s)
