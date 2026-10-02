"""Build Lanco FCU monthly loan balances from the GL Balance Sheets.

Maps each loan GL account to a pool, splits the mixed 702100 "MBL Mortgage"
account between Real Estate (Commercial MBL Mortgage codes 100/101) and
Business Loans (Commercial Mtg codes 110-115) by the Aires loan-level
proportion, and nets the participation-sold contra accounts against their
pool. Writes a wide monthly-balance file (pools x 36 month-ends) and wires
``monthly_balance`` into the config.

NOTE (analyst review): the 702100 RE/Business split ratio and the
GL_ACCOUNT_POOL map are the two judgment calls — adjust here if Lanco's
methodology differs.
"""
import os, calendar
import yaml, openpyxl, pandas as pd

os.environ.setdefault('CECL_WORKSPACE_ROOT', r'Z:\Shared\TCT Files\CECL - CM Files')
WS = os.environ['CECL_WORKSPACE_ROOT']
SHORT = 'lanco_fcu'
CLIENT = (r'Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients'
          r'\Lanco FCU 16657\Portfolio Management\CECL Migration IDLR')
cfg_path = os.path.join(WS, 'client_configs', SHORT + '.yaml')
cfg = yaml.safe_load(open(cfg_path, encoding='utf-8'))

# ── 1) 702100 RE/Business split ratio from Aires commercial-mtg codes ──
aires = pd.read_excel(os.path.join(CLIENT, '2025', 'Dec 2025',
                                   'Aires Loan Data 12-31-2025 WO Charge-offs.xlsx'), header=0)
aires.columns = [str(c).strip() for c in aires.columns]
bal = pd.to_numeric(aires['CURRENT BAL'], errors='coerce').fillna(0.0)
code = aires['Loan Type Code'].astype(str).str.strip()
RE_CM = {'100', '101'}                                    # Commercial MBL Mortgage -> Real Estate
BUS_CM = {'110', '111', '112', '113', '114', '115'}       # Commercial Mtg -> Business
re_cm = bal[code.isin(RE_CM)].sum()
bus_cm = bal[code.isin(BUS_CM)].sum()
re_frac = re_cm / (re_cm + bus_cm) if (re_cm + bus_cm) else 0.6
print(f"702100 split from Aires: RE commercial (100/101)=${re_cm:,.0f}  "
      f"Business commercial (110-115)=${bus_cm:,.0f}  -> RE fraction={re_frac:.3f}")

# ── 2) GL account -> pool map (participation contras net into their pool) ──
GL_POOL = {
    701000: 'Consumer Unsecured', 701001: 'Consumer Secured',
    701100: 'Real Estate', 701110: 'Real Estate', 701120: 'Real Estate',
    701130: 'Real Estate', 701140: 'Real Estate', 701150: 'Real Estate',
    701200: 'Consumer Secured', 701300: 'Consumer Secured', 701400: 'Consumer Secured',
    701450: 'Consumer Indirect', 701500: 'Consumer Indirect',
    701600: 'Consumer Unsecured', 701700: 'Consumer Unsecured',
    701900: 'Consumer Unsecured', 701950: 'Consumer Unsecured',
    702000: 'Business Loans', 702025: 'Business Loans', 702050: 'Business Loans',
    # 702100 handled specially (split RE/Business).
    702110: 'Real Estate',       # RE Loan Part Sold (contra, negative)
    702111: 'Business Loans', 702112: 'Business Loans', 702114: 'Business Loans',
    702119: 'Business Loans', 702120: 'Business Loans', 702130: 'Business Loans',
    702140: 'Business Loans',
    703000: 'Business Loans', 703141: 'Consumer Unsecured', 703143: 'Real Estate',
    703144: 'Consumer Secured', 703145: 'Consumer Unsecured',
    703146: 'Consumer Secured', 703148: 'Consumer Secured',
    704101: 'Credit Cards', 704102: 'Credit Cards',
    704110: 'Student Loans', 704120: 'Student Loans',
    704140: 'Real Estate', 704150: 'Real Estate', 704160: 'Real Estate', 704170: 'Real Estate',
}
POOLS = ['Real Estate', 'Consumer Secured', 'Consumer Indirect',
         'Consumer Unsecured', 'Credit Cards', 'Student Loans', 'Business Loans']

def num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(',', '').replace('$', '')
    neg = (s.startswith('<') and s.endswith('>')) or (s.startswith('(') and s.endswith(')'))
    s = s.strip('<>()')
    try:
        return -float(s) if neg else float(s)
    except ValueError:
        return 0.0

def sheet_to_date(sn):
    m, d, y = sn.split('.')
    y = 2000 + int(y)
    return '%04d-%02d-%02d' % (y, int(m), calendar.monthrange(y, int(m))[1])

# ── 3) process each GL sheet ──
wb = openpyxl.load_workbook(os.path.join(CLIENT, 'Balance Sheets - Use This One Lano FCU.xlsx'),
                            data_only=True, read_only=True)
data = {}  # date -> {pool: bal}
for sn in wb.sheetnames:
    try:
        d = sheet_to_date(sn)
    except Exception:
        continue
    ws = wb[sn]
    by = {p: 0.0 for p in POOLS}
    for row in ws.iter_rows(values_only=True):
        try:
            a = int(str(row[0]).strip())
        except (ValueError, TypeError):
            continue
        v = num(row[2])
        if v == 0:
            continue
        if a == 702100:  # mixed commercial mortgage -> split
            by['Real Estate'] += v * re_frac
            by['Business Loans'] += v * (1 - re_frac)
        elif a in GL_POOL:
            by[GL_POOL[a]] += v
    data[d] = by

dates = sorted(data)
rows = []
for p in POOLS:
    row = {'Pool': p}
    for d in dates:
        row[d] = round(data[d].get(p, 0.0), 2)
    rows.append(row)
wide = pd.DataFrame(rows, columns=['Pool'] + dates)
out_dir = os.path.join(WS, 'Raw_Uploads', SHORT)
os.makedirs(out_dir, exist_ok=True)
out = os.path.join(out_dir, 'Monthly Loan Balances by Type.xlsx')
with pd.ExcelWriter(out, engine='openpyxl') as xw:
    wide.to_excel(xw, sheet_name='Sheet1', index=False)
print('WROTE', out, '| months:', len(dates), dates[0], '->', dates[-1])
dec = data.get('2025-12-31', {})
print('Dec-2025 GL-net total: %.2f (target $86,744,194.24)' % sum(dec.values()))
for p in POOLS:
    print('   %-20s %14.2f' % (p, dec.get(p, 0.0)))

cfg['monthly_balance'] = {'source': 'single', 'sheet': 'Sheet1', 'header_row': 1,
                          'pool_name_col': 'A', 'first_date_col': 'B',
                          'filename': 'Monthly Loan Balances by Type.xlsx', 'saved_path': out,
                          'pool_map': {p: p for p in POOLS}}
with open(cfg_path, 'w', encoding='utf-8') as f:
    yaml.safe_dump(cfg, f, sort_keys=False, allow_unicode=True)
print('config monthly_balance wired')
