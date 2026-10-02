import os, glob, yaml
os.environ['CECL_WORKSPACE_ROOT'] = r'Z:\Shared\TCT Files\CECL - CM Files'
import pandas as pd
from collections import defaultdict, Counter
ARCH = r'Z:\Shared\TCT Files\CECL - CM Files\Archive\credit_union_of_richmond'
cfg = yaml.safe_load(open(r'Z:\Shared\TCT Files\CECL - CM Files\client_configs\credit_union_of_richmond.yaml', encoding='utf-8'))
pm = {str(k): v for k, v in cfg['pool_map'].items()}

def old_pool(code):
    s = str(code).strip()
    for k in (s, s.upper(), s.lower()):
        if k in pm: return pm[k]
    try: return pm.get(str(int(float(s))))
    except: return None

def find(name):
    h = glob.glob(os.path.join(ARCH, '**', name), recursive=True)
    return h[0] if h else None

by_member = defaultdict(list)
for fn in ['AIRES 12-31-2025.xlsx', 'AIRES 01-31-2026.xlsx']:
    p = find(fn)
    if not p: continue
    df = pd.read_excel(p, header=0)
    acol = [c for c in df.columns if str(c).strip().lower() == 'account number'][0]
    for _, r in df.iterrows():
        acct = str(r[acol]).strip()
        member = acct.split('L')[0].split('l')[0]
        pool = old_pool(r.get('LOAN TYPE'))
        bal = pd.to_numeric(r.get('LOAN BALANCE'), errors='coerce')
        if member and pool and pd.notna(bal):
            by_member[member].append((float(bal), pool))

tally = defaultdict(Counter)
matched = unmatched = 0
for fn in ['AIRESV2 02-28-2026.xlsx', 'AIRESV2 03-31-2026.xlsx']:
    p = find(fn)
    if not p: continue
    df = pd.read_excel(p, header=0)
    for _, r in df.iterrows():
        acct = str(r['Account Number']).strip().split('.')[0]
        member = acct[:-3] if len(acct) > 3 else acct
        code = r['Loan Type Code']
        try: code = str(int(float(code)))
        except: code = str(code).strip()
        bal = pd.to_numeric(r['Current Balance'], errors='coerce')
        cands = by_member.get(member)
        if not cands or pd.isna(bal):
            unmatched += 1; continue
        best = min(cands, key=lambda t: abs(t[0] - float(bal)))
        if abs(best[0] - float(bal)) <= max(3000.0, 0.25 * abs(float(bal)) + 1):
            tally[code][best[1]] += 1; matched += 1
        else:
            unmatched += 1
print('matched=%d unmatched=%d' % (matched, unmatched))
print('%-6s %6s %-28s %s' % ('code', 'n', 'pool', 'purity'))
v2map = {}
for code in sorted(tally, key=lambda c: -sum(tally[c].values())):
    cnt = tally[code]; pool, n = cnt.most_common(1)[0]; tot = sum(cnt.values())
    v2map[code] = pool
    print('%-6s %6d %-28s %3.0f%%  %s' % (code, tot, pool, 100*n/tot, dict(cnt)))
print('\nV2MAP =', v2map)
