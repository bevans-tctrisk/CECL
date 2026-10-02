"""Quick regression: ensure DQ overlay works for Mountain CU too."""
import sys, os
sys.path.insert(0, '.')
os.environ.setdefault('CECL_WORKSPACE_ROOT', r'Z:\Shared\TCT Files\CECL - CM Files')

import yaml
from pathlib import Path
from generate_report import _load_dq_history_from_db

for short in ('mountain_cu', 'tcp_cu'):
    p = Path(rf'Z:\Shared\TCT Files\CECL - CM Files\client_configs\{short}.yaml')
    if not p.exists():
        print(f"{short}: YAML missing"); continue
    cfg = yaml.safe_load(p.read_text(encoding='utf-8'))
    cfg.setdefault('short_name', short)
    try:
        dq = _load_dq_history_from_db(cfg)
    except Exception as e:
        print(f"{short}: ERROR {e}"); continue
    if not dq:
        print(f"{short}: empty"); continue
    yrs = sorted(dq.keys())
    print(f"{short}: {len(yrs)} years, pools per year:")
    for y in yrs[-3:]:
        rows = [(p, v) for p, v in dq[y].items()]
        print(f"  {y}: " + ", ".join(f"{p}={v*100:.2f}%" for p, v in rows[:6]))
