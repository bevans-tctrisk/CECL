"""Debug Phase 9.35f loan_balances per pool."""
from __future__ import annotations

import os
import sys

os.environ.setdefault("CECL_WORKSPACE_ROOT",
                      r"Z:\Shared\TCT Files\CECL - CM Files")
sys.path.insert(0, r"c:\Dev\CECL")

from cecl_ui.app import create_app  # noqa: E402
from cecl_ui.services import balance_check, config_service  # noqa: E402


def main() -> int:
    short = sys.argv[1] if len(sys.argv) > 1 else "shuford_fcu"
    app = create_app()
    with app.app_context():
        cfg = config_service.load_client_config(
            os.environ["CECL_WORKSPACE_ROOT"], short)
        print(f"file_pattern: {cfg.get('file_pattern')}")
        print(f"default_pool: {cfg.get('default_pool')}")
        print(f"pool_map sample: {list((cfg.get('pool_map') or {}).items())[:5]}")
        st = balance_check._build_state_for_run(cfg, short)
        files = st["sample_uploads"]["loan_data_files"]
        print(f"files matched: {len(files)}")
        for f in files:
            cm = f.get("column_mappings") or {}
            print(f"  {f['name']!r}  cm={list(cm.items())[:4]}")
        lb = balance_check.loan_balances_by_pool(st)
        print(f"\nloan_balances ok={lb.get('ok')}")
        bp = lb.get("by_pool_code") or {}
        print(f"by_pool keys: {list(bp.keys())}")
        for pool, codes in bp.items():
            print(f"  pool={pool!r}: {len(codes)} code(s)")
            for c, b in codes[:5]:
                print(f"      code={c!r:>20}  bal=${b:,.0f}")
        print("\nfiles summary:")
        for f in (lb.get("files") or []):
            print(f"  {f.get('name')} err={f.get('error')} rows={f.get('rows')}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
