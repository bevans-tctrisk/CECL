"""Verify load_standalone_impaired now finds TCP's 6-2026 file."""
import os
os.environ["CECL_WORKSPACE_ROOT"] = r"Z:\Shared\TCT Files\CECL - CM Files"
import sys
sys.path.insert(0, r"c:\Dev\CECL")

import yaml
from pathlib import Path

ws = Path(r"Z:\Shared\TCT Files\CECL - CM Files")
cfg_path = ws / "client_configs" / "tcp_cu.yaml"
cfg = yaml.safe_load(cfg_path.read_text(encoding="utf-8"))
print(f"data_directory: {cfg.get('data_directory')}")
print(f"credit_union: {cfg.get('credit_union')}")

import generate_report

# Simulate the run for 2026-06 without needing the full df.
result = generate_report.load_standalone_impaired(cfg, "2026-06-30", df=None)
print(f"\nResult keys: {list(result.keys())}")
if result:
    print(f"  acl_impaired: {result.get('acl_impaired')}")
    print(f"  total_spec_id: {result.get('total_spec_id')}")
    print(f"  spec_id_by_pool pool count: {len(result.get('spec_id_by_pool') or {})}")
else:
    print("  (empty dict - file not found or unreadable)")
