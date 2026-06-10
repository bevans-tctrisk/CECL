"""Set up Test Nucor Emp CU as a clone of Nucor Emp CU.

Mirrors the Test Nova rebuild approach:
1. Clone client_configs/nucor_emp_cu.yaml -> test_nucor_emp_cu.yaml with name + path substitutions
2. Mirror Raw_Uploads/nucor_emp_cu/ -> Raw_Uploads/test_nucor_emp_cu/
3. Clone wizard_drafts/nucor_emp_cu__migration.json -> test_nucor_emp_cu__migration.json with name substitutions
"""
from __future__ import annotations
import json
import os
import shutil
from pathlib import Path

WS = Path(r'Z:\Shared\TCT Files\CECL - CM Files')

SRC_SHORT = 'nucor_emp_cu'
DST_SHORT = 'test_nucor_emp_cu'
SRC_NAME = 'Nucor Emp CU'
DST_NAME = 'Test Nucor Emp CU'

# 1. Clone YAML
src_yaml = WS / 'client_configs' / f'{SRC_SHORT}.yaml'
dst_yaml = WS / 'client_configs' / f'{DST_SHORT}.yaml'
text = src_yaml.read_text(encoding='utf-8')
# Substitutions - careful: only swap the credit_union name and data_directory short_name
text = text.replace(f'credit_union: {SRC_NAME}', f'credit_union: {DST_NAME}')
text = text.replace(f'/Raw_Uploads/{SRC_SHORT}', f'/Raw_Uploads/{DST_SHORT}')
text = text.replace(f'\\Raw_Uploads\\{SRC_SHORT}', f'\\Raw_Uploads\\{DST_SHORT}')
dst_yaml.write_text(text, encoding='utf-8')
print(f'Wrote {dst_yaml.name} ({dst_yaml.stat().st_size:,} bytes)')

# 2. Mirror Raw_Uploads folder
src_uploads = WS / 'Raw_Uploads' / SRC_SHORT
dst_uploads = WS / 'Raw_Uploads' / DST_SHORT
if dst_uploads.exists():
    print(f'Removing existing {dst_uploads}')
    shutil.rmtree(dst_uploads)
print(f'Mirroring {src_uploads} -> {dst_uploads}')
shutil.copytree(src_uploads, dst_uploads)
n = sum(1 for _ in dst_uploads.rglob('*') if _.is_file())
print(f'Copied {n} files into {dst_uploads.name}')

# 3. Clone wizard draft with name substitutions
src_draft = WS / 'wizard_drafts' / f'{SRC_SHORT}__migration.json'
dst_draft = WS / 'wizard_drafts' / f'{DST_SHORT}__migration.json'
data = json.loads(src_draft.read_text(encoding='utf-8'))
# Top-level identity
data['credit_union'] = DST_NAME
data['short_name'] = DST_SHORT
# data_directory string
if isinstance(data.get('data_directory'), str):
    data['data_directory'] = data['data_directory'].replace(f'\\Raw_Uploads\\{SRC_SHORT}', f'\\Raw_Uploads\\{DST_SHORT}').replace(f'/Raw_Uploads/{SRC_SHORT}', f'/Raw_Uploads/{DST_SHORT}')
# Keep the temp-cached hist_scan paths intact (they're CU-agnostic file caches)
# But we don't need to touch them - the aggregator reads them by path, doesn't care about CU
# _draft_meta
meta = data.setdefault('_draft_meta', {})
meta['model'] = 'migration'
meta.pop('completed_at', None)  # mark as in-progress draft so it's distinct
dst_draft.write_text(json.dumps(data, indent=2), encoding='utf-8')
print(f'Wrote {dst_draft.name} ({dst_draft.stat().st_size:,} bytes)')

print('\nDone.')
