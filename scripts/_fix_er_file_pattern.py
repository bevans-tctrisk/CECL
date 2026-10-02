"""Hand-patch ER YAML: regenerate file_pattern from current March file,
add optional trailing .NN suffix so older baseline (DEC 2025.01.xlsx) also matches.

Backs up original to .bak.phase931.
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import re
import shutil
import yaml
from cecl_ui.services import sample_parser as sp

CFG = Path(r"Z:\Shared\TCT Files\CECL - CM Files\client_configs\emergency_responders_cu.yaml")
BAK = CFG.with_suffix(".yaml.bak.phase931")

# 1) Generate clean pattern from a current sample file.
guesses = sp.guess_filename_patterns("airesln v2 March 2026.xlsx")
base = guesses["file_pattern"]
print(f"Generator output:\n  {base}")

# 2) Extend with optional ".NN" suffix between year and extension
#    to also match Vizo IDLR-style baseline files like
#    "airesln V2 DEC 2025.01.xlsx". Inject (?:\.\d{1,3})? right before
#    the closing \.(xlsx|xls)$ group.
new_pattern = re.sub(
    r"(\\\.\(xlsx\|xls\)\$)$",
    r"(?:\\.\\d{1,3})?\1",
    base,
)
print(f"\nPatched pattern (optional .NN suffix):\n  {new_pattern}")

# 3) Quick sanity check against actual files.
rx = re.compile(new_pattern)
folder = Path(r"Z:\Shared\TCT Files\CECL - CM Files\Raw_Uploads\emergency_responders_cu")
print(f"\nMatch test against {folder}:")
hits = 0
for p in sorted(folder.iterdir()):
    if not p.is_file():
        continue
    m = bool(rx.match(p.name))
    if m:
        hits += 1
    print(f"  {'MATCH' if m else 'no   '} {p.name}")
print(f"\nTotal matches: {hits}")

if hits == 0:
    print("\nABORT — pattern matches zero files; not patching YAML.")
    raise SystemExit(1)

# 4) Load YAML, back up, patch both top-level and per-extract patterns,
#    save.
if not BAK.exists():
    shutil.copy2(CFG, BAK)
    print(f"\nBackup written: {BAK}")
else:
    print(f"\nBackup already exists at {BAK} (not overwritten).")

cfg = yaml.safe_load(CFG.read_text())
old_top = cfg.get("file_pattern", "")
cfg["file_pattern"] = new_pattern
print(f"\nTop-level file_pattern:\n  old: {old_top}\n  new: {new_pattern}")

for i, ex in enumerate(cfg.get("loan_data_extracts") or []):
    old = ex.get("file_pattern", "")
    ex["file_pattern"] = new_pattern
    print(f"\nExtract[{i}] file_pattern:\n  old: {old}\n  new: {new_pattern}")

CFG.write_text(yaml.safe_dump(cfg, sort_keys=False, allow_unicode=True))
print("\nYAML saved.")
