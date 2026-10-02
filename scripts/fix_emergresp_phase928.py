"""Phase 9.28 — Emergency Responders CU file_pattern repair.

The wizard-generated file_pattern was corrupted by an over-greedy
post-process collapse step in sample_parser._generalize_filename_pattern
(now fixed). The corruption produced a malformed nested character class
``[\\s[\\s_\\-]+\\.]+`` that requires literal ``]`` characters in
filenames — so nothing matched and ``run_import`` returned 0 rows.

This script:
  1. Loads the YAML at Z:\\...\\client_configs\\emergency_responders_cu.yaml
  2. Regenerates ``file_pattern`` for top-level + each loan_data_extract
     using the now-fixed sample_parser.guess_filename_patterns based on
     the actual staged sample filename in Raw_Uploads.
  3. Writes a .bak_phase928 backup, then saves the repaired YAML.
"""
from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path

import yaml

sys.path.insert(0, r"c:\Dev\CECL")

from cecl_ui.services.sample_parser import guess_filename_patterns  # noqa: E402

WS = Path(r"Z:\Shared\TCT Files\CECL - CM Files")
CFG_PATH = WS / "client_configs" / "emergency_responders_cu.yaml"
UPLOAD_DIR = WS / "Raw_Uploads" / "emergency_responders_cu"

# Only Aires loan extracts — pick a representative sample filename.
# All AIRES files in this CU follow ``airesln V<n> <MONTH> <YYYY>.01.xlsx``.
SAMPLE_NAME = "airesln V2 DEC 2025.01.xlsx"


def main() -> int:
    if not CFG_PATH.exists():
        print(f"ERROR: YAML not found at {CFG_PATH}")
        return 1

    raw = CFG_PATH.read_text(encoding="utf-8")
    cfg = yaml.safe_load(raw)

    new_pat = guess_filename_patterns(SAMPLE_NAME)["file_pattern"]
    print(f"Re-derived file_pattern from {SAMPLE_NAME!r}:")
    print(f"  {new_pat}")

    rx = re.compile(new_pat)
    matches = []
    if UPLOAD_DIR.exists():
        for f in UPLOAD_DIR.iterdir():
            if f.is_file() and f.suffix.lower() in {".xls", ".xlsx", ".xlsm"}:
                if rx.match(f.name):
                    matches.append(f.name)
    print(f"  Matches {len(matches)} file(s) in Raw_Uploads:")
    for m in matches:
        print(f"    {m}")
    if not matches:
        print("  WARNING: pattern matches zero files. Aborting (no change).")
        return 2

    old_top = cfg.get("file_pattern", "")
    cfg["file_pattern"] = new_pat
    print(f"\nTop-level file_pattern:")
    print(f"  OLD: {old_top}")
    print(f"  NEW: {new_pat}")

    extracts = cfg.get("loan_data_extracts") or []
    for i, ex in enumerate(extracts):
        if not isinstance(ex, dict):
            continue
        old_ex = ex.get("file_pattern", "")
        ex["file_pattern"] = new_pat
        print(f"\nloan_data_extracts[{i}].file_pattern:")
        print(f"  OLD: {old_ex}")
        print(f"  NEW: {new_pat}")

    bak = CFG_PATH.with_suffix(".yaml.bak_phase928")
    shutil.copy2(CFG_PATH, bak)
    print(f"\nBackup written: {bak}")

    CFG_PATH.write_text(
        yaml.safe_dump(cfg, sort_keys=False, allow_unicode=True),
        encoding="utf-8",
    )
    print(f"Saved: {CFG_PATH}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
