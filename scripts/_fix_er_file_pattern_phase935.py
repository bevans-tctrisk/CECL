"""Phase 9.35b: regenerate Emergency Responders CU file_pattern.

The ER YAML's top-level + per-extract ``file_pattern`` was stuck in a
broken form (Phase 9.28 malformed nested char class ``[\\s[\\s_\\-]+\\.]+``
plus a stale literal ``\\.01\\.`` baseline-suffix). Matched ZERO files at
runtime, so:

  * ``import_data.process_client`` only imported by accident (older runs
    when the pattern still worked); current import would be a no-op.
  * ``generate_impdet_report._load_extract_enrichment`` couldn't resolve
    any extract file in the data folder / Archive, so the Improved/
    Deteriorated Loans report's enrichment columns (Loan Suffix, Loan
    Type, Open Date, Interest Rate, Days Delinquent, Original Loan
    Amount, Total Available Credit) were ALL blank — and downstream
    formula columns that reference those cells were blank too.

Fix: re-derive a clean pattern via ``sample_parser.guess_filename_patterns``
seeded on the OLDEST staged baseline (``airesln V2 DEC 2025.01.xlsx``) so
the Phase 9.31 ``(?:\\.\\d{1,3})?`` optional-suffix rule fires. Apply to
both the top-level and per-extract ``file_pattern`` fields. Back up to
``.bak.phase935`` first. Idempotent: re-running is a no-op if the YAML
already has the regenerated pattern.
"""
from __future__ import annotations

import os
import re
import shutil
import sys
from pathlib import Path

import yaml

# Bootstrap so we can import the wizard service module.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from cecl_ui.services import sample_parser  # noqa: E402

YAML_PATH = Path(
    r"Z:\Shared\TCT Files\CECL - CM Files\client_configs\emergency_responders_cu.yaml"
)
BACKUP_PATH = YAML_PATH.with_suffix(YAML_PATH.suffix + ".bak.phase935")
SEED_FILENAME = "airesln V2 DEC 2025.01.xlsx"


def main() -> int:
    if not YAML_PATH.exists():
        print(f"ER YAML not found: {YAML_PATH}")
        return 1

    cfg = yaml.safe_load(YAML_PATH.read_text())
    old_top = cfg.get("file_pattern", "")
    extracts = cfg.get("loan_data_extracts") or []
    old_ext = (extracts[0] or {}).get("file_pattern", "") if extracts else ""

    new_pat = sample_parser.guess_filename_patterns(SEED_FILENAME)["file_pattern"]
    print(f"Regenerated file_pattern: {new_pat}")

    rx = re.compile(new_pat)
    staged = [
        "airesln V2 DEC 2025.01.xlsx",
        "airesln v2 jan 2026.xlsx",
        "airesln v2 March 2026.xlsx",
    ]
    misses = [n for n in staged if not rx.match(n)]
    if misses:
        print(f"ERROR: regenerated pattern does not match: {misses}")
        return 2

    if old_top == new_pat and old_ext == new_pat:
        print("YAML already has the regenerated pattern (idempotent no-op).")
        return 0

    if not BACKUP_PATH.exists():
        shutil.copy2(YAML_PATH, BACKUP_PATH)
        print(f"Backed up YAML -> {BACKUP_PATH.name}")
    else:
        print(f"Backup already exists: {BACKUP_PATH.name}")

    cfg["file_pattern"] = new_pat
    if extracts:
        extracts[0]["file_pattern"] = new_pat
        cfg["loan_data_extracts"] = extracts

    YAML_PATH.write_text(yaml.safe_dump(cfg, sort_keys=False))
    print(f"Wrote updated YAML to {YAML_PATH}")
    print("OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
