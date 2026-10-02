"""Phase — SCI FCU credit-score-no-movement fix.

Root cause: the wizard saved `member_account: {mode: delimiter, delimiter: '-'}`
for the Aires Monthly Loans extract, but the actual `Account Number` column
carries contiguous IDs like `2340L24` (member `2340` + suffix `L24`) with NO
delimiter. `derive_member_account` split-on-dash returns the whole string;
`pd.to_numeric` coerces to NaN -> 0, so every credit-pull join key is 0 and
zero of 628 credit-pull scores land. Fallback `current_fico = original_fico`
fires for every loan.

Fix: switch `member_account.mode` to `fixed_suffix` (suffix_length=3) so
the last 3 chars are treated as the suffix. Both extract-level and top-level
copies are updated for consistency. Idempotent; backs YAML up to
`.bak.phase_sci_credit_pull`.
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

import yaml

YAML_PATH = Path(r"Z:\Shared\TCT Files\CECL - CM Files\client_configs\sci_fcu.yaml")
BACKUP_PATH = YAML_PATH.with_suffix(YAML_PATH.suffix + ".bak.phase_sci_credit_pull")

TARGET_EXTRACT_MARKER = "Aires"  # match by file_pattern substring


def _needs_fix(ma: dict | None) -> bool:
    if not ma:
        return False
    if str(ma.get("mode") or "").strip().lower() != "delimiter":
        return False
    return True


def main() -> int:
    if not YAML_PATH.exists():
        print(f"YAML not found: {YAML_PATH}")
        return 2

    if not BACKUP_PATH.exists():
        shutil.copy2(YAML_PATH, BACKUP_PATH)
        print(f"Backup written: {BACKUP_PATH}")
    else:
        print(f"Backup already exists (idempotent): {BACKUP_PATH}")

    cfg = yaml.safe_load(YAML_PATH.read_text(encoding="utf-8"))
    changed = False

    # Extract-level fix (Aires loan extract only)
    for ex in cfg.get("loan_data_extracts") or []:
        fp = str(ex.get("file_pattern") or "")
        if TARGET_EXTRACT_MARKER.lower() not in fp.lower():
            continue
        ma = ex.get("member_account") or {}
        if _needs_fix(ma):
            print(f"Extract '{fp[:60]}...': {ma} -> mode=fixed_suffix")
            ma["mode"] = "fixed_suffix"
            # Keep suffix_length=3 (matches actual `L24`/`L21`/etc suffixes)
            if int(ma.get("suffix_length") or 0) <= 0:
                ma["suffix_length"] = 3
            ex["member_account"] = ma
            changed = True
        else:
            print(f"Extract '{fp[:60]}...': already {ma}")

    # Top-level fix (only if it currently mirrors the bad delimiter/3 default
    # AND does NOT conflict with the Credit Card extract which uses
    # fixed_suffix/0). Change top-level to the Aires shape so any code that
    # falls back to the top-level config gets the correct behaviour on the
    # main extract.
    top_ma = cfg.get("member_account") or {}
    if _needs_fix(top_ma):
        print(f"Top-level member_account: {top_ma} -> mode=fixed_suffix")
        top_ma["mode"] = "fixed_suffix"
        if int(top_ma.get("suffix_length") or 0) <= 0:
            top_ma["suffix_length"] = 3
        cfg["member_account"] = top_ma
        changed = True
    else:
        print(f"Top-level member_account already {top_ma}")

    if not changed:
        print("Nothing to change.")
        return 0

    YAML_PATH.write_text(
        yaml.safe_dump(cfg, sort_keys=False, allow_unicode=True, width=1000),
        encoding="utf-8",
    )
    print(f"YAML updated: {YAML_PATH}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
