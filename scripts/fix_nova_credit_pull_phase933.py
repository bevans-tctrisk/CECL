"""Phase 9.33 — NOVA CU credit-pull recovery.

Restores archived NOVA loan extracts to Raw_Uploads, generalizes the
nova_cu.yaml file_pattern so all 4 IDLR job IDs / dates / Loan(s?)
variants match, and updates report_period to 2026-03 so snapshot
fallbacks anchor correctly.

Backs the YAML to .bak.phase933 and the archived loan extracts get
COPIED (not moved) — the originals stay in Archive/ as a safety net.
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

import yaml

WS = Path(r"Z:\Shared\TCT Files\CECL - CM Files")
YAML_PATH = WS / "client_configs" / "nova_cu.yaml"
ARCHIVE_DIR = WS / "Archive" / "nova_cu"
RAW_UPLOAD_DIR = WS / "Raw_Uploads" / "nova_cu"

# Wildcarded to cover all 4 known IDLR prefixes (784404 / 805588 /
# 824838 / 846394), MMDDYYYY date, Loan/Loans, V2/v2 case variants,
# and space-or-underscore separators (browser uploads underscore).
NEW_TOP_PATTERN = (
    r"(?i)^\d+[\s_]+\-[\s_]+Nova[\s_]+Aires[\s_]+Loans?[\s_]+"
    r"\d{8}[\s_]+[Vv]\d{1,2}\.(xlsx|xls)$"
)
# Per-extract pattern is the same body, no anchors — extracts are
# matched via re.search semantics in _patterns_match_any.
NEW_EXTRACT_PATTERN = (
    r"(?i)\d+[\s_]+\-[\s_]+Nova[\s_]+Aires[\s_]+Loans?[\s_]+"
    r"\d{8}[\s_]+[Vv]\d{1,2}\.(xlsx|xls)$"
)

# Sanity-check filenames the new pattern MUST match.
SANITY_FILES = [
    "784404_-_Nova_Aires_Loans_12312025_V2.xlsx",
    "805588 - Nova Aires Loans 01312026 V2.xlsx",
    "824838 - Nova Aires Loan 02282026 v2.xlsx",
    "846394 - Nova Aires Loan 03312026 v2.xlsx",
]


def main() -> int:
    print("=" * 64)
    print("Phase 9.33 NOVA CU credit-pull recovery")
    print("=" * 64)

    # ---- 1. Sanity test the new pattern -------------------------------
    rx = re.compile(NEW_TOP_PATTERN)
    for name in SANITY_FILES:
        ok = bool(rx.match(name))
        print(f"  pattern test: {name:55s} -> {ok}")
        if not ok:
            print("  ABORT: new pattern fails sanity check.")
            return 1

    # ---- 2. Back up + rewrite YAML ------------------------------------
    if not YAML_PATH.exists():
        print(f"  ABORT: YAML not found: {YAML_PATH}")
        return 1
    backup = YAML_PATH.with_suffix(YAML_PATH.suffix + ".bak.phase933")
    if not backup.exists():
        shutil.copy2(YAML_PATH, backup)
        print(f"  YAML backed up: {backup.name}")
    else:
        print(f"  YAML backup already present: {backup.name}")

    cfg = yaml.safe_load(YAML_PATH.read_text(encoding="utf-8"))
    old_top = cfg.get("file_pattern")
    cfg["file_pattern"] = NEW_TOP_PATTERN
    print(f"  top-level file_pattern: {old_top!r}")
    print(f"                      -> {NEW_TOP_PATTERN!r}")

    extracts = cfg.get("loan_data_extracts") or []
    if extracts:
        old_ex = (extracts[0] or {}).get("file_pattern")
        extracts[0]["file_pattern"] = NEW_EXTRACT_PATTERN
        print(f"  extract[0].file_pattern: {old_ex!r}")
        print(f"                       -> {NEW_EXTRACT_PATTERN!r}")
    else:
        print("  (no loan_data_extracts entries — top-level used)")

    old_period = cfg.get("report_period")
    cfg["report_period"] = "2026-03"
    print(f"  report_period: {old_period!r} -> '2026-03'")

    YAML_PATH.write_text(
        yaml.safe_dump(cfg, sort_keys=False, default_flow_style=False),
        encoding="utf-8",
    )
    print("  YAML written.")

    # ---- 3. Restore archived loan extracts to Raw_Uploads -------------
    if not ARCHIVE_DIR.is_dir():
        print(f"  ABORT: Archive dir not found: {ARCHIVE_DIR}")
        return 1
    RAW_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

    restored = 0
    skipped = 0
    for src in sorted(ARCHIVE_DIR.iterdir()):
        if not src.is_file():
            continue
        if not rx.match(src.name):
            continue
        dst = RAW_UPLOAD_DIR / src.name
        if dst.exists():
            print(f"  skip (already in Raw_Uploads): {src.name}")
            skipped += 1
            continue
        shutil.copy2(src, dst)
        print(f"  restored: {src.name}")
        restored += 1
    print(f"  restored={restored}  skipped={skipped}")

    print("=" * 64)
    print("Done. Next: run import + reports.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
