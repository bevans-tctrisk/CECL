"""Phase 9.38 — fix First Area FCU file_pattern (both top-level + per-extract).

Symptoms: March 2026 report fails with "No loan data for 2026-03-31" because only
the Dec 2025 baseline imported. Top-level pattern hardcoded literal "AiresV3" (doesn't
match any current file), per-extract pattern hardcoded literal "2025 December Aires
file V" (only matches December baseline).

Filenames observed in Raw_Uploads/first_area_fcu/:
  - 2025 December Aires file V2.xlsx (baseline, archived)
  - January 2026 Aires V2.xlsx
  - February 2026 Aires V2.xlsx
  - March 2026 Aires2.xlsx   (NOTE: no space, no V — direct "Aires2")

Fix: replace both top-level + per-extract with a permissive pattern matching:
  YEAR MONTH AIRES [file] [V]##   OR   MONTH YEAR AIRES [V]##

Idempotent. Backs up YAML to .bak.phase938 once.
"""
from pathlib import Path
import shutil
import re
import yaml

YAML_PATH = Path(r"Z:\Shared\TCT Files\CECL - CM Files\client_configs\first_area_fcu.yaml")

NEW_PATTERN = (
    r"(?i)^"
    r"(?:"
    r"20\d{2}[\s_\-]+"
    r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)"
    r"(?:uary|ruary|ch|il|e|y|ust|tember|ober|ember)?"
    r"|"
    r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)"
    r"(?:uary|ruary|ch|il|e|y|ust|tember|ober|ember)?"
    r"[\s_\-]+20\d{2}"
    r")"
    r"[\s_\-]+aires?(?:[\s_\-]+file)?[\s_\-]*[Vv]?\d{0,3}"
    r"\.(xlsx|xls)$"
)


def main() -> int:
    if not YAML_PATH.exists():
        print(f"YAML not found: {YAML_PATH}")
        return 1

    # Compile-test new pattern against expected filenames
    rx = re.compile(NEW_PATTERN)
    test_cases = [
        ("2025 December Aires file V2.xlsx", True),
        ("2025_December_Aires_file_V2.xlsx", True),
        ("January 2026 Aires V2.xlsx", True),
        ("February 2026 Aires V2.xlsx", True),
        ("March 2026 Aires2.xlsx", True),
        ("April 2026 Aires V3.xlsx", True),
        ("Some other file.xlsx", False),
        ("2026 March Aires V2.xlsx", True),
    ]
    print("Pattern self-test:")
    failed = False
    for name, expected in test_cases:
        got = bool(rx.search(name))
        marker = " OK" if got == expected else "FAIL"
        print(f"  [{marker}] {name!r} expected={expected} got={got}")
        if got != expected:
            failed = True
    if failed:
        print("Pattern self-test failed — aborting before writing YAML.")
        return 2

    # Backup
    bak = YAML_PATH.with_suffix(YAML_PATH.suffix + ".bak.phase938")
    if not bak.exists():
        shutil.copy2(YAML_PATH, bak)
        print(f"Backup written to {bak.name}")
    else:
        print(f"Backup already exists at {bak.name} (idempotent — keeping original)")

    cfg = yaml.safe_load(YAML_PATH.read_text(encoding="utf-8"))
    old_top = cfg.get("file_pattern")
    old_ex = None
    extracts = cfg.get("loan_data_extracts") or []
    if extracts:
        old_ex = extracts[0].get("file_pattern")

    cfg["file_pattern"] = NEW_PATTERN
    if extracts:
        extracts[0]["file_pattern"] = NEW_PATTERN
        cfg["loan_data_extracts"] = extracts

    # Write back preserving YAML structure
    YAML_PATH.write_text(yaml.safe_dump(cfg, sort_keys=False), encoding="utf-8")

    print()
    print("Wrote updated YAML:")
    print(f"  TOP (old): {old_top}")
    print(f"  TOP (new): {NEW_PATTERN}")
    print(f"  EX  (old): {old_ex}")
    print(f"  EX  (new): {NEW_PATTERN}")

    # Verify Raw_Uploads files now match
    raw_dir = Path(r"Z:\Shared\TCT Files\CECL - CM Files\Raw_Uploads\first_area_fcu")
    aires_files = [p for p in raw_dir.iterdir() if p.is_file() and "aires" in p.name.lower()]
    print()
    print(f"Raw_Uploads aires files ({len(aires_files)}):")
    for p in sorted(aires_files):
        print(f"  match={bool(rx.search(p.name))}  {p.name}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
