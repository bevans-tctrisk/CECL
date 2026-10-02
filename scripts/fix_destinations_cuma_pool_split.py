"""Phase 9.22 live remediation for Destinations CU.

The Destinations CU wizard draft was created before the per-extract
``pool_code_split`` override existed.  Their CUMA loan-data extract is
silently being routed through the CU-level ``pool_code_split: '/'``,
which truncates the mortgage codes ``15/15 ARM``, ``5/1 ARM``,
``10/1 ARM`` down to ``15``, ``5``, ``10``.  Those numeric stubs then
fall into ``default_pool: Ignore``, and the mortgage balances vanish
from the report.

This script:

1. Loads the live Destinations CU draft.
2. Finds the CUMA loan-data extract under
   ``state["sample_uploads"]["loan_data_files"]``.
3. Sets ``entry["pool_code_split"] = ""`` so the new per-extract
   override forces a no-split read of that file only.
4. Re-runs ``sample_parser.extract_pool_codes`` with ``split_char=""``
   to get the un-truncated codes.
5. Removes the polluted ``'15' -> ''`` (or ``'15' -> 'Ignore'``) entry
   from ``state["pool_map"]`` (it was the truncated code that should
   never have been there).
6. Adds the newly-discovered CUMA codes to ``state["pool_map"]`` with
   a blank value so the user picks the destination pool on Step 3.
7. Persists the draft via :mod:`cecl_ui.services.wizard_drafts`.
"""

from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from cecl_ui.services import sample_parser  # noqa: E402

WORKSPACE = Path(
    os.environ.get("WORKSPACE_ROOT",
                   r"Z:\Shared\TCT Files\CECL - CM Files")
)
DRAFT = WORKSPACE / "wizard_drafts" / "destinations_cu__migration.json"


def _find_cuma(state: dict) -> dict | None:
    files = ((state.get("sample_uploads") or {}).get("loan_data_files")
             or [])
    for f in files:
        n = (f.get("name") or "").lower()
        if "cuma" in n or "mortgage" in n or "mtg" in n:
            return f
    return None


def main() -> int:
    if not DRAFT.exists():
        print(f"!! Draft not found at {DRAFT}")
        return 1

    backup = DRAFT.with_suffix(".json.bak_phase922")
    if not backup.exists():
        shutil.copy2(DRAFT, backup)
        print(f"== Backup written: {backup}")
    else:
        print(f"== Backup already present: {backup}")

    state = json.loads(DRAFT.read_text(encoding="utf-8"))

    cuma = _find_cuma(state)
    if not cuma:
        print("!! No CUMA-like extract in draft. Nothing to do.")
        return 1

    print()
    print("== Before:")
    print(f"   CUMA entry name        = {cuma.get('name')}")
    print(f"   CUMA pool_code_split   = "
          f"{cuma.get('pool_code_split', '<inherit (top-level / )>')!r}")
    print(f"   top-level pool_code_split = "
          f"{state.get('pool_code_split')!r}")
    print(f"   pool_map size          = "
          f"{len(state.get('pool_map') or {})}")
    print(f"   '15' in pool_map?      = "
          f"{'15' in (state.get('pool_map') or {})}")

    path = cuma.get("path") or cuma.get("saved_path")
    cm = cuma.get("column_mappings") or {}
    pool_col = cm.get("loan_pool_code")
    header_row = cuma.get("header_row")

    if not path or not Path(path).exists():
        print(f"!! CUMA file missing on disk: {path}")
        print("   Cannot re-extract codes. Setting the override anyway "
              "so the next run picks them up correctly.")
        codes_full: list[str] = []
    elif not pool_col:
        print("!! CUMA entry has no loan_pool_code mapping; "
              "cannot re-extract.")
        codes_full = []
    else:
        codes_full = sample_parser.extract_pool_codes(
            path, column_name=pool_col,
            header_row=header_row, split_char="",
        )

    # Apply the per-extract override.
    cuma["pool_code_split"] = ""

    pool_map = state.setdefault("pool_map", {})

    # Drop the truncated '15' entry (and any same-prefix truncations of
    # codes we just recovered) -- they were artefacts of the bug.
    truncated_prefixes = {
        c.split("/", 1)[0].strip()
        for c in codes_full if "/" in c
    }
    dropped = []
    for stale in list(truncated_prefixes):
        if stale and stale in pool_map and stale not in codes_full:
            dropped.append((stale, pool_map.pop(stale)))

    added = []
    for code in codes_full:
        if code and code not in pool_map:
            pool_map[code] = ""
            added.append(code)

    state["pool_map"] = pool_map

    # Persist.
    DRAFT.write_text(
        json.dumps(state, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    print()
    print("== After:")
    print(f"   CUMA pool_code_split   = {cuma.get('pool_code_split')!r}")
    print(f"   Re-extracted codes     = {codes_full}")
    print(f"   Dropped truncated keys = {dropped}")
    print(f"   Added pool_map keys    = {added}")
    print(f"   pool_map size          = {len(pool_map)}")
    print()
    print("Done. User should reload Step 3 (Loan Code Mapping) in the "
          "wizard and pick the correct pool for each CUMA mortgage "
          "code.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
