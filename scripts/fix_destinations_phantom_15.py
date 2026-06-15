"""Phase 9.26 cleanup for Destinations CU live wizard draft.

Removes the phantom ``'15'`` truncation artifact from ``pool_map`` and
re-extracts ``pool_code_suggestions`` for the CUMA MTG Servicing extract
with ``split_char=""`` so the suggestion list carries the full composite
labels (``'15/15 ARM'``) instead of the truncated leading numbers.

Idempotent. Backs up to ``<draft>.bak.phase926`` before mutating.
"""
from __future__ import annotations
import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from cecl_ui.services import sample_parser  # noqa: E402

DRAFT = Path(
    r"Z:\Shared\TCT Files\CECL - CM Files\wizard_drafts"
    r"\destinations_cu__migration.json"
)
BACKUP = DRAFT.with_suffix(DRAFT.suffix + ".bak.phase926")


def main() -> None:
    if not DRAFT.exists():
        print(f"[FAIL] draft not found: {DRAFT}")
        sys.exit(1)

    if not BACKUP.exists():
        shutil.copy2(DRAFT, BACKUP)
        print(f"[backup] {BACKUP}")
    else:
        print(f"[backup] already present: {BACKUP}")

    state = json.loads(DRAFT.read_text(encoding="utf-8"))

    pool_map = state.get("pool_map") or {}
    changed = False

    if "15" in pool_map:
        old_val = pool_map.pop("15")
        print(f"[fix] removed pool_map['15'] (was -> {old_val!r})")
        changed = True
    else:
        print("[skip] pool_map['15'] not present (already cleaned)")

    su = state.get("sample_uploads") or {}
    for entry in su.get("loan_data_files") or []:
        name = entry.get("name") or ""
        if "CUMA" not in name.upper() and "MTG" not in name.upper():
            continue
        analysis = entry.get("analysis") or {}
        cur_sugg = list(analysis.get("pool_code_suggestions") or [])
        path = entry.get("path") or ""
        col = (analysis.get("column_suggestions") or {}).get("loan_pool_code")
        if not (path and col and Path(path).exists()):
            print(
                f"[warn] cannot re-extract for {name}: path missing or "
                f"loan_pool_code column missing"
            )
            continue
        try:
            fresh = sample_parser.extract_pool_codes(
                path,
                column_name=col,
                header_row=entry.get("header_row"),
                split_char="",
            )
        except Exception as exc:  # noqa: BLE001
            print(f"[warn] re-extract failed for {name}: {exc}")
            continue
        if fresh and fresh != cur_sugg:
            analysis["pool_code_suggestions"] = list(fresh)
            entry["analysis"] = analysis
            print(f"[fix] {name}: pool_code_suggestions {cur_sugg} -> {list(fresh)}")
            changed = True
        else:
            print(f"[skip] {name}: pool_code_suggestions already current ({len(cur_sugg)})")

    if changed:
        DRAFT.write_text(json.dumps(state, indent=2), encoding="utf-8")
        print(f"[saved] {DRAFT}")
    else:
        print("[noop] nothing changed")


if __name__ == "__main__":
    main()
