"""Phase 9.27 — Patch the live Emergency Responders CU draft on Z: drive
to recover the historical-balance layout (pool_name_col A -> B,
parsed_pool_labels [] -> 32 items, seed pool_map).

Idempotent: re-runnable; only acts when source=single AND parsed_pool_labels
empty AND saved_path exists. Backs up the draft to .bak_phase927 first.
"""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

DRAFT = Path(
    r"Z:\Shared\TCT Files\CECL - CM Files\wizard_drafts"
    r"\emergency_responders_cu__migration.json"
)


def main() -> int:
    if not DRAFT.exists():
        print(f"Draft not found: {DRAFT}")
        return 1
    backup = DRAFT.with_suffix(DRAFT.suffix + ".bak_phase927")
    shutil.copy2(DRAFT, backup)
    print(f"Backup: {backup}")

    state = json.loads(DRAFT.read_text(encoding="utf-8"))
    mb_before = state.get("monthly_bal", {}) or {}
    print("BEFORE:")
    print(f"  source={mb_before.get('source')!r}")
    print(f"  pool_name_col={mb_before.get('pool_name_col')!r}")
    print(f"  parsed_pool_labels={len(mb_before.get('parsed_pool_labels') or [])}")
    print(f"  pool_map keys={len(mb_before.get('pool_map') or {})}")

    from cecl_ui.services import auto_setup

    msgs = auto_setup.selfheal_single_hist_bal_layout(state)
    print()
    print("Self-heal messages:")
    for m in msgs:
        print(f"  - {m}")

    mb_after = state.get("monthly_bal", {}) or {}
    print()
    print("AFTER:")
    print(f"  pool_name_col={mb_after.get('pool_name_col')!r}")
    print(f"  parsed_pool_labels={len(mb_after.get('parsed_pool_labels') or [])}")
    pm = mb_after.get("pool_map") or {}
    mapped = sum(1 for v in pm.values() if v)
    print(f"  pool_map keys={len(pm)} (auto-mapped: {mapped})")
    if mb_after.get("parsed_pool_labels"):
        print("  first 5 labels:", (mb_after["parsed_pool_labels"] or [])[:5])

    DRAFT.write_text(json.dumps(state, indent=2), encoding="utf-8")
    print()
    print("Draft saved.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
