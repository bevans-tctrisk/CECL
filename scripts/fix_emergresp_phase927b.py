"""Apply Phase 9.27b auto-mapping to the live Emergency Responders draft.

Backs up the existing draft to ``.bak_phase927b`` then runs
``auto_setup.selfheal_auto_map_balance_labels`` to fill blank
``monthly_bal.pool_map`` entries using keyword-category heuristics.

Idempotent: re-running on an already-mapped draft is a no-op.
"""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from cecl_ui.services import auto_setup  # noqa: E402

DRAFT = Path(
    r"Z:\Shared\TCT Files\CECL - CM Files\wizard_drafts\\"
    r"emergency_responders_cu__migration.json"
)


def main() -> int:
    if not DRAFT.exists():
        print(f"ERROR: draft not found: {DRAFT}")
        return 1
    backup = DRAFT.with_suffix(DRAFT.suffix + ".bak_phase927b")
    if not backup.exists():
        shutil.copy2(DRAFT, backup)
        print(f"Backup written: {backup}")
    else:
        print(f"Backup already exists: {backup}")

    state = json.loads(DRAFT.read_text())

    mb = state.get("monthly_bal") or {}
    pm_before = mb.get("pool_map") or {}
    blanks_before = sum(1 for v in pm_before.values() if not v)

    msgs = auto_setup.selfheal_auto_map_balance_labels(state)
    print()
    print("Messages:")
    for m in msgs:
        print(f"  {m}")
    if not msgs:
        print("  (no-op — pool_map already populated or no labels/pools)")

    pm_after = state.get("monthly_bal", {}).get("pool_map") or {}
    blanks_after = sum(1 for v in pm_after.values() if not v)
    filled = blanks_before - blanks_after

    print()
    print(f"Blank entries: {blanks_before} -> {blanks_after} (filled {filled})")

    if filled:
        DRAFT.write_text(json.dumps(state, indent=2))
        print(f"Draft saved: {DRAFT}")

    print()
    print("Final pool_map:")
    for k, v in sorted(pm_after.items()):
        marker = "->" if v else "  "
        print(f"  {k!r:45s} {marker} {v!r}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
