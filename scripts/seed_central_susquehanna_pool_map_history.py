"""One-shot: seed ``state["_pool_map_history"]`` for the
central_susquehanna_comm_fcu wizard draft so a future extract
delete-and-re-add cycle can restore pool_map assignments transparently.

Idempotent. Reads the current live draft, snapshots every non-blank
pool_map entry into ``_pool_map_history``, and saves.
"""

from __future__ import annotations

import json
import os
import shutil
from pathlib import Path

WORKSPACE_ROOT = r"Z:\Shared\TCT Files\CECL - CM Files"
os.environ.setdefault("CECL_WORKSPACE_ROOT", WORKSPACE_ROOT)

DRAFT_PATH = Path(WORKSPACE_ROOT) / "wizard_drafts" / "central_susquehanna_comm_fcu__migration.json"
BACKUP_SUFFIX = ".bak.phase_pool_map_history"


def main() -> int:
    if not DRAFT_PATH.exists():
        print(f"Draft not found: {DRAFT_PATH}")
        return 1

    backup = DRAFT_PATH.with_suffix(DRAFT_PATH.suffix + BACKUP_SUFFIX)
    if not backup.exists():
        shutil.copy2(DRAFT_PATH, backup)
        print(f"Backed up draft to: {backup.name}")
    else:
        print(f"Backup already exists: {backup.name} (leaving alone)")

    payload = json.loads(DRAFT_PATH.read_text(encoding="utf-8"))
    state = payload.get("state") or payload

    pm = state.get("pool_map") or {}
    hist = state.setdefault("_pool_map_history", {})
    if not isinstance(hist, dict):
        hist = {}
        state["_pool_map_history"] = hist

    added = 0
    updated = 0
    for code, pool in pm.items():
        if not isinstance(code, str):
            continue
        val = (pool or "").strip() if isinstance(pool, str) else ""
        if not val:
            continue
        prev = hist.get(code)
        if prev is None:
            hist[code] = val
            added += 1
        elif prev != val:
            hist[code] = val
            updated += 1

    print(f"pool_map entries: {len(pm)}")
    print(f"_pool_map_history: added {added}, updated {updated}, total {len(hist)}")
    for code, pool in sorted(hist.items()):
        print(f"  {code!r:>8} -> {pool!r}")

    DRAFT_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    print(f"Saved: {DRAFT_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
