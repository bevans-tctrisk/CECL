"""One-shot: remap TCP CU's polluted pool_map using TCP's own TCT Edits file.

Loads `TCP Loan Type Codes with Pools- TCT Edits.xlsx` from Z:\\, builds a precise
code -> "Pool for Report" mapping, then patches the active TCP CU Flask session
file at %TEMP%\\cecl_ui_sessions\\bb143c82ee0a42377a3cabfd842a07ef.

Session file format: 4-byte big-endian uint32 timestamp header + raw pickle.

Idempotent: re-running with already-clean pool_map is a no-op.
"""
from __future__ import annotations

import os
import pickle
import struct
import sys
from pathlib import Path

import openpyxl

TCT_EDITS = Path(
    r"Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients"
    r"\TCP CU 65384\Portfolio Management\CECL Migration IDLR"
    r"\TCP Loan Type Codes with Pools- TCT Edits.xlsx"
)
SESSION_FILE = Path(os.environ["TEMP"]) / "cecl_ui_sessions" / "bb143c82ee0a42377a3cabfd842a07ef"

FOREIGN_BUCKETS = {"1st Mortgage", "2nd Mortgage/HE", "Unsecured", "Secured"}


def _norm_code(v: object) -> str:
    s = "" if v is None else str(v).strip()
    # Preserve leading zeros: TCT file stores "01", "02", "1", "2" all as strings already.
    return s


def load_remap() -> dict[str, str]:
    """Returns {loan_code: pool_for_report}. Pool col is index 1."""
    wb = openpyxl.load_workbook(TCT_EDITS, data_only=True, read_only=True)
    ws = wb["Loan Codes"]
    remap: dict[str, str] = {}
    rows = list(ws.iter_rows(values_only=True))
    # row 1 is header: ['Loan Type Code', 'Pool for Report', 'Description', 'Pool']
    for row in rows[1:]:
        if not row:
            continue
        code = _norm_code(row[0])
        pool = _norm_code(row[1])
        if code and pool:
            remap[code] = pool
    return remap


def patch_session(remap: dict[str, str]) -> None:
    raw = SESSION_FILE.read_bytes()
    # 4-byte big-endian uint32 timestamp header + pickle
    ts_bytes = raw[:4]
    payload = pickle.loads(raw[4:])
    state = payload.get("setup_state") or {}
    pm = state.get("pool_map") or {}
    if not isinstance(pm, dict):
        print("WARN: pool_map is not a dict; aborting")
        return

    settings = state.get("pool_settings") or []
    valid_names = {p.get("name") for p in settings if p.get("name")}
    valid_names |= {"Ignore", "Exclude"}

    changes: list[tuple[str, str, str]] = []
    for code, current in list(pm.items()):
        if current in FOREIGN_BUCKETS or (current and current not in valid_names):
            target = remap.get(_norm_code(code))
            if target and target in valid_names:
                pm[code] = target
                changes.append((code, current, target))
            else:
                pm[code] = "Ignore"
                changes.append((code, current, "Ignore"))

    if not changes:
        print("No polluted entries found. Session is clean.")
        return

    state["pool_map"] = pm
    payload["setup_state"] = state

    # Backup before write
    backup = SESSION_FILE.with_suffix(".bak")
    backup.write_bytes(raw)
    print(f"Backup written: {backup}")

    new_pickle = pickle.dumps(payload)
    SESSION_FILE.write_bytes(ts_bytes + new_pickle)
    print(f"Session patched: {SESSION_FILE}")
    print(f"Changes: {len(changes)}")
    for code, old, new in sorted(changes):
        print(f"  {code:>3}  {old!r:>22}  ->  {new!r}")


def main() -> int:
    if not TCT_EDITS.exists():
        print(f"ERROR: TCT Edits file not found: {TCT_EDITS}")
        return 1
    if not SESSION_FILE.exists():
        print(f"ERROR: Session file not found: {SESSION_FILE}")
        return 1

    remap = load_remap()
    print(f"Loaded {len(remap)} code->pool entries from TCT Edits")
    patch_session(remap)
    return 0


if __name__ == "__main__":
    sys.exit(main())
