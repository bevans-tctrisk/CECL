"""Scan ALL Flask session pickle files for the 4 foreign pool names."""
import os
import pickle
import struct
from pathlib import Path
import tempfile

SESS = Path(tempfile.gettempdir()) / "cecl_ui_sessions"
needles = ["Unsecured", "Secured", "1st Mortgage", "2nd Mortgage/HE"]

results = []
for f in SESS.iterdir():
    if not f.is_file():
        continue
    try:
        with f.open("rb") as fh:
            ts_bytes = fh.read(4)
            if len(ts_bytes) < 4:
                continue
            payload = pickle.load(fh)
    except Exception as exc:
        continue
    if not isinstance(payload, dict):
        continue
    state = (payload.get("setup_state") or {})
    pm = state.get("pool_map") or {}
    if not isinstance(pm, dict):
        continue
    pm_values = set(pm.values())
    matches = [n for n in needles if n in pm_values]
    cu = state.get("credit_union") or "(none)"
    short = state.get("short_name") or "(none)"
    if matches:
        results.append((f.name, f.stat().st_mtime, cu, short, matches, len(pm)))

results.sort(key=lambda x: x[1], reverse=True)
print(f"Scanned {sum(1 for _ in SESS.iterdir())} session files; {len(results)} contain any foreign pool name.")
print()
for name, mt, cu, short, matches, pmlen in results:
    import datetime
    when = datetime.datetime.fromtimestamp(mt).strftime("%Y-%m-%d %H:%M:%S")
    print(f"{when}  {name[:20]}  cu={cu!r:30}  short={short!r:25}  pool_map={pmlen}  matches={matches}")
