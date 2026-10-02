"""Track how monthly_bal.pool_name_col evolved across Emergency Responders draft backups."""
import json
from pathlib import Path

folder = Path(r"Z:\Shared\TCT Files\CECL - CM Files\wizard_drafts")
files = sorted(
    [f for f in folder.iterdir() if "emergency" in f.name.lower()],
    key=lambda p: p.stat().st_mtime,
)
for f in files:
    try:
        d = json.loads(f.read_text())
    except Exception as e:
        print(f"{f.name}: ERROR {e}")
        continue
    mb = d.get("monthly_bal") or {}
    src = mb.get("source")
    pnc = mb.get("pool_name_col")
    labels = mb.get("parsed_pool_labels") or []
    pmap = mb.get("pool_map") or {}
    sp = mb.get("saved_path") or ""
    print(
        f"{f.stat().st_mtime:.0f} {f.name}\n"
        f"   src={src!r} pnc={pnc!r} labels={len(labels)} pool_map={len(pmap)} "
        f"saved_path={sp[-60:]!r}"
    )
