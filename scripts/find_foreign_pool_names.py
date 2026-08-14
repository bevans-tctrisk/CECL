import json, os, sys, glob
os.chdir(r"c:\Dev\CECL")
for f in sorted(glob.glob("wizard_drafts/*.json")):
    try:
        d = json.load(open(f, encoding="utf-8"))
    except Exception as e:
        print(f"{f}: cannot load - {e}")
        continue
    pm = d.get("pool_map") or {}
    vals = sorted({(v or "").strip() for v in pm.values()})
    ps = [(p or {}).get("name") for p in (d.get("pool_settings") or [])]
    print(f"{f}:")
    print(f"  CU: {d.get('credit_union')!r}  short={d.get('short_name')!r}  charter={d.get('charter_number')!r}")
    print(f"  pool_settings ({len(ps)}): {ps}")
    print(f"  pool_map ({len(pm)} entries) distinct values: {vals}")
    # Hunt for the 4 foreign names that polluted TCP
    foreign = {"Unsecured", "Secured", "1st Mortgage", "2nd Mortgage/HE"}
    if foreign & set(vals):
        print(f"  *** MATCHES foreign names: {sorted(foreign & set(vals))}")
    print()
