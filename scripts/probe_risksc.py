"""Probe RISKSC values for the impaired members in ceclce/cecloe."""
import json
from pathlib import Path
import pandas as pd

state = json.loads(Path(r"Z:\Shared\TCT Files\CECL - CM Files\wizard_drafts\destinations_cu__migration.json").read_text(encoding="utf-8"))
su = (state.get("sample_uploads") or {}).get("loan_data_files") or []
needles = [("15031","521"),("14601","560"),("4961","780"),("26791","560"),("19491","566"),("25244","560")]

for entry in su:
    nm = entry.get("name") or ""
    if "ceclce" in nm or "cecloe" in nm:
        df = pd.read_excel(entry["path"], dtype=object)
        df.columns = [str(c) for c in df.columns]
        risksc_col = next((c for c in df.columns if c.strip().upper() == "RISKSC"), None)
        print(f"{nm}: RISKSC col = {risksc_col!r}")
        for m, s in needles:
            mask = (df["ACCTBS"].astype(str).str.strip() == m) & \
                   (df["ACTTYP"].astype(str).str.strip().str.zfill(3) == s.zfill(3))
            if mask.any():
                r = df[mask].iloc[0]
                v = r.get(risksc_col) if risksc_col else None
                print(f"  {m}-{s}: RISKSC={v!r}")
