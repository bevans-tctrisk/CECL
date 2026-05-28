"""One-off diagnostic: print the U29/U30/U31/U33 and AY2..AY5 / AZ2..AZ5
values for camden's 2025-12 and 2026-03 SCALE workbooks."""
import openpyxl
from pathlib import Path

ws_root = Path(r"Z:\Shared\TCT Files\CECL - CM Files\Generated_Reports\camden_firemen_s_fcu")

for period in ["2025-12", "2026-03"]:
    files = sorted((ws_root / period).glob("*Vizo.xlsx"))
    if not files:
        print(f"--- {period}: no Vizo file ---")
        continue
    p = files[0]
    print(f"--- {period}: {p.name} ---")
    for data_only in [True, False]:
        wb = openpyxl.load_workbook(p, data_only=data_only, read_only=False)
        sc = wb["Scale Calculation"]
        hd = wb["Historical Data"]
        print(f"  data_only={data_only}:")
        print(f"    SC!U29={sc['U29'].value!r}")
        print(f"    SC!U30={sc['U30'].value!r}")
        print(f"    SC!U31={sc['U31'].value!r}")
        print(f"    SC!U33={sc['U33'].value!r}")
        print(f"    HD!AY2={hd['AY2'].value!r}  AZ2={hd['AZ2'].value!r}")
        print(f"    HD!AY3={hd['AY3'].value!r}  AZ3={hd['AZ3'].value!r}")
        print(f"    HD!AY4={hd['AY4'].value!r}  AZ4={hd['AZ4'].value!r}")
        print(f"    HD!AY5={hd['AY5'].value!r}  AZ5={hd['AZ5'].value!r}")
        wb.close()
