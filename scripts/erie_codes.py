import openpyxl, os
from collections import defaultdict

WARM = r"Z:\Shared\Clients\Erie FCU\Portfolio Management (CM, ID, Warm)\2025-11 CECL-Migration-WARM - Erie FCU.xlsx"
QDIR = r"Z:\Shared\Clients\Erie FCU\Client Access\Portfolio Management (CM, ID, Warm)\2025-11"
POOLMAP = r"Z:\Shared\Clients\Erie FCU\Client Access\Portfolio Management (CM, ID, Warm)\Erie FCU - Pool Mapping.xlsx"

# collateral code (int) -> pool
wb = openpyxl.load_workbook(POOLMAP, read_only=True, data_only=True)
coll = {}
for row in wb["Loan Codes"].iter_rows(min_row=2, values_only=True):
    if row[0] is not None and row[0] != "Collateral Code":
        coll[str(row[0]).strip()] = row[1]
wb.close()
print("collateral map has", len(coll), "codes")

# distinct export codes
codes = defaultdict(lambda: [0, 0.0])
with open(os.path.join(QDIR, "TCT.EXPORT_202511.txt"), encoding="latin-1") as f:
    for line in f:
        p = line.rstrip("\n").split("|")
        if len(p) < 11:
            continue
        try:
            bal = float(p[3])
        except ValueError:
            continue
        codes[p[2]][0] += 1
        codes[p[2]][1] += bal

print("\ndistinct export codes:", len(codes))
unmapped = []
for code in sorted(codes):
    key = code.lstrip("0") or "0"   # strip leading zeros to match int keys
    pool = coll.get(key) or coll.get(code)
    if code == "MC":
        pool = "Credit Card"
    cnt, bal = codes[code]
    flag = "" if pool else "  <<< UNMAPPED"
    if not pool:
        unmapped.append(code)
    print(f"  {code!r:8} -> {pool!r:20} {cnt:6d} loans ${bal:,.2f}{flag}")
print("\nUNMAPPED codes:", unmapped)

# Pool totals via mapping
pooltot = defaultdict(float)
for code, (cnt, bal) in codes.items():
    key = code.lstrip("0") or "0"
    pool = ("Credit Card" if code == "MC" else coll.get(key)) or "UNMAPPED"
    pooltot[pool] += bal
print("\nExport pool totals:")
for pool, bal in sorted(pooltot.items(), key=lambda x: -x[1]):
    print(f"   {pool:20} ${bal:,.2f}")
