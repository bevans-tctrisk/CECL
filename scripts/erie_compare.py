import openpyxl
import sys

PERIOD = sys.argv[1] if len(sys.argv) > 1 else "2025-11-30"
_ym = PERIOD[:7]
GEN = rf"Z:\Shared\TCT Files\CECL - CM Files\Reports\{PERIOD}_CECL_Migration_Erie_FCU_TCT_Model.xlsx"
WARM = rf"Z:\Shared\Clients\Erie FCU\Portfolio Management (CM, ID, Warm)\{_ym} CECL-Migration-WARM - Erie FCU.xlsx"


KNOWN = {"Auto", "Share Secured", "Unsecured", "Real Estate", "Credit Card",
         "Off Books Real Estate", "Business", "Business Credit Card", "Student Loans",
         "Participation Loans", "Courtesy Pay", "Negative Share"}


def pool_totals(path, sheet):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    out = {}
    cur = None
    for row in ws.iter_rows(values_only=True):
        a = (str(row[0]).strip() if row[0] is not None else "")
        if a in KNOWN and (row[1] in (None, "")):
            cur = a
        elif a == "Total" and cur:
            out[cur] = (row[1], row[2], row[10])  # Balance, SpecID, TotalAllowance
            cur = None
    wb.close()
    return out


gen = pool_totals(GEN, "ACL Env by Pool Mgmt Adj")
warm = pool_totals(WARM, "ACL Env by Pool Mgmt Adj")

def fmt(x):
    return f"{x:,.2f}" if isinstance(x, (int, float)) else str(x)

print(f"{'Pool':24} {'metric':10} {'GENERATED':>18} {'WARM':>18} {'DIFF':>15}")
for pool in ["Auto", "Share Secured", "Unsecured", "Real Estate", "Credit Card"]:
    g = gen.get(pool); w = warm.get(pool)
    if not g or not w:
        print(f"{pool:24}  MISSING gen={bool(g)} warm={bool(w)}")
        continue
    for j, label in enumerate(["Balance", "SpecID", "TotalAllow"]):
        gv = g[j] or 0; wv = w[j] or 0
        diff = (gv - wv) if isinstance(gv, (int, float)) and isinstance(wv, (int, float)) else "?"
        d = f"{diff:+,.2f}" if isinstance(diff, (int, float)) else diff
        print(f"{pool:24} {label:10} {fmt(gv):>18} {fmt(wv):>18} {d:>15}")
    print()
