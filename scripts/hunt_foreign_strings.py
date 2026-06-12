"""Hunt for the foreign pool-name strings in TCP source files."""
import openpyxl
from pathlib import Path

SRC = Path(r"Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access"
           r"\Clients\TCP CU 65384\Portfolio Management\CECL Migration IDLR")

needles = ("1st Mortgage", "2nd Mortgage/HE", "2nd Mortgage")
for f in SRC.glob("*.xlsx"):
    wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
    for ws in wb.worksheets:
        for ridx, row in enumerate(ws.iter_rows(values_only=True), 1):
            cells = [("" if c is None else str(c)) for c in row]
            line = " | ".join(cells)
            for needle in needles:
                if needle in line:
                    print(f"{f.name} :: {ws.title} :: row {ridx}")
                    print(f"   {line[:300]}")
                    break
    wb.close()
