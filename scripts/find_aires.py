"""Inspect the TCP AIRES sample file for any column that looks like a 'Pool'
column with the foreign names.
"""
import openpyxl
from pathlib import Path

# the saved sample is at C:\Users\BRIANE~1\... let me look in TEMP
import tempfile
import os

# First try: source folder original
candidates = [
    Path(r"Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients\TCP CU 65384\Portfolio Management\CECL Migration IDLR"),
]
for d in candidates:
    if d.exists():
        for sub in d.rglob("TCP Aires*.xlsx"):
            print(sub)

# Also check user TEMP for staged sample
print("\n-- staged sample --")
for d in [Path(os.environ.get("TEMP", ""))]:
    if d.exists():
        for sub in d.rglob("TCP Aires*.xlsx"):
            print(sub)
