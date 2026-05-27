"""Quick check of generated report formatting."""
import openpyxl

wb = openpyxl.load_workbook(r'Reports\2025-12-31 CECL-Migration-WARM - Franklin_Trust_FCU.xlsx')

def dump_sheet(name, max_row=25, max_col=15):
    ws = wb[name]
    print(f'\n=== {name} ===')
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value is not None:
                f = cell.font
                b = 'B' if f.bold else '-'
                color = f.color.rgb if f.color and f.color.rgb else '-'
                val = repr(cell.value)[:60]
                nf = cell.number_format if cell.number_format != 'General' else ''
                print(f'  R{r}C{c}: {f.name}/{f.size}/{b}/c={color} nf={nf} val={val}')
    print(f'  Merges: {[str(m) for m in ws.merged_cells.ranges][:20]}')
    print(f'  Col widths: ', end='')
    for col in 'ABCDEFGHIJKLMN':
        w = ws.column_dimensions[col].width
        if w:
            print(f'{col}={w:.1f} ', end='')
    print()

ws = wb['ACL Env by Pool Mgmt Adj']
print('ACL Col widths: ', end='')
for col in 'ABCDEFGHIJK':
    w = ws.column_dimensions[col].width
    print(f'{col}={w:.1f} ', end='')
print()
