#!/usr/bin/env python3
"""Quick analysis of TBKQ template and Data sheet for cell mapping."""
import xlrd

wb = xlrd.open_workbook("excel-files/TBKQ-phuclong.xls")

# --- TBKQ template: all non-empty cells ---
tbkq = wb.sheet_by_name('TBKQ')
print(f"=== TBKQ: {tbkq.nrows}r x {tbkq.ncols}c ===")
for r in range(tbkq.nrows):
    for c in range(tbkq.ncols):
        cell = tbkq.cell(r, c)
        if cell.ctype not in (0, 6):  # not empty/blank
            n = c + 1
            cl = ''
            while n > 0:
                n, rem = divmod(n - 1, 26)
                cl = chr(65 + rem) + cl
            typ = ['E','T','N','D','B','ERR','BL'][cell.ctype]
            print(f"  {cl}{r+1:2d} [{typ}] = {str(cell.value)[:80]}")

# --- Data sheet: ALL column headers ---
print(f"\n=== DATA: ALL HEADERS (Row 2) ===")
data = wb.sheet_by_name('Data')
print(f"Rows: {data.nrows}, Cols: {data.ncols}")
for c in range(data.ncols):
    h = data.cell_value(1, c)
    v4 = data.cell_value(3, c) if data.nrows > 3 else ''
    if h or v4:
        n = c + 1
        cl = ''
        while n > 0:
            n, rem = divmod(n - 1, 26)
            cl = chr(65 + rem) + cl
        print(f"  {cl:3s}: header='{h}' | row4='{v4}'")

# --- Row 3 codes ---
print(f"\n=== DATA: Row 3 numeric codes ===")
for c in range(data.ncols):
    v = data.cell_value(2, c)
    if v:
        n = c + 1
        cl = ''
        while n > 0:
            n, rem = divmod(n - 1, 26)
            cl = chr(65 + rem) + cl
        print(f"  {cl}: {v}")

# --- bodymail ---
print(f"\n=== BODYMAIL ===")
bm = wb.sheet_by_name('bodymail')
for r in range(bm.nrows):
    for c in range(bm.ncols):
        cell = bm.cell(r, c)
        if cell.ctype not in (0, 6):
            print(f"  A{r+1}: {str(cell.value)[:100]}")

wb.release_resources()
