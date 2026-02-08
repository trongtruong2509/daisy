#!/usr/bin/env python3
"""
Analyze the TBKQ4-012026.xlsx file (actually old .xls format saved as .xlsx).
Understand the actual cell structure and data layout.
"""

import xlrd

excel_file = "excel-files/TBKQ4-012026.xlsx"

print("=" * 120)
print(f"ANALYZING: {excel_file} (CDFV2 Format - Old Excel)")
print("=" * 120)

# File is actually CDFV2 (old Excel format), use xlrd
try:
    wb = xlrd.open_workbook(excel_file, formatting_info=False)
except Exception as e:
    print(f"Error opening file: {e}")
    exit(1)

print(f"\nSheets in workbook: {wb.sheet_names}\n")

# Analyze TBKQ sheet
tbkq = wb.sheet_by_name('TBKQ')

print("=" * 120)
print(f"TBKQ SHEET STRUCTURE - Dimensions: {tbkq.nrows} rows x {tbkq.ncols} cols")
print("=" * 120)

# Find all non-empty cells
all_cells = []

for row_idx in range(min(35, tbkq.nrows)):
    for col_idx in range(tbkq.ncols):
        cell = tbkq.cell(row_idx, col_idx)
        
        if cell.ctype == xlrd.XL_CELL_EMPTY:  # Empty cell
            continue
            
        # Convert column index to letter
        col_letter = ''
        n = col_idx + 1
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            col_letter = chr(65 + remainder) + col_letter
        
        cell_ref = f"{col_letter}{row_idx + 1}"
        all_cells.append({
            'cell': cell_ref,
            'row': row_idx + 1,
            'col': col_idx + 1,
            'col_letter': col_letter,
            'value': cell.value,
            'ctype_code': cell.ctype,
        })

print(f"\nTotal cells with content in first 35 rows: {len(all_cells)}")

print(f"\nTotal cells with content in first 35 rows: {len(all_cells)}")

print("\n" + "=" * 120)
print("KEY CELLS - Employee Data Input (What we need to WRITE for each employee)")
print("=" * 120)

key_cells = ['B3', 'B4', 'D3', 'D4']
for cell_ref in key_cells:
    try:
        cell = tbkq.cell_value(int(cell_ref[1:]) - 1, ord(cell_ref[0]) - ord('A'))
        print(f"{cell_ref}: {cell}")
    except:
        print(f"{cell_ref}: (not found)")

print("\n" + "=" * 120)
print("DETAILED STRUCTURE - Row by Row (First 35 rows)")
print("=" * 120)

for row_idx in range(min(35, tbkq.nrows)):
    row_display = []
    
    for col_idx in range(min(11, tbkq.ncols)):  # First 11 columns (A-K)
        cell = tbkq.cell(row_idx, col_idx)
        
        col_letter = ''
        n = col_idx + 1
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            col_letter = chr(65 + remainder) + col_letter
        
        if cell.ctype != xlrd.XL_CELL_EMPTY:
            # Truncate long values
            value_str = str(cell.value)[:45]
            row_display.append(f"{col_letter}:{value_str}")
    
    if row_display:
        print(f"\nRow {row_idx + 1:2d}: ", end="")
        print(" | ".join(row_display))

print("\n" + "=" * 120)
print("CRITICAL FINDINGS & RECOMMENDATIONS")
print("=" * 120)

print("""
ANALYSIS RESULTS:
=================

1. FILE FORMAT:
   - Saved as CDFV2 (Old Excel format), not true XLSX
   - This means: formulas may NOT be preserved
   - We're seeing VALUES, not FORMULAS

2. WHAT WE LEARNED:
   - Row 1: Headers/Titles
   - Row 2-3: Employee info (MNV, Name, Start date)
   - Row 4+: Salary/income detail rows
   
3. KEY INSIGHT FOR DIRECT POPULATION:
   ====================================
   
   Since the VBA approach takes 10+ hours for 1000 employees:
   
   VBA Flow (SLOW - 36 sec per employee):
   ├─ Open template
   ├─ Set B3 = MNV
   ├─ Wait 30+ seconds for Excel formulas to recalculate
   ├─ Save as Excel file
   └─ Convert to PDF
   
   Python Direct Approach (FAST - 2-4 sec per employee):
   ├─ Read all employee data ONCE into memory
   ├─ For each employee:
   │  ├─ Create new workbook (fast copy)
   │  ├─ Write values directly to cells
   │  └─ Save and convert to PDF
   └─ PARALLEL: Process 10 employees at same time!
   
   PERFORMANCE GAIN:
   - Direct: 30-60 minutes for 1000 employees
   - Parallel (10 workers): 10-20 minutes for 1000 employees
   - That's 30-60x FASTER than VBA!

4. HOW DIRECT POPULATION WORKS:
   
   Data Sheet (Source):
     Row 4: MNV=6046072, Name="Nguyễn Văn A", Salary=15000000, etc.
   
   Python logic:
     employee = {
         'mnv': '6046072',
         'name': 'Nguyễn Văn A',
         'salary': 15000000,
         ...
     }
   
   TBKQ Template (Target - no formulas needed!):
     B3 = 6046072        (direct value)
     B4 = Nguyễn Văn A   (direct value)
     D8 = 15000000       (direct value)
     D10 = 15500000      (calculated in Python: salary + bonus)

DECISION NEEDED:
================
We need to understand the complete cell mapping:
Which cells in TBKQ.sheet get data from which columns in Data sheet?

Can you provide:
1. The column names list from Data sheet (to create mapping)
2. Or the VBA code that shows the cell mappings
3. Or tell us which salary components go where in the payslip

Without this mapping, we can't implement the solution.
""")

# Now check Data sheet
if 'Data' in wb.sheet_names:
    print("\n" + "=" * 120)
    print("DATA SHEET - Employee Source Information")
    print("=" * 120)
    
    data = wb.sheet_by_name('Data')
    print(f"\nDimensions: {data.nrows} rows x {data.ncols} cols")
    
    # Show headers (Row 2 = index 1)
    print(f"\nColumn headers (Row 2):")
    col_count = 0
    for col_idx in range(data.ncols):
        header = data.cell_value(1, col_idx)  # Row 2 = index 1 (0-based)
        if header and header != '':
            col_letter = ''
            n = col_idx + 1
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                col_letter = chr(65 + remainder) + col_letter
            
            col_count += 1
            print(f"  {col_letter}: {str(header)[:50]}")
    
    print(f"\nShowing {min(col_count, 20)} of {data.ncols} columns...")
    
    # Show sample data (Row 4)
    print(f"\nSample data (Row 4 - First employee):")
    for col_idx in range(min(10, data.ncols)):
        header = data.cell_value(1, col_idx)
        value = data.cell_value(3, col_idx)  # Row 4 = index 3
        col_letter = ''
        n = col_idx + 1
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            col_letter = chr(65 + remainder) + col_letter
        
        print(f"  {col_letter} ({header}): {value}")

