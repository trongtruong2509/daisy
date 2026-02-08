#!/usr/bin/env python3
"""
Analyze TBKQ sheet formulas to understand what data is being looked up.
This helps decide between formula-based vs direct population approach.
"""

import xlrd
import os

excel_file = "excel-files/TBKQ-phuclong.xls"
if not os.path.exists(excel_file):
    print(f"Error: {excel_file} not found!")
    exit(1)

print(f"Analyzing formulas in {excel_file}...\n")

# Open workbook
workbook = xlrd.open_workbook(excel_file, formatting_info=False)

# Get TBKQ sheet
tbkq_sheet = workbook.sheet_by_name('TBKQ')

print(f"TBKQ Sheet: {tbkq_sheet.nrows} rows x {tbkq_sheet.ncols} columns\n")

# Check cells for formulas
# xlrd doesn't directly expose formulas, but we can check cell types
# and look at values that might indicate formula results

print("=" * 80)
print("ANALYZING TBKQ TEMPLATE STRUCTURE")
print("=" * 80)

# Sample key cells to understand structure
key_cells = [
    ('B3', 2, 1),   # Employee MNV input cell
    ('B4', 3, 1),   # Likely employee name
    ('B5', 4, 1),   # Likely other employee info
    ('B6', 5, 1),
    ('B7', 6, 1),
    ('B8', 7, 1),
    ('B9', 8, 1),
    ('B10', 9, 1),
]

print("\nKey cells in TBKQ template:")
print("-" * 80)
for cell_name, row, col in key_cells:
    if row < tbkq_sheet.nrows and col < tbkq_sheet.ncols:
        cell = tbkq_sheet.cell(row, col)
        value = cell.value
        cell_type = cell.ctype  # 0=empty, 1=text, 2=number, 3=date, 4=boolean, 5=error, 6=blank
        print(f"Cell {cell_name} (R{row+1}C{col+1}): Type={cell_type}, Value='{value}'")

print("\n" + "=" * 80)
print("SCANNING ALL NON-EMPTY CELLS")
print("=" * 80)

non_empty_cells = []
for row_idx in range(min(30, tbkq_sheet.nrows)):  # First 30 rows
    for col_idx in range(min(15, tbkq_sheet.ncols)):  # First 15 columns
        cell = tbkq_sheet.cell(row_idx, col_idx)
        if cell.ctype != 0 and cell.ctype != 6:  # Not empty or blank
            # Convert to Excel column letter
            col_letter = ''
            n = col_idx + 1
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                col_letter = chr(65 + remainder) + col_letter
            
            cell_ref = f"{col_letter}{row_idx + 1}"
            non_empty_cells.append({
                'ref': cell_ref,
                'row': row_idx + 1,
                'col': col_idx + 1,
                'type': cell.ctype,
                'value': str(cell.value)[:60]  # Truncate long values
            })

# Group by row for better readability
current_row = -1
for cell in non_empty_cells:
    if cell['row'] != current_row:
        current_row = cell['row']
        print(f"\n--- Row {current_row} ---")
    
    type_map = {0: 'Empty', 1: 'Text', 2: 'Number', 3: 'Date', 4: 'Bool', 5: 'Error', 6: 'Blank'}
    type_str = type_map.get(cell['type'], str(cell['type']))
    print(f"  {cell['ref']:6s} ({type_str:6s}): {cell['value']}")

print("\n" + "=" * 80)
print("ANALYSIS SUMMARY")
print("=" * 80)
print("""
To understand what formulas exist in the original .xlsx file,
we need to open it in Excel or use openpyxl (which preserves formulas).

However, based on VBA code analysis:
- VBA sets B3 = Employee MNV
- This triggers VLOOKUP formulas that populate the template
- VBA then saves as PDF

KEY QUESTION: Is there an .xlsx version with formulas intact?
The current .xls file may have lost formula information.
""")

print("\nRECOMMENDATION:")
print("-" * 80)
print("""
For 1000+ employees with current VBA taking 10+ hours:

DIRECT POPULATION APPROACH (Recommended):
1. Analyze original .xlsx to map: which cells contain which data
2. Build a mapping in Python: {cell_ref: data_source}
   Example: B3 -> MNV, B4 -> Name, D10 -> Salary Component X
3. For each employee:
   - Read employee data once (in memory)
   - Create new workbook from template
   - Populate all cells directly (no formulas)
   - Save as PDF using Excel COM or reportlab
4. Can process in parallel (e.g., 10 employees at a time)

ESTIMATED PERFORMANCE:
- Current VBA: 10+ hours for 1000 employees (~36 seconds per employee)
- Direct population: 30-60 minutes for 1000 employees (~2-4 seconds per employee)
- With parallel processing: 10-20 minutes for 1000 employees

BENEFITS:
- 10-30x faster execution
- No formula recalculation overhead
- Can parallelize safely
- More predictable performance
- Better error handling
""")
