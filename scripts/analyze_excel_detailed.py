#!/usr/bin/env python3
"""
Detailed analysis of Data sheet columns.
"""

import xlrd
from pathlib import Path

file_path = Path(__file__).parent / "excel-files" / "TBKQ-phuclong.xls"
wb = xlrd.open_workbook(file_path)

# Analyze Data sheet
data_sheet = wb.sheet_by_name("Data")

print("="*80)
print("DATA SHEET - DETAILED COLUMN ANALYSIS")
print("="*80)
print(f"\nDimensions: {data_sheet.nrows} rows x {data_sheet.ncols} columns")

# Row 2 has the main headers (Vietnamese)
print("\n" + "="*80)
print("ROW 2 - MAIN COLUMN HEADERS")
print("="*80)
for col_idx in range(data_sheet.ncols):
    col_letter = xlrd.colname(col_idx)
    value = data_sheet.cell(1, col_idx).value  # Row 2 is index 1
    if value:
        print(f"{col_letter:3s} | {value}")

# Row 3 has numeric codes
print("\n" + "="*80)
print("ROW 3 - NUMERIC CODES")
print("="*80)
for col_idx in range(data_sheet.ncols):
    col_letter = xlrd.colname(col_idx)
    value = data_sheet.cell(2, col_idx).value  # Row 3 is index 2
    if value:
        print(f"{col_letter:3s} | {value}")

# Show sample data from row 4
print("\n" + "="*80)
print("ROW 4 - SAMPLE EMPLOYEE DATA")
print("="*80)
for col_idx in range(min(20, data_sheet.ncols)):  # First 20 columns
    col_letter = xlrd.colname(col_idx)
    header = data_sheet.cell(1, col_idx).value
    value = data_sheet.cell(3, col_idx).value  # Row 4 is index 3
    print(f"{col_letter:3s} | {header:30s} | {value}")

# Find specific columns
print("\n" + "="*80)
print("SEARCHING FOR KEY COLUMNS")
print("="*80)

keywords = ['password', 'email', 'mnv', 'họ', 'tên']
for col_idx in range(data_sheet.ncols):
    col_letter = xlrd.colname(col_idx)
    header = str(data_sheet.cell(1, col_idx).value).lower()
    
    for keyword in keywords:
        if keyword in header:
            print(f"Found '{keyword}' in column {col_letter}: {data_sheet.cell(1, col_idx).value}")
            break

# Check for column AZ (column index 51)
print("\n" + "="*80)
print("COLUMN AZ (INDEX 51) - PASSWORD COLUMN")
print("="*80)
if data_sheet.ncols > 51:
    col_letter = xlrd.colname(51)
    header = data_sheet.cell(1, 51).value
    sample = data_sheet.cell(3, 51).value if data_sheet.nrows > 3 else "N/A"
    print(f"Column: {col_letter}")
    print(f"Header: {header}")
    print(f"Sample value: {sample}")
else:
    print(f"Data sheet only has {data_sheet.ncols} columns (AZ would be column 52)")

# Analyze TBKQ sheet structure
print("\n" + "="*80)
print("TBKQ SHEET - PAYSLIP TEMPLATE STRUCTURE")
print("="*80)
tbkq_sheet = wb.sheet_by_name("TBKQ")
print(f"Dimensions: {tbkq_sheet.nrows} rows x {tbkq_sheet.ncols} columns")
print("\nKey cells:")
for row_idx in range(min(15, tbkq_sheet.nrows)):
    row_data = []
    for col_idx in range(min(5, tbkq_sheet.ncols)):
        value = tbkq_sheet.cell(row_idx, col_idx).value
        if value:
            row_data.append(f"{xlrd.colname(col_idx)}{row_idx+1}: {str(value)[:40]}")
    if row_data:
        print(f"Row {row_idx+1}: {' | '.join(row_data)}")

# Analyze bodymail sheet
print("\n" + "="*80)
print("BODYMAIL SHEET - EMAIL TEMPLATE")
print("="*80)
bodymail_sheet = wb.sheet_by_name("bodymail")
print(f"Dimensions: {bodymail_sheet.nrows} rows x {bodymail_sheet.ncols} columns")
print("\nEmail template content:")
for row_idx in range(bodymail_sheet.nrows):
    value = bodymail_sheet.cell(row_idx, 0).value
    if value:
        print(f"A{row_idx+1}: {value[:80]}")

# Analyze bang luong sheet
print("\n" + "="*80)
print("BANG LUONG SHEET - SALARY DATA")
print("="*80)
bang_luong = wb.sheet_by_name("bang luong")
print(f"Dimensions: {bang_luong.nrows} rows x {bang_luong.ncols} columns")
print("\nFirst 30 column headers (Row 2):")
for col_idx in range(min(30, bang_luong.ncols)):
    col_letter = xlrd.colname(col_idx)
    header = bang_luong.cell(1, col_idx).value  # Row 2 is index 1
    if header:
        print(f"{col_letter:3s} | {header}")
