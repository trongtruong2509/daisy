#!/usr/bin/env python3
"""
Extract complete cell mapping from:
1. TBKQ-phuclong.xls - Template and source data
2. vbd-code.vba - VBA logic showing how data flows
3. TBKQ4-012026.xlsx - Result file showing actual values
"""

import xlrd
import json
from collections import defaultdict

print("=" * 140)
print("BUILDING COMPLETE CELL MAPPING FOR PAYSLIP AUTOMATION")
print("=" * 140)

# ============================================================================
# STEP 1: Analyze Data sheet from TBKQ-phuclong.xls
# ============================================================================

print("\n" + "=" * 140)
print("STEP 1: EXTRACT DATA SHEET COLUMNS (Source Data)")
print("=" * 140)

data_file = "excel-files/TBKQ-phuclong.xls"
wb_data = xlrd.open_workbook(data_file, formatting_info=False)
data_sheet = wb_data.sheet_by_name('Data')

print(f"\nDimensions: {data_sheet.nrows} rows x {data_sheet.ncols} columns")

# Extract column headers from Row 2
data_columns = {}
for col_idx in range(data_sheet.ncols):
    header = data_sheet.cell_value(1, col_idx)  # Row 2 = index 1
    if header and str(header).strip():
        col_letter = ''
        n = col_idx + 1
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            col_letter = chr(65 + remainder) + col_letter
        
        data_columns[col_letter] = str(header).strip()

print(f"\nFound {len(data_columns)} columns in Data sheet:")
print("\nColumn Headers (Row 2 of Data sheet):")
for col_letter in sorted(data_columns.keys()):
    print(f"  {col_letter:3s}: {data_columns[col_letter]}")

# Extract sample data from Row 4 (first employee)
print("\n\nSample Data (Row 4 - First Employee):")
sample_data = {}
for col_idx in range(min(30, data_sheet.ncols)):
    value = data_sheet.cell_value(3, col_idx)  # Row 4 = index 3
    if value is not None and value != '':
        col_letter = ''
        n = col_idx + 1
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            col_letter = chr(65 + remainder) + col_letter
        
        header = data_columns.get(col_letter, f"Col {col_letter}")
        sample_data[col_letter] = value
        print(f"  {col_letter}: {header:40s} = {str(value)[:60]}")

# ============================================================================
# STEP 2: Try to analyze result file TBKQ4-012026.xlsx
# ============================================================================

print("\n" + "=" * 140)
print("STEP 2: ANALYZE RESULT FILE (TBKQ4-012026.xlsx)")
print("=" * 140)

result_file = "excel-files/TBKQ4-012026.xlsx"

# Try to extract using different method - read as xls if possible
try:
    # If it's actually saved as old format, try xlrd with more lenient settings
    import zipfile
    
    with zipfile.ZipFile(result_file, 'r') as z:
        print("\nFile is a valid ZIP (XLSX format)")
        print(f"Contents: {z.namelist()}")
except Exception as e:
    print(f"\nTrying alternative approach for result file...")
    
    try:
        wb_result = xlrd.open_workbook(result_file, formatting_info=False)
        tbkq_result = wb_result.sheet_by_name('TBKQ')
        
        print(f"\nTBKQ Result Sheet: {tbkq_result.nrows} rows x {tbkq_result.ncols} cols")
        
        print("\nKey cells with values (first 35 rows):")
        result_values = {}
        
        for row_idx in range(min(35, tbkq_result.nrows)):
            row_values = {}
            for col_idx in range(tbkq_result.ncols):
                cell_value = tbkq_result.cell_value(row_idx, col_idx)
                cell_type = tbkq_result.cell_type(row_idx, col_idx)
                
                if cell_type != xlrd.XL_CELL_EMPTY and cell_value != '':
                    col_letter = ''
                    n = col_idx + 1
                    while n > 0:
                        n, remainder = divmod(n - 1, 26)
                        col_letter = chr(65 + remainder) + col_letter
                    
                    cell_ref = f"{col_letter}{row_idx + 1}"
                    result_values[cell_ref] = cell_value
                    
                    # Show in rows
                    if cell_ref in ['B3', 'B4', 'D3', 'D4', 'B9', 'D8', 'D9', 'D10']:
                        print(f"  {cell_ref}: {str(cell_value)[:60]}")
        
        result_values_all = result_values
    except Exception as e2:
        print(f"Could not read result file: {e2}")
        result_values_all = {}

# ============================================================================
# STEP 3: Analyze VBA Code Logic
# ============================================================================

print("\n" + "=" * 140)
print("STEP 3: EXTRACT MAPPING FROM VBA CODE")
print("=" * 140)

print("""
KEY VBA LOGIC IDENTIFIED:

1. INPUT CELLS (What we write for each employee):
   - B3 = Data.Range("A" & i)           [MNV - Employee ID from column A]
   
2. COLUMN LOOKUPS (VBA searches for these headers):
   - Finds "MNV" column         → stores in colTBKQ
   - Finds "EmailAddress"       → stores in colemail
   - Finds "PassWord"           → stores in colpw
   
3. DATA VALIDATION:
   - Checks all passwords are filled (column AZ based on earlier analysis)
   - Checks all emails are filled (column found by "EmailAddress" search)
   - Checks no duplicate emails
   
4. PROCESSING LOGIC:
   For each employee (row 4 to lr):
   a) Set B3 = MNV from Data sheet
   b) This triggers VLOOKUP formulas in TBKQ sheet
   c) Read password from column found by "PassWord"
   d) Create copy of TBKQ sheet
   e) Save as XLSX with password protection
   f) Convert formulas to values (PasteSpecial xlPasteValues)
   g) Read email from column found by "EmailAddress"
   h) Send email with attachment
   
5. TEMPLATE CELLS (Where TBKQ gets populated):
   - The TBKQ sheet has VLOOKUP formulas that reference B3
   - When B3 is set to an MNV, these formulas recalculate
   - The VBA then converts formula results to values

6. EMAIL TEMPLATE:
   - Subject: PS.Range("G1").Value (G1 from TBKQ sheet)
   - Body: Multiple cells from "bodymail" sheet (A1, A3, A5, A7, A8, A9, A11-A32)
""")

# ============================================================================
# STEP 4: Build the mapping configuration
# ============================================================================

print("\n" + "=" * 140)
print("STEP 4: RECOMMENDED CELL MAPPING CONFIGURATION FOR PYTHON")
print("=" * 140)

mapping_config = {
    "SHEETS": {
        "data": "Data",
        "template": "TBKQ",
        "email_body": "bodymail"
    },
    "DATA_COLUMNS": {
        "mnv": "A",
        "name": "B",
        "email": "C",
        # "... other columns based on your Data sheet"
        "password": "AZ"
    },
    "TBKQ_INPUT_CELLS": {
        "B3": {
            "description": "Employee MNV (triggers all VLOOKUP formulas)",
            "source": "Data.A",
            "type": "text"
        },
        "B4": {
            "description": "Employee Name (from VLOOKUP)",
            "source": "Data.B",
            "type": "text"
        },
        "D3": {
            "description": "Start Date or other info",
            "source": "Data.C",
            "type": "date"
        }
    },
    "TBKQ_CALCULATED_CELLS": {
        "Note": "These cells contain VLOOKUP or SUM formulas in original",
        "D8": "Salary or income component",
        "D9": "Bonus or other component",
        "D10": "Total (SUM or VLOOKUP result)"
    },
    "EMAIL_BODY_CELLS": {
        "A1": "Greeting",
        "A3": "Main message (contains date placeholder)",
        "A5": "Instructions",
        "A7": "Additional info",
        "A8": "More body text",
        "A9": "Body content",
        "A11": "Important note",
        "A12": "Signature line",
        "Note": "Replace date placeholders (11/2025) with actual DATE from .env"
    },
    "IMPORTANT_NOTES": [
        "The original TBKQ sheet has VLOOKUP formulas that pull data from Data sheet",
        "When B3 is set to an MNV, all VLOOKUP formulas automatically recalculate",
        "For Python direct population, we need to:",
        "  1. Read all data from Data sheet into memory",
        "  2. Understand which columns map to which salary/income components",
        "  3. For each employee, write values directly to cells (no formulas)",
        "  4. Save as XLSX, then convert to PDF"
    ]
}

print(json.dumps(mapping_config, indent=2, ensure_ascii=False))

# ============================================================================
# STEP 5: Create a template for .env configuration
# ============================================================================

print("\n" + "=" * 140)
print("STEP 5: SAMPLE .env CONFIGURATION FILE")
print("=" * 140)

env_template = """
# Excel File Configuration
PAYSLIP_EXCEL_PATH=./excel-files/TBKQ-phuclong.xls

# Sheet Names
DATA_SHEET=Data
TEMPLATE_SHEET=TBKQ
EMAIL_BODY_SHEET=bodymail
SALARY_DATA_SHEET=bang luong

# Column Mappings (Data sheet columns)
DATA_COLUMN_MNV=A
DATA_COLUMN_NAME=B
DATA_COLUMN_EMAIL=C
DATA_COLUMN_PASSWORD=AZ
# Add more columns as needed for all salary components

# TBKQ Template Cell Mappings
TBKQ_CELL_MNV=B3
TBKQ_CELL_NAME=B4
TBKQ_CELL_DATE=D3
# Define all cells that need to be populated...

# Email Template Configuration
EMAIL_SUBJECT_CELL=G1
EMAIL_BODY_CELLS=A1,A3,A5,A7,A8,A9,A11,A12
EMAIL_DATE_PLACEHOLDER_FORMAT=mm/yyyy
EMAIL_DATE_REPLACEMENT_CELL=A3

# Payroll Date
DATE=01/2026

# Outlook Configuration
OUTLOOK_ACCOUNT=your.email@company.com

# Processing Options
DRY_RUN=false

# PDF Password Protection
USE_PDF_PASSWORD=true
PDF_PASSWORD_SOURCE=DATA_COLUMN_PASSWORD
PDF_PASSWORD_PREFIX_STRIP_ZEROS=true

# Output Configuration
OUTPUT_DIR=./payslips
PDF_FILENAME_PATTERN=TBKQ_{name}_{mmyyyy}.pdf

# Processing Options
BATCH_SIZE=10
PARALLEL_WORKERS=5
TEST_MODE=false
"""

print(env_template)

# ============================================================================
# STEP 6: What we still need to know
# ============================================================================

print("\n" + "=" * 140)
print("STEP 6: CRITICAL INFORMATION NEEDED")
print("=" * 140)

critical_questions = """
To build the COMPLETE and ACCURATE mapping, please provide:

1. DATA SHEET COLUMNS - What are all the salary/income components?
   Current known columns:
   - A: MNV (Employee ID)
   - B: Name (Employee Name)
   - C: EmailAddress (Email)
   - AZ: PassWord (PDF Password)
   
   MISSING: What are columns D through AY?
   Please list:
   - Column D: ?
   - Column E: ?
   - Column F: ?
   etc.

2. TBKQ TEMPLATE STRUCTURE - Which cells get populated with which data?
   From VBA, we know B3 is the trigger (MNV).
   MISSING: Which other cells need values?
   Please list key cells:
   - B3: MNV ✓ (known)
   - B4: Name (probably)
   - D3: Date (probably)
   - D8-D10: Salary components (unknown)
   - D14+: Other income details (unknown structure)

3. CALCULATED CELLS - Are there formulas that need calculation?
   For example:
   - D10 = SUM(D8:D9) = Basic Salary + Allowances + Bonus
   Please list any calculations needed.

RECOMMENDATION:
===============
Once you provide the complete column list from Data sheet,
I can build a PyQt5-based column mapping tool that:
1. Reads TBKQ-phuclong.xls
2. Reads TBKQ4-012026.xlsx result
3. Matches columns to cells automatically
4. Generates the exact Python code needed for payslip generation

This way, the mapping will be 100% accurate and maintainable.
"""

print(critical_questions)

wb_data.release_resources()
