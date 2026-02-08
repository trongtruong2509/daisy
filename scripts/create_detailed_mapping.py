#!/usr/bin/env python3
"""
Create detailed mapping by reading TBKQ template and comparing with example output.
Since result file is encrypted, use TBKQ template directly and infer from structure.
"""

import xlrd
import json

print("=" * 140)
print("DETAILED MAPPING ANALYSIS - TBKQ TEMPLATE STRUCTURE")
print("=" * 140)

# Read the TBKQ template
template_file = "excel-files/TBKQ-phuclong.xls"
wb = xlrd.open_workbook(template_file, formatting_info=False)
tbkq = wb.sheet_by_name('TBKQ')

print(f"\nTBKQ Template: {tbkq.nrows} rows x {tbkq.ncols} columns")

# Extract all cells with content
print("\n" + "=" * 140)
print("COMPLETE CELL STRUCTURE (First 35 rows)")
print("=" * 140)

cell_structure = {}

for row_idx in range(min(35, tbkq.nrows)):
    row_content = []
    
    for col_idx in range(tbkq.ncols):
        cell = tbkq.cell(row_idx, col_idx)
        
        if cell.ctype != xlrd.XL_CELL_EMPTY and cell.value != '':
            col_letter = ''
            n = col_idx + 1
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                col_letter = chr(65 + remainder) + col_letter
            
            cell_ref = f"{col_letter}{row_idx + 1}"
            cell_value = str(cell.value)[:100]  # Truncate
            
            cell_structure[cell_ref] = {
                'value': cell_value,
                'ctype': cell.ctype
            }
            
            row_content.append(f"{col_letter}:{cell_value}")
    
    if row_content:
        print(f"\nRow {row_idx + 1:2d}: {' | '.join(row_content)}")

# Identify key cells
print("\n" + "=" * 140)
print("KEY CELLS FOR MAPPING")
print("=" * 140)

print("""
From VBA analysis and TBKQ inspection:

1. TRIGGER/INPUT CELLS:
   - B3: Employee MNV (this is the key trigger cell)
   
2. LIKELY DATA PULL CELLS (need VLOOKUP or direct values):
   - B4: Employee Name
   - D3: Start Date or Department
   - D4: Department or Job Title
   - B9+: Salary detail rows
   - D8+: Income detail rows
   
3. EMAIL TEMPLATE CELLS (in bodymail sheet):
   - Subject: G1 from TBKQ sheet
   - Body: A1, A3, A5, A7, A8, A9, A11, A12, etc. from bodymail sheet
""")

# Read the bodymail sheet
print("\n" + "=" * 140)
print("EMAIL BODY TEMPLATE (bodymail sheet)")
print("=" * 140)

bodymail = wb.sheet_by_name('bodymail')
print(f"\nbodymail: {bodymail.nrows} rows x {bodymail.ncols} columns")

print("\nEmail body cells:")
for row_idx in range(bodymail.nrows):
    for col_idx in range(min(3, bodymail.ncols)):
        cell = bodymail.cell(row_idx, col_idx)
        if cell.ctype != xlrd.XL_CELL_EMPTY:
            col_letter = ''
            n = col_idx + 1
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                col_letter = chr(65 + remainder) + col_letter
            
            cell_ref = f"{col_letter}{row_idx + 1}"
            cell_value = str(cell.value)[:80]
            print(f"  {cell_ref}: {cell_value}")

# Create comprehensive mapping
print("\n" + "=" * 140)
print("COMPLETE CONFIGURATION FOR PYTHON IMPLEMENTATION")
print("=" * 140)

complete_mapping = {
    "metadata": {
        "version": "1.0",
        "created": "2026-02-06",
        "description": "Cell mapping for Phuc Long payslip automation"
    },
    
    "sheets": {
        "data": "Data",
        "template": "TBKQ", 
        "email_body": "bodymail",
        "salary_detail": "bang luong"
    },
    
    "data_sheet_columns": {
        "A": {
            "header": "MNV",
            "python_name": "employee_id",
            "type": "text",
            "required": True,
            "usage": "Primary key, used as lookup key"
        },
        "B": {
            "header": "Họ và tên",
            "python_name": "name",
            "type": "text",
            "required": True,
            "usage": "Employee name"
        },
        "C": {
            "header": "EmailAddress",
            "python_name": "email",
            "type": "email",
            "required": True,
            "usage": "Email for payslip delivery"
        },
        "D": {
            "header": "Chức danh",
            "python_name": "job_title",
            "type": "text",
            "required": False,
            "usage": "Job title"
        },
        "E": {
            "header": "Số tài khoản",
            "python_name": "account_number",
            "type": "text",
            "required": False,
            "usage": "Bank account number"
        },
        "F": {
            "header": "Ngày bắt đầu làm việc",
            "python_name": "start_date",
            "type": "date",
            "required": False,
            "usage": "Employment start date"
        },
        "G": {
            "header": "Master code",
            "python_name": "master_code",
            "type": "text",
            "required": False,
            "usage": "Master identifier"
        },
        "H": {
            "header": "Ban/Phòng/Bộ phận",
            "python_name": "department",
            "type": "text",
            "required": False,
            "usage": "Department/Section"
        },
        "I": {
            "header": "Ngân hàng",
            "python_name": "bank",
            "type": "text",
            "required": False,
            "usage": "Bank name"
        },
        "J": {
            "header": "Mức lương Gross",
            "python_name": "gross_salary",
            "type": "number",
            "required": False,
            "usage": "Gross salary"
        },
        "K": {"header": "Công chuẩn", "python_name": "standard_hours", "type": "number"},
        "L": {"header": "Công làm việc thực tế", "python_name": "actual_hours", "type": "number"},
        "M": {"header": "Số ngày nghỉ có hưởng lương", "python_name": "paid_leave_days", "type": "number"},
        "N": {"header": "Tổng công tính lương", "python_name": "total_hours_salary", "type": "number"},
        "O": {"header": "Lương cơ bản + Thưởng YTCLCV", "python_name": "base_salary", "type": "number"},
        "P": {"header": "Thưởng khoán/KPI", "python_name": "kpi_bonus", "type": "number"},
        "Q": {"header": "Hỗ trợ ăn ca", "python_name": "meal_allowance", "type": "number"},
        "R": {"header": "Tiền làm thêm ngoài giờ", "python_name": "overtime", "type": "number"},
        "S": {"header": "Trợ cấp làm đêm", "python_name": "night_shift_allowance", "type": "number"},
        "T": {"header": "Thưởng dự án", "python_name": "project_bonus", "type": "number"},
        "U": {"header": "Hỗ trợ chi phí đi lại/gửi xe", "python_name": "transport_allowance", "type": "number"},
        "V": {"header": "Thâm niên", "python_name": "seniority_bonus", "type": "number"},
        "W": {"header": "Hỗ trợ nghỉ việc", "python_name": "separation_allowance", "type": "number"},
        "X": {"header": "Truy thu/Truy lĩnh chịu thuế", "python_name": "taxable_recovery", "type": "number"},
        "Y": {"header": "Truy thu/Truy lĩnh không chịu thuế", "python_name": "non_taxable_recovery", "type": "number"},
        "Z": {"header": "Khấu trừ khác", "python_name": "other_deductions", "type": "number"},
        "AA": {"header": "Khấu trừ đồng phục", "python_name": "uniform_deduction", "type": "number"},
        "AB": {"header": "Trích nộp Bảo hiểm bắt buộc", "python_name": "insurance_deduction", "type": "number"},
        "AC": {"header": "Trích nộp phí đoàn viên Công đoàn", "python_name": "union_fee", "type": "number"},
        "AD": {"header": "THUẾ TNCN", "python_name": "income_tax", "type": "number"},
        "AE": {"header": "Tổng thu nhập chịu thuế", "python_name": "taxable_income", "type": "number"},
        "AF": {"header": "Tổng giảm trừ gia cảnh", "python_name": "family_deduction", "type": "number"},
        "AG": {"header": "Tổng thu nhập tính thuế", "python_name": "taxable_amount", "type": "number"},
        "AH": {"header": "THỰC LĨNH", "python_name": "net_payment", "type": "number"},
        "AZ": {
            "header": "PassWord",
            "python_name": "password",
            "type": "text",
            "required": True,
            "usage": "PDF password (strip leading zeros from employee_id)"
        }
    },
    
    "tbkq_template_mapping": {
        "NOTE": "B3 is the trigger cell that must be set to MNV. This should trigger VLOOKUP formulas in the template.",
        "B3": {
            "description": "Employee MNV - TRIGGER CELL for all VLOOKUPs",
            "input": "Data.A (employee_id)",
            "type": "text",
            "required": True
        },
        "NOTE_2": "Other cells in TBKQ will be populated by VLOOKUP formulas. We need to manually identify each once you run the template in Excel."
    },
    
    "email_configuration": {
        "subject_cell": "G1",
        "subject_source": "TBKQ.G1",
        "body_cells": ["A1", "A3", "A5", "A7", "A8", "A9", "A11", "A12"],
        "body_sheet": "bodymail",
        "date_replacement": {
            "placeholder_pattern": r"\\d{1,2}/\\d{4}",
            "replacement_value": "${DATE}",
            "cell_to_replace": "A3",
            "date_format": "MM/YYYY"
        }
    }
}

print(json.dumps(complete_mapping, indent=2, ensure_ascii=False))

# Save to JSON file
with open('cell_mapping.json', 'w', encoding='utf-8') as f:
    json.dump(complete_mapping, f, indent=2, ensure_ascii=False)

print("\n✓ Mapping saved to cell_mapping.json")

wb.release_resources()

# Create sample .env file
print("\n" + "=" * 140)
print("GENERATED .env FILE")
print("=" * 140)

env_content = """# Payslip Tool Configuration for Phuc Long

# Excel File Configuration
PAYSLIP_EXCEL_PATH=./excel-files/TBKQ-phuclong.xls

# Sheet Names (configurable for flexibility)
DATA_SHEET=Data
TEMPLATE_SHEET=TBKQ
EMAIL_BODY_SHEET=bodymail
SALARY_DATA_SHEET=bang luong

# Data Sheet Column Mappings
# These define which columns in Data sheet contain which information
DATA_COLUMN_MNV=A
DATA_COLUMN_NAME=B
DATA_COLUMN_EMAIL=C
DATA_COLUMN_JOB_TITLE=D
DATA_COLUMN_ACCOUNT_NUMBER=E
DATA_COLUMN_START_DATE=F
DATA_COLUMN_DEPARTMENT=H
DATA_COLUMN_BANK=I
DATA_COLUMN_GROSS_SALARY=J
DATA_COLUMN_PASSWORD=AZ
DATA_COLUMN_NET_PAYMENT=AH

# TBKQ Template Cell Mappings
# These define which cells in the TBKQ template should be populated
TBKQ_TRIGGER_CELL=B3

# Email Configuration
EMAIL_SUBJECT_CELL=G1
EMAIL_BODY_CELLS=A1,A3,A5,A7,A8,A9,A11,A12
EMAIL_DATE_PLACEHOLDER=mm/yyyy
EMAIL_DATE_REPLACEMENT_CELL=A3

# Payroll Date (format: MM/YYYY)
DATE=01/2026

# Outlook Configuration
OUTLOOK_ACCOUNT=your.email@company.com

# PDF Configuration
PDF_PASSWORD_ENABLED=true
PDF_PASSWORD_STRIP_LEADING_ZEROS=true
PDF_PASSWORD_SOURCE=employee_id

# Output Configuration
OUTPUT_DIR=./payslips
PDF_FILENAME_PATTERN=TBKQ_{employee_name}_{mmyyyy}.pdf

# Processing Options
DRY_RUN=false
BATCH_SIZE=10
PARALLEL_WORKERS=5
MAX_RETRIES=3

# Logging
LOG_LEVEL=INFO
LOG_FILE=./payslips/payslip_send.log

# State Tracking (for idempotency)
STATE_FILE=./payslips/sent_state.json
"""

with open('tools/payslip-phuclong/.env.example', 'w') as f:
    f.write(env_content)

print(env_content)
print("\n✓ Sample .env saved to tools/payslip-phuclong/.env.example")

