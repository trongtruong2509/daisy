"""
Generate batch test data for payslip tool.

Copies TBKQ-phuclong.xls to a new test file and populates it with 200+ employees.

Structure:
- bang luong sheet: Master employee data (MNV in column L, Name in column M)
- Data sheet: Mapping with MNV in column A, formulas pulling from bang luong, Email in column C
"""

import random
import sys
import shutil
from pathlib import Path

# Ensure project root is importable
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from office.utils.com import is_available, get_win32com_client

if not is_available():
    print("ERROR: pywin32 not available. Required for .xls files.")
    print("Install with: pip install pywin32")
    sys.exit(1)


# Vietnamese names for generating realistic test data
FIRST_NAMES = [
    "Nguyễn", "Trần", "Hoàng", "Phạm", "Huỳnh", "Đinh", "Bùi", "Đặng", "Vũ", "Võ",
    "Dương", "Lý", "Phan", "Tô", "Tạ", "Lê", "Đỗ", "Đào", "Từ", "Lâm"
]

MIDDLE_NAMES = [
    "Văn", "Thái", "Hữu", "Huy", "Hoài", "Kiến", "Minh", "Thanh", "Hùng", "Tiến",
    "Đình", "Anh", "Hải", "Duy", "Phong", "Tuấn", "Trí", "Quang", "Nhân", "Đức"
]

LAST_NAMES = [
    "A", "B", "C", "D", "E", "Sơn", "Hải", "Hào", "Hân", "Hiền",
    "Hồng", "Hương", "Huy", "Huỳnh", "Huỳnh", "Hùng", "Hạnh", "Hải", "Hân", "Hiền"
]

TEST_EMAILS = ["tht.tts@gmail.com", "trong.h.truong@gmail.com"]


def generate_vietnamese_name():
    """Generate a random Vietnamese name."""
    first = random.choice(FIRST_NAMES)
    middle = random.choice(MIDDLE_NAMES)
    last = random.choice(LAST_NAMES)
    return f"{first} {middle} {last}"


def generate_batch_test_file(source_file: Path, output_file: Path, num_employees: int = 200):
    """
    Generate batch test file with employees in both bang luong and Data sheets.
    
    Uses batch writes for better performance.
    """
    if not source_file.exists():
        print(f"ERROR: Source file not found: {source_file}")
        return False
    
    print(f"[1/4] Copying Excel file...")
    shutil.copy2(source_file, output_file)
    print(f"      ✓ Copied")
    
    print(f"[2/4] Opening Excel file via COM...")
    excel = None
    workbook = None
    try:
        excel = get_win32com_client().DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False  # Disable screen updates for speed
        
        workbook = excel.Workbooks.Open(str(output_file), UpdateLinks=0)
        bang_luong = workbook.Sheets("bang luong")
        data_sheet = workbook.Sheets("Data")
        print(f"      ✓ Opened")
    except Exception as e:
        print(f"      ✗ Failed to open: {e}")
        return False
    
    try:
        print(f"[3/4] Populating {num_employees} employees...")
        
        # Column positions
        col_l = 12  # MNV in bang luong
        col_m = 13  # Name in bang luong
        col_a = 1   # MNV in Data
        col_c = 3   # Email in Data
        
        # bang luong: Data starts at row 3
        # Data sheet: Data starts at row 4
        bl_start_row = 3
        data_start_row = 4
        
        # Pre-generate all data first
        print(f"      - Pre-generating {num_employees} rows...")
        mnv_counter = 6046073
        bl_data = []
        data_mnv = []
        data_email = []
        
        for i in range(num_employees):
            mnv = mnv_counter + i
            name = generate_vietnamese_name()
            email = random.choice(TEST_EMAILS)
            
            bl_data.append([mnv, name])
            data_mnv.append([mnv])
            data_email.append([email])
        
        # Write bang luong data (column L:M)
        print(f"      - Writing bang luong data...")
        bl_range = bang_luong.Range(
            f"L{bl_start_row}:M{bl_start_row + num_employees - 1}"
        )
        bl_range.Value = bl_data
        
        # Write Data sheet MNV (column A)
        print(f"      - Writing Data sheet MNV...")
        data_mnv_range = data_sheet.Range(
            f"A{data_start_row}:A{data_start_row + num_employees - 1}"
        )
        data_mnv_range.Value = data_mnv
        
        # Write Data sheet Email (column C)
        print(f"      - Writing Data sheet Email...")
        data_email_range = data_sheet.Range(
            f"C{data_start_row}:C{data_start_row + num_employees - 1}"
        )
        data_email_range.Value = data_email
        
        print(f"      ✓ Generated {num_employees} employees")
        
        # Count email distribution
        email_count = {email: data_email.count([email]) for email in TEST_EMAILS}
        flat_emails = [e[0] for e in data_email]
        email_count = {email: flat_emails.count(email) for email in TEST_EMAILS}
        print(f"      - Email distribution:")
        for email, count in email_count.items():
            pct = (count / num_employees) * 100
            print(f"        • {email}: {count} ({pct:.1f}%)")
        
    finally:
        excel.ScreenUpdating = True  # Re-enable screen updates
        print(f"[4/4] Saving file...")
        try:
            workbook.Save()
            print(f"      ✓ Saved")
        except Exception as e:
            print(f"      ✗ Save failed: {e}")
            return False
        finally:
            try:
                workbook.Close(SaveChanges=False)
            except:
                pass
            try:
                excel.Quit()
            except:
                pass
    
    print(f"\n✅ SUCCESS: {output_file}")
    print(f"   Employees: {num_employees}")
    print(f"   MNV range: {mnv_counter}-{mnv_counter + num_employees - 1}")
    print(f"   bang luong: rows {bl_start_row}-{bl_start_row + num_employees - 1}")
    print(f"   Data sheet: rows {data_start_row}-{data_start_row + num_employees - 1}")
    print(f"   Emails: {', '.join(TEST_EMAILS)}")
    return True


if __name__ == "__main__":
    excel_files_dir = Path(__file__).parent.parent.parent.parent / "excel-files"
    source_file = excel_files_dir / "TBKQ-phuclong.xls"
    output_file = excel_files_dir / "TBKQ-phuclong-batch-200.xls"
    
    print("=" * 80)
    print("BATCH TEST DATA GENERATOR")
    print("=" * 80)
    print(f"Source: {source_file}")
    print(f"Output: {output_file}")
    print("=" * 80)
    
    success = generate_batch_test_file(source_file, output_file, num_employees=210)
    sys.exit(0 if success else 1)
