# Batch Test Data

This directory contains batch test files for testing the payslip tool with larger datasets.

## Files

### TBKQ-phuclong-batch-200.xls
- **Employees**: 210 (rows 4-213)
- **MNV Range**: 6046073 to 6046282
- **Email Distribution**:
  - 50% assigned to `tht.tts@gmail.com`
  - 50% assigned to `trong.h.truong@gmail.com`
- **Generated**: 2026-02-08 by `generate_batch_test_data.py`

### Original File
- **TBKQ-phuclong.xls**: Contains only 1 employee (Nguyễn Văn A, MNV: 6046072)

## How to Test

### Option 1: Update .env File
Update `.env` in `tools/payslip-phuclong-ecom/` to use the batch file:

```env
PAYSLIP_EXCEL_PATH=..\..\excel-files\TBKQ-phuclong-batch-200.xls
```

### Option 2: Use Environment Variable
Set the file path before running:

```powershell
$env:PAYSLIP_EXCEL_PATH = "..\..\excel-files\TBKQ-phuclong-batch-200.xls"
python main.py
```

### Option 3: Use Template Config
Copy the template config:

```powershell
Copy-Item .env.batch-test .env
```

## Testing Account Selection

The batch file has employees randomly assigned to two email addresses:

- **tht.tts@gmail.com** (test account 1)
- **trong.h.truong@gmail.com** (test account 2)

When you run payslip processing with the batch file:

1. The tool will prompt for `OUTLOOK_ACCOUNT` (first time)
2. Select account [2] (`trong.h.truong@gmail.com`)
3. All 210 payslips will be sent
4. Check the logs to verify account usage
5. Check received emails to confirm they came from the correct sender

## Expected Log Output

With account [2] selected, you should see:

```
[ACCOUNT-SEARCH] Account [2] = 'trong.h.truong@gmail.com' ✓ MATCH
[ACCOUNT-SELECTED] ✓ Found requested account: trong.h.truong@gmail.com
[SEND-METHOD] Account object valid: trong.h.truong@gmail.com
[SEND-VERIFY] Mail item's SendUsingAccount = 'trong.h.truong@gmail.com'  ← Key line!
[SENT] Email sent - To: tht.tts@gmail.com, Subject: ...
```

The **[SEND-VERIFY]** line should show the correct account being used.

## Statistics

After running with the batch file, check the output summary:

```
Payslip processing complete: {
  'total': 210,
  'generated': 210,
  'converted': 210,
  'sent': 210,
  'errors': 0,
  ...
}
```

This indicates all 210 payslips were successfully processed and sent.

## Regenerating Test Data

To regenerate the batch test file with different data:

```powershell
cd tools\payslip-phuclong-ecom
python scripts\generate_batch_test_data.py
```

The script supports customization of:
- Number of employees
- Email addresses (modify `TEST_EMAILS` list)
- Vietnamese name data (modify name lists)
- Column positions (modify `col_a`, `col_b`, etc.)

## Notes

- The password field is copied from the original employee record
- Employee MNV values are sequential to avoid duplicates
- Vietnamese names are randomly generated from realistic name lists
- Email addresses are evenly distributed between the two test accounts
