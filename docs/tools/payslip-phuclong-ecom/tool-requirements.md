## Tool Requirements — Payslip Generator & Email Sender (`payslip-phuclong-ecom`)

### Overview

This tool automates the monthly payslip generation and distribution workflow for Phuc Long employees. It reads employee payroll data from a shared Excel workbook, generates individual payslip files via Excel COM, converts them to password-protected PDFs, and sends personalised emails via Outlook.

---

### Functional Requirements

#### REQ-PAY-01: Read employee data from the source Excel workbook

The tool must read the following fields from the configured data sheet (`DATA_SHEET`, default `"Data"`) using `ExcelComReader`:

| Column (configurable)         | Field               | Notes                                                                  |
| ----------------------------- | ------------------- | ---------------------------------------------------------------------- |
| `COL_MNV` (default `A`)       | Employee ID (`mnv`) | Required; used as filename key and PDF password base                   |
| `COL_NAME` (default `B`)      | Full name           | Required                                                               |
| `COL_EMAIL` (default `C`)     | Email address       | Required; must pass email format validation                            |
| `COL_PASSWORD` (default `AZ`) | PDF password        | Required; stripped of leading zeros if `PDF_PASSWORD_STRIP_ZEROS=true` |

Data rows start at `DATA_START_ROW` (default 4), with headers at `DATA_HEADER_ROW` (default 2). Rows where all configured columns are empty must be skipped.

#### REQ-PAY-02: Validate all employee data before generation — fail fast

After reading, all employee records must be validated by `DataValidator` before any file is generated or email is sent. Validation must check:

- `mnv` is not empty.
- `name` is not empty.
- `email` is a valid email address.
- `password` is not empty.
- Duplicate emails are flagged as errors unless `ALLOW_DUPLICATE_EMAILS=true`.

If any validation error exists, the tool must print all errors via `cprint(..., level="ERROR")` and exit with code 1. Warnings (non-blocking) must be shown but must not stop execution.

#### REQ-PAY-03: Generate payslip XLSX files via the VBA macro approach

For each employee, `PayslipGenerator` must:

1. Open the source workbook with Excel COM.
2. Set `TBKQ!B3 = MNV` to trigger VLOOKUP / XLOOKUP formula recalculation.
3. Force a full recalculation (`CalculateFull()`).
4. Copy the `TBKQ` sheet to a new workbook.
5. Paste as values (removing all formulas).
6. Save the new workbook as `.xlsx` in the configured `OUTPUT_DIR`.

This replicates the original VBA macro exactly and must not attempt to evaluate formula logic in Python.

#### REQ-PAY-04: Output filename pattern is configurable

The output filename must follow the pattern set by `PDF_FILENAME_PATTERN` (default `"TBKQ_{name}_{mmyyyy}"`). Template variables `{name}` and `{mmyyyy}` are substituted at runtime. Characters unsafe for Windows filenames in the employee name must be replaced with `_`.

When two employees share the same name, a numeric suffix (`_1`, `_2`, …) must be appended to disambiguate filenames.

#### REQ-PAY-05: Convert XLSX to password-protected PDF

After all XLSX files are generated, `PdfConverter` must convert each file to PDF with:

- Password derived from the employee's `password` field.
- Leading zeros stripped when `PDF_PASSWORD_STRIP_ZEROS=true` (default).
- Source XLSX deleted after successful conversion when `cleanup_xlsx=True` (default).

If `PDF_PASSWORD_ENABLED=false`, the PDF must be generated without a password.

#### REQ-PAY-06: Send personalised emails via `OutlookSender`

Each employee must receive one email with:

- **Subject**: read from `TEMPLATE_SHEET!EMAIL_SUBJECT_CELL` (default `TBKQ!G1`) or overridden by `EMAIL_SUBJECT` in `.env`.
- **Body**: composed from configurable cells in `EMAIL_BODY_SHEET` (default `bodymail`). The date placeholder in the body must be substituted with `DATE` from config.
- **Attachment**: the employee's PDF payslip.
- **From**: `OUTLOOK_ACCOUNT`.

The email body source cells are configurable via `EMAIL_BODY_CELLS` (default: `A1, A3, A5, A7, A9, A11, A12`).

#### REQ-PAY-07: Batch processing with configurable size

Employees must be processed in batches of at most `BATCH_SIZE` (default 50). Between batches, Excel COM must be released and garbage collected to prevent COM memory exhaustion. A `time.sleep(1)` must be observed between batches to allow COM resource cleanup.

#### REQ-PAY-08: Resume after crash via `StateTracker`

Before sending each email, the tool must check `StateTracker.is_processed(employee_id)`. After a successful send (or dry-run), it must call `mark_processed()`. On restart, already-processed employees must be skipped automatically.

At startup, if a state file exists the tool must present an interactive menu:

- **Resume** — continue from the last checkpoint.
- **Start fresh** — clear all state and output files and restart.
- **Exit**.

#### REQ-PAY-09: Dry-run mode is the default

`DRY_RUN` defaults to `true`. In dry-run mode, payslip files must still be generated and converted to PDF, but no email is sent via `OutlookSender`. The tool must clearly label each log line and console output with `[DRY RUN]`.

#### REQ-PAY-10: Pre-run configuration summary and confirmation

Before processing begins, the tool must display a summary box (via `cprint_summary_box_lite()`) showing:

- Excel file name, payroll date, number of employees, Outlook account, dry-run status, PDF password status, output directory.

In live mode (`DRY_RUN=false`), the user must explicitly confirm before emails are sent.

#### REQ-PAY-11: End-of-run summary

After all processing the tool must display a summary box (via `cprint_summary_box()`) with:

- Total employees, successfully generated payslips, successfully sent emails, skipped (already processed), errors.

Exit codes: `0` on full success or dry-run completion, `1` on partial or full failure.

---

### Configuration Reference

All settings can be placed in `tools/payslip-phuclong-ecom/.env`. Absent values are prompted interactively at startup.

| Key                        | Default    | Description                                          |
| -------------------------- | ---------- | ---------------------------------------------------- |
| `PAYSLIP_EXCEL_PATH`       | (prompted) | Path to the source Excel file                        |
| `DATE`                     | (prompted) | Payroll month in `MM/YYYY` format                    |
| `OUTLOOK_ACCOUNT`          | (prompted) | SMTP address of the sending account                  |
| `DRY_RUN`                  | `true`     | When `true`, no emails are sent                      |
| `BATCH_SIZE`               | `50`       | Employees per processing batch                       |
| `DATA_SHEET`               | `Data`     | Sheet name for employee data                         |
| `TEMPLATE_SHEET`           | `TBKQ`     | Sheet name for the payslip template                  |
| `EMAIL_BODY_SHEET`         | `bodymail` | Sheet name for the email body content                |
| `PDF_PASSWORD_ENABLED`     | `true`     | Apply password to generated PDFs                     |
| `PDF_PASSWORD_STRIP_ZEROS` | `true`     | Strip leading zeros from PDF password                |
| `KEEP_PDF_PAYSLIPS`        | `false`    | Keep PDF files after sending                         |
| `ALLOW_DUPLICATE_EMAILS`   | `false`    | Allow multiple employees with the same email address |
| `LOG_LEVEL`                | `INFO`     | Logging verbosity                                    |

---

### Non-functional Constraints

- The tool is Windows-only; `pywin32` and `pikepdf` are required dependencies.
- No COM object must be passed between threads. Each batch runs serially in a single thread.
- State and output directories are created automatically.
- The tool must exit gracefully if Outlook is not running or the account is not found.
- Integration tests must use `@pytest.mark.integration` and are excluded from the default run.
