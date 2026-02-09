# Payslip Email Automation Tool

**Project:** Payslip Phuc Long
**Location:** `<workspace>/tools/payslip-phuclong`
**Version:** 1.0

---

# 1. Purpose

This document defines the functional and non-functional requirements for the Payslip Email Automation Tool.

The purpose of the tool is to:

- Generate individual password-protected PDF payslips from an Excel template.
- Send personalized payslip emails to employees via Outlook.
- Ensure safe, idempotent, validated, and auditable email distribution.

This specification is intended for the QA/Test team to design and execute test cases.

---

# 2. Scope

The tool:

- Reads employee and payroll data from a configured Excel file.
- Generates PDF payslips using a Excel template via COM automation.
- Convert to PDF via Excel COM automation

- Sends email using a specified Outlook account.
- Logs execution details and provides execution summary.
- Prevents duplicate sending.
- Supports DRY_RUN mode.

The tool must support 2000+ employees in production.

---

# 3. System Configuration

All configurable values must be defined in:

- Global `.env`
- Local `.env` (tool-specific override)

Local `.env` takes precedence over global `.env`.

**Configuration Management:**

- Uses `core/config_manager.py` (`ConfigManager` class) for centralized configuration handling
- Provides typed getters: `get()`, `get_bool()`, `get_int()`, `get_path()`, `get_list()`
- Built-in validators: `validate_date()`, `validate_email()`, `validate_file_path()`
- Interactive prompting: `prompt_for_value()` with validation
- Persistence: `save_to_env()` for saving prompted values

---

# 4. Functional Requirements

---

## 4.1 Input Configuration Requirements

### FR-01: Excel Path

- Excel file path must be provided via:

  ```
  PAYSLIP_EXCEL_PATH=<path>
  ```

- If not found or invalid → script will prompt user interactively for the path.
- Interactive prompt includes validation to ensure file exists.
- User can provide relative path (resolved from tool directory) or absolute path.
- After successful prompt, value can be saved to `.env` for future use.

---

### FR-02: Sheet Configuration (All configurable in `.env`)

The following must be configurable:

| Item                                        | Configurable via .env |
| ------------------------------------------- | --------------------- |
| Data sheet name                             | YES                   |
| TBKQ template sheet name                    | YES                   |
| Email body sheet name                       | YES                   |
| Column letters (MNV, Name, Email, Password) | YES                   |
| Email body cell locations                   | YES                   |

No hardcoded sheet names or column letters are allowed.

---

### FR-03: DATE Configuration

- `.env` variable:

  ```
  DATE=MM/YYYY
  ```

- Must validate format (MM/YYYY where MM is 01-12).
- Used for:
  - Email body replacement
  - PDF naming
  - Output directory structure

- If not set → script will prompt user interactively.
- Interactive prompt includes format validation.
- After successful prompt, value can be saved to `.env` for future use.

Invalid format → re-prompt or terminate if non-interactive.

---

### FR-04: Outlook Account Selection

- `.env` must define:

  ```
  OUTLOOK_ACCOUNT=<email>
  ```

- Script must send emails ONLY from this account.
- If account not found in Outlook profile → terminate.

**Implementation:**

- Uses `office.outlook.client.get_outlook_accounts()` to enumerate available accounts
- If not configured in `.env` → script presents interactive menu:
  - Lists all available Outlook accounts (if any detected)
  - User selects by number
  - If no accounts detected, prompts for manual email entry
- Validates selected/entered email format
- After successful selection, value can be saved to `.env` for future use
- Validates account exists before execution begins

---

### FR-05: Email Subject Customization

- `.env` variable:

  ```
  EMAIL_SUBJECT=<string>
  ```

- Must support DATE placeholder replacement if applicable.

---

### FR-06: DRY_RUN Mode

- `.env`:

  ```
  DRY_RUN=true|false
  ```

- If true:
  - No actual email is sent.
  - System simulates sending.
  - Logs and summary behave normally.

---

### FR-06a: PDF Cleanup Configuration

- `.env`:

  ```
  KEEP_PDF_PAYSLIPS=true|false
  ```

- If false (default):
  - PDF files are deleted after successful email send.
  - Saves disk space for large batches.
- If true:
  - PDF files are retained in output directory.

---

# 5. Data Processing Requirements

---

## 5.1 Employee Source of Truth

### FR-07: Employee Source

- Employee list MUST be read from `Data` sheet.
- `bang luong` sheet must be ignored.

---

## 5.2 Data Sheet Structure

### FR-08: Data Sheet Layout

- Row 2 = Header row.
- Row 4+ = Employee data.
- Row 3 numeric codes are irrelevant.
- Script must start reading employees from row 4.

---

## 5.3 Column Mapping

Columns are configurable via `.env`, example:

```
COLUMN_MNV=A
COLUMN_NAME=B
COLUMN_EMAIL=C
COLUMN_PASSWORD=AZ
```

Script must:

- Dynamically read based on config.
- Fail if configured column does not exist.

---

# 6. Validation Requirements (Pre-Send Validation Phase)

## 6.1 Global Validation Rule

### FR-09: All-or-Nothing Policy

Before sending any email:

- Validate ALL employees.
- If ANY critical error exists → terminate.
- No emails should be sent if validation fails.

---

## 6.2 Email Validation

### FR-10:

- Email must not be empty.
- Must match valid email regex.
- Duplicate emails across employees must be detected.
- If duplicates found → terminate execution.

Invalid email rows must be skipped only if:

- Email field empty or invalid format.

However, duplicate emails are treated as fatal error.

---

## 6.3 Name Validation

### FR-11:

- Name must be non-empty.
- No advanced name format validation required.

---

# 7. Payslip Generation Requirements

---

## 7.1 Generation Strategy

### FR-12: Excel COM–Based Generation

The system must use **Microsoft Excel COM automation** for:

- Opening the source Excel file
- Copying or referencing the `TBKQ` template sheet
- Populating employee-specific data
- Triggering Excel recalculation
- Generating the employee-specific Excel payslip file
- Converting the payslip to PDF

`openpyxl` must NOT be used for template population.

Excel application must be installed on the machine executing the script.

If Excel COM initialization fails → execution must terminate.

---

## 7.2 Template Handling

### FR-13: TBKQ Template Usage

- The payslip template sheet name must be configurable via `.env`.
- For each employee:
  - The system must create a new Excel payslip file based on the `TBKQ` template.
  - The system must set the employee identifier (MNV) into the designated cell (configurable via `.env`, e.g., `TBKQ_MNV_CELL=B3`).
  - Excel formulas inside the template must calculate automatically.

The tool must rely on Excel’s internal formulas and recalculation engine.

---

## 7.3 Data Population Mechanism

### FR-14: Formula-Driven Population

- The `TBKQ` template contains formulas (e.g., VLOOKUP) that reference the `Data` sheet.
- The script must:
  1. Set the configured MNV cell (e.g., B3).
  2. Trigger recalculation (if needed).
  3. Ensure all dependent cells update correctly.

The script must not manually map or copy salary component cells.

If recalculation fails → employee processing must be marked as failure.

---

## 7.4 Excel File Generation

### FR-15: Intermediate Excel File

For each employee:

- The system must generate an Excel file before PDF conversion.
- File naming format (configurable pattern, default):

```
TBKQ_<Name>_<mmyyyy>.xlsx
```

Example:

```
TBKQ_NguyenVanA_112025.xlsx
```

The intermediate Excel file may be:

- Stored temporarily and deleted after PDF generation, OR
- Persisted (implementation-specific)

Behavior must be consistent and documented in README.

**Output Directory Structure:**

- All output files are organized in date-based subdirectories:
  ```
  output/
    <mmyyyy>/
      TBKQ_Employee1_mmyyyy.pdf
      TBKQ_Employee2_mmyyyy.pdf
      sent_results_mmyyyy.csv
  ```
- This prevents file conflicts when processing multiple payroll periods.

---

## 7.5 PDF Conversion

### FR-16: PDF Export

- The generated Excel payslip must be converted to PDF using Excel COM:
  - `ExportAsFixedFormat` method.

- PDF file naming format:

```
TBKQ_<Name>_<mmyyyy>.pdf
```

If PDF export fails → employee must be marked as failure.

---

## 7.6 Password Protection

### FR-17: PDF Password Protection

- After PDF generation, the file must be password-protected.
- Password source:
  - Data sheet column defined in `.env` (default AZ).

- Password equals employee MNV with leading zeros stripped.

If password value is missing → execution must terminate.

If password protection fails → employee must be marked as failure.

**PDF Lifecycle Management:**

- After successful email send, PDFs may be automatically deleted (configurable via `KEEP_PDF_PAYSLIPS=false`)
- This saves disk space when processing large employee counts (1000+)
- Failed or skipped emails retain their PDFs for retry/investigation

---

## 7.7 Excel Resource Management

### FR-18: COM Resource Handling

The system must:

- Properly close:
  - Workbook objects
  - Excel Application instance

- Release COM objects
- Prevent orphaned Excel processes

After execution, no background Excel processes should remain.

Failure to release COM resources is considered a defect.

---

## 7.8 Performance Considerations

### FR-19: Large Dataset Handling

For 1000+ employees:

- Excel must not remain open per employee unnecessarily.
- System must avoid opening/closing Excel per employee unless required.
- COM usage must be optimized to prevent excessive memory consumption.

---

## 7.9 Error Handling During Generation

The system must detect and handle:

- Missing TBKQ sheet
- Missing MNV cell
- Excel recalculation failure
- ExportAsFixedFormat failure
- File write permission issues
- COM exceptions

Errors must:

- Be logged
- Be reflected in execution summary

---

# 8. Email Composition Requirements

---

## 8.1 Email Body

### FR-20:

- Email body read from exact configured cells (e.g. A1, A3, etc).
- Cell locations configurable in `.env`.
- DATE placeholder must be replaced dynamically.

---

## 8.2 Attachment

### FR-21:

- Each email must contain:
  - Exactly one attachment.
  - Password-protected PDF.

- If attachment missing → mark as failure.

---

# 9. Execution Flow Requirements

---

## 9.1 Pre-Execution Summary

### FR-22:

Before confirmation:

- Display total number of valid emails to be sent.
- DO NOT display employee list.

---

## 9.2 Confirmation

### FR-23:

- User must explicitly confirm before sending.
- If user declines → terminate safely.

---

## 9.3 Post-Execution Summary

### FR-24:

After execution, display:

- Total employees processed
- Successfully sent
- Failed
- Skipped
- Execution time

---

# 10. Logging Requirements

---

## FR-25: Logging

The system must:

- Log each email attempt:
  - Timestamp
  - Employee MNV
  - Email
  - Status (Success / Failure / Skipped)
  - Error message (if any)

- Log to:
  - Console (colored, user-facing output via custom CONSOLE level)
  - Log file (all levels including CONSOLE=25)

- Logging implementation:
  - Uses `core/console.py` with `cprint()` for dual console+file output
  - Custom CONSOLE log level (25) between INFO and WARNING
  - All user-facing messages automatically logged to file
  - Supports structured logging levels: INFO, PHASE, SUCCESS, ERROR, WARNING, BANNER, SUMMARY

- Result tracking:
  - CSV file format: `sent_results_<mmyyyy>.csv`
  - Columns: employee_id, employee_name, email_address, payslip_filename, sent_status, timestamp, error_message
  - Located in date-specific subdirectory: `output/<mmyyyy>/`

---

# 11. Idempotency Requirements

---

## FR-26: No Duplicate Sending

The tool must ensure:

- If an employee has already received payslip for the same DATE:
  - Email must NOT be sent again.

- System must track sending state (implementation-specific).
- Re-running script must not resend previously successful emails.

---

# 12. Error Handling Requirements

System must handle:

- Missing Excel file
- Missing sheet
- Missing column
- Invalid DATE
- Outlook connectivity failure
- PDF generation failure
- Email send failure

Errors must:

- Be logged
- Be reflected in summary

---

# 13. Testing Requirements

---

## 13.1 Unit Tests

### TR-01:

Must cover:

- Email sending logic (mocked)
- Attachment generation
- Email validation
- Duplicate detection
- Idempotency behavior
- DRY_RUN behavior
- Error scenarios
- Configuration loading and validation
- CSV result writer
- State management (checkpoint tracking)
- Helper function isolation (enabled by modular refactoring)

**Testability Enhancements:**

- `main.py` refactored into discrete helper functions:
  - `load_and_validate_config()`
  - `read_employee_data()`
  - `validate_employee_data()`
  - `generate_payslips()`
  - `convert_to_pdf()`
  - `compose_emails()`
  - `send_emails()`
- Each function can be unit tested independently
- Foundation modules (`core/console`, `core/config_manager`, `office/excel`) are independently testable

Must use:

- Mock Outlook
- Mock Excel COM interactions
- Mock file system operations

No real file or email interaction.

---

## 13.2 Integration Tests

### TR-02:

Separate test suite:

- Uses real Excel file (sample file with 2 employees).

- Uses real Outlook account.

- Clearly marked as:

  ```
  Integration Tests
  ```

- Executed manually.

---

# 14. Performance Requirements

- Must handle 1000+ employees.
- Execution time under acceptable operational threshold (exact SLA not specified).
- Memory usage must not grow unbounded.

---

# 15. Documentation Requirements

### FR-27: README.md

Must include:

- Setup instructions
- .env configuration explanation (including new options: KEEP_PDF_PAYSLIPS)
- Required dependencies
- How to run
- How DRY_RUN works
- How idempotency works
- How to run integration tests
- Date-based output directory structure
- CSV result file format specification

---

# 16. Architecture & Code Organization

**Modular Foundation Layer:**

The tool leverages reusable foundation modules from `daisy/`:

- `core/console.py` - Standardized colored console output with automatic file logging
- `core/config_manager.py` - Centralized configuration management
- `core/logger.py` - Dual console+file logging with custom CONSOLE level (25)
- `core/state.py` - State tracking for idempotency
- `office/excel/` - Excel COM abstraction layer (reader, converter, utilities)
- `office/outlook/` - Outlook COM abstraction layer

**Tool Structure:**

- `main.py` - Orchestrator with helper functions for each phase (60 lines, clean workflow)
- `config.py` - Tool-specific configuration using ConfigManager
- `excel_reader.py` - Payslip-specific Excel reading logic
- `payslip_generator.py` - Excel COM payslip generation
- `email_composer.py` - Email body composition from template
- `validator.py` - Employee data validation
- `utils.py` - Progress tracking, state management, result writing (CSV)

**Key Design Principles:**

- Separation of concerns - each module has single responsibility
- Reusable foundation - common utilities extracted to core/
- Testable components - helper functions enable unit testing
- Fail-fast validation - all data validated before processing
- Idempotent execution - state tracking prevents duplicate sends

---

# 17. Out of Scope

- Advanced name validation
- Payroll recalculation logic
- Editing Excel formulas
- UI application (CLI only)

---

# 18. Excel Abstraction Layer

**Implementation Details:**

The system uses a reusable Excel COM abstraction layer (`office/excel/`):

- **ExcelComReader** - Generic Excel file reading with context manager
  - Methods: `read_cell()`, `read_cells()`, `read_range()`, `get_sheet()`, `recalculate()`
  - Proper COM lifecycle management (prevents orphaned Excel processes)

- **PdfConverter** - Excel to PDF conversion with password protection
  - Uses `ExportAsFixedFormat` COM method
  - Supports password protection via `pikepdf`
  - Batch processing with progress callbacks

- **Utilities** (`office/excel/utils.py`):
  - `col_letter_to_index()` / `index_to_col_letter()` - Column conversion
  - `safe_cell_value()` - Handles COM error values (#N/A, #VALUE!, etc.)
  - `normalize_numeric_string()` - Numeric string cleanup
  - `xlookup_to_index_match()` - XLOOKUP → INDEX/MATCH conversion (Excel 2019 compatibility)

This abstraction enables:

- Code reuse across multiple tools
- Easier unit testing (mock at abstraction boundary)
- Consistent COM resource management
- Excel version compatibility handling

---

# 19. Test Data Clarification

- Provided Excel file (2 employees) is TEST DATA.
- Unit tests must use synthetic fake Excel file.
- Real production file will contain 1000+ employees.

---

# 20. Acceptance Criteria Summary

The system is accepted when:

- Emails are sent correctly.
- PDF is password protected.
- DATE substitution works.
- Validation blocks invalid runs.
- Idempotency prevents re-sending.
- DRY_RUN works correctly.
- All unit and integration tests pass.
- CSV result file is properly formatted.
- Date-based output directories function correctly.
- PDF cleanup works according to configuration.
- Console output uses standardized cprint with automatic file logging.
- Foundation modules (console, config_manager, excel abstraction) are reusable.

---
