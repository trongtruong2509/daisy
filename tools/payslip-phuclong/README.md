# Payslip Generator & Distributor - Phuc Long

Automated payslip generation and email distribution tool for Phuc Long company. Reads employee data from Excel, generates password-protected PDF payslips, and sends them via Outlook.

## Features

- **Direct Population Approach** — Fast payslip generation without VLOOKUP recalculation (30-60x faster than VBA)
- **Password Protection** — Each PDF is protected with employee's ID (with leading zeros stripped)
- **Email Automation** — Sends personalized emails with attachments via Outlook COM
- **Idempotent** — Prevents duplicate sends on re-run using state tracking
- **Fail-Fast Validation** — Validates all data before any processing begins
- **Comprehensive Logging** — Detailed logs for auditing and troubleshooting
- **Dry-Run Mode** — Test without actually sending emails
- **Fully Configurable** — All sheet names, columns, cell mappings via `.env`

## Prerequisites

### System Requirements
- **Windows 10/11** with Microsoft Outlook Desktop installed
- **Python 3.9+**
- **Microsoft Excel** (for PDF conversion via COM)

### Python Dependencies
```bash
pip install -r ../../requirements.txt
```

Key dependencies:
- `pywin32` — Windows COM automation (Outlook + Excel)
- `openpyxl` — Excel file manipulation
- `xlrd` — Legacy .xls file reading
- `pikepdf` — PDF password protection
- `python-dotenv` — Environment configuration

## Installation

1. **Clone the repository** (if not already done):
   ```bash
   git clone https://github.com/trongtruong2509/daisy.git
   cd daisy
   ```

2. **Run the setup script** (creates virtual environment and installs dependencies):
   ```bash
   setup.bat
   ```

3. **Navigate to the tool directory** (optional):
   ```bash
   cd tools/payslip-phuclong
   ```

4. **Create local configuration**:
   ```bash
   cp .env.example .env
   ```

## Configuration

Edit the `.env` file with your settings:

### Required Settings

```bash
# Path to Excel file (relative or absolute)
PAYSLIP_EXCEL_PATH=../../excel-files/TBKQ-phuclong.xls

# Payroll date (MM/YYYY format)
DATE=01/2026

# Outlook account to send from
OUTLOOK_ACCOUNT=your.email@company.com
```

### Optional Settings

```bash
# Dry-run mode (test without sending)
DRY_RUN=true

# Output directories
OUTPUT_DIR=./output
LOG_DIR=./logs
STATE_DIR=./state

# PDF settings
PDF_PASSWORD_ENABLED=true
PDF_PASSWORD_STRIP_LEADING_ZEROS=true
PDF_FILENAME_PATTERN=TBKQ_{name}_{mmyyyy}.pdf

# Email template cells (from bodymail sheet)
EMAIL_BODY_CELLS=A1,A3,A5,A7,A9,A11,A12
```

### Cell Mapping (Advanced)

The tool uses configurable cell mappings to populate the TBKQ template. The default mappings are pre-configured in `.env.example`:

```bash
# Direct mappings: TBKQ cell → Data column
TBKQ_MAP_B3=A     # Employee ID (MNV)
TBKQ_MAP_B4=B     # Employee Name
TBKQ_MAP_D53=AH   # Net Payment

# Calculated cells: formulas or sums
TBKQ_CALC_D16==D17+D21    # Total income
TBKQ_CALC_D38=U+W+X+Y     # Sum of columns U,W,X,Y
```

## Usage

### Quick Start (Windows)

**From repository root:**
```cmd
REM Interactive menu
run.bat

REM Direct tool invocation
run.bat payslip-phuclong
```

**From this tool directory:**
```cmd
REM Tool wrapper (calls master launcher)
run.bat
```

### Basic Usage

1. **Prepare your Excel file** with employee data in the Data sheet
2. **Configure `.env`** with the correct date and Outlook account
3. **Run in dry-run mode first** (to verify):
   ```cmd
   run.bat payslip-phuclong
   ```
4. **Review the pre-execution summary** and confirm
5. **Check the output** in `./output/` and logs in `./logs/`

### Production Run

Once verified in dry-run mode, disable dry-run:

1. Edit `.env`:
   ```bash
   DRY_RUN=false
   ```

2. Run the tool:
   ```cmd
   run.bat payslip-phuclong
   ```

3. **Confirm when prompted**:
   ```
   Proceed with payslip generation and email sending? (yes/no): yes
   ```

### Python Direct Invocation (Advanced)

If you prefer to use Python directly:

```bash
# Activate virtual environment first
venv\Scripts\activate.bat

# Run the tool
python main.py  # with DRY_RUN=true

# Step 2: Verify generated files
ls -lh output/

# Step 3: Production run
# (Edit .env: DRY_RUN=false)
python main.py

# Step 4: Check results
tail -50 logs/payslip_*.log
```

## Excel File Structure

The tool expects an Excel file with these sheets:

### Data Sheet
- **Row 2**: Column headers
- **Row 4+**: Employee data
- **Required columns**:
  - `A` — MNV (Employee ID)
  - `B` — Name
  - `C` — EmailAddress
  - `AZ` — PassWord

### TBKQ Sheet
- Template for individual payslips
- Cell `G1` — Email subject
- Cells are populated via direct mapping (no formulas)

### bodymail Sheet
- Email template cells (A1, A3, A5, A7, A9, A11, A12)
- Cell `A3` — Contains date placeholder (will be replaced)

## Output

The tool generates:

1. **Excel files** (`.xlsx`) — Filled payslip templates (temporary, cleaned up after PDF conversion)
2. **PDF files** — Password-protected payslips (`TBKQ_{name}_{mmyyyy}.pdf`)
3. **Log files** — Detailed execution logs (`./logs/payslip_*.log`)
4. **State files** — Tracking for idempotency (`./state/payslip_send_*.json`)

## Workflow

```
1. Configuration Loading
   ↓
2. Excel Reading (Data, TBKQ, bodymail)
   ↓
3. Data Validation (email format, duplicates, required fields)
   ↓
4. Pre-Execution Summary & User Confirmation
   ↓
5. Payslip Generation (direct population, calculated cells)
   ↓
6. PDF Conversion (Excel COM → PDF → password protection)
   ↓
7. Email Composition & Sending (via Outlook)
   ↓
8. Post-Execution Summary
```

## Testing

### Run Unit Tests

```bash
# Install pytest (if not already installed)
pip install pytest

# Run all tests
python -m pytest tests/ -v

# Run specific test file
python -m pytest tests/test_validator.py -v

# Run with coverage
python -m pytest tests/ --cov=. --cov-report=html
```

### Test Coverage

The test suite includes:
- **69 tests** covering all modules
- **Config validation** — date format, required fields
- **Email validation** — format, duplicates, required fields
- **Payslip generation** — cell mapping, calculated cells, date updates
- **Email composition** — HTML body, date replacement, cell ordering

## Troubleshooting

### Common Issues

#### 1. "ModuleNotFoundError: No module named 'dotenv'"
```bash
pip install python-dotenv
```

#### 2. "Excel COM not available"
- Ensure Microsoft Excel is installed on Windows
- Run on Windows (not WSL/Linux) for COM support
- Alternative: Convert .xls to .xlsx and use openpyxl-only mode

#### 3. "Outlook COM not available"
- Ensure Microsoft Outlook Desktop is installed
- Run on Windows (not WSL/Linux)
- Outlook must be configured with the account specified in `.env`

#### 4. "Validation failed: Duplicate email"
- Check Data sheet for duplicate email addresses
- Fix duplicates and re-run

#### 5. "No employee data found"
- Verify `DATA_START_ROW` is correct (default: 4)
- Check that Data sheet has employee rows with MNV values

#### 6. "Template preparation failed"
- Ensure Excel is installed (for .xls files)
- Or convert source file to .xlsx format first

### Debug Mode

Enable verbose logging:
```bash
LOG_LEVEL=DEBUG
```

Check logs:
```bash
tail -100 logs/payslip_*.log
```

### State Management

If you need to re-send emails (override idempotency):
```bash
# Delete state files
rm -rf ./state/

# Or delete specific date state
rm ./state/payslip_send_012026.json
```

## Performance

For **1000 employees**:
- **Payslip generation**: ~5-10 minutes
- **PDF conversion**: ~10-15 minutes (depends on Excel COM)
- **Email sending**: ~5-10 minutes (depends on Outlook)
- **Total**: ~20-35 minutes (vs. VBA 10+ hours)

## Security

- **Passwords**: PDFs protected with employee ID (leading zeros stripped)
- **State files**: Track sent emails to prevent duplicates
- **Dry-run mode**: Test before production sends
- **Logging**: All operations logged for audit trail
- **.env**: Never commit `.env` to version control (already in `.gitignore`)

## Architecture

```
tools/payslip-phuclong/
├── main.py                 # Orchestrator
├── config.py               # Configuration management
├── excel_reader.py         # Read .xls/.xlsx files
├── validator.py            # Data validation (fail-fast)
├── payslip_generator.py    # Direct population approach
├── pdf_converter.py        # Excel COM + pikepdf
├── email_composer.py       # HTML email builder
├── .env.example            # Configuration reference
└── tests/                  # Unit tests (69 tests)
    ├── conftest.py
    ├── test_config.py
    ├── test_validator.py
    ├── test_payslip_generator.py
    ├── test_email_composer.py
    └── test_excel_reader.py
```

## Foundation Integration

This tool uses the Daisy Foundation codebase:
- `core.config` — Configuration loading
- `core.logger` — Logging with colors and timestamps
- `core.state` — State tracking for idempotency
- `core.retry` — Retry logic for COM operations
- `office.outlook.sender` — Safe email sending
- `office.outlook.models` — `NewEmail`, `Importance` enums

## License

Internal tool for Phuc Long company.

## Support

For issues or questions:
1. Check logs in `./logs/`
2. Review this README
3. Run tests: `python -m pytest tests/ -v`
4. Contact: [maintainer email]

---

**Last Updated**: February 2026  
**Version**: 1.0.0
