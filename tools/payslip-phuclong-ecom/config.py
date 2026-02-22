"""
Configuration management for the Payslip tool (Excel COM variant).

Loads settings from local .env (tool-specific) with fallback to global .env.
All Excel sheet names, column mappings, and cell references are configurable.

This variant uses Excel COM for all formula evaluation, so it does NOT need
TBKQ cell-to-Data column mappings or calculated cell formulas.
"""

import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

# Add project root to path for foundation imports
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.config_manager import ConfigManager
from core.console import cprint
from office.outlook import get_outlook_accounts


# ── Outlook Account Functions ───────────────────────────────

def _prompt_for_outlook_account() -> str:
    """
    Prompt user to select Outlook account from configured profiles.

    Uses get_outlook_accounts() from office.outlook module.
    If no accounts found, falls back to manual email entry.

    Returns:
        Selected or entered email address.
    """
    accounts = get_outlook_accounts()

    print()
    if not accounts:
        # Fallback: ask user to manually enter email
        cprint("Choose Outlook accounts", level="WARNING")
        return ConfigManager.prompt_for_value(
            "OUTLOOK_ACCOUNT",
            "No Outlook accounts detected. Enter email manually.",
            "user@company.com",
            validator=ConfigManager.validate_email,
        )

    # Show menu of available accounts
    cprint("Choose Outlook accounts", level="WARNING")
    for i, account in enumerate(accounts, 1):
        print(f"  [{i}] {account}")

    # Prompt for selection
    while True:
        try:
            choice = input("\u2192 Selected: ").strip()
            if not choice:
                print("  Selection cannot be empty. Please try again.")
                continue

            choice_num = int(choice)
            if 1 <= choice_num <= len(accounts):
                return accounts[choice_num - 1]
            else:
                cprint(f"Invalid selection. Please enter 1-{len(accounts)}", level="ERROR")
                continue

        except ValueError:
            cprint("Please enter a valid number.", level="ERROR")
            continue


@dataclass
class PayslipConfig:
    """Configuration for the Payslip tool (Excel COM variant)."""

    # Excel file path
    excel_path: Path = field(default_factory=lambda: Path(""))

    # Sheet names
    data_sheet: str = "Data"
    template_sheet: str = "TBKQ"
    email_body_sheet: str = "bodymail"

    # Data sheet structure
    data_header_row: int = 2
    data_start_row: int = 4

    # Data sheet column mappings
    col_mnv: str = "A"
    col_name: str = "B"
    col_email: str = "C"
    col_password: str = "AZ"

    # Email configuration
    email_subject: str = ""
    email_subject_cell: str = "G1"
    email_body_cells: List[str] = field(
        default_factory=lambda: ["A1", "A3", "A5", "A7", "A9", "A11", "A12"]
    )
    email_date_cell: str = "A3"

    # Payroll date (MM/YYYY)
    date: str = ""

    # Outlook settings
    outlook_account: str = ""

    # Processing options
    dry_run: bool = True
    batch_size: int = 50
    allow_duplicate_emails: bool = False

    # PDF options
    pdf_password_enabled: bool = True
    pdf_password_strip_zeros: bool = True
    pdf_filename_pattern: str = "TBKQ_{name}_{mmyyyy}"
    keep_pdf_payslips: bool = False

    # Output paths
    output_dir: Path = field(default_factory=lambda: Path("./output"))
    log_dir: Path = field(default_factory=lambda: Path("./logs"))
    state_dir: Path = field(default_factory=lambda: Path("./state"))
    log_level: str = "INFO"

    def __post_init__(self):
        """Validate and normalize."""
        if isinstance(self.excel_path, str):
            self.excel_path = Path(self.excel_path)
        if isinstance(self.output_dir, str):
            self.output_dir = Path(self.output_dir)
        if isinstance(self.log_dir, str):
            self.log_dir = Path(self.log_dir)
        if isinstance(self.state_dir, str):
            self.state_dir = Path(self.state_dir)
        self.log_level = self.log_level.upper()

    def validate(self) -> List[str]:
        """Validate configuration, return list of errors."""
        errors = []
        if not self.excel_path or not self.excel_path.is_file():
            errors.append(f"PAYSLIP_EXCEL_PATH not set or file not found: {self.excel_path}")
        if not self.date:
            errors.append("DATE is required (format: MM/YYYY)")
        elif not re.match(r"^\d{2}/\d{4}$", self.date):
            errors.append(f"DATE format invalid: '{self.date}'. Expected MM/YYYY")
        else:
            month = int(self.date[:2])
            if month < 1 or month > 12:
                errors.append(f"DATE month invalid: {month}")
        if not self.outlook_account:
            errors.append("OUTLOOK_ACCOUNT is required")
        return errors

    def ensure_directories(self):
        """Create output directories if they don't exist."""
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self.state_dir.mkdir(parents=True, exist_ok=True)

    @property
    def date_mm(self) -> str:
        """Month part of DATE (e.g., '01')."""
        return self.date.split("/")[0] if self.date else ""

    @property
    def date_yyyy(self) -> str:
        """Year part of DATE (e.g., '2026')."""
        return self.date.split("/")[1] if self.date else ""

    @property
    def date_mmyyyy(self) -> str:
        """Date without separator (e.g., '012026')."""
        return self.date_mm + self.date_yyyy


def load_config(
    tool_dir: Optional[Path] = None,
    global_env: Optional[Path] = None,
) -> PayslipConfig:
    """
    Load payslip configuration from .env files.

    Priority: local .env (tool dir) > global .env (project root) > defaults.
    """
    if tool_dir is None:
        tool_dir = Path(__file__).resolve().parent
    else:
        tool_dir = Path(tool_dir).resolve()

    if global_env is None:
        global_env = Path(__file__).resolve().parent.parent.parent / ".env"

    # Load .env files with priority ordering via ConfigManager
    mgr = ConfigManager()
    mgr.load_env([global_env, tool_dir / ".env"])

    # Build config using ConfigManager typed getters
    config = PayslipConfig(
        excel_path=mgr.get_path("PAYSLIP_EXCEL_PATH", base_dir=tool_dir),
        data_sheet=mgr.get("DATA_SHEET", "Data"),
        template_sheet=mgr.get("TEMPLATE_SHEET", "TBKQ"),
        email_body_sheet=mgr.get("EMAIL_BODY_SHEET", "bodymail"),
        data_header_row=mgr.get_int("DATA_HEADER_ROW", 2),
        data_start_row=mgr.get_int("DATA_START_ROW", 4),
        col_mnv=mgr.get("DATA_COLUMN_MNV", "A"),
        col_name=mgr.get("DATA_COLUMN_NAME", "B"),
        col_email=mgr.get("DATA_COLUMN_EMAIL", "C"),
        col_password=mgr.get("DATA_COLUMN_PASSWORD", "AZ"),
        email_subject=mgr.get("EMAIL_SUBJECT", ""),
        email_subject_cell=mgr.get("EMAIL_SUBJECT_CELL", "G1"),
        email_body_cells=mgr.get_list(
            "EMAIL_BODY_CELLS", default=["A1", "A3", "A5", "A7", "A9", "A11", "A12"]
        ),
        email_date_cell=mgr.get("EMAIL_DATE_CELL", "A3"),
        date=mgr.get("DATE", ""),
        outlook_account=mgr.get("OUTLOOK_ACCOUNT", ""),
        dry_run=mgr.get_bool("DRY_RUN", True),
        batch_size=mgr.get_int("BATCH_SIZE", 50),
        allow_duplicate_emails=mgr.get_bool("ALLOW_DUPLICATE_EMAILS", False),
        pdf_password_enabled=mgr.get_bool("PDF_PASSWORD_ENABLED", True),
        pdf_password_strip_zeros=mgr.get_bool("PDF_PASSWORD_STRIP_LEADING_ZEROS", True),
        pdf_filename_pattern=mgr.get("PDF_FILENAME_PATTERN", "TBKQ_{name}_{mmyyyy}"),
        keep_pdf_payslips=mgr.get_bool("KEEP_PDF_PAYSLIPS", False),
        output_dir=tool_dir / Path(mgr.get("OUTPUT_DIR", "./output")),
        log_dir=tool_dir / Path(mgr.get("LOG_DIR", "./logs")),
        state_dir=tool_dir / Path(mgr.get("STATE_DIR", "./state")),
        log_level=mgr.get("LOG_LEVEL", "INFO"),
    )

    # Prompt for missing critical values
    local_env = tool_dir / ".env"

    # Check if any prompts are needed
    needs_prompt = (
        (not config.excel_path or not config.excel_path.is_file()) or
        not config.date or
        not config.outlook_account
    )

    if needs_prompt:
        cprint("Configuration", level="PHASE")
        print()

    if not config.excel_path or not config.excel_path.is_file():
        excel_path_str = mgr.prompt_for_value(
            "PAYSLIP_EXCEL_PATH",
            "Excel file path not found or invalid. Please input the excel file path",
            "D:\\path\\to\\excel\\file\\TBKQ-phuclong.xls",
            validator=lambda v: ConfigManager.validate_file_path(v, tool_dir),
        )
        # prompt_for_value() and validate_file_path() already normalize quotes
        excel_path = Path(excel_path_str)
        if not excel_path.is_absolute():
            excel_path = tool_dir / excel_path
        config.excel_path = excel_path

    if not config.date:
        date_str = mgr.prompt_for_value(
            "DATE",
            "Set Payroll Date (MM/YYYY)",
            "01/2026",
            validator=ConfigManager.validate_date,
        )
        config.date = date_str

    if not config.outlook_account:
        outlook_account = _prompt_for_outlook_account()
        config.outlook_account = outlook_account

    # Restructure output_dir to include date-based subfolder: output/<MMYYYY>/
    if config.date:
        config.output_dir = config.output_dir / config.date_mmyyyy

    return config
