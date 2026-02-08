"""
Configuration management for the Payslip tool (Excel COM variant).

Loads settings from local .env (tool-specific) with fallback to global .env.
All Excel sheet names, column mappings, and cell references are configurable.

This variant uses Excel COM for all formula evaluation, so it does NOT need
TBKQ cell-to-Data column mappings or calculated cell formulas.
"""

import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Tuple

from dotenv import load_dotenv, set_key


RESET = "\033[0m"

def _print_with_color(text: str, color_code: int = 92):
    """Print text with ANSI color codes."""
    print(f"\033[{color_code}m{text}{RESET}")

def _str_to_bool(value: str) -> bool:
    """Convert string to boolean."""
    if not value:
        return False
    return value.strip().lower() in ("true", "1", "yes", "on")


def _parse_cell_list(value: str) -> List[str]:
    """Parse comma-separated cell references like 'A1,A3,A5'."""
    if not value:
        return []
    return [c.strip() for c in value.split(",") if c.strip()]


def _prompt_for_value(
    key: str,
    description: str,
    example: str = "",
    validator=None,
) -> str:
    """
    Prompt user for a configuration value with optional validation.
    
    Args:
        key: Config key name
        description: User-facing description
        example: Example value to show
        validator: Optional callable(value) -> (is_valid: bool, error_msg: str)
                  Returns tuple of (is_valid, error_message)
    """
    _print_with_color(f"{description}", 33)
    if example:
        print(f"Example: {example}")
    while True:
        value = input(f"{key}: ").strip()
        if not value:
            print("Value cannot be empty. Please try again.")
            continue
        
        # Validate if validator provided
        if validator:
            is_valid, error_msg = validator(value)
            if not is_valid:
                print(f"❌ Invalid input: {error_msg}")
                continue
        
        return value


def _save_to_env(env_file: Path, key: str, value: str) -> None:
    """Save a key-value pair to .env file."""
    try:
        set_key(str(env_file), key, value)
        print(f"  Saved {key} to .env file")
    except Exception as e:
        print(f"  Warning: Could not save to .env: {e}")


# ── Validation Functions ────────────────────────────────────

def _validate_date_format(value: str) -> tuple:
    """Validate DATE format (MM/YYYY with valid month 1-12)."""
    if not re.match(r"^\d{2}/\d{4}$", value):
        return False, "Expected format: MM/YYYY (e.g., 01/2026)"
    
    try:
        month = int(value[:2])
        if month < 1 or month > 12:
            return False, f"Month must be 01-12, got {month}"
    except ValueError:
        return False, "Invalid month value"
    
    return True, ""


def _validate_email_format(value: str) -> tuple:
    """Validate email format."""
    email_pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    if not re.match(email_pattern, value):
        return False, "Invalid email format (e.g., user@company.com)"
    return True, ""


def _validate_file_path(value: str, tool_dir: Path) -> tuple:
    """Validate file path exists."""
    file_path = Path(value)
    
    # Make relative paths relative to tool_dir
    if not file_path.is_absolute():
        file_path = tool_dir / file_path
    
    if not file_path.exists():
        return False, f"File not found: {file_path}"
    
    if not file_path.is_file():
        return False, f"Path is not a file: {file_path}"
    
    return True, ""


# ── Outlook Account Functions ───────────────────────────────

def _get_outlook_accounts() -> List[str]:
    """
    Get list of configured accounts from Outlook COM.
    
    Returns:
        List of email addresses, or empty list if Outlook not available.
    """
    accounts = []
    try:
        import win32com.client
        
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        
        # Iterate accounts using indexed access (same as sender.py)
        # to ensure consistent ordering
        for i in range(1, ns.Accounts.Count + 1):
            account = ns.Accounts.Item(i)
            email = getattr(account, "SmtpAddress", "")
            if email:
                accounts.append(email)
        
        # Remove duplicates while preserving order
        seen = set()
        unique_accounts = []
        for email in accounts:
            if email not in seen:
                unique_accounts.append(email)
                seen.add(email)
        
        return unique_accounts
    
    except Exception as e:
        # Outlook not available or error occurred
        _print_with_color(f"  Note: Could not retrieve Outlook accounts: {e}", 31)
        return []


def _prompt_for_outlook_account() -> str:
    """
    Prompt user to select Outlook account from configured profiles.
    
    If no accounts found, falls back to manual email entry.
    
    Returns:
        Selected or entered email address.
    """
    accounts = _get_outlook_accounts()
    
    if not accounts:
        # Fallback: ask user to manually enter email
        print("\nOutlook account email is required.")
        validator = _validate_email_format
    else:
        # Show menu of available accounts
        _print_with_color("\nOutlook account email is required. Choose your Outlook profile:", 33)
        for i, account in enumerate(accounts, 1):
            print(f"  [{i}] {account}")
        
        # Prompt for selection
        while True:
            try:
                choice = input("\nYour account (enter number): ").strip()
                if not choice:
                    print("Selection cannot be empty. Please try again.")
                    continue
                
                choice_num = int(choice)
                if 1 <= choice_num <= len(accounts):
                    return accounts[choice_num - 1]
                else:
                    print(f"Invalid selection. Please enter a number between 1 and {len(accounts)}")
                    continue
                    
            except ValueError:
                print("Please enter a valid number.")
                continue
    
    # Manual entry fallback
    while True:
        value = input("\nOUTLOOK_ACCOUNT: ").strip()
        if not value:
            print("Value cannot be empty. Please try again.")
            continue
        
        is_valid, error_msg = _validate_email_format(value)
        if not is_valid:
            print(f"❌ Invalid input: {error_msg}")
            continue
        
        return value


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
        if not self.excel_path or not self.excel_path.exists():
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

    # Load global .env first (lower priority)
    if global_env.exists():
        load_dotenv(global_env, override=False)

    # Load local .env (higher priority, overrides global)
    local_env = tool_dir / ".env"
    if local_env.exists():
        load_dotenv(local_env, override=True)

    # Handle Excel path - make it relative to tool directory if relative
    excel_path_str = os.getenv("PAYSLIP_EXCEL_PATH", "")
    excel_path = Path(excel_path_str)
    if excel_path_str and not excel_path.is_absolute():
        excel_path = tool_dir / excel_path

    config = PayslipConfig(
        excel_path=excel_path,
        data_sheet=os.getenv("DATA_SHEET", "Data"),
        template_sheet=os.getenv("TEMPLATE_SHEET", "TBKQ"),
        email_body_sheet=os.getenv("EMAIL_BODY_SHEET", "bodymail"),
        data_header_row=int(os.getenv("DATA_HEADER_ROW", "2")),
        data_start_row=int(os.getenv("DATA_START_ROW", "4")),
        col_mnv=os.getenv("DATA_COLUMN_MNV", "A"),
        col_name=os.getenv("DATA_COLUMN_NAME", "B"),
        col_email=os.getenv("DATA_COLUMN_EMAIL", "C"),
        col_password=os.getenv("DATA_COLUMN_PASSWORD", "AZ"),
        email_subject=os.getenv("EMAIL_SUBJECT", ""),
        email_subject_cell=os.getenv("EMAIL_SUBJECT_CELL", "G1"),
        email_body_cells=_parse_cell_list(
            os.getenv("EMAIL_BODY_CELLS", "A1,A3,A5,A7,A9,A11,A12")
        ),
        email_date_cell=os.getenv("EMAIL_DATE_CELL", "A3"),
        date=os.getenv("DATE", ""),
        outlook_account=os.getenv("OUTLOOK_ACCOUNT", ""),
        dry_run=_str_to_bool(os.getenv("DRY_RUN", "true")),
        batch_size=int(os.getenv("BATCH_SIZE", "50")),
        allow_duplicate_emails=_str_to_bool(os.getenv("ALLOW_DUPLICATE_EMAILS", "false")),
        pdf_password_enabled=_str_to_bool(os.getenv("PDF_PASSWORD_ENABLED", "true")),
        pdf_password_strip_zeros=_str_to_bool(
            os.getenv("PDF_PASSWORD_STRIP_LEADING_ZEROS", "true")
        ),
        pdf_filename_pattern=os.getenv("PDF_FILENAME_PATTERN", "TBKQ_{name}_{mmyyyy}"),
        output_dir=tool_dir / Path(os.getenv("OUTPUT_DIR", "./output")),
        log_dir=tool_dir / Path(os.getenv("LOG_DIR", "./logs")),
        state_dir=tool_dir / Path(os.getenv("STATE_DIR", "./state")),
        log_level=os.getenv("LOG_LEVEL", "INFO"),
    )

    # Prompt for missing critical values
    local_env = tool_dir / ".env"

    if not config.excel_path or not config.excel_path.exists():
        excel_path_str = _prompt_for_value(
            "PAYSLIP_EXCEL_PATH",
            "Excel file path not found or invalid.",
            "../../excel-files/TBKQ-phuclong.xls",
            validator=lambda v: _validate_file_path(v, tool_dir)
        )
        excel_path = Path(excel_path_str)
        if not excel_path.is_absolute():
            excel_path = tool_dir / excel_path
        config.excel_path = excel_path
        # _save_to_env(local_env, "PAYSLIP_EXCEL_PATH", excel_path_str)

    if not config.date:
        date_str = _prompt_for_value(
            "DATE",
            "Payroll date (MM/YYYY) is required.",
            "01/2026",
            validator=_validate_date_format
        )
        config.date = date_str
        # _save_to_env(local_env, "DATE", date_str)

    if not config.outlook_account:
        outlook_account = _prompt_for_outlook_account()
        config.outlook_account = outlook_account
        # _save_to_env(local_env, "OUTLOOK_ACCOUNT", outlook_account)

    return config
