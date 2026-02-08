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
from typing import List, Optional

from dotenv import load_dotenv


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

    return config
