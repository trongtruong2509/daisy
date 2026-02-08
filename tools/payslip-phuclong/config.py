"""
Configuration management for the Payslip tool.

Loads settings from local .env (tool-specific) with fallback to global .env.
All Excel sheet names, column mappings, and cell references are configurable.
"""

import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

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


def _parse_cell_mapping(env_prefix: str = "TBKQ_MAP_") -> Dict[str, str]:
    """
    Parse TBKQ cell mapping from environment variables.

    Variables like TBKQ_MAP_B3=A mean: TBKQ cell B3 gets value from Data column A.
    """
    mapping = {}
    for key, value in os.environ.items():
        if key.startswith(env_prefix) and value:
            cell_ref = key[len(env_prefix):]
            mapping[cell_ref] = value.strip()
    return mapping


def _parse_calc_mapping(env_prefix: str = "TBKQ_CALC_") -> Dict[str, str]:
    """
    Parse calculated cell formulas from environment variables.

    Variables like TBKQ_CALC_D16=D17+D21 mean: TBKQ cell D16 = sum of D17 and D21.
    """
    calcs = {}
    for key, value in os.environ.items():
        if key.startswith(env_prefix) and value:
            cell_ref = key[len(env_prefix):]
            calcs[cell_ref] = value.strip()
    return calcs


# Default TBKQ cell-to-Data-column mapping (discovered from template analysis)
DEFAULT_CELL_MAPPING = {
    # Employee info
    "B3": "A",     # MNV (Employee ID)
    "B4": "B",     # Employee Name
    "D3": "F",     # Start Date
    # Work info (row 10)
    "B10": "J",    # Gross salary (Mức lương)
    "C10": "K",    # Standard hours (Công chuẩn)
    "D10": "N",    # Total hours for salary (Tổng công tính lương)
    # Income details
    "D18": "O",    # Base salary + YTCLCV
    "D24": "P",    # KPI bonus
    "D28": "T",    # Project bonus (Thưởng chiến dịch/dự án)
    "D31": "Q",    # Meal allowance
    "D32": "R",    # Overtime
    "D33": "S",    # Night shift allowance
    "D37": "V",    # Seniority
    # Deductions
    "D44": "Z",    # Other deductions (Khấu trừ các khoản chi hộ)
    "D48": "AA",   # Uniform deduction
    "D50": "AB",   # Insurance
    "D51": "AC",   # Union fee
    "D52": "AD",   # Income tax
    # Net payment
    "D53": "AH",   # Net payment (THỰC LĨNH)
}

# Default calculated cells (sum of Data columns or other TBKQ cells)
# Format: "cell": "COL1+COL2" for data columns, or "cell": "=CELL1+CELL2" for TBKQ cells
DEFAULT_CALC_MAPPING = {
    "D38": "U+W+X+Y",          # Other misc payments (columns U,W,X,Y from Data)
    "D17": "=D18",              # Basic income = base salary
    "D22": "=D24+D28",          # Bonus total = KPI + project
    "D30": "=D31+D32+D33+D37+D38",  # Other payments subtotal
    "D21": "=D22+D30",          # Other income total
    "D16": "=D17+D21",          # Total income increase
    "D40": "=D44+D45+D48",      # Deductions subtotal
    "D45": "0",                 # Advance deductions (default 0)
    "D49": "=D50+D51",          # Insurance + Union
    "D39": "=D40+D49+D52",      # Total deductions
}


@dataclass
class PayslipConfig:
    """Configuration for the Payslip tool."""

    # Excel file path
    excel_path: Path = field(default_factory=lambda: Path(""))

    # Sheet names
    data_sheet: str = "Data"
    template_sheet: str = "TBKQ"
    email_body_sheet: str = "bodymail"

    # Data sheet structure
    data_header_row: int = 2        # Row containing column headers (1-based)
    data_start_row: int = 4         # First row with employee data (1-based)

    # Data sheet column mappings
    col_mnv: str = "A"
    col_name: str = "B"
    col_email: str = "C"
    col_password: str = "AZ"

    # TBKQ cell mapping (TBKQ cell → Data column)
    cell_mapping: Dict[str, str] = field(default_factory=lambda: dict(DEFAULT_CELL_MAPPING))
    # Calculated cells
    calc_mapping: Dict[str, str] = field(default_factory=lambda: dict(DEFAULT_CALC_MAPPING))

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
    pdf_filename_pattern: str = "TBKQ_{name}_{mmyyyy}.pdf"

    # Output paths
    output_dir: Path = field(default_factory=lambda: Path("./output"))
    log_dir: Path = field(default_factory=lambda: Path("./logs"))
    state_dir: Path = field(default_factory=lambda: Path("./state"))
    log_level: str = "INFO"

    # Title/header cells to update with date
    title_cell: str = "G1"
    info_cell: str = "A2"
    period_cell: str = "A10"

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


def load_payslip_config(
    tool_dir: Optional[Path] = None,
    global_env: Optional[Path] = None,
) -> PayslipConfig:
    """
    Load payslip configuration from .env files.

    Priority: local .env (tool dir) > global .env (project root) > defaults.

    Args:
        tool_dir: Path to tool directory containing local .env.
        global_env: Path to global .env file.

    Returns:
        PayslipConfig with loaded values.
    """
    # Determine paths
    if tool_dir is None:
        tool_dir = Path(__file__).resolve().parent
    if global_env is None:
        global_env = Path(__file__).resolve().parent.parent.parent / ".env"

    # Load global .env first (lower priority)
    if global_env.exists():
        load_dotenv(global_env, override=False)

    # Load local .env (higher priority, overrides global)
    local_env = tool_dir / ".env"
    if local_env.exists():
        load_dotenv(local_env, override=True)

    # Parse cell mapping from env vars (with defaults)
    env_cell_mapping = _parse_cell_mapping("TBKQ_MAP_")
    cell_mapping = dict(DEFAULT_CELL_MAPPING)
    cell_mapping.update(env_cell_mapping)

    env_calc_mapping = _parse_calc_mapping("TBKQ_CALC_")
    calc_mapping = dict(DEFAULT_CALC_MAPPING)
    calc_mapping.update(env_calc_mapping)

    config = PayslipConfig(
        # Excel file
        excel_path=Path(os.getenv("PAYSLIP_EXCEL_PATH", "")),
        # Sheet names
        data_sheet=os.getenv("DATA_SHEET", "Data"),
        template_sheet=os.getenv("TEMPLATE_SHEET", "TBKQ"),
        email_body_sheet=os.getenv("EMAIL_BODY_SHEET", "bodymail"),
        # Data structure
        data_header_row=int(os.getenv("DATA_HEADER_ROW", "2")),
        data_start_row=int(os.getenv("DATA_START_ROW", "4")),
        # Column mappings
        col_mnv=os.getenv("DATA_COLUMN_MNV", "A"),
        col_name=os.getenv("DATA_COLUMN_NAME", "B"),
        col_email=os.getenv("DATA_COLUMN_EMAIL", "C"),
        col_password=os.getenv("DATA_COLUMN_PASSWORD", "AZ"),
        # Cell mappings
        cell_mapping=cell_mapping,
        calc_mapping=calc_mapping,
        # Email
        email_subject=os.getenv("EMAIL_SUBJECT", ""),
        email_subject_cell=os.getenv("EMAIL_SUBJECT_CELL", "G1"),
        email_body_cells=_parse_cell_list(
            os.getenv("EMAIL_BODY_CELLS", "A1,A3,A5,A7,A9,A11,A12")
        ),
        email_date_cell=os.getenv("EMAIL_DATE_CELL", "A3"),
        # Date
        date=os.getenv("DATE", ""),
        # Outlook
        outlook_account=os.getenv("OUTLOOK_ACCOUNT", ""),
        # Processing
        dry_run=_str_to_bool(os.getenv("DRY_RUN", "true")),
        batch_size=int(os.getenv("BATCH_SIZE", "50")),
        # PDF
        pdf_password_enabled=_str_to_bool(os.getenv("PDF_PASSWORD_ENABLED", "true")),
        pdf_password_strip_zeros=_str_to_bool(
            os.getenv("PDF_PASSWORD_STRIP_LEADING_ZEROS", "true")
        ),
        pdf_filename_pattern=os.getenv(
            "PDF_FILENAME_PATTERN", "TBKQ_{name}_{mmyyyy}.pdf"
        ),
        # Paths
        output_dir=Path(os.getenv("OUTPUT_DIR", "./output")),
        log_dir=Path(os.getenv("LOG_DIR", "./logs")),
        state_dir=Path(os.getenv("STATE_DIR", "./state")),
        log_level=os.getenv("LOG_LEVEL", "INFO"),
        # Title cells
        title_cell=os.getenv("TBKQ_TITLE_CELL", "G1"),
        info_cell=os.getenv("TBKQ_INFO_CELL", "A2"),
        period_cell=os.getenv("TBKQ_PERIOD_CELL", "A10"),
    )

    return config
