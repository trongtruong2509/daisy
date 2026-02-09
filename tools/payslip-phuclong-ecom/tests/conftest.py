"""
Test configuration and shared fixtures for payslip-phuclong-ecom tests.

Provides:
- Sample employee data factories
- Mock config objects
- Temporary directory setup
- COM mock helpers
"""

import csv
import json
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List
from unittest.mock import MagicMock, patch

import pytest

# Add project root AND tool directory to path for imports
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent
TOOL_DIR = Path(__file__).resolve().parent.parent

for p in [str(PROJECT_ROOT), str(TOOL_DIR)]:
    if p not in sys.path:
        sys.path.insert(0, p)


# ── Employee Data Factories ────────────────────────────────


def make_employee(
    row: int = 4,
    mnv: str = "001",
    name: str = "Nguyen Van A",
    email: str = "a@company.com",
    password: str = "001",
) -> Dict[str, Any]:
    """Create a single employee dict."""
    return {
        "row": row,
        "mnv": mnv,
        "name": name,
        "email": email,
        "password": password,
    }


def make_employees(count: int = 3) -> List[Dict[str, Any]]:
    """Create a list of sample employees."""
    employees = []
    for i in range(1, count + 1):
        employees.append(
            make_employee(
                row=i + 3,
                mnv=f"{i:03d}",
                name=f"Employee {i}",
                email=f"emp{i}@company.com",
                password=f"{i:03d}",
            )
        )
    return employees


# ── Fixtures ────────────────────────────────────────────────


@pytest.fixture
def sample_employee():
    """Single sample employee."""
    return make_employee()


@pytest.fixture
def sample_employees():
    """List of 3 sample employees."""
    return make_employees(3)


@pytest.fixture
def tmp_output_dir(tmp_path):
    """Temporary output directory."""
    out = tmp_path / "output" / "012026"
    out.mkdir(parents=True)
    return out


@pytest.fixture
def tmp_state_dir(tmp_path):
    """Temporary state directory."""
    state = tmp_path / "state"
    state.mkdir(parents=True)
    return state


@pytest.fixture
def tmp_log_dir(tmp_path):
    """Temporary log directory."""
    logs = tmp_path / "logs"
    logs.mkdir(parents=True)
    return logs


@pytest.fixture
def mock_config(tmp_path, tmp_output_dir, tmp_state_dir, tmp_log_dir):
    """
    Create a mock PayslipConfig with temporary directories.

    Provides all attributes used by main.py helper functions.
    """
    # Create a dummy Excel file
    excel_file = tmp_path / "test.xls"
    excel_file.touch()

    config = MagicMock()
    config.excel_path = excel_file
    config.data_sheet = "Data"
    config.template_sheet = "TBKQ"
    config.email_body_sheet = "bodymail"
    config.data_header_row = 2
    config.data_start_row = 4
    config.col_mnv = "A"
    config.col_name = "B"
    config.col_email = "C"
    config.col_password = "AZ"
    config.email_subject = "Test Subject"
    config.email_subject_cell = "G1"
    config.email_body_cells = ["A1", "A3", "A5"]
    config.email_date_cell = "A3"
    config.date = "01/2026"
    config.date_mm = "01"
    config.date_yyyy = "2026"
    config.date_mmyyyy = "012026"
    config.outlook_account = "test@company.com"
    config.dry_run = True
    config.batch_size = 50
    config.allow_duplicate_emails = False
    config.pdf_password_enabled = True
    config.pdf_password_strip_zeros = True
    config.pdf_filename_pattern = "TBKQ_{name}_{mmyyyy}"
    config.keep_pdf_payslips = False
    config.output_dir = tmp_output_dir
    config.log_dir = tmp_log_dir
    config.state_dir = tmp_state_dir
    config.log_level = "INFO"
    config.validate.return_value = []
    config.ensure_directories = MagicMock()

    return config


@pytest.fixture
def mock_email_template():
    """Sample email template cells."""
    return {
        "A1": "Dear Employee,",
        "A3": "Your payslip for tháng 01/2026 is attached.",
        "A5": "Password: Your employee ID",
        "A7": "Best regards,",
        "A9": "HR Department",
    }


@pytest.fixture
def mock_com_worksheet():
    """Create a mock COM worksheet."""
    ws = MagicMock()

    def mock_range(cell_ref):
        cell = MagicMock()
        cell.Value = f"MockValue_{cell_ref}"
        return cell

    ws.Range = mock_range
    ws.Cells = MagicMock()
    ws.Rows = MagicMock()
    ws.Rows.Count = 65536
    return ws


@pytest.fixture
def mock_excel_app():
    """Create a mock Excel COM Application."""
    app = MagicMock()
    app.Visible = False
    app.DisplayAlerts = False
    return app
