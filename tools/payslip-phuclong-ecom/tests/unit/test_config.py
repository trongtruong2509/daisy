"""
Unit tests for config.py — PayslipConfig and load_config.

Covers:
- PayslipConfig defaults
- PayslipConfig validation (date format, missing fields)
- Date properties (date_mm, date_yyyy, date_mmyyyy)
- Directory creation
- Post-init type normalization
- load_config with .env files
- Interactive prompting (mocked)
"""

import os
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
TOOL_DIR = Path(__file__).resolve().parent.parent.parent
for p in [str(PROJECT_ROOT), str(TOOL_DIR)]:
    if p not in sys.path:
        sys.path.insert(0, p)

from config import PayslipConfig, load_config


class TestPayslipConfigDefaults:
    """Tests for PayslipConfig default values."""

    def test_default_sheets(self):
        cfg = PayslipConfig()
        assert cfg.data_sheet == "Data"
        assert cfg.template_sheet == "TBKQ"
        assert cfg.email_body_sheet == "bodymail"

    def test_default_columns(self):
        cfg = PayslipConfig()
        assert cfg.col_mnv == "A"
        assert cfg.col_name == "B"
        assert cfg.col_email == "C"
        assert cfg.col_password == "AZ"

    def test_default_processing_options(self):
        cfg = PayslipConfig()
        assert cfg.dry_run is True
        assert cfg.pdf_password_enabled is True
        assert cfg.keep_pdf_payslips is False

    def test_default_data_rows(self):
        cfg = PayslipConfig()
        assert cfg.data_header_row == 2
        assert cfg.data_start_row == 4


class TestPayslipConfigValidation:
    """Tests for PayslipConfig.validate()."""

    def test_valid_config(self, tmp_path):
        excel = tmp_path / "test.xls"
        excel.touch()
        cfg = PayslipConfig(
            excel_path=excel,
            date="01/2026",
            outlook_account="test@co.com",
        )
        errors = cfg.validate()
        assert len(errors) == 0

    def test_missing_excel_path(self, tmp_path):
        cfg = PayslipConfig(
            excel_path=tmp_path / "nonexistent.xls",
            date="01/2026",
            outlook_account="test@co.com",
        )
        errors = cfg.validate()
        assert any("PAYSLIP_EXCEL_PATH" in e for e in errors)

    def test_nonexistent_excel_path(self, tmp_path):
        cfg = PayslipConfig(
            excel_path=tmp_path / "missing.xls",
            date="01/2026",
            outlook_account="test@co.com",
        )
        errors = cfg.validate()
        assert any("PAYSLIP_EXCEL_PATH" in e for e in errors)

    def test_missing_date(self, tmp_path):
        excel = tmp_path / "test.xls"
        excel.touch()
        cfg = PayslipConfig(
            excel_path=excel,
            outlook_account="test@co.com",
        )
        errors = cfg.validate()
        assert any("DATE" in e for e in errors)

    def test_invalid_date_format(self, tmp_path):
        excel = tmp_path / "test.xls"
        excel.touch()
        cfg = PayslipConfig(
            excel_path=excel,
            date="2026-01",
            outlook_account="test@co.com",
        )
        errors = cfg.validate()
        assert any("DATE format" in e for e in errors)

    def test_invalid_date_month(self, tmp_path):
        excel = tmp_path / "test.xls"
        excel.touch()
        cfg = PayslipConfig(
            excel_path=excel,
            date="13/2026",
            outlook_account="test@co.com",
        )
        errors = cfg.validate()
        assert any("month" in e.lower() for e in errors)

    def test_missing_outlook_account(self, tmp_path):
        excel = tmp_path / "test.xls"
        excel.touch()
        cfg = PayslipConfig(
            excel_path=excel,
            date="01/2026",
        )
        errors = cfg.validate()
        assert any("OUTLOOK_ACCOUNT" in e for e in errors)

    def test_multiple_errors_reported(self, tmp_path):
        cfg = PayslipConfig(excel_path=tmp_path / "missing.xls")
        errors = cfg.validate()
        # Should have: missing excel, missing date, missing outlook
        assert len(errors) >= 3


class TestPayslipConfigDateProperties:
    """Tests for date-related properties."""

    def test_date_mm(self):
        cfg = PayslipConfig(date="03/2026")
        assert cfg.date_mm == "03"

    def test_date_yyyy(self):
        cfg = PayslipConfig(date="03/2026")
        assert cfg.date_yyyy == "2026"

    def test_date_mmyyyy(self):
        cfg = PayslipConfig(date="12/2025")
        assert cfg.date_mmyyyy == "122025"

    def test_empty_date_properties(self):
        cfg = PayslipConfig()
        assert cfg.date_mm == ""
        assert cfg.date_yyyy == ""
        assert cfg.date_mmyyyy == ""


class TestPayslipConfigPostInit:
    """Tests for __post_init__ type normalization."""

    def test_string_to_path_conversion(self):
        cfg = PayslipConfig(
            excel_path="./test.xls",
            output_dir="./out",
            log_dir="./logs",
            state_dir="./state",
        )
        assert isinstance(cfg.excel_path, Path)
        assert isinstance(cfg.output_dir, Path)
        assert isinstance(cfg.log_dir, Path)
        assert isinstance(cfg.state_dir, Path)

    def test_log_level_uppercase(self):
        cfg = PayslipConfig(log_level="debug")
        assert cfg.log_level == "DEBUG"


class TestPayslipConfigDirectories:
    """Tests for ensure_directories."""

    def test_creates_directories(self, tmp_path):
        cfg = PayslipConfig(
            output_dir=tmp_path / "out",
            log_dir=tmp_path / "logs",
            state_dir=tmp_path / "state",
        )
        cfg.ensure_directories()
        assert (tmp_path / "out").exists()
        assert (tmp_path / "logs").exists()
        assert (tmp_path / "state").exists()


class TestLoadConfig:
    """Tests for load_config function."""

    @patch("config._prompt_for_outlook_account", return_value="test@co.com")
    @patch("config.ConfigManager")
    def test_load_config_from_env(self, MockCM, mock_prompt, tmp_path):
        """Test that load_config reads from .env and creates config."""
        # Create env file
        env_file = tmp_path / ".env"
        env_file.write_text(
            "PAYSLIP_EXCEL_PATH=test.xls\n"
            "DATE=01/2026\n"
            "OUTLOOK_ACCOUNT=user@co.com\n"
            "DRY_RUN=true\n"
        )

        # Create the excel file
        excel_file = tmp_path / "test.xls"
        excel_file.touch()

        # Setup mock ConfigManager
        mgr = MagicMock()
        MockCM.return_value = mgr
        mgr.get.side_effect = lambda k, d="": {
            "DATA_SHEET": "Data",
            "TEMPLATE_SHEET": "TBKQ",
            "EMAIL_BODY_SHEET": "bodymail",
            "DATA_COLUMN_MNV": "A",
            "DATA_COLUMN_NAME": "B",
            "DATA_COLUMN_EMAIL": "C",
            "DATA_COLUMN_PASSWORD": "AZ",
            "EMAIL_SUBJECT": "",
            "EMAIL_SUBJECT_CELL": "G1",
            "EMAIL_DATE_CELL": "A3",
            "OUTLOOK_ACCOUNT": "user@co.com",
            "PDF_FILENAME_PATTERN": "TBKQ_{name}_{mmyyyy}",
            "OUTPUT_DIR": "./output",
            "LOG_DIR": "./logs",
            "STATE_DIR": "./state",
            "LOG_LEVEL": "INFO",
            "DATE": "01/2026",
        }.get(k, d)
        mgr.get_bool.side_effect = lambda k, d=False: {
            "DRY_RUN": True,
            "ALLOW_DUPLICATE_EMAILS": False,
            "PDF_PASSWORD_ENABLED": True,
            "PDF_PASSWORD_STRIP_LEADING_ZEROS": True,
            "KEEP_PDF_PAYSLIPS": False,
        }.get(k, d)
        mgr.get_int.side_effect = lambda k, d=0: {
            "DATA_HEADER_ROW": 2,
            "DATA_START_ROW": 4,
            "BATCH_SIZE": 50,
        }.get(k, d)
        mgr.get_path.return_value = excel_file
        mgr.get_list.return_value = ["A1", "A3", "A5"]

        # Mock the validators
        MockCM.validate_date = MagicMock()
        MockCM.validate_email = MagicMock()
        MockCM.validate_file_path = MagicMock()

        config = load_config(tool_dir=tmp_path, global_env=tmp_path / "global.env")

        assert isinstance(config, PayslipConfig)
        assert config.data_sheet == "Data"
        assert config.dry_run is True
