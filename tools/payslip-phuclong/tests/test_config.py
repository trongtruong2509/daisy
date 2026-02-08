"""Unit tests for config module."""

import os
import sys
from pathlib import Path

import pytest

TOOL_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(TOOL_DIR))


class TestPayslipConfig:
    """Tests for PayslipConfig."""

    def test_validate_missing_excel(self):
        from config import PayslipConfig

        config = PayslipConfig(
            excel_path=Path("/nonexistent/file.xls"),
            date="01/2026",
            outlook_account="test@example.com",
        )
        errors = config.validate()
        assert any("PAYSLIP_EXCEL_PATH" in e for e in errors)

    def test_validate_missing_date(self):
        from config import PayslipConfig

        config = PayslipConfig(
            excel_path=Path(__file__),
            date="",
            outlook_account="test@example.com",
        )
        errors = config.validate()
        assert any("DATE" in e for e in errors)

    def test_validate_bad_date_format(self):
        from config import PayslipConfig

        config = PayslipConfig(
            excel_path=Path(__file__),
            date="2026/01",
            outlook_account="test@example.com",
        )
        errors = config.validate()
        assert any("DATE" in e for e in errors)

    def test_validate_invalid_month(self):
        from config import PayslipConfig

        config = PayslipConfig(
            excel_path=Path(__file__),
            date="13/2026",
            outlook_account="test@example.com",
        )
        errors = config.validate()
        assert any("month" in e for e in errors)

    def test_validate_missing_outlook(self):
        from config import PayslipConfig

        config = PayslipConfig(
            excel_path=Path(__file__),
            date="01/2026",
            outlook_account="",
        )
        errors = config.validate()
        assert any("OUTLOOK_ACCOUNT" in e for e in errors)

    def test_valid_config(self):
        from config import PayslipConfig

        config = PayslipConfig(
            excel_path=Path(__file__),
            date="01/2026",
            outlook_account="test@example.com",
        )
        errors = config.validate()
        assert len(errors) == 0

    def test_date_properties(self):
        from config import PayslipConfig

        config = PayslipConfig(date="03/2025")
        assert config.date_mm == "03"
        assert config.date_yyyy == "2025"
        assert config.date_mmyyyy == "032025"

    def test_ensure_directories(self, tmp_path):
        from config import PayslipConfig

        config = PayslipConfig(
            output_dir=tmp_path / "out",
            log_dir=tmp_path / "logs",
            state_dir=tmp_path / "state",
        )
        config.ensure_directories()
        assert (tmp_path / "out").exists()
        assert (tmp_path / "logs").exists()
        assert (tmp_path / "state").exists()

    def test_default_cell_mapping(self):
        from config import DEFAULT_CELL_MAPPING

        assert "B3" in DEFAULT_CELL_MAPPING
        assert DEFAULT_CELL_MAPPING["B3"] == "A"
        assert DEFAULT_CELL_MAPPING["D53"] == "AH"
        assert len(DEFAULT_CELL_MAPPING) >= 19

    def test_default_calc_mapping(self):
        from config import DEFAULT_CALC_MAPPING

        assert "D16" in DEFAULT_CALC_MAPPING
        assert DEFAULT_CALC_MAPPING["D16"] == "=D17+D21"
        assert DEFAULT_CALC_MAPPING["D38"] == "U+W+X+Y"


class TestLoadConfig:
    """Tests for load_payslip_config."""

    def test_load_with_env_vars(self, tmp_path):
        os.environ["PAYSLIP_EXCEL_PATH"] = str(tmp_path / "test.xls")
        os.environ["DATE"] = "06/2025"
        os.environ["OUTLOOK_ACCOUNT"] = "sender@company.com"

        from config import load_payslip_config

        config = load_payslip_config(tool_dir=tmp_path)
        assert config.date == "06/2025"
        assert config.outlook_account == "sender@company.com"

        # Cleanup
        for key in ["PAYSLIP_EXCEL_PATH", "DATE", "OUTLOOK_ACCOUNT"]:
            os.environ.pop(key, None)

    def test_parse_cell_list(self):
        from config import _parse_cell_list

        result = _parse_cell_list("A1,A3,A5")
        assert result == ["A1", "A3", "A5"]

    def test_parse_cell_list_empty(self):
        from config import _parse_cell_list

        assert _parse_cell_list("") == []

    def test_str_to_bool(self):
        from config import _str_to_bool

        assert _str_to_bool("true") is True
        assert _str_to_bool("True") is True
        assert _str_to_bool("1") is True
        assert _str_to_bool("false") is False
        assert _str_to_bool("0") is False
        assert _str_to_bool("") is False
