"""
Unit tests for PasteSpecial retry logic in PayslipGenerator._generate_one.

Verifies that the retry_with_backoff wrapper handles transient COM errors
during the Copy + PasteSpecial operation.
"""

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch, call

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
TOOL_DIR = Path(__file__).resolve().parent.parent.parent
for p in [str(PROJECT_ROOT), str(TOOL_DIR)]:
    if p not in sys.path:
        sys.path.insert(0, p)

from payslip_generator import PayslipGenerator
from tests.conftest import make_employee


def _make_mock_excel():
    """Create mock Excel COM objects for _generate_one."""
    mock_excel = MagicMock()
    mock_wb = MagicMock()

    # Data worksheet
    data_ws = MagicMock()
    data_cell = MagicMock()
    data_cell.Value = "001"
    data_ws.Range.return_value = data_cell

    # TBKQ worksheet
    tbkq_ws = MagicMock()

    mock_wb.Sheets.side_effect = lambda name: {
        "Data": data_ws,
        "TBKQ": tbkq_ws,
    }[name]

    # New workbook after copy
    new_wb = MagicMock()
    new_ws = MagicMock()
    mock_excel.ActiveWorkbook = new_wb
    new_wb.ActiveSheet = new_ws

    new_ws.Range.return_value = MagicMock(Value="tháng 01/2025")
    new_ws.Cells = MagicMock()
    new_ws.Buttons = MagicMock()
    new_ws.Columns = MagicMock()
    new_ws.Outline = MagicMock()
    new_ws.PageSetup = MagicMock()
    new_wb.Names = []

    return mock_excel, mock_wb, new_ws


class TestPasteSpecialRetry:
    """Tests for PasteSpecial retry mechanism."""

    @patch("payslip_generator.retry_with_backoff")
    def test_retry_with_backoff_called(self, mock_retry, tmp_path):
        """Verify retry_with_backoff is called during generation."""
        gen = PayslipGenerator(tmp_path, "01/2026")
        mock_excel, mock_wb, new_ws = _make_mock_excel()

        # Make retry_with_backoff execute the function immediately
        mock_retry.side_effect = lambda func, **kwargs: func()

        emp = make_employee(row=4, mnv="001")
        gen._generate_one(mock_excel, mock_wb, "TBKQ", "Data", "A", emp)

        # retry_with_backoff should have been called once for PasteSpecial
        mock_retry.assert_called_once()
        # Check that config was passed with correct parameters
        call_kwargs = mock_retry.call_args
        assert call_kwargs.kwargs["config"].max_attempts == 3
        assert "PasteSpecial" in call_kwargs.kwargs["operation_name"]

    @patch("payslip_generator.retry_with_backoff")
    def test_retry_config_parameters(self, mock_retry, tmp_path):
        """Verify RetryConfig parameters for PasteSpecial."""
        gen = PayslipGenerator(tmp_path, "01/2026")
        mock_excel, mock_wb, new_ws = _make_mock_excel()

        mock_retry.side_effect = lambda func, **kwargs: func()

        emp = make_employee(row=4, mnv="001")
        gen._generate_one(mock_excel, mock_wb, "TBKQ", "Data", "A", emp)

        retry_config = mock_retry.call_args.kwargs["config"]
        assert retry_config.max_attempts == 3
        assert retry_config.base_delay == 1.0
        assert retry_config.max_delay == 5.0

    @patch("payslip_generator.retry_with_backoff")
    def test_operation_name_includes_employee(self, mock_retry, tmp_path):
        """Verify operation_name includes employee name and MNV."""
        gen = PayslipGenerator(tmp_path, "01/2026")
        mock_excel, mock_wb, new_ws = _make_mock_excel()

        mock_retry.side_effect = lambda func, **kwargs: func()

        emp = make_employee(row=4, mnv="007", name="James Bond")
        gen._generate_one(mock_excel, mock_wb, "TBKQ", "Data", "A", emp)

        op_name = mock_retry.call_args.kwargs["operation_name"]
        assert "James Bond" in op_name
        assert "007" in op_name

    @patch("payslip_generator.retry_with_backoff")
    def test_generate_one_accepts_name_suffix(self, mock_retry, tmp_path):
        """Verify _generate_one uses name_suffix in output path."""
        gen = PayslipGenerator(tmp_path, "01/2026")
        mock_excel, mock_wb, new_ws = _make_mock_excel()

        mock_retry.side_effect = lambda func, **kwargs: func()

        emp = make_employee(row=4, mnv="001", name="Alice")
        result = gen._generate_one(
            mock_excel, mock_wb, "TBKQ", "Data", "A", emp,
            name_suffix="_2",
        )

        # Output path should contain the suffix
        assert result is not None
        assert "_2_" in result.name
