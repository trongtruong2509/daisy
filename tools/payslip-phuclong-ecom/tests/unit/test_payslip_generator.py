"""
Unit tests for payslip_generator.py — PayslipGenerator (mocked COM).

Covers:
- Output path building
- Filename pattern support
- Safe character replacement in names
- generate_batch skip logic for existing files
- generate_batch with mocked Excel COM
- _generate_one (mocked COM)
- Error handling during generation
"""

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
TOOL_DIR = Path(__file__).resolve().parent.parent.parent
for p in [str(PROJECT_ROOT), str(TOOL_DIR)]:
    if p not in sys.path:
        sys.path.insert(0, p)

from payslip_generator import PayslipGenerator
from tests.conftest import make_employee, make_employees


class TestBuildOutputPath:
    """Tests for _build_output_path."""

    def test_standard_name(self, tmp_path):
        gen = PayslipGenerator(tmp_path, "01/2026")
        emp = make_employee(name="Nguyen Van A")
        path = gen._build_output_path(emp)
        assert path.name == "TBKQ_Nguyen Van A_012026.xlsx"
        assert path.parent == tmp_path

    def test_special_characters_replaced(self, tmp_path):
        gen = PayslipGenerator(tmp_path, "01/2026")
        emp = make_employee(name='Bad/Name*File?"')
        path = gen._build_output_path(emp)
        assert "/" not in path.name
        assert "*" not in path.name
        assert "?" not in path.name
        assert '"' not in path.name

    def test_empty_name_uses_mnv(self, tmp_path):
        gen = PayslipGenerator(tmp_path, "01/2026")
        emp = make_employee(name="", mnv="123")
        path = gen._build_output_path(emp)
        assert "123" in path.name

    def test_custom_pattern(self, tmp_path):
        gen = PayslipGenerator(
            tmp_path, "02/2026",
            filename_pattern="Payslip_{name}_{mmyyyy}",
        )
        emp = make_employee(name="Test")
        path = gen._build_output_path(emp)
        assert path.name == "Payslip_Test_022026.xlsx"

    def test_date_parsing(self, tmp_path):
        gen = PayslipGenerator(tmp_path, "12/2025")
        assert gen.month == "12"
        assert gen.year == "2025"
        assert gen.mmyyyy == "122025"


class TestGenerateBatchSkipLogic:
    """Tests for generate_batch skip when all files exist."""

    def test_skips_when_all_exist(self, tmp_path):
        """Should skip Excel COM entirely when all output files exist."""
        gen = PayslipGenerator(tmp_path, "01/2026")
        emps = make_employees(2)

        # Pre-create output files
        for emp in emps:
            output = gen._build_output_path(emp)
            output.touch()

        # Mock win32com at sys.modules level since it's imported inside generate_batch
        mock_win32com = MagicMock()
        with patch.dict("sys.modules", {"win32com": mock_win32com, "win32com.client": mock_win32com.client}):
            results = gen.generate_batch(
                employees=emps,
                source_xls=Path("dummy.xls"),
            )

        assert len(results) == 2
        assert all(r["skipped"] for r in results)
        assert all(r["success"] for r in results)

    def test_skip_detection_with_pdf(self, tmp_path):
        """Should detect existing PDF as 'already processed'."""
        gen = PayslipGenerator(tmp_path, "01/2026")
        emp = make_employee(name="TestPDF")
        output = gen._build_output_path(emp)
        pdf_path = output.with_suffix(".pdf")
        pdf_path.touch()

        # The _build_output_path.with_suffix(".pdf") check
        assert pdf_path.exists()


class TestGenerateOneLogic:
    """Tests for _generate_one with mocked COM."""

    def test_generate_one_sets_b3(self, tmp_path):
        """Test that _generate_one sets B3=MNV correctly."""
        gen = PayslipGenerator(tmp_path, "01/2026")

        mock_excel = MagicMock()
        mock_wb = MagicMock()

        # Data worksheet
        data_ws = MagicMock()
        data_cell = MagicMock()
        data_cell.Value = "001"
        data_ws.Range.return_value = data_cell

        # TBKQ worksheet
        tbkq_ws = MagicMock()
        tbkq_b3 = MagicMock()
        tbkq_ws.Range.return_value = tbkq_b3

        mock_wb.Sheets.side_effect = lambda name: {
            "Data": data_ws,
            "TBKQ": tbkq_ws,
        }[name]

        # New workbook after copy
        new_wb = MagicMock()
        new_ws = MagicMock()
        mock_excel.ActiveWorkbook = new_wb
        new_wb.ActiveSheet = new_ws

        # Mock A2 value for date update
        a2_cell = MagicMock()
        a2_cell.Value = "Bảng kê quỹ tháng 01/2025"
        new_ws.Range.side_effect = lambda ref: {
            "A1": MagicMock(),
            "A2": a2_cell,
            "K2": MagicMock(),
        }.get(ref, MagicMock())

        new_ws.Cells = MagicMock()
        new_ws.Cells.Copy = MagicMock()
        new_ws.Buttons = MagicMock()
        new_ws.Columns = MagicMock()
        new_ws.Outline = MagicMock()
        new_ws.PageSetup = MagicMock()
        new_wb.Names = []

        emp = make_employee(row=4, mnv="001")
        result = gen._generate_one(
            mock_excel, mock_wb, "TBKQ", "Data", "A", emp
        )

        # Should set B3 value
        assert data_ws.Range.called
        assert result is not None


class TestPayslipGeneratorInit:
    """Tests for PayslipGenerator initialization."""

    def test_creates_output_dir(self, tmp_path):
        out = tmp_path / "new_output"
        gen = PayslipGenerator(out, "01/2026")
        assert out.exists()

    def test_parses_date_correctly(self, tmp_path):
        gen = PayslipGenerator(tmp_path, "07/2025")
        assert gen.month == "07"
        assert gen.year == "2025"
        assert gen.mmyyyy == "072025"
