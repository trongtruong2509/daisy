"""
Unit tests for excel_reader.py — ExcelReader (mocked COM).

Covers:
- read_employees with mocked COM worksheet
- Name fallback from 'bang luong' sheet
- Email template reading
- Email subject reading
- Context manager protocol
"""

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch, PropertyMock

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
TOOL_DIR = Path(__file__).resolve().parent.parent.parent
for p in [str(PROJECT_ROOT), str(TOOL_DIR)]:
    if p not in sys.path:
        sys.path.insert(0, p)


class TestExcelReaderReadEmployees:
    """Tests for ExcelReader.read_employees with mocked COM."""

    @patch("excel_reader.ExcelComReader")
    def test_reads_employees(self, MockComReader):
        from excel_reader import ExcelReader

        # Setup mock worksheet
        mock_reader = MagicMock()
        MockComReader.return_value = mock_reader
        ws = MagicMock()
        mock_reader.get_sheet.return_value = ws

        # Mock last row detection: Cells(65536, col).End(xlUp).Row
        end_cell = MagicMock()
        end_cell.Row = 6  # 3 employees: rows 4, 5, 6
        ws.Cells.return_value.End.return_value = end_cell
        ws.Rows.Count = 65536

        # Mock cell values for 3 employees
        def mock_range(cell_ref):
            cell = MagicMock()
            data = {
                "A4": 1, "B4": "Employee 1", "C4": "emp1@co.com", "AZ4": "001",
                "A5": 2, "B5": "Employee 2", "C5": "emp2@co.com", "AZ5": "002",
                "A6": 3, "B6": "Employee 3", "C6": "emp3@co.com", "AZ6": "003",
            }
            cell.Value = data.get(cell_ref)
            return cell

        ws.Range = mock_range

        reader = ExcelReader(Path("dummy.xls"))
        reader._reader = mock_reader

        employees = reader.read_employees(
            data_sheet="Data",
            header_row=2,
            start_row=4,
            col_mnv="A",
            col_name="B",
            col_email="C",
            col_password="AZ",
        )

        assert len(employees) == 3
        assert employees[0]["mnv"] == "1"
        assert employees[0]["name"] == "Employee 1"
        assert employees[0]["email"] == "emp1@co.com"

    @patch("excel_reader.ExcelComReader")
    def test_skips_empty_mnv_rows(self, MockComReader):
        from excel_reader import ExcelReader

        mock_reader = MagicMock()
        MockComReader.return_value = mock_reader
        ws = MagicMock()
        mock_reader.get_sheet.return_value = ws

        end_cell = MagicMock()
        end_cell.Row = 5
        ws.Cells.return_value.End.return_value = end_cell
        ws.Rows.Count = 65536

        def mock_range(cell_ref):
            cell = MagicMock()
            data = {
                "A4": 1, "B4": "Employee 1", "C4": "emp1@co.com", "AZ4": "001",
                "A5": None, "B5": "", "C5": "", "AZ5": "",  # Empty row
            }
            cell.Value = data.get(cell_ref)
            return cell

        ws.Range = mock_range

        reader = ExcelReader(Path("dummy.xls"))
        reader._reader = mock_reader
        employees = reader.read_employees("Data", 2, 4, "A", "B", "C", "AZ")

        assert len(employees) == 1

    @patch("excel_reader.ExcelComReader")
    def test_fill_missing_names_from_bang_luong(self, MockComReader):
        from excel_reader import ExcelReader

        mock_reader = MagicMock()
        MockComReader.return_value = mock_reader

        # Data sheet: employee with missing name (XLOOKUP failed)
        data_ws = MagicMock()
        end_cell = MagicMock()
        end_cell.Row = 4
        data_ws.Cells.return_value.End.return_value = end_cell
        data_ws.Rows.Count = 65536

        def data_range(cell_ref):
            cell = MagicMock()
            data = {
                "A4": 1, "B4": None, "C4": "emp1@co.com", "AZ4": "001",
            }
            cell.Value = data.get(cell_ref)
            return cell

        data_ws.Range = data_range

        # bang luong sheet: has the name
        bl_ws = MagicMock()
        bl_end = MagicMock()
        bl_end.Row = 2
        bl_ws.Cells.side_effect = lambda r, c: MagicMock(
            Value={"1_12": 1.0, "1_13": "Employee One", "2_12": 2.0, "2_13": "Employee Two"}
            .get(f"{r}_{c}"),
            End=MagicMock(return_value=bl_end),
        )

        def get_sheet(name):
            if name == "Data":
                return data_ws
            elif name == "bang luong":
                return bl_ws
            raise Exception(f"Sheet {name} not found")

        mock_reader.get_sheet.side_effect = get_sheet

        reader = ExcelReader(Path("dummy.xls"))
        reader._reader = mock_reader
        employees = reader.read_employees("Data", 2, 4, "A", "B", "C", "AZ")

        assert len(employees) == 1
        assert employees[0]["name"] == "Employee One"


class TestExcelReaderEmailTemplate:
    """Tests for email template reading."""

    @patch("excel_reader.ExcelComReader")
    def test_read_email_template(self, MockComReader):
        from excel_reader import ExcelReader

        mock_reader = MagicMock()
        MockComReader.return_value = mock_reader
        ws = MagicMock()
        mock_reader.get_sheet.return_value = ws

        def mock_range(cell_ref):
            cell = MagicMock()
            data = {
                "A1": "Dear Employee,",
                "A3": "Your payslip for tháng 01/2026.",
                "A5": "Password hint",
            }
            cell.Value = data.get(cell_ref, None)
            return cell

        ws.Range = mock_range

        reader = ExcelReader(Path("dummy.xls"))
        reader._reader = mock_reader

        template = reader.read_email_template(
            sheet_name="bodymail",
            body_cells=["A1", "A3", "A5"],
            date_cell="A3",
        )

        assert len(template) == 3
        assert template["A1"] == "Dear Employee,"
        assert "tháng" in template["A3"]

    @patch("excel_reader.ExcelComReader")
    def test_read_email_subject(self, MockComReader):
        from excel_reader import ExcelReader

        mock_reader = MagicMock()
        MockComReader.return_value = mock_reader
        ws = MagicMock()
        mock_reader.get_sheet.return_value = ws

        cell = MagicMock()
        cell.Value = "Phiếu lương tháng 01/2026"
        ws.Range.return_value = cell

        reader = ExcelReader(Path("dummy.xls"))
        reader._reader = mock_reader

        subject = reader.read_email_subject("TBKQ", "G1")
        assert "Phiếu lương" in subject


class TestExcelReaderContextManager:
    """Tests for context manager protocol."""

    @patch("excel_reader.ExcelComReader")
    def test_context_manager(self, MockComReader):
        from excel_reader import ExcelReader

        mock_reader = MagicMock()
        MockComReader.return_value = mock_reader

        with ExcelReader(Path("dummy.xls")) as reader:
            pass

        mock_reader.open.assert_called_once()
        mock_reader.close.assert_called_once()
