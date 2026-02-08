"""Unit tests for excel_reader module."""

import sys
from pathlib import Path

import pytest

TOOL_DIR = Path(__file__).resolve().parent.parent
PROJECT_ROOT = TOOL_DIR.parent.parent
sys.path.insert(0, str(TOOL_DIR))

from excel_reader import ExcelReader, _col_letter_to_index, _cell_to_row_col

SAMPLE_EXCEL = PROJECT_ROOT / "excel-files" / "TBKQ-phuclong.xls"


class TestColumnConversion:
    """Test column letter/index conversion."""

    def test_a_is_0(self):
        assert _col_letter_to_index("A") == 0

    def test_b_is_1(self):
        assert _col_letter_to_index("B") == 1

    def test_z_is_25(self):
        assert _col_letter_to_index("Z") == 25

    def test_aa_is_26(self):
        assert _col_letter_to_index("AA") == 26

    def test_az_is_51(self):
        assert _col_letter_to_index("AZ") == 51


class TestCellParsing:
    """Test cell reference parsing."""

    def test_a1(self):
        row, col = _cell_to_row_col("A1")
        assert row == 0
        assert col == 0

    def test_b3(self):
        row, col = _cell_to_row_col("B3")
        assert row == 2
        assert col == 1

    def test_az2(self):
        row, col = _cell_to_row_col("AZ2")
        assert row == 1
        assert col == 51

    def test_invalid(self):
        with pytest.raises(ValueError):
            _cell_to_row_col("123")


class TestIndexToColLetter:
    """Test index-to-column-letter conversion."""

    def test_basic_letters(self):
        assert ExcelReader._index_to_col_letter(0) == "A"
        assert ExcelReader._index_to_col_letter(1) == "B"
        assert ExcelReader._index_to_col_letter(25) == "Z"

    def test_double_letters(self):
        assert ExcelReader._index_to_col_letter(26) == "AA"
        assert ExcelReader._index_to_col_letter(27) == "AB"
        assert ExcelReader._index_to_col_letter(51) == "AZ"


@pytest.mark.skipif(
    not SAMPLE_EXCEL.exists(),
    reason="Sample Excel file not available"
)
class TestExcelReaderWithSample:
    """Integration tests using the sample Excel file."""

    def test_open_close(self):
        reader = ExcelReader(SAMPLE_EXCEL)
        reader.open()
        assert reader._workbook is not None
        reader.close()

    def test_context_manager(self):
        with ExcelReader(SAMPLE_EXCEL) as reader:
            assert reader._workbook is not None

    def test_read_employees(self):
        with ExcelReader(SAMPLE_EXCEL) as reader:
            employees = reader.read_employees(
                data_sheet="Data",
                header_row=2,
                start_row=4,
                col_mnv="A",
                col_name="B",
                col_email="C",
                col_password="AZ",
            )
        assert len(employees) >= 1
        emp = employees[0]
        assert emp["mnv"] == "6046072"
        assert emp["email"] == "tht.tts@gmail.com"
        assert "columns" in emp

    def test_read_email_template(self):
        with ExcelReader(SAMPLE_EXCEL) as reader:
            template = reader.read_email_template(
                sheet_name="bodymail",
                body_cells=["A1", "A3", "A5", "A7"],
                date_cell="A3",
            )
        assert "A1" in template
        assert "Kính gửi" in template["A1"]

    def test_read_email_subject(self):
        with ExcelReader(SAMPLE_EXCEL) as reader:
            subject = reader.read_email_subject(
                sheet_name="TBKQ",
                subject_cell="G1",
            )
        assert subject  # non-empty
        assert "THÔNG BÁO" in subject or "PHIẾU LƯƠNG" in subject

    def test_read_template_structure(self):
        with ExcelReader(SAMPLE_EXCEL) as reader:
            structure = reader.read_template_structure("TBKQ")
        assert "labels" in structure
        assert "f_hints" in structure
        # Should have F-column hints (mapping references)
        f_hints = structure["f_hints"]
        assert len(f_hints) > 0

    def test_normalize_mnv(self):
        assert ExcelReader._normalize_mnv(6046072.0) == "6046072"
        assert ExcelReader._normalize_mnv("6046072") == "6046072"
        assert ExcelReader._normalize_mnv(None) == ""

    def test_normalize_password(self):
        assert ExcelReader._normalize_password(6046072.0) == "6046072"
        assert ExcelReader._normalize_password("006046072") == "6046072"
        assert ExcelReader._normalize_password("0000") == "0"
        assert ExcelReader._normalize_password(None) == ""

    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            ExcelReader(Path("/nonexistent/file.xls"))
