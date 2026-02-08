"""
Excel reader for payslip data.

Reads employee data from the Data sheet, email template from bodymail,
and TBKQ template structure using xlrd (.xls) or openpyxl (.xlsx).
"""

import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl.utils import column_index_from_string

logger = logging.getLogger(__name__)


def _col_letter_to_index(col: str) -> int:
    """
    Convert column letter to 0-based index.

    Examples: A->0, B->1, Z->25, AA->26, AZ->51
    """
    return column_index_from_string(col) - 1


def _cell_to_row_col(cell_ref: str) -> Tuple[int, int]:
    """
    Parse cell reference like 'B3' to (row_0based, col_0based).

    Args:
        cell_ref: Cell reference (e.g., 'B3', 'AA10').

    Returns:
        Tuple of (row_index_0based, col_index_0based).
    """
    match = re.match(r"^([A-Z]+)(\d+)$", cell_ref.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    col_str, row_str = match.groups()
    return int(row_str) - 1, _col_letter_to_index(col_str)


class ExcelReader:
    """
    Reads payslip-related data from an Excel file.

    Supports both .xls (via xlrd) and .xlsx (via openpyxl).
    """

    def __init__(self, excel_path: Path):
        self.excel_path = Path(excel_path)
        if not self.excel_path.exists():
            raise FileNotFoundError(f"Excel file not found: {self.excel_path}")

        self._ext = self.excel_path.suffix.lower()
        self._workbook = None
        self._is_xls = self._ext in (".xls",)

    def open(self):
        """Open the Excel workbook."""
        if self._is_xls:
            try:
                import xlrd
            except ImportError:
                raise ImportError("xlrd is required for .xls files: pip install xlrd")
            self._workbook = xlrd.open_workbook(str(self.excel_path))
        else:
            import openpyxl

            self._workbook = openpyxl.load_workbook(
                str(self.excel_path), data_only=True, read_only=True
            )
        logger.info(f"Opened Excel file: {self.excel_path}")

    def close(self):
        """Close the workbook."""
        if self._workbook and not self._is_xls:
            self._workbook.close()
        self._workbook = None

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        return False

    def _get_sheet(self, sheet_name: str):
        """Get a sheet by name."""
        if self._workbook is None:
            raise RuntimeError("Workbook not opened. Call open() first.")
        if self._is_xls:
            return self._workbook.sheet_by_name(sheet_name)
        else:
            if sheet_name not in self._workbook.sheetnames:
                raise ValueError(f"Sheet '{sheet_name}' not found")
            return self._workbook[sheet_name]

    def _get_cell_value(self, sheet, row: int, col: int) -> Any:
        """
        Get cell value by 0-based row and col index.

        For xlrd, handles error type cells by returning None.
        """
        if self._is_xls:
            import xlrd

            ctype = sheet.cell_type(row, col)
            if ctype == xlrd.XL_CELL_ERROR:
                return None
            if ctype == xlrd.XL_CELL_EMPTY:
                return None
            value = sheet.cell_value(row, col)
            # Handle dates
            if ctype == xlrd.XL_CELL_DATE:
                try:
                    date_tuple = xlrd.xldate_as_tuple(value, self._workbook.datemode)
                    return value  # Return as float for now
                except Exception:
                    return value
            return value
        else:
            cell = sheet.cell(row=row + 1, column=col + 1)
            return cell.value

    def _get_cell_by_ref(self, sheet, cell_ref: str) -> Any:
        """Get cell value by reference like 'A1', 'B3', etc."""
        row, col = _cell_to_row_col(cell_ref)
        return self._get_cell_value(sheet, row, col)

    def read_employees(
        self,
        data_sheet: str,
        header_row: int,
        start_row: int,
        col_mnv: str,
        col_name: str,
        col_email: str,
        col_password: str,
    ) -> List[Dict[str, Any]]:
        """
        Read employee data from the Data sheet.

        Args:
            data_sheet: Name of the Data sheet.
            header_row: 1-based row number of headers.
            start_row: 1-based row number of first employee.
            col_mnv: Column letter for MNV.
            col_name: Column letter for Name.
            col_email: Column letter for Email.
            col_password: Column letter for Password.

        Returns:
            List of employee dicts with keys: mnv, name, email, password,
            and all other columns keyed by column letter.
        """
        sheet = self._get_sheet(data_sheet)

        # Determine dimensions
        if self._is_xls:
            nrows = sheet.nrows
            ncols = sheet.ncols
        else:
            nrows = sheet.max_row or 0
            ncols = sheet.max_column or 0

        # Read headers to build column letter → header name mapping
        header_row_idx = header_row - 1  # 0-based
        headers = {}
        for c in range(ncols):
            val = self._get_cell_value(sheet, header_row_idx, c)
            if val is not None:
                col_letter = self._index_to_col_letter(c)
                headers[col_letter] = str(val).strip()

        logger.info(f"Data sheet: {nrows} rows, {ncols} cols, {len(headers)} headers")

        # Key column indices
        mnv_idx = _col_letter_to_index(col_mnv)
        name_idx = _col_letter_to_index(col_name)
        email_idx = _col_letter_to_index(col_email)
        pw_idx = _col_letter_to_index(col_password)

        employees = []
        start_row_idx = start_row - 1  # 0-based

        for r in range(start_row_idx, nrows):
            mnv_val = self._get_cell_value(sheet, r, mnv_idx)

            # Skip empty rows (no MNV)
            if mnv_val is None or str(mnv_val).strip() == "":
                continue

            # Build employee dict with ALL column data
            emp = {
                "row": r + 1,  # 1-based row number
                "mnv": self._normalize_mnv(mnv_val),
                "name": str(self._get_cell_value(sheet, r, name_idx) or "").strip(),
                "email": str(self._get_cell_value(sheet, r, email_idx) or "").strip(),
                "password": self._normalize_password(
                    self._get_cell_value(sheet, r, pw_idx)
                ),
            }

            # Read all columns and store by column letter
            emp["columns"] = {}
            for c in range(ncols):
                col_letter = self._index_to_col_letter(c)
                val = self._get_cell_value(sheet, r, c)
                emp["columns"][col_letter] = val

            employees.append(emp)

        logger.info(f"Read {len(employees)} employees from {data_sheet}")
        return employees

    def read_email_template(
        self,
        sheet_name: str,
        body_cells: List[str],
        date_cell: str,
    ) -> Dict[str, Any]:
        """
        Read email body template from bodymail sheet.

        Args:
            sheet_name: Name of the bodymail sheet.
            body_cells: List of cell references to read (e.g., ['A1', 'A3']).
            date_cell: Cell containing date placeholder text.

        Returns:
            Dict with cell reference → value mapping.
        """
        sheet = self._get_sheet(sheet_name)

        template = {}
        for cell_ref in body_cells:
            val = self._get_cell_by_ref(sheet, cell_ref)
            template[cell_ref] = str(val) if val is not None else ""

        logger.info(
            f"Read email template from {sheet_name}: "
            f"{len(template)} cells ({', '.join(body_cells)})"
        )
        return template

    def read_email_subject(
        self, sheet_name: str, subject_cell: str
    ) -> str:
        """
        Read email subject from the TBKQ sheet.

        Args:
            sheet_name: Name of the template sheet (TBKQ).
            subject_cell: Cell reference for subject (e.g., 'G1').

        Returns:
            Email subject string.
        """
        sheet = self._get_sheet(sheet_name)
        val = self._get_cell_by_ref(sheet, subject_cell)
        subject = str(val) if val else ""
        logger.info(f"Email subject from {sheet_name}!{subject_cell}: {subject}")
        return subject

    def read_template_structure(
        self, sheet_name: str
    ) -> Dict[str, Any]:
        """
        Read TBKQ template structure for reference (labels, hints, etc.)

        Returns:
            Dict with 'labels' and 'f_column_hints' for debugging.
        """
        sheet = self._get_sheet(sheet_name)

        if self._is_xls:
            nrows = sheet.nrows
        else:
            nrows = sheet.max_row or 0

        labels = {}
        f_hints = {}

        for r in range(nrows):
            # Column A labels
            a_val = self._get_cell_value(sheet, r, 0)
            if a_val is not None:
                labels[f"A{r + 1}"] = str(a_val)

            # Column F hints (mapping references)
            f_val = self._get_cell_value(sheet, r, 5)
            if f_val is not None:
                f_hints[f"F{r + 1}"] = str(f_val)

        return {"labels": labels, "f_hints": f_hints}

    @staticmethod
    def _normalize_mnv(value: Any) -> str:
        """Normalize MNV to string, preserving leading zeros."""
        if value is None:
            return ""
        if isinstance(value, float) and value == int(value):
            return str(int(value))
        return str(value).strip()

    @staticmethod
    def _normalize_password(value: Any) -> str:
        """Normalize password: strip leading zeros per spec."""
        if value is None:
            return ""
        if isinstance(value, float) and value == int(value):
            pw = str(int(value))
        else:
            pw = str(value).strip()
        return pw.lstrip("0") or "0"

    @staticmethod
    def _index_to_col_letter(index: int) -> str:
        """Convert 0-based column index to letter(s). 0->A, 25->Z, 26->AA."""
        result = ""
        while True:
            result = chr(65 + (index % 26)) + result
            index = index // 26 - 1
            if index < 0:
                break
        return result
