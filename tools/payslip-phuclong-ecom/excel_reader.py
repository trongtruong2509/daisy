"""
Excel reader for payslip data (Excel COM variant).

Uses Excel COM to read employee data from the Data sheet. Since all formula
evaluation is handled by Excel COM during payslip generation, this reader
only needs to extract employee metadata (MNV, Name, Email, Password).
"""

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


def _col_letter_to_index(col: str) -> int:
    """Convert column letter to 1-based index. A=1, B=2, ..., AZ=52."""
    result = 0
    for ch in col.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


class ExcelReader:
    """
    Reads employee metadata and email template from an Excel file via COM.

    Only reads MNV, Name, Email, Password from the Data sheet, and the
    email template from the bodymail sheet. All formula-dependent values
    are resolved by Excel COM during payslip generation.

    If Data sheet formulas return errors (XLOOKUP not supported), falls
    back to reading Name from the 'bang luong' sheet.
    """

    # COM error integer values
    _COM_ERROR_VALUES = {
        -2146826281,  # #DIV/0!
        -2146826246,  # #N/A
        -2146826259,  # #NAME?
        -2146826288,  # #NULL!
        -2146826252,  # #NUM!
        -2146826265,  # #REF!
        -2146826273,  # #VALUE!
    }

    def __init__(self, excel_path: Path):
        self.excel_path = Path(excel_path).resolve()
        if not self.excel_path.exists():
            raise FileNotFoundError(f"Excel file not found: {self.excel_path}")

        self._excel = None
        self._workbook = None

    def open(self):
        """Open the Excel workbook via COM with formula recalculation."""
        import win32com.client as win32

        self._excel = win32.Dispatch("Excel.Application")
        self._excel.Visible = False
        self._excel.DisplayAlerts = False

        self._workbook = self._excel.Workbooks.Open(
            str(self.excel_path),
            ReadOnly=True,
            UpdateLinks=0,
        )

        # Force full recalculation so all formulas (VLOOKUP, XLOOKUP) resolve
        try:
            self._excel.Calculation = -4105  # xlCalculationAutomatic
            self._workbook.Application.CalculateFull()
        except Exception as e:
            logger.warning(f"Formula recalculation warning: {e}")

        logger.info(f"Opened Excel file via COM: {self.excel_path}")

    def close(self):
        """Close the workbook and Excel COM."""
        if self._workbook:
            try:
                self._workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if self._excel:
            try:
                self._excel.Quit()
            except Exception:
                pass
        self._excel = None
        self._workbook = None

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        return False

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
        Read employee metadata from the Data sheet.

        Only reads MNV, Name, Email, and Password columns.
        All salary/payslip data is resolved by Excel COM during generation.

        Args:
            data_sheet: Name of the Data sheet.
            header_row: 1-based row number of headers.
            start_row: 1-based row number of first employee.
            col_mnv: Column letter for MNV.
            col_name: Column letter for Name.
            col_email: Column letter for Email.
            col_password: Column letter for Password.

        Returns:
            List of employee dicts with keys: row, mnv, name, email, password.
        """
        ws = self._workbook.Sheets(data_sheet)

        # Find last row with data in MNV column
        mnv_col_idx = _col_letter_to_index(col_mnv)
        last_row = ws.Cells(ws.Rows.Count, mnv_col_idx).End(-4162).Row  # xlUp

        logger.info(
            f"Data sheet: rows {start_row}-{last_row}, "
            f"columns MNV={col_mnv}, Name={col_name}, Email={col_email}, PW={col_password}"
        )

        employees = []
        for r in range(start_row, last_row + 1):
            mnv_val = ws.Range(f"{col_mnv}{r}").Value
            if mnv_val is None or str(mnv_val).strip() == "":
                continue

            name_val = self._safe_value(ws.Range(f"{col_name}{r}").Value)
            email_val = self._safe_value(ws.Range(f"{col_email}{r}").Value)
            pw_val = self._safe_value(ws.Range(f"{col_password}{r}").Value)

            emp = {
                "row": r,
                "mnv": self._normalize_mnv(mnv_val),
                "name": str(name_val).strip() if name_val else "",
                "email": str(email_val).strip() if email_val else "",
                "password": self._normalize_password(pw_val),
            }
            employees.append(emp)

        # If Data sheet formulas failed (XLOOKUP errors), fill names from bang luong
        self._fill_missing_names(employees)

        logger.info(f"Read {len(employees)} employees from {data_sheet}")
        return employees

    def _safe_value(self, value) -> Any:
        """Return None for COM error values, otherwise return as-is."""
        if isinstance(value, int) and value in self._COM_ERROR_VALUES:
            return None
        return value

    def _fill_missing_names(self, employees: List[Dict[str, Any]]):
        """
        Fill missing employee names by looking up in 'bang luong' sheet.

        When Data sheet uses XLOOKUP formulas that fail, employee names
        will be None. This method reads names from 'bang luong' using MNV.
        """
        missing = [e for e in employees if not e.get("name")]
        if not missing:
            return

        logger.info(
            f"{len(missing)} employees have missing names, "
            f"attempting lookup from 'bang luong' sheet"
        )

        try:
            bl_ws = self._workbook.Sheets("bang luong")
        except Exception:
            logger.warning("Could not access 'bang luong' sheet for name lookup")
            return

        # Build MNV -> Name lookup from bang luong
        # Column L = MNV (index 12), Column M = Name (index 13)
        last_row = bl_ws.Cells(bl_ws.Rows.Count, 12).End(-4162).Row  # xlUp
        mnv_name_map = {}
        for r in range(1, last_row + 1):
            bl_mnv = bl_ws.Cells(r, 12).Value  # Column L
            bl_name = bl_ws.Cells(r, 13).Value  # Column M
            if bl_mnv is not None:
                mnv_str = self._normalize_mnv(bl_mnv)
                if bl_name and isinstance(bl_name, str) and len(bl_name.strip()) >= 2:
                    mnv_name_map[mnv_str] = bl_name.strip()

        for emp in missing:
            name = mnv_name_map.get(emp["mnv"])
            if name:
                emp["name"] = name
                logger.debug(f"Filled name for MNV {emp['mnv']}: {name}")
            else:
                logger.warning(f"Could not find name for MNV {emp['mnv']} in bang luong")

    def read_email_template(
        self,
        sheet_name: str,
        body_cells: List[str],
        date_cell: str,
    ) -> Dict[str, str]:
        """
        Read email body template from bodymail sheet.

        Args:
            sheet_name: Name of the bodymail sheet.
            body_cells: List of cell references to read.
            date_cell: Cell containing date placeholder text.

        Returns:
            Dict with cell reference -> value mapping.
        """
        ws = self._workbook.Sheets(sheet_name)
        template = {}
        for cell_ref in body_cells:
            val = ws.Range(cell_ref).Value
            template[cell_ref] = str(val) if val is not None else ""

        logger.info(
            f"Read email template from {sheet_name}: "
            f"{len(template)} cells ({', '.join(body_cells)})"
        )
        return template

    def read_email_subject(
        self, sheet_name: str, subject_cell: str
    ) -> str:
        """Read email subject from the TBKQ sheet."""
        ws = self._workbook.Sheets(sheet_name)
        val = ws.Range(subject_cell).Value
        subject = str(val) if val else ""
        logger.info(f"Email subject from {sheet_name}!{subject_cell}: {subject}")
        return subject

    @staticmethod
    def _normalize_mnv(value) -> str:
        """Normalize MNV to string, strip leading zeros."""
        if isinstance(value, float):
            value = int(value)
        return str(value).strip().lstrip("0") or "0"

    @staticmethod
    def _normalize_password(value) -> str:
        """Normalize password to string, strip leading zeros."""
        if value is None:
            return ""
        if isinstance(value, float):
            value = int(value)
        return str(value).strip().lstrip("0") or "0"
