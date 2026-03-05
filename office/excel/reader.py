"""
Generic Excel COM reader for Office Automation Foundation.

Provides a context-managed interface to read data from Excel files
via COM automation. Handles workbook lifecycle, formula recalculation,
and safe value extraction.

Usage:
    from office.excel import ExcelComReader

    with ExcelComReader(Path("data.xlsx")) as reader:
        value = reader.read_cell("Sheet1", "A1")
        rows = reader.read_range("Sheet1", start_row=2, end_row=100,
                                 columns={"A": "id", "B": "name"})
"""

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

from office.utils.com import com_initialized, ensure_com_available, get_win32com_client
from office.utils.helpers import create_excel_background, safe_quit_excel
from office.excel.utils import col_letter_to_index, safe_cell_value

logger = logging.getLogger(__name__)


class ExcelComReader:
    """
    Generic Excel file reader via COM automation.

    Opens a workbook in read-only mode, supports formula recalculation,
    and provides methods to read cells, ranges, and sheet metadata.
    """

    def __init__(self, excel_path: Path):
        """
        Args:
            excel_path: Path to the Excel file to read.

        Raises:
            FileNotFoundError: If the file does not exist.
        """
        ensure_com_available()
        self.excel_path = Path(excel_path).resolve()
        if not self.excel_path.exists():
            raise FileNotFoundError(f"Excel file not found: {self.excel_path}")

        self._excel = None
        self._workbook = None
        self._was_already_running = False
        self._com_ctx = None

    def open(self, read_only: bool = True, recalculate: bool = True) -> None:
        """
        Open the Excel workbook via COM.

        Args:
            read_only: Open in read-only mode.
            recalculate: Force full formula recalculation after opening.
        """
        self._com_ctx = com_initialized()
        self._com_ctx.__enter__()

        self._excel, self._was_already_running = create_excel_background()
        self._excel.Visible = False
        self._excel.DisplayAlerts = False

        self._workbook = self._excel.Workbooks.Open(
            str(self.excel_path),
            ReadOnly=read_only,
            UpdateLinks=0,
        )

        if recalculate:
            try:
                self._excel.Calculation = -4105  # xlCalculationAutomatic
                self._workbook.Application.CalculateFull()
            except Exception as e:
                logger.warning(f"Formula recalculation warning: {e}")

        logger.info(f"Opened Excel file via COM: {self.excel_path}")

    def close(self) -> None:
        """Close the workbook and release Excel COM."""
        if self._workbook:
            try:
                self._workbook.Close(SaveChanges=False)
            except Exception:
                pass
        safe_quit_excel(self._excel, self._was_already_running)
        self._excel = None
        self._workbook = None

        if self._com_ctx:
            self._com_ctx.__exit__(None, None, None)
            self._com_ctx = None

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        return False

    @property
    def workbook(self):
        """Access the underlying COM workbook object."""
        return self._workbook

    @property
    def excel_app(self):
        """Access the underlying COM Excel application object."""
        return self._excel

    def get_sheet(self, sheet_name: str):
        """
        Get a worksheet by name.

        Args:
            sheet_name: Name of the worksheet.

        Returns:
            COM worksheet object.
        """
        return self._workbook.Sheets(sheet_name)

    def get_sheet_names(self) -> List[str]:
        """Get list of all sheet names in the workbook."""
        return [self._workbook.Sheets(i).Name
                for i in range(1, self._workbook.Sheets.Count + 1)]

    def read_cell(self, sheet_name: str, cell_ref: str) -> Any:
        """
        Read a single cell value.

        Args:
            sheet_name: Name of the worksheet.
            cell_ref: Cell reference (e.g., "A1", "B3").

        Returns:
            Cell value, or None for COM errors.
        """
        ws = self._workbook.Sheets(sheet_name)
        value = ws.Range(cell_ref).Value
        return safe_cell_value(value)

    def read_cells(self, sheet_name: str, cell_refs: List[str]) -> Dict[str, Any]:
        """
        Read multiple cell values from a sheet.

        Args:
            sheet_name: Name of the worksheet.
            cell_refs: List of cell references (e.g., ["A1", "A3", "A5"]).

        Returns:
            Dict mapping cell reference to value.
        """
        ws = self._workbook.Sheets(sheet_name)
        result = {}
        for ref in cell_refs:
            value = ws.Range(ref).Value
            result[ref] = str(value) if value is not None else ""
        return result

    def get_last_row(self, sheet_name: str, column: str) -> int:
        """
        Find the last row with data in a specific column.

        Args:
            sheet_name: Name of the worksheet.
            column: Column letter (e.g., "A").

        Returns:
            1-based row number of the last data row.
        """
        ws = self._workbook.Sheets(sheet_name)
        col_idx = col_letter_to_index(column)
        return ws.Cells(ws.Rows.Count, col_idx).End(-4162).Row  # xlUp

    def read_range(
        self,
        sheet_name: str,
        start_row: int,
        end_row: int,
        columns: Dict[str, str],
    ) -> List[Dict[str, Any]]:
        """
        Read a range of rows with specified columns.

        Args:
            sheet_name: Name of the worksheet.
            start_row: First row to read (1-based).
            end_row: Last row to read (1-based, inclusive).
            columns: Dict mapping column letter to field name,
                     e.g., {"A": "id", "B": "name", "C": "email"}.

        Returns:
            List of dicts, one per row, with field names as keys.
        """
        ws = self._workbook.Sheets(sheet_name)
        rows = []

        for r in range(start_row, end_row + 1):
            row_data = {"row": r}
            for col_letter, field_name in columns.items():
                value = ws.Range(f"{col_letter}{r}").Value
                row_data[field_name] = safe_cell_value(value)
            rows.append(row_data)

        return rows

    def recalculate(self) -> None:
        """Force full recalculation of all formulas."""
        try:
            self._workbook.Application.CalculateFull()
        except Exception as e:
            logger.warning(f"Recalculation warning: {e}")
