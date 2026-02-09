"""
Excel-specific module for Office Automation Foundation.

Provides abstraction over Excel COM automation for reading,
writing, and converting Excel files.
"""

from office.excel.reader import ExcelComReader
from office.excel.converter import PdfConverter
from office.excel.utils import (
    col_letter_to_index,
    index_to_col_letter,
    safe_cell_value,
    normalize_numeric_string,
    COM_ERROR_VALUES,
)

__all__ = [
    "ExcelComReader",
    "PdfConverter",
    "col_letter_to_index",
    "index_to_col_letter",
    "safe_cell_value",
    "normalize_numeric_string",
    "COM_ERROR_VALUES",
]
