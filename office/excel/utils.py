"""
Shared utility functions for Excel COM operations.

Provides column conversion, value normalization, and COM error handling
used across the Excel abstraction layer.
"""

import re
from typing import Any, Optional, Set

# COM error integer values returned by Excel for formula errors
COM_ERROR_VALUES: Set[int] = {
    -2146826281,  # #DIV/0!
    -2146826246,  # #N/A
    -2146826259,  # #NAME?
    -2146826288,  # #NULL!
    -2146826252,  # #NUM!
    -2146826265,  # #REF!
    -2146826273,  # #VALUE!
}

# Regex to replace XLOOKUP(lookup, search_range, result_range)
# with INDEX(result_range, MATCH(lookup, search_range, 0))
#
# Matches both forms:
#   _xlfn.XLOOKUP(...)  — compatibility prefix written by Office 2019/2016
#   XLOOKUP(...)        — native form used by Office 365 / Microsoft 365
XLOOKUP_PATTERN = re.compile(
    r"(?:_xlfn\.)?XLOOKUP\("
    r"([^,]+),"       # lookup_value
    r"([^,]+),"       # search_range
    r"([^,\)]+)"      # result_range
    r"\)"
)


def col_letter_to_index(col: str) -> int:
    """
    Convert column letter(s) to 1-based index.

    Examples:
        A -> 1, B -> 2, Z -> 26, AA -> 27, AZ -> 52
    """
    result = 0
    for ch in col.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def index_to_col_letter(index: int) -> str:
    """
    Convert 1-based column index to letter(s).

    Examples:
        1 -> A, 2 -> B, 26 -> Z, 27 -> AA, 52 -> AZ
    """
    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def safe_cell_value(value: Any) -> Any:
    """
    Return None for COM error values, otherwise return as-is.

    Excel COM returns negative integers for formula errors (#N/A, #REF!, etc.).
    This function converts those error codes to None for safe processing.
    """
    if isinstance(value, int) and value in COM_ERROR_VALUES:
        return None
    return value


def normalize_numeric_string(value: Any, strip_leading_zeros: bool = True) -> str:
    """
    Normalize a numeric value to a string representation.

    Handles float-to-int conversion (e.g., 123.0 -> "123") and
    optional stripping of leading zeros.

    Args:
        value: The value to normalize (can be float, int, str, or None).
        strip_leading_zeros: Whether to strip leading zeros.

    Returns:
        Normalized string, or empty string if value is None.
    """
    if value is None:
        return ""
    if isinstance(value, float):
        value = int(value)
    result = str(value).strip()
    if strip_leading_zeros:
        result = result.lstrip("0") or "0"
    return result


def xlookup_to_index_match(formula: str) -> str:
    """
    Replace _xlfn.XLOOKUP(lookup, search, result) with
    INDEX(result, MATCH(lookup+0, search, 0)) in a formula string.

    The +0 forces type coercion (text -> number) so MATCH works
    when lookup is text but search column contains numbers.

    Handles multiple XLOOKUP calls in the same formula.
    """
    def _replacer(m):
        lookup = m.group(1)
        search = m.group(2)
        result = m.group(3)
        return f"INDEX({result},MATCH({lookup}+0,{search},0))"

    return XLOOKUP_PATTERN.sub(_replacer, formula)


def fix_xlookup_formulas(worksheet, logger=None) -> int:
    """
    Replace all _xlfn.XLOOKUP formulas in a worksheet with
    INDEX/MATCH equivalents for compatibility with older Excel versions.

    Args:
        worksheet: COM worksheet object.
        logger: Optional logger for progress messages.

    Returns:
        Number of formulas fixed.
    """
    used = worksheet.UsedRange
    row_count = used.Rows.Count
    col_count = used.Columns.Count
    start_row = used.Row
    start_col = used.Column
    fixed_count = 0

    for r in range(start_row, start_row + row_count):
        for c in range(start_col, start_col + col_count):
            cell = worksheet.Cells(r, c)
            formula = cell.Formula
            if formula and isinstance(formula, str) and "XLOOKUP" in formula.upper():
                new_formula = xlookup_to_index_match(formula)
                cell.Formula = new_formula
                fixed_count += 1

    if logger:
        if fixed_count > 0:
            logger.info(f"Replaced {fixed_count} XLOOKUP formulas with INDEX/MATCH")
        else:
            logger.debug("No XLOOKUP formulas found")

    return fixed_count
