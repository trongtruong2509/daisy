## Excel Automation Requirements for Daisy Automation Platform

### Rationale

Reading data from complex Excel workbooks that contain VLOOKUP, XLOOKUP, and other formula-dependent calculations requires COM automation — openpyxl cannot evaluate formulas. Centralising Excel COM logic in `office/excel/` prevents scattered `win32com.client.Dispatch()` calls across tools and ensures consistent resource lifecycle, formula recalculation, and safe cleanup.

### Requirements

#### REQ-XL-01: `ExcelComReader` is the standard Excel reading interface

All Excel COM reading must go through `office.excel.reader.ExcelComReader`. Tools must not call `win32com.client.Dispatch("Excel.Application")` or `win32com.client.GetObject()` directly (see REQ-COM-03).

```python
from office.excel import ExcelComReader

with ExcelComReader(Path("data.xlsx")) as reader:
    value = reader.read_cell("Sheet1", "A1")
    rows  = reader.read_range("Sheet1", start_row=2, end_row=100,
                               columns={"A": "id", "B": "name"})
```

#### REQ-XL-02: Context manager protocol is mandatory

All use of `ExcelComReader` and `PdfConverter` must follow the `with` statement (context manager) protocol. Calling `open()` / `close()` outside of a `with` block is forbidden in production code (see REQ-COM-07).

#### REQ-XL-03: Formula recalculation on open

`ExcelComReader.open()` must trigger a full formula recalculation (`CalculateFull()`) after opening the workbook. This ensures formula-dependent values (VLOOKUP, XLOOKUP, etc.) are correct regardless of the saved state of the file. Recalculation may be disabled via `recalculate=False` for pure data reads where formulas are not relevant.

#### REQ-XL-04: Read-only by default

`ExcelComReader.open()` must default to `read_only=True`. Workbooks opened for inspection must never be modified. Only dedicated write-path classes (e.g. `PayslipGenerator`) may open files with `ReadOnly=False`.

#### REQ-XL-05: Safe cell value extraction via `safe_cell_value()`

Raw COM cell values may be `None`, COM date objects, floats, or integers depending on the cell type. `office.excel.utils.safe_cell_value(cell)` must normalise these values to Python-native types. Tools must use this utility rather than accessing `.Value` directly.

#### REQ-XL-06: Column letter to index via `col_letter_to_index()`

Any code that maps Excel column letters (e.g. `"AZ"`) to 1-based column indices must use `office.excel.utils.col_letter_to_index()`. Inline character arithmetic for column indexing is forbidden.

#### REQ-XL-07: `ExcelComReader` raises `FileNotFoundError` on missing input

If the target file does not exist, `ExcelComReader.__init__()` must raise `FileNotFoundError` with a clear message. The error must be raised at construction time — not deferred to `open()`.

#### REQ-XL-08: `get_or_create_excel()` helper manages application lifecycle

`office.excel.utils.helpers.get_or_create_excel()` (or equivalent) must implement the detect-or-create pattern described in REQ-COM-08. `ExcelComReader` and `PdfConverter` must both use this helper in their `__enter__` methods. The returned flag `was_already_running` must be stored and used in `__exit__` to decide whether to call `excel.Quit()`.

#### REQ-XL-09: `DisplayAlerts = False` and `Visible = False`

All COM Excel instances created by the platform must have `DisplayAlerts = False` and `Visible = False` set immediately after acquisition. This prevents modal dialogs from blocking the automation process and keeps background Excel instances hidden from the user.

#### REQ-XL-10: `ExcelComReader.read_range()` returns a list of dicts

The `read_range()` method must return `list[dict[str, Any]]`, where each dict maps the caller-supplied column label (not the Excel column letter) to the cell value. This decouples tool code from spreadsheet column positions.

```python
rows = reader.read_range(
    sheet_name="Data",
    start_row=4,
    end_row=200,
    columns={"A": "employee_id", "B": "name", "C": "email"},
)
# rows[0] == {"employee_id": "EMP001", "name": "Nguyen Van A", "email": "a@company.com"}
```

Rows where all mapped columns are empty must be omitted from the result.

### Non-functional constraints

- `office/excel/` modules must only be imported on Windows where `pywin32` is available. Importing on a non-Windows system must raise a clear `ImportError` with an actionable message.
- Unit tests must mock `ExcelComReader` at the boundary. They must not require a live Excel installation.
- Integration tests that launch Excel COM must be tagged `@pytest.mark.integration` and excluded from the default test run.
