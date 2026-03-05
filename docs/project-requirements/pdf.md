## PDF Conversion Requirements for Daisy Automation Platform

### Rationale

Payslip and report distribution requires PDF files rather than Excel workbooks so recipients cannot alter the data. Excel's built-in export (via COM) produces the highest-fidelity output — matching printer layout, page breaks, and header/footer settings exactly. Password protection prevents unauthorised access to sensitive payslip data. Centralising this logic in `office/excel/converter.py` means tools never need to know about COM export APIs or `pikepdf` internals.

### Requirements

#### REQ-PDF-01: `PdfConverter` is the sole PDF generation interface

All Excel-to-PDF conversion must go through `office.excel.converter.PdfConverter`. Tools must not call `workbook.ExportAsFixedFormat()`, `pikepdf`, or any other PDF library directly.

```python
from office.excel.converter import PdfConverter

with PdfConverter(output_dir=Path("./output")) as converter:
    pdf_path = converter.convert_to_pdf(
        xlsx_path=Path("payslip_EMP001.xlsx"),
        password="012345",
    )
```

#### REQ-PDF-02: Context manager protocol is mandatory

`PdfConverter` must be used exclusively via the `with` statement (see REQ-COM-07). Calling `_init_excel()` / `_cleanup_excel()` directly in production code is forbidden.

#### REQ-PDF-03: Excel COM renders the PDF

`PdfConverter` must use Excel COM (`ExportAsFixedFormat`) to export the workbook to PDF. This ensures formula-driven formatting, conditional formatting, and page layout are preserved identically to what the user sees in Excel. Alternative libraries (e.g. `LibreOffice`, `openpyxl` PDF export) must not be used.

#### REQ-PDF-04: Password protection via `pikepdf`

When `password_enabled=True` (the default), `PdfConverter.convert_to_pdf()` must apply an owner password to the generated PDF using `pikepdf`. The password is the employee's identifier (e.g. employee number) with an option to strip leading zeros (`strip_leading_zeros=True`). The password must not be embedded in the PDF filename or stored in a log file.

```python
# password "00123" with strip_leading_zeros=True → applied password "123"
# password "45678" with strip_leading_zeros=True → applied password "45678"
```

#### REQ-PDF-05: XLSX source cleanup after conversion

When `cleanup_xlsx=True` (the default), the source `.xlsx` file must be deleted after a successful PDF conversion. Cleanup must not occur if the conversion fails. This prevents accumulation of intermediate files.

#### REQ-PDF-06: Output filename derived from source filename

The output PDF filename must be the source `.xlsx` filename with the extension changed to `.pdf`. Tools must not pass an explicit output filename — the naming convention is enforced by `PdfConverter`.

```
payslip_EMP001.xlsx  →  payslip_EMP001.pdf
```

#### REQ-PDF-07: `convert_batch()` for multi-file conversion

`PdfConverter` must expose a `convert_batch(items)` method accepting a list of `(xlsx_path, password)` tuples. It must call `convert_to_pdf()` for each item, collect successes and failures, and return a summary without raising on individual item failure. This allows the caller to continue processing remaining payslips even if one conversion fails.

#### REQ-PDF-08: `get_or_create_excel()` lifecycle compliance

`PdfConverter` must use `office.utils.helpers.get_or_create_excel()` (or the equivalent helper) in `__enter__` and must respect the `was_already_running` flag in `__exit__` (see REQ-COM-08 and REQ-XL-08).

#### REQ-PDF-09: Background Excel instance — hidden and quiet

The Excel instance used for PDF export must have `Visible = False` and `DisplayAlerts = False` set immediately after acquisition (see REQ-XL-09). The user's existing Excel windows must remain unaffected.

### Non-functional constraints

- `pikepdf` is a required dependency. If it is not installed, `PdfConverter` must raise `ImportError` with an actionable message on construction.
- `pywin32` / `pythoncom` must be available (Windows only). Non-Windows environments must receive a clear `ImportError`.
- Unit tests must mock `PdfConverter` at the boundary. Integration tests must be tagged `@pytest.mark.integration`.
- Generated PDFs must pass basic integrity checks (i.e. `pikepdf.open()` without error) as part of integration test assertions.
