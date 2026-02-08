"""
Payslip generator using openpyxl for direct population.

Creates individual payslip Excel files by copying the TBKQ template
and filling cells with employee data from the Data sheet.
No VLOOKUP formulas - all values are written directly.
"""

import logging
import re
import shutil
from pathlib import Path
from typing import Any, Dict, List, Optional

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

logger = logging.getLogger(__name__)


def _col_letter_to_index(col: str) -> int:
    """Column letter to 1-based index for openpyxl."""
    return column_index_from_string(col)


def _cell_to_row_col(cell_ref: str):
    """Parse 'B3' to (row_1based, col_1based) for openpyxl."""
    match = re.match(r"^([A-Z]+)(\d+)$", cell_ref.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    col_str, row_str = match.groups()
    return int(row_str), _col_letter_to_index(col_str)


class PayslipGenerator:
    """
    Generates payslip Excel files from a template.

    Flow:
    1. Use a prepared .xlsx template (TBKQ sheet only)
    2. For each employee, copy template and fill data directly
    3. Save as .xlsx (later converted to PDF)
    """

    def __init__(
        self,
        template_path: Path,
        output_dir: Path,
        cell_mapping: Dict[str, str],
        calc_mapping: Dict[str, str],
        date_str: str,
        filename_pattern: str = "TBKQ_{name}_{mmyyyy}.xlsx",
    ):
        """
        Args:
            template_path: Path to TBKQ template .xlsx file.
            output_dir: Directory for generated payslip files.
            cell_mapping: TBKQ cell -> Data column mapping (e.g., {"B3": "A"}).
            calc_mapping: Calculated cell formulas (e.g., {"D16": "=D17+D21"}).
            date_str: Payroll date in MM/YYYY format.
            filename_pattern: Pattern for output filenames.
        """
        self.template_path = Path(template_path)
        self.output_dir = Path(output_dir)
        self.cell_mapping = cell_mapping
        self.calc_mapping = calc_mapping
        self.date_str = date_str
        self.filename_pattern = filename_pattern

        self.output_dir.mkdir(parents=True, exist_ok=True)

        # Parse date
        parts = date_str.split("/")
        self.month = parts[0] if len(parts) >= 1 else ""
        self.year = parts[1] if len(parts) >= 2 else ""
        self.mmyyyy = self.month + self.year

    def prepare_template(
        self,
        source_xls: Path,
        template_sheet: str = "TBKQ",
    ) -> Path:
        """
        Prepare a clean .xlsx template from the source .xls file.

        Uses win32com (Excel COM) to:
        1. Open the .xls file
        2. Copy the TBKQ sheet to a new workbook
        3. Clear formulas (paste values)
        4. Save as .xlsx template

        This is a one-time operation per run.

        Args:
            source_xls: Path to source Excel file (.xls or .xlsx).
            template_sheet: Sheet name to extract.

        Returns:
            Path to the prepared template .xlsx file.
        """
        template_out = self.output_dir / "_template.xlsx"

        if template_out.exists():
            logger.info(f"Template already prepared: {template_out}")
            return template_out

        logger.info(f"Preparing template from {source_xls}...")

        try:
            import win32com.client as win32
            import pythoncom

            pythoncom.CoInitialize()

            excel = win32.gencache.EnsureDispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            try:
                wb = excel.Workbooks.Open(str(source_xls.resolve()))
                ws = wb.Sheets(template_sheet)

                # Copy sheet to new workbook
                ws.Copy()
                new_wb = excel.ActiveWorkbook

                new_ws = new_wb.ActiveSheet

                # Paste as values to remove formulas
                new_ws.Cells.Copy()
                new_ws.Cells.PasteSpecial(Paste=-4163)  # xlPasteValues
                excel.CutCopyMode = False

                # Delete columns F and G (mapping hints, not needed in output)
                # Column G first (higher index), then F
                new_ws.Columns("G").Delete()
                new_ws.Columns("F").Delete()

                # Clear cells used by macro buttons (K2:L2)
                try:
                    new_ws.Range("I2", "J2").ClearContents()
                except Exception:
                    pass

                # Delete any buttons
                try:
                    for btn in new_ws.Buttons:
                        btn.Delete()
                except Exception:
                    pass

                # Set print area
                new_ws.PageSetup.PrintArea = "$A$1:$E$61"

                # Save as xlsx (no password)
                new_wb.SaveAs(
                    str(template_out.resolve()),
                    FileFormat=51,  # xlOpenXMLWorkbook
                )
                new_wb.Close(SaveChanges=False)

            finally:
                wb.Close(SaveChanges=False)
                excel.Quit()
                pythoncom.CoUninitialize()

            logger.info(f"Template prepared: {template_out}")
            self.template_path = template_out
            return template_out

        except ImportError:
            logger.warning(
                "win32com not available. Creating template from scratch with openpyxl."
            )
            return self._prepare_template_openpyxl(source_xls, template_sheet)

    def _prepare_template_openpyxl(
        self, source_xls: Path, template_sheet: str
    ) -> Path:
        """
        Fallback: prepare template using openpyxl only (for .xlsx sources).

        Limited: won't work with .xls files, won't preserve all formatting.
        """
        template_out = self.output_dir / "_template.xlsx"

        if source_xls.suffix.lower() == ".xls":
            raise RuntimeError(
                f"Cannot process .xls file without win32com. "
                f"Please convert {source_xls} to .xlsx first, or install pywin32."
            )

        wb = load_workbook(str(source_xls))
        # Keep only the template sheet
        for name in wb.sheetnames:
            if name != template_sheet:
                del wb[name]

        ws = wb.active
        # Clear formula cells (replace with None)
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = None

        wb.save(str(template_out))
        wb.close()

        logger.info(f"Template prepared (openpyxl fallback): {template_out}")
        self.template_path = template_out
        return template_out

    def generate_payslip(
        self,
        employee: Dict[str, Any],
    ) -> Optional[Path]:
        """
        Generate a single payslip Excel file.

        Args:
            employee: Employee data dict with 'columns' sub-dict.

        Returns:
            Path to generated .xlsx file, or None on failure.
        """
        mnv = employee.get("mnv", "")
        name = employee.get("name", "")
        columns = employee.get("columns", {})

        # Generate output filename
        safe_name = re.sub(r'[\\/*?:"<>|]', "_", name) if name else mnv
        filename = self.filename_pattern.replace("{name}", safe_name).replace(
            "{mmyyyy}", self.mmyyyy
        )
        # Replace .pdf with .xlsx for the intermediate file
        xlsx_filename = filename.replace(".pdf", ".xlsx")
        output_path = self.output_dir / xlsx_filename

        try:
            # Copy template
            shutil.copy2(str(self.template_path), str(output_path))

            # Open and fill
            wb = load_workbook(str(output_path))
            ws = wb.active

            # Step 1: Fill direct mappings from Data columns
            filled_cells = {}
            for tbkq_cell, data_col in self.cell_mapping.items():
                value = columns.get(data_col)
                if value is not None:
                    row, col = _cell_to_row_col(tbkq_cell)
                    ws.cell(row=row, column=col, value=value)
                    filled_cells[tbkq_cell] = value

            # Step 2: Calculate computed cells
            self._fill_calculated_cells(ws, filled_cells, columns)

            # Step 3: Update date-dependent cells
            self._update_date_cells(ws)

            wb.save(str(output_path))
            wb.close()

            logger.debug(
                f"Generated payslip for {name} (MNV: {mnv}): {output_path}"
            )
            return output_path

        except Exception as e:
            logger.error(f"Failed to generate payslip for {name}: {e}")
            if output_path.exists():
                output_path.unlink()
            return None

    def _fill_calculated_cells(
        self,
        ws,
        filled_cells: Dict[str, Any],
        columns: Dict[str, Any],
    ):
        """
        Fill calculated cells based on calc_mapping.

        Formulas:
        - "COL1+COL2+..." : Sum of Data columns (e.g., "U+W+X+Y")
        - "=CELL1+CELL2" : Sum of other TBKQ cells (e.g., "=D17+D21")
        - "0" or number : Direct value
        """
        # First pass: calculate cells that reference Data columns directly
        for tbkq_cell, formula in self.calc_mapping.items():
            if formula.startswith("="):
                # References other TBKQ cells — handle in second pass
                continue
            if formula.replace(".", "").isdigit():
                # Direct numeric value
                row, col = _cell_to_row_col(tbkq_cell)
                value = float(formula)
                ws.cell(row=row, column=col, value=value)
                filled_cells[tbkq_cell] = value
                continue

            # Sum of Data columns (e.g., "U+W+X+Y")
            col_refs = [c.strip() for c in formula.split("+")]
            total = 0.0
            for col_ref in col_refs:
                val = columns.get(col_ref)
                if val is not None:
                    try:
                        total += float(val)
                    except (ValueError, TypeError):
                        pass

            row, col = _cell_to_row_col(tbkq_cell)
            ws.cell(row=row, column=col, value=total)
            filled_cells[tbkq_cell] = total

        # Second pass: calculate cells that reference other TBKQ cells
        # May need multiple passes for dependent chains
        max_passes = 5
        for _ in range(max_passes):
            unresolved = False
            for tbkq_cell, formula in self.calc_mapping.items():
                if not formula.startswith("="):
                    continue
                if tbkq_cell in filled_cells:
                    continue

                # Parse "=D17+D21" → ["D17", "D21"]
                cell_refs = [c.strip() for c in formula[1:].split("+")]

                # Check if all referenced cells are resolved
                all_resolved = all(ref in filled_cells for ref in cell_refs)
                if not all_resolved:
                    unresolved = True
                    continue

                total = sum(
                    float(filled_cells.get(ref, 0) or 0) for ref in cell_refs
                )
                row, col = _cell_to_row_col(tbkq_cell)
                ws.cell(row=row, column=col, value=total)
                filled_cells[tbkq_cell] = total

            if not unresolved:
                break

    def _update_date_cells(self, ws):
        """Update date-dependent cells in the payslip."""
        if not self.month or not self.year:
            return

        # G1 → title with month (but after F/G deletion, this is now E1)
        # Actually, the template preparation deletes columns F and G,
        # so what was G1 is now E1. However, the title cell reference in
        # config still says "G1". We need to handle this.
        #
        # IMPORTANT: After template prep deletes cols F & G from the template,
        # the remaining columns shift. But we write to the TEMPLATE copy which
        # already has the deletion applied. So we should NOT adjust here —
        # the template already has the right column layout.
        #
        # The title and info cells (A2, A10) are in column A, so they don't shift.
        # We'll update them with the correct month/year.

        # A2 - Info subtitle (update month/year reference)
        a2_val = ws.cell(row=2, column=1).value
        if a2_val and isinstance(a2_val, str):
            # Replace month/year references like "tháng 11/2025"
            updated = re.sub(
                r"tháng\s+\d{1,2}/\d{4}",
                f"tháng {self.month}/{self.year}",
                a2_val,
            )
            ws.cell(row=2, column=1, value=updated)

    def generate_batch(
        self, employees: List[Dict[str, Any]]
    ) -> List[Dict[str, Any]]:
        """
        Generate payslips for all employees.

        Args:
            employees: List of employee data dicts.

        Returns:
            List of result dicts with 'employee', 'xlsx_path', 'success'.
        """
        results = []
        total = len(employees)

        for i, emp in enumerate(employees, 1):
            name = emp.get("name", "N/A")
            mnv = emp.get("mnv", "N/A")

            logger.info(f"[{i}/{total}] Generating payslip for {name} (MNV: {mnv})")

            xlsx_path = self.generate_payslip(emp)

            results.append(
                {
                    "employee": emp,
                    "xlsx_path": xlsx_path,
                    "success": xlsx_path is not None,
                }
            )

        success = sum(1 for r in results if r["success"])
        failed = total - success
        logger.info(
            f"Payslip generation complete: {success} success, {failed} failed"
        )
        return results
