"""
Payslip generator using Excel COM — replicates the VBA macro approach.

For each employee:
  1. Open source workbook with Excel COM
  2. Set TBKQ!B3 = MNV to trigger VLOOKUP/XLOOKUP formulas
  3. Let Excel recalculate
  4. Copy the TBKQ sheet to a new workbook
  5. Paste as values, delete helper columns, clean up
  6. Save as .xlsx

This approach lets Excel handle ALL formula evaluation, guaranteeing
correct values regardless of formula complexity (VLOOKUP, XLOOKUP, etc.).
"""

import gc
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

from core.logger import get_logger
from office.excel.utils import xlookup_to_index_match, fix_xlookup_formulas, XLOOKUP_PATTERN

logger = get_logger(__name__)


class PayslipGenerator:
    """
    Generates payslip Excel files using the VBA macro approach via COM.

    Uses Excel COM to set B3=MNV, recalculate, copy sheet, paste as values.
    """

    def __init__(
        self,
        output_dir: Path,
        date_str: str,
        filename_pattern: str = "TBKQ_{name}_{mmyyyy}",
    ):
        """
        Args:
            output_dir: Directory for generated payslip files.
            date_str: Payroll date in MM/YYYY format.
            filename_pattern: Pattern for output filenames.
        """
        self.output_dir = Path(output_dir)
        self.date_str = date_str
        self.filename_pattern = filename_pattern
        self.output_dir.mkdir(parents=True, exist_ok=True)

        parts = date_str.split("/")
        self.month = parts[0] if len(parts) >= 1 else ""
        self.year = parts[1] if len(parts) >= 2 else ""
        self.mmyyyy = self.month + self.year

    def _build_output_path(self, employee: Dict[str, Any]) -> Path:
        """Build the expected output file path for an employee's payslip."""
        name = employee.get("name", "")
        mnv = employee.get("mnv", "")
        safe_name = re.sub(r'[\\/*?:"<>|]', "_", name) if name else mnv
        filename = (
            self.filename_pattern
            .replace("{name}", safe_name)
            .replace("{mmyyyy}", self.mmyyyy)
            + ".xlsx"
        )
        return self.output_dir / filename

    def generate_batch(
        self,
        employees: List[Dict[str, Any]],
        source_xls: Path,
        template_sheet: str = "TBKQ",
        data_sheet: str = "Data",
        col_mnv: str = "A",
        progress_callback=None,
    ) -> List[Dict[str, Any]]:
        """
        Generate payslips by replicating the VBA macro approach.

        Opens the source workbook once, then for each employee:
        1. Set TBKQ!B3 = MNV (triggers formula recalculation)
        2. Copy TBKQ sheet to a new workbook
        3. Paste as values, clean up, save as .xlsx

        Args:
            employees: List of employee data dicts.
            source_xls: Path to the source Excel file.
            template_sheet: Name of the TBKQ sheet.
            data_sheet: Name of the Data sheet.
            col_mnv: Column letter for MNV in Data sheet.

        Returns:
            List of result dicts with 'employee', 'xlsx_path', 'success'.
        """
        import win32com.client as win32

        results = []
        total = len(employees)

        # Resume optimization: skip Excel COM entirely if all files exist
        needs_generation = any(
            not self._build_output_path(emp).exists()
            and not self._build_output_path(emp).with_suffix(".pdf").exists()
            for emp in employees
        )
        if not needs_generation:
            logger.info("All payslips already exist, skipping Excel COM")
            for i, emp in enumerate(employees, 1):
                output_path = self._build_output_path(emp)
                results.append({
                    "employee": emp,
                    "xlsx_path": output_path if output_path.exists() else None,
                    "success": True,
                    "skipped": True,
                })
                if progress_callback:
                    progress_callback(i, total, emp.get("name", ""), skipped=True)
            return results

        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            wb = excel.Workbooks.Open(
                str(Path(source_xls).resolve()),
                UpdateLinks=0,
            )

            # Fix XLOOKUP formulas in Data sheet (unsupported in Excel 2016)
            # Replace with INDEX/MATCH equivalents before generation
            self._fix_xlookup_formulas(wb, data_sheet)
            excel.CalculateFullRebuild()

            for i, emp in enumerate(employees, 1):
                mnv = emp.get("mnv", "")
                name = emp.get("name", "")

                # Resume support: skip if output already exists
                output_path = self._build_output_path(emp)
                if output_path.exists() or output_path.with_suffix(".pdf").exists():
                    logger.debug(f"[{i}/{total}] Skipping {name} - output already exists")
                    results.append({
                        "employee": emp,
                        "xlsx_path": output_path if output_path.exists() else None,
                        "success": True,
                        "skipped": True,
                    })
                    if progress_callback:
                        progress_callback(i, total, name, skipped=True)
                    continue

                logger.info(
                    f"[{i}/{total}] Generating payslip for {name} (MNV: {mnv})"
                )

                try:
                    xlsx_path = self._generate_one(
                        excel, wb, template_sheet, data_sheet, col_mnv, emp
                    )
                    results.append({
                        "employee": emp,
                        "xlsx_path": xlsx_path,
                        "success": xlsx_path is not None,
                        "skipped": False,
                    })
                    if progress_callback:
                        progress_callback(i, total, name, skipped=False)
                except Exception as e:
                    logger.error(f"Failed to generate payslip for {name}: {e}")
                    results.append({
                        "employee": emp,
                        "xlsx_path": None,
                        "success": False,
                        "skipped": False,
                    })
                    if progress_callback:
                        progress_callback(i, total, name, skipped=False)

            wb.Close(SaveChanges=False)
        finally:
            try:
                excel.Quit()
            except Exception:
                pass
            del excel
            gc.collect()
            time.sleep(2)

        success = sum(1 for r in results if r["success"])
        failed = total - success
        logger.info(
            f"Payslip generation complete: {success} success, {failed} failed"
        )
        return results

    def _generate_one(
        self,
        excel,
        source_wb,
        template_sheet: str,
        data_sheet: str,
        col_mnv: str,
        employee: Dict[str, Any],
    ) -> Optional[Path]:
        """
        Generate a single payslip by setting B3=MNV and copying the sheet.

        Follows the EXACT VBA macro order:
          1. Set B3 = MNV (triggers formula recalculation)
          2. Copy TBKQ sheet to new workbook
          3. Paste as values immediately (breaks external links)
          4. Delete buttons, columns, clean up
          5. SaveAs .xlsx
          6. Close
        """
        mnv = employee.get("mnv", "")
        name = employee.get("name", "")
        row = employee.get("row", 0)

        # Build output filename
        output_path = self._build_output_path(employee)

        # Step 1: Set TBKQ!B3 = MNV from Data sheet
        data_ws = source_wb.Sheets(data_sheet)
        ws = source_wb.Sheets(template_sheet)
        mnv_value = data_ws.Range(f"{col_mnv}{row}").Value

        # Ensure B3 is text to match Data!A column type
        ws.Range("B3").NumberFormat = "@"
        ws.Range("B3").Value = str(mnv_value) if mnv_value else mnv_value

        # Step 2: Recalculate TBKQ sheet
        ws.Calculate()

        # Step 3: Copy TBKQ sheet to a new workbook
        ws.Copy()
        new_wb = excel.ActiveWorkbook
        new_ws = new_wb.ActiveSheet

        # Step 4: IMMEDIATELY paste as values to break external references
        # This must happen FIRST to avoid external link dialogs/hangs
        new_ws.Cells.Copy()
        new_ws.Range("A1").PasteSpecial(Paste=-4163)  # xlPasteValues
        excel.CutCopyMode = False

        # Step 5: Delete macro buttons
        try:
            new_ws.Buttons.Delete()
        except Exception:
            pass

        # Step 6: Clear macro button cells (K2:L2)
        try:
            new_ws.Range("K2", "L2").ClearContents()
        except Exception:
            pass

        # Step 7: Delete columns F and G (helper/mapping columns)
        # Delete G first (higher index) to preserve F's position
        new_ws.Columns("G").Delete()
        new_ws.Columns("F").Delete()

        # Step 8: Collapse outline levels
        try:
            new_ws.Outline.ShowLevels(RowLevels=0, ColumnLevels=1)
        except Exception:
            pass

        # Step 9: Delete named ranges from new workbook
        try:
            for named in list(new_wb.Names):
                named.Delete()
        except Exception:
            pass

        # Step 10: Set print area
        try:
            new_ws.PageSetup.PrintArea = "$A$1:$E$61"
        except Exception:
            pass

        # Step 11: Update date in A2 cell
        a2_val = new_ws.Range("A2").Value
        if a2_val and isinstance(a2_val, str) and self.month and self.year:
            updated = re.sub(
                r"tháng\s+\d{1,2}/\d{4}",
                f"tháng {self.month}/{self.year}",
                a2_val,
            )
            new_ws.Range("A2").Value = updated

        # Step 12: Save as .xlsx
        excel.DisplayAlerts = False
        new_wb.SaveAs(
            str(output_path.resolve()),
            FileFormat=51,  # xlOpenXMLWorkbook
        )
        new_wb.Saved = True
        new_wb.Close()

        logger.debug(f"Generated payslip: {output_path}")
        return output_path

    @staticmethod
    def _xlookup_to_index_match(formula: str) -> str:
        """
        Replace _xlfn.XLOOKUP(lookup, search, result) with
        INDEX(result, MATCH(lookup+0, search, 0)) in a formula string.

        Delegates to office.excel.utils.xlookup_to_index_match.
        """
        return xlookup_to_index_match(formula)

    def _fix_xlookup_formulas(self, wb, data_sheet: str):
        """
        Replace all _xlfn.XLOOKUP formulas in the Data sheet with
        INDEX/MATCH equivalents so they work in Excel versions that
        don't support XLOOKUP (e.g., Excel 2016).

        Delegates to office.excel.utils.fix_xlookup_formulas.
        """
        ws = wb.Sheets(data_sheet)
        fix_xlookup_formulas(ws, logger=logger)
