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

            # Fix XLOOKUP formulas in both Data and TBKQ sheets.
            #
            # Why both sheets:
            #   - TBKQ template contains XLOOKUP formulas that look up salary
            #     data from 'bang luong' using B3 (MNV) as the lookup key.
            #   - fix_xlookup_formulas replaces XLOOKUP with INDEX/MATCH and
            #     adds +0 coercion so numeric MNV in 'bang luong' is matched
            #     even when B3 is set as text.
            #
            # Why the fix was needed for Office 365:
            #   - Office 2019 cannot evaluate XLOOKUP → uses cached cell values
            #     (no recalculation) → appears correct.
            #   - Office 365 evaluates XLOOKUP natively → recalculates live →
            #     returns #N/A because B3='6046072' (text) != 6046072 (number).
            #   - Office 365 stores formulas as plain XLOOKUP(...) without the
            #     _xlfn. prefix, so the old pattern never matched on 365.
            # Only trigger a full workbook recalculation when XLOOKUP formulas
            # were actually replaced. On Office 365, CalculateFullRebuild() forces
            # recalculation of ALL sheets (including Data!K:K salary formulas).
            # Those Data-sheet formulas may fail on 365's engine, propagating
            # #N/A into every TBKQ INDEX/MATCH that reads from them.
            #
            # When no XLOOKUP exists (typical case), preserve the Data sheet's
            # cached values from the last XLS save — they are already correct.
            # Per-employee ws.Calculate() in _generate_one() recalculates only
            # the TBKQ sheet after B3 is set, which reads Data's cached values.
            xlookup_fixed = self._fix_xlookup_formulas(wb, data_sheet)
            xlookup_fixed += self._fix_xlookup_formulas(wb, template_sheet)
            if xlookup_fixed > 0:
                logger.debug(
                    f"Replaced {xlookup_fixed} XLOOKUP formula(s) — triggering CalculateFullRebuild"
                )
                excel.CalculateFullRebuild()
            else:
                logger.debug(
                    "No XLOOKUP replacements — skipping CalculateFullRebuild to preserve Data sheet cache"
                )

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

        # Assign B3 the MNV value, matching the type used in 'bang luong' column L
        # so that VLOOKUP(B3, 'bang luong'!...) can find the employee.
        #
        # ROOT CAUSE of #N/A on Office 365:
        #   Excel 365 recalculates VLOOKUP live. If B3 type != 'bang luong' L type,
        #   VLOOKUP exact match returns #N/A. Office 2019 used cached values so the
        #   type mismatch was never evaluated.
        #
        # Strategy: read the first non-empty value from 'bang luong' col L to
        # determine if MNV is stored as text or number, then set B3 to match.
        ws.Range("B3").NumberFormat = "General"

        # COM returns MNV from Data!A as float (e.g. 6046072.0) — convert to int.
        if mnv_value is not None:
            try:
                float_val = float(mnv_value)
                mnv_value = int(float_val) if float_val == int(float_val) else float_val
            except (TypeError, ValueError):
                pass

        # Detect 'bang luong' L column type and set B3 to match
        bl_mnv_is_text = False
        try:
            bl_ws = source_wb.Sheets("bang luong")
            last_bl_row = bl_ws.Cells(bl_ws.Rows.Count, 12).End(-4162).Row
            for r in range(1, min(20, last_bl_row + 1)):
                sample = bl_ws.Cells(r, 12).Value  # col L = index 12
                if sample is not None:
                    bl_mnv_is_text = isinstance(sample, str)
                    logger.debug(
                        f"[GEN-DEBUG] bang luong L{r} sample: value={sample!r}, "
                        f"type={type(sample).__name__} → "
                        f"{'TEXT — will set B3 as text' if bl_mnv_is_text else 'NUMBER — will set B3 as number'}"
                    )
                    break
        except Exception as ex:
            logger.debug(f"[GEN-DEBUG] Could not read bang luong type sample: {ex}")

        if bl_mnv_is_text:
            # 'bang luong' stores MNV as text → B3 must also be text
            ws.Range("B3").NumberFormat = "@"
            ws.Range("B3").Value = str(int(mnv_value)) if isinstance(mnv_value, (int, float)) else str(mnv_value)
            logger.debug(f"[GEN-DEBUG] B3 set as TEXT: {ws.Range('B3').Value!r}")
        else:
            # 'bang luong' stores MNV as number → B3 must also be number
            ws.Range("B3").Value = mnv_value
            logger.debug(f"[GEN-DEBUG] B3 set as NUMBER: {ws.Range('B3').Value!r}")

        # Diagnostic: dump C10 formula and the named range "TBKQ" definition
        # so we can see EXACTLY what VLOOKUP is doing and where it looks.
        try:
            c10_formula = ws.Range("C10").Formula
            logger.debug(f"[GEN-DEBUG] C10 formula: {c10_formula!r}")
        except Exception as ex:
            logger.debug(f"[GEN-DEBUG] Cannot read C10 formula: {ex}")
        try:
            tbkq_range_def = source_wb.Names("TBKQ").RefersTo
            logger.debug(f"[GEN-DEBUG] Named range TBKQ = {tbkq_range_def!r}")
        except Exception as ex:
            logger.debug(f"[GEN-DEBUG] Named range TBKQ not found or error: {ex}")

        # Diagnostic: inspect Data sheet col A type to understand what
        # the TBKQ named range contains as lookup key values.
        try:
            d_row = data_ws.Range(f"{col_mnv}{row}").Value
            d_formula = data_ws.Range(f"{col_mnv}{row}").Formula
            logger.debug(
                f"[GEN-DEBUG] Data!{col_mnv}{row} value={d_row!r} "
                f"({type(d_row).__name__}), formula={d_formula!r}"
            )
        except Exception as ex:
            logger.debug(f"[GEN-DEBUG] Cannot inspect Data sheet: {ex}")

        # Step 2: Full recalculation of TBKQ sheet in source workbook
        ws.Calculate()

        # Log sample cell values from TBKQ after recalculation.
        # COM_ERROR_VALUES are negative ints for #N/A, #REF!, #VALUE! etc.
        # If values here are already #N/A the issue is in B3/VLOOKUP.
        # If values here are correct but PDF is #N/A, issue is in Copy/paste step.
        _com_errors = {-2146826246, -2146826281, -2146826259, -2146826288,
                       -2146826252, -2146826265, -2146826273}
        def _cell_repr(sheet, addr):
            try:
                v = sheet.Range(addr).Value
                if isinstance(v, int) and v in _com_errors:
                    return f"#ERR({v})"
                return repr(v)
            except Exception as ex:
                return f"ERROR({ex})"

        logger.debug(
            f"[GEN-DEBUG] TBKQ values after Calculate(): "
            f"B3={_cell_repr(ws,'B3')}, "
            f"C5={_cell_repr(ws,'C5')}, C8={_cell_repr(ws,'C8')}, "
            f"C10={_cell_repr(ws,'C10')}, D10={_cell_repr(ws,'D10')}"
        )

        # Step 3: Disable auto-calculation BEFORE Copy so the new workbook does
        # NOT immediately recalculate its formulas (which become broken external
        # references after copy on Office 365, producing #N/A).
        #
        # Office 2019: After ws.Copy(), new workbook retains CACHED values from
        #   the source — auto-recalculation did not fire before PasteSpecial.
        # Office 365: Immediately triggers full recalculation after ws.Copy(),
        #   external VLOOKUP refs fail → #N/A captured by PasteSpecial.
        #
        # Fix: suspend calculation around the copy+paste block.
        try:
            excel.Calculation = -4135  # xlCalculationManual
            excel.CalculateBeforeSave = False
        except Exception as e:
            logger.debug(f"[GEN-DEBUG] Could not set Calculation=Manual: {e}")

        # Copy TBKQ sheet to a new workbook
        ws.Copy()
        new_wb = excel.ActiveWorkbook
        new_ws = new_wb.ActiveSheet

        # Log new workbook cell state BEFORE paste — the critical diagnostic.
        # Correct = same values as source TBKQ above.
        # All #ERR = Office 365 recalculated external refs before we got here
        #            (fix above did not work, deeper issue).
        logger.debug(
            f"[GEN-DEBUG] new_ws values BEFORE paste (formula state): "
            f"B3={_cell_repr(new_ws,'B3')}, "
            f"C5={_cell_repr(new_ws,'C5')}, C8={_cell_repr(new_ws,'C8')}, "
            f"C10={_cell_repr(new_ws,'C10')}, D10={_cell_repr(new_ws,'D10')}"
        )

        # Step 4: Paste as values — captures current cell state into literal values
        new_ws.Cells.Copy()
        new_ws.Range("A1").PasteSpecial(Paste=-4163)  # xlPasteValues
        excel.CutCopyMode = False

        # Restore automatic calculation
        try:
            excel.Calculation = -4105  # xlCalculationAutomatic
        except Exception:
            pass

        # Log new workbook AFTER paste — should show correct values now.
        logger.debug(
            f"[GEN-DEBUG] new_ws values AFTER paste (captured values): "
            f"B3={_cell_repr(new_ws,'B3')}, "
            f"C5={_cell_repr(new_ws,'C5')}, C8={_cell_repr(new_ws,'C8')}, "
            f"C10={_cell_repr(new_ws,'C10')}, D10={_cell_repr(new_ws,'D10')}"
        )

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

    def _fix_xlookup_formulas(self, wb, sheet_name: str) -> int:
        """
        Replace all XLOOKUP formulas in the given sheet with INDEX/MATCH
        equivalents so they work in Excel versions that don't support XLOOKUP.

        Returns the number of formulas replaced.
        Delegates to office.excel.utils.fix_xlookup_formulas.
        """
        ws = wb.Sheets(sheet_name)
        return fix_xlookup_formulas(ws, logger=logger)
