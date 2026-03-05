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

from office.utils.com import com_initialized, get_win32com_client
from office.utils.helpers import create_excel_background, safe_quit_excel

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
        batch_size: int,
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
            batch_size: Max employees per Excel session (0 = all in one session).
                        Restart Excel between batches to prevent COM OOM.

        Returns:
            List of result dicts with 'employee', 'xlsx_path', 'success'.
        """
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

        # Fix B: split employees into chunks and restart Excel between them.
        # batch_size=0 means process all in one session (legacy behaviour).
        effective_batch_size = batch_size if batch_size > 0 else total
        chunks = [
            employees[i:i + effective_batch_size]
            for i in range(0, total, effective_batch_size)
        ]

        if len(chunks) > 1:
            logger.info(
                f"Processing {total} employees in {len(chunks)} batch(es) "
                f"of up to {effective_batch_size} each (COM restart between batches)"
            )

        processed_count = 0  # global index across all chunks

        for chunk_idx, chunk in enumerate(chunks, 1):
            if len(chunks) > 1:
                logger.info(
                    f"[Batch {chunk_idx}/{len(chunks)}] Starting Excel "
                    f"({len(chunk)} employees)"
                )

            com_ctx = com_initialized()
            com_ctx.__enter__()

            excel, was_already_running = create_excel_background()
            excel.Visible = False
            excel.DisplayAlerts = False

            try:
                wb = excel.Workbooks.Open(
                    str(Path(source_xls).resolve()),
                    UpdateLinks=0,
                )

                # Fix XLOOKUP formulas in the TBKQ template sheet ONLY.
                #
                # Why TBKQ only (NOT Data sheet):
                #   The Data sheet holds pre-computed salary values that were
                #   cached when the XLS file was last saved by HR in Excel. Those
                #   cached values are correct and must NOT be recalculated here.
                #
                #   Replacing XLOOKUP in the Data sheet and then calling
                #   CalculateFullRebuild() forces Data sheet formulas to re-evaluate
                #   with the +0 coercion. If the lookup key type in 'bang luong'
                #   mismatches (text vs. number), every Data!K:K cell returns #N/A
                #   and TBKQ formulas that read from Data!K:K also return #N/A.
                #
                # Why TBKQ may still need fixing:
                #   TBKQ may contain XLOOKUP formulas that look up B3 directly
                #   against 'bang luong'. On Office 365 these are evaluated live.
                #   INDEX/MATCH replacement with +0 coercion makes them type-safe.
                #   Per-employee ws.Calculate() in _generate_one() then recalculates
                #   the TBKQ sheet only, reading Data sheet cached values.
                #
                # CalculateFullRebuild is intentionally NOT called. Recalculating the
                # entire workbook would re-evaluate Data sheet formulas whose results
                # are already correct from the XLS save.
                xlookup_fixed = self._fix_xlookup_formulas(wb, template_sheet)
                if xlookup_fixed > 0:
                    logger.debug(
                        f"Replaced {xlookup_fixed} XLOOKUP formula(s) in {template_sheet} — "
                        "per-employee ws.Calculate() will apply them"
                    )
                else:
                    logger.debug(
                        f"No XLOOKUP formulas in {template_sheet} — Data sheet cache preserved"
                    )

                for emp in chunk:
                    processed_count += 1
                    global_i = processed_count
                    mnv = emp.get("mnv", "")
                    name = emp.get("name", "")

                    # Resume support: skip if output already exists
                    output_path = self._build_output_path(emp)
                    if output_path.exists() or output_path.with_suffix(".pdf").exists():
                        logger.debug(
                            f"[{global_i}/{total}] Skipping {name} - output already exists"
                        )
                        results.append({
                            "employee": emp,
                            "xlsx_path": output_path if output_path.exists() else None,
                            "success": True,
                            "skipped": True,
                        })
                        if progress_callback:
                            progress_callback(global_i, total, name, skipped=True)
                        continue

                    logger.info(
                        f"[{global_i}/{total}] Generating payslip for {name} (MNV: {mnv})"
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
                            progress_callback(global_i, total, name, skipped=False)
                    except Exception as e:
                        logger.error(f"Failed to generate payslip for {name}: {e}")
                        results.append({
                            "employee": emp,
                            "xlsx_path": None,
                            "success": False,
                            "skipped": False,
                        })
                        if progress_callback:
                            progress_callback(global_i, total, name, skipped=False)

                    # Fix A: release per-employee COM proxies before next iteration
                    gc.collect()

                wb.Close(SaveChanges=False)
            finally:
                safe_quit_excel(excel, was_already_running)
                del excel
                com_ctx.__exit__(None, None, None)
                gc.collect()
                time.sleep(2)

            if chunk_idx < len(chunks):
                logger.info(
                    f"[Batch {chunk_idx}/{len(chunks)}] Complete. "
                    f"Excel restarted. Starting batch {chunk_idx + 1}/{len(chunks)}."
                )

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

        # Set B3 = MNV, matching the exact type stored in Data!A (the lookup
        # key column used by MATCH($B$3, Data!$A:$A, 0) in TBKQ formulas).
        #
        # Why type matters on Office 365:
        #   Excel 365 recalculates MATCH live. MATCH(text, number_column, 0)
        #   returns #N/A even when values look identical. We read mnv_value
        #   directly from Data!A{row}, so its Python type is authoritative:
        #     str  → B3 must be text  (NumberFormat "@")
        #     int/float → B3 must be a number (NumberFormat "General")
        #
        # COM returns numeric cells as float (e.g. 6046072.0) — convert to int.
        if mnv_value is not None and isinstance(mnv_value, float):
            try:
                int_val = int(mnv_value)
                if mnv_value == int_val:
                    mnv_value = int_val
            except (TypeError, ValueError):
                pass

        ws.Range("B3").NumberFormat = "General"
        if isinstance(mnv_value, str):
            # Data!A stores MNV as text → B3 must also be text
            ws.Range("B3").NumberFormat = "@"
            ws.Range("B3").Value = mnv_value
            logger.debug(f"[GEN-DEBUG] B3 set as TEXT: {ws.Range('B3').Value!r}")
        else:
            # Data!A stores MNV as number → B3 must also be number
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

        # Step 11b: Update date in A10 cell.
        # A10 holds an Excel date value (pywintypes.datetime) displayed with
        # format "mm-yyyy".  It is merged A10:A12, so MergeArea.Cells(1,1) is
        # used for both read and write.  We replace the date with the 1st day
        # of the configured month/year.
        if self.month and self.year:
            try:
                import datetime as _dt
                a10_range = new_ws.Range("A10")
                is_merged = bool(a10_range.MergeCells)
                a10_top = a10_range.MergeArea.Cells(1, 1) if is_merged else a10_range
                a10_val = a10_top.Value
                logger.debug(
                    f"[GEN-DEBUG] A10 raw: merged={is_merged}, "
                    f"value={a10_val!r} (type={type(a10_val).__name__})"
                )
                # Use day=15 to avoid UTC ±day boundary issues on timezones
                # ahead of UTC (e.g. UTC+7).  The mm-yyyy format only displays
                # month and year, so the exact day has no visible effect.
                new_date = _dt.datetime(int(self.year), int(self.month), 15)
                a10_top.Value = new_date
                logger.debug(f"[GEN-DEBUG] A10 updated: {a10_val!r} → {new_date!r}")
            except Exception as e:
                logger.debug(f"[GEN-DEBUG] A10 update failed: {e}")

        # Step 12: Save as .xlsx
        excel.DisplayAlerts = False
        new_wb.SaveAs(
            str(output_path.resolve()),
            FileFormat=51,  # xlOpenXMLWorkbook
        )
        new_wb.Saved = True
        new_wb.Close()

        # Fix A: explicitly release COM proxy references so Python's garbage
        # collector can decrement the RCW ref count immediately, not on some
        # future GC cycle. Without this, 700+ iterations accumulate hundreds
        # of unreleased COM proxies and the Excel process runs OOM.
        del new_ws
        del new_wb
        del data_ws
        del ws

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
