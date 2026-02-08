"""
PDF conversion and password protection for payslip files.

Converts Excel .xlsx payslips to PDF using:
1. Excel COM (win32com) for high-fidelity PDF conversion
2. pikepdf for password protection
"""

import logging
from pathlib import Path
from typing import List, Optional

logger = logging.getLogger(__name__)


class PdfConverter:
    """
    Converts payslip .xlsx files to password-protected PDFs.

    Uses Excel COM for rendering (preserves formatting) and
    pikepdf for password protection.
    """

    def __init__(
        self,
        output_dir: Path,
        password_enabled: bool = True,
        strip_leading_zeros: bool = True,
        cleanup_xlsx: bool = True,
    ):
        """
        Args:
            output_dir: Directory for PDF output files.
            password_enabled: Whether to apply password protection.
            strip_leading_zeros: Strip leading zeros from passwords.
            cleanup_xlsx: Remove intermediate .xlsx files after conversion.
        """
        self.output_dir = Path(output_dir)
        self.password_enabled = password_enabled
        self.strip_leading_zeros = strip_leading_zeros
        self.cleanup_xlsx = cleanup_xlsx

        self.output_dir.mkdir(parents=True, exist_ok=True)

        self._excel = None
        self._initialized = False

    def __enter__(self):
        self._init_excel()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._cleanup_excel()
        return False

    def _init_excel(self):
        """Initialize Excel COM application for PDF export."""
        if self._initialized:
            return

        try:
            import win32com.client as win32
            import pythoncom

            pythoncom.CoInitialize()
            self._excel = win32.gencache.EnsureDispatch("Excel.Application")
            self._excel.Visible = False
            self._excel.DisplayAlerts = False
            self._initialized = True
            logger.info("Excel COM initialized for PDF conversion")
        except ImportError:
            logger.warning(
                "win32com not available. PDF conversion will not work. "
                "Install pywin32 on Windows."
            )
        except Exception as e:
            logger.error(f"Failed to initialize Excel COM: {e}")

    def _cleanup_excel(self):
        """Quit Excel COM application."""
        if self._excel:
            try:
                self._excel.Quit()
            except Exception:
                pass
            self._excel = None
            self._initialized = False

            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass

            logger.debug("Excel COM cleaned up")

    def convert_to_pdf(
        self,
        xlsx_path: Path,
        password: Optional[str] = None,
        pdf_filename: Optional[str] = None,
    ) -> Optional[Path]:
        """
        Convert a single .xlsx file to a password-protected PDF.

        Args:
            xlsx_path: Path to the .xlsx payslip file.
            password: Password for PDF protection (if enabled).
            pdf_filename: Custom PDF filename. Defaults to same name with .pdf.

        Returns:
            Path to the generated PDF file, or None on failure.
        """
        xlsx_path = Path(xlsx_path)
        if not xlsx_path.exists():
            logger.error(f"XLSX file not found: {xlsx_path}")
            return None

        # Determine PDF path
        if pdf_filename:
            pdf_path = self.output_dir / pdf_filename
        else:
            pdf_path = self.output_dir / xlsx_path.with_suffix(".pdf").name

        try:
            # Step 1: Convert xlsx to PDF via Excel COM
            raw_pdf = self._excel_to_pdf(xlsx_path, pdf_path)
            if raw_pdf is None:
                return None

            # Step 2: Apply password protection
            if self.password_enabled and password:
                protected_pdf = self._protect_pdf(raw_pdf, password)
                if protected_pdf is None:
                    return raw_pdf  # Return unprotected if protection fails
                pdf_path = protected_pdf

            # Step 3: Cleanup intermediate xlsx
            if self.cleanup_xlsx and xlsx_path.exists():
                xlsx_path.unlink()
                logger.debug(f"Deleted intermediate file: {xlsx_path}")

            return pdf_path

        except Exception as e:
            logger.error(f"PDF conversion failed for {xlsx_path}: {e}")
            return None

    def _excel_to_pdf(self, xlsx_path: Path, pdf_path: Path) -> Optional[Path]:
        """
        Convert Excel file to PDF using COM.

        Uses ExportAsFixedFormat for high-fidelity conversion.
        """
        if not self._initialized:
            self._init_excel()

        if not self._excel:
            logger.error("Excel COM not available for PDF conversion")
            return None

        wb = None
        try:
            wb = self._excel.Workbooks.Open(str(xlsx_path.resolve()))
            ws = wb.ActiveSheet

            # ExportAsFixedFormat: Type=0 (PDF)
            ws.ExportAsFixedFormat(
                Type=0,  # xlTypePDF
                Filename=str(pdf_path.resolve()),
                Quality=0,  # xlQualityStandard
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )

            logger.debug(f"Converted to PDF: {pdf_path}")
            return pdf_path

        except Exception as e:
            logger.error(f"Excel-to-PDF conversion failed: {e}")
            return None

        finally:
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass

    def _protect_pdf(self, pdf_path: Path, password: str) -> Optional[Path]:
        """
        Add password protection to a PDF using pikepdf.

        The password is applied as user password (required to open).
        """
        try:
            import pikepdf
        except ImportError:
            logger.warning(
                "pikepdf not available. PDF will not be password-protected. "
                "Install pikepdf: pip install pikepdf"
            )
            return None

        try:
            # Temp path for the protected version
            protected_path = pdf_path.with_suffix(".protected.pdf")

            with pikepdf.open(pdf_path) as pdf:
                pdf.save(
                    str(protected_path),
                    encryption=pikepdf.Encryption(
                        owner=password,
                        user=password,
                        R=6,  # AES-256 encryption
                    ),
                )

            # Replace original with protected version
            pdf_path.unlink()
            protected_path.rename(pdf_path)

            logger.debug(f"Password-protected PDF: {pdf_path}")
            return pdf_path

        except Exception as e:
            logger.error(f"PDF password protection failed: {e}")
            if protected_path.exists():
                protected_path.unlink()
            return None

    def convert_batch(
        self,
        items: List[dict],
    ) -> List[dict]:
        """
        Batch convert multiple payslip files to PDF.

        Args:
            items: List of dicts with 'xlsx_path', 'employee' keys.
                   Each employee dict should have 'password' and 'name'.

        Returns:
            Updated list with 'pdf_path' added to each dict.
        """
        total = len(items)
        success = 0
        failed = 0

        for i, item in enumerate(items, 1):
            xlsx_path = item.get("xlsx_path")
            emp = item.get("employee", {})
            name = emp.get("name", "N/A")
            password = emp.get("password", "")

            if not xlsx_path or not Path(xlsx_path).exists():
                logger.warning(
                    f"[{i}/{total}] Skipping {name}: no xlsx file"
                )
                item["pdf_path"] = None
                failed += 1
                continue

            logger.info(f"[{i}/{total}] Converting PDF for {name}")

            # Build PDF filename from xlsx name
            pdf_filename = Path(xlsx_path).with_suffix(".pdf").name

            pdf_path = self.convert_to_pdf(
                Path(xlsx_path),
                password=password,
                pdf_filename=pdf_filename,
            )

            item["pdf_path"] = pdf_path
            if pdf_path:
                success += 1
            else:
                failed += 1

        logger.info(
            f"PDF conversion complete: {success} success, {failed} failed"
        )
        return items
