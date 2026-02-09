"""
PDF conversion and password protection for Excel files.

Converts Excel .xlsx files to PDF using:
1. Excel COM (win32com) for high-fidelity PDF rendering
2. pikepdf for password protection

This is a generic utility — not tied to any specific tool.

Usage:
    from office.excel import PdfConverter

    with PdfConverter(output_dir=Path("./output")) as converter:
        pdf_path = converter.convert_to_pdf(
            xlsx_path=Path("report.xlsx"),
            password="secret123",
        )
"""

import logging
from pathlib import Path
from typing import List, Optional

logger = logging.getLogger(__name__)


class PdfConverter:
    """
    Converts Excel .xlsx files to password-protected PDFs.

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
            output_dir: Directory for generated PDF files.
            password_enabled: Whether to apply password protection.
            strip_leading_zeros: Strip leading zeros from passwords.
            cleanup_xlsx: Delete source .xlsx after conversion.
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

    def _init_excel(self) -> None:
        """Initialize Excel COM for PDF export."""
        if self._initialized:
            return

        try:
            import win32com.client as win32

            self._excel = win32.Dispatch("Excel.Application")
            self._excel.Visible = False
            self._excel.DisplayAlerts = False
            self._initialized = True
            logger.info("Excel COM initialized for PDF conversion")
        except Exception as e:
            logger.error(f"Failed to initialize Excel COM: {e}")

    def _cleanup_excel(self) -> None:
        """Quit Excel COM."""
        if self._excel:
            try:
                self._excel.Quit()
            except Exception:
                pass
            self._excel = None
            self._initialized = False

    def convert_to_pdf(
        self,
        xlsx_path: Path,
        password: Optional[str] = None,
        pdf_filename: Optional[str] = None,
    ) -> Optional[Path]:
        """
        Convert a single .xlsx file to a password-protected PDF.

        Args:
            xlsx_path: Path to the .xlsx file.
            password: Password for PDF protection.
            pdf_filename: Custom PDF filename (optional).

        Returns:
            Path to the generated PDF file, or None on failure.
        """
        xlsx_path = Path(xlsx_path)
        if not xlsx_path.exists():
            logger.error(f"XLSX file not found: {xlsx_path}")
            return None

        if pdf_filename:
            pdf_path = self.output_dir / pdf_filename
        else:
            pdf_path = self.output_dir / xlsx_path.with_suffix(".pdf").name

        try:
            # Convert xlsx to PDF via Excel COM
            raw_pdf = self._excel_to_pdf(xlsx_path, pdf_path)
            if raw_pdf is None:
                return None

            # Apply password protection
            if self.password_enabled and password:
                protected_pdf = self._protect_pdf(raw_pdf, password)
                if protected_pdf is None:
                    return raw_pdf
                pdf_path = protected_pdf

            # Cleanup intermediate xlsx
            if self.cleanup_xlsx and xlsx_path.exists():
                xlsx_path.unlink()
                logger.debug(f"Deleted intermediate file: {xlsx_path}")

            return pdf_path

        except Exception as e:
            logger.error(f"PDF conversion failed for {xlsx_path}: {e}")
            return None

    def _excel_to_pdf(self, xlsx_path: Path, pdf_path: Path) -> Optional[Path]:
        """Convert Excel file to PDF using COM ExportAsFixedFormat."""
        if not self._initialized:
            self._init_excel()
        if not self._excel:
            logger.error("Excel COM not available for PDF conversion")
            return None

        wb = None
        try:
            wb = self._excel.Workbooks.Open(str(xlsx_path.resolve()))
            ws = wb.ActiveSheet

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
        """Add password protection to a PDF using pikepdf."""
        try:
            import pikepdf
        except ImportError:
            logger.warning("pikepdf not available. PDF will not be password-protected.")
            return None

        protected_path = pdf_path.with_suffix(".protected.pdf")
        try:
            with pikepdf.open(pdf_path) as pdf:
                pdf.save(
                    str(protected_path),
                    encryption=pikepdf.Encryption(
                        owner=password,
                        user=password,
                        R=6,
                    ),
                )

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
        progress_callback=None,
    ) -> List[dict]:
        """
        Batch convert payslip files to PDF.

        Args:
            items: List of dicts with 'xlsx_path', 'employee' keys.
                   Each employee dict should have 'password' and 'name'.
            progress_callback: Optional callback(current, total, name, skipped).

        Returns:
            Updated list with 'pdf_path' and 'pdf_skipped' added.
        """
        total = len(items)
        success = 0
        failed = 0
        skipped = 0

        for i, item in enumerate(items, 1):
            xlsx_path = item.get("xlsx_path")
            emp = item.get("employee", {})
            name = emp.get("name", "N/A")
            password = emp.get("password", "")

            # Determine expected PDF path for existence check
            if xlsx_path:
                pdf_filename = Path(xlsx_path).with_suffix(".pdf").name
                expected_pdf = self.output_dir / pdf_filename

                # Resume support: skip if PDF already exists
                if expected_pdf.exists():
                    logger.debug(f"[{i}/{total}] Skipping {name} - PDF already exists")
                    item["pdf_path"] = expected_pdf
                    item["pdf_skipped"] = True
                    success += 1
                    skipped += 1
                    if progress_callback:
                        progress_callback(i, total, name, skipped=True)
                    continue

            if not xlsx_path or not Path(xlsx_path).exists():
                logger.warning(f"[{i}/{total}] Skipping {name}: no xlsx file")
                item["pdf_path"] = None
                item["pdf_skipped"] = False
                failed += 1
                if progress_callback:
                    progress_callback(i, total, name, skipped=False)
                continue

            logger.info(f"[{i}/{total}] Converting PDF for {name}")

            pdf_filename = Path(xlsx_path).with_suffix(".pdf").name
            pdf_path = self.convert_to_pdf(
                Path(xlsx_path),
                password=password,
                pdf_filename=pdf_filename,
            )

            item["pdf_path"] = pdf_path
            item["pdf_skipped"] = False
            if pdf_path:
                success += 1
            else:
                failed += 1

            if progress_callback:
                progress_callback(i, total, name, skipped=False)

        logger.info(
            f"PDF conversion complete: {success} success "
            f"({skipped} skipped), {failed} failed"
        )
        return items
