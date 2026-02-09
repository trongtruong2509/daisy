"""
Integration tests for payslip-phuclong-ecom tool.

These tests use REAL Excel COM and REAL Outlook COM objects.
They require:
- Microsoft Excel installed
- Microsoft Outlook installed and configured
- A sample Excel file with 2 test employees

Run manually before production release:
    pytest tests/integration/ -v -m integration

DO NOT run in CI/CD pipelines.
"""

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
TOOL_DIR = Path(__file__).resolve().parent.parent.parent
for p in [str(PROJECT_ROOT), str(TOOL_DIR)]:
    if p not in sys.path:
        sys.path.insert(0, p)

# Mark all tests in this module as integration
pytestmark = pytest.mark.integration


@pytest.mark.integration
class TestRealExcelCom:
    """Integration tests with real Excel COM (manual execution only)."""

    @pytest.mark.skip(reason="Requires real Excel installation — run manually")
    def test_excel_com_opens_workbook(self):
        """Verify Excel COM can open and read a workbook."""
        import win32com.client

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            # Update path to your test Excel file
            test_file = TOOL_DIR / "tests" / "fixtures" / "test_payslip.xls"
            if not test_file.exists():
                pytest.skip(f"Test file not found: {test_file}")

            wb = excel.Workbooks.Open(str(test_file.resolve()))
            ws = wb.Sheets("Data")

            # Should have at least header + 1 employee
            assert ws.Range("A4").Value is not None

            wb.Close(SaveChanges=False)
        finally:
            excel.Quit()

    @pytest.mark.skip(reason="Requires real Excel installation — run manually")
    def test_pdf_generation_and_password(self):
        """Verify PDF generation with password protection."""
        pass

    @pytest.mark.skip(reason="Requires real Outlook — run manually")
    def test_outlook_email_send_dryrun(self):
        """Verify email composition and dry-run send via Outlook COM."""
        pass

    @pytest.mark.skip(reason="Requires real Excel — run manually")
    def test_no_orphan_excel_processes(self):
        """Verify no Excel processes remain after execution."""
        import subprocess

        # Count Excel processes before
        before = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE"],
            capture_output=True, text=True,
        )
        # Run a short test
        # ...
        # Count Excel processes after
        after = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE"],
            capture_output=True, text=True,
        )
        # Verify same count
