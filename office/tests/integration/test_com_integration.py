"""
Integration tests for COM handling — requires live Excel and Outlook.

These tests verify:
- REQ-COM-01: com_initialized() actually calls CoInitialize/CoUninitialize
- REQ-COM-08: Existing Excel/Outlook sessions are preserved
- REQ-COM-02: office/ classes properly manage COM lifecycle

Run manually:
    cd office
    ..\\venv\\Scripts\\pytest tests/integration/ -v -m integration

DO NOT run in CI/CD — requires Windows with Office installed.
"""

import subprocess
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

pytestmark = pytest.mark.integration


def _count_processes(name: str) -> int:
    """Count running processes by image name."""
    result = subprocess.run(
        ["tasklist", "/FI", f"IMAGENAME eq {name}"],
        capture_output=True, text=True,
    )
    # Each process line starts with the image name
    return result.stdout.lower().count(name.lower())


@pytest.mark.integration
class TestComInitializedIntegration:
    """Integration tests for com_initialized() with real COM."""

    def test_com_initialized_basic(self):
        """com_initialized() should succeed on Windows with pywin32."""
        from office.utils.com import com_initialized, is_available
        if not is_available():
            pytest.skip("pywin32 not installed")

        with com_initialized():
            # Inside the context, COM should be initialized
            pass  # No error means success

    def test_nested_com_initialized(self):
        """Nested com_initialized() should work (ref-counted)."""
        from office.utils.com import com_initialized, is_available
        if not is_available():
            pytest.skip("pywin32 not installed")

        with com_initialized():
            with com_initialized():
                pass  # No error means COM ref-counting works


@pytest.mark.integration
class TestExcelLifecycleIntegration:
    """Integration tests verifying Excel COM lifecycle (REQ-COM-08)."""

    def test_excel_com_reader_does_not_kill_existing_excel(self, tmp_path):
        """
        If Excel is already running, ExcelComReader must NOT call Quit().

        This test:
        1. Notes initial Excel process count
        2. Opens ExcelComReader (which may attach to existing or create new)
        3. Closes ExcelComReader
        4. Verifies Excel process count is >= initial (existing session preserved)
        """
        from office.utils.com import is_available
        if not is_available():
            pytest.skip("pywin32 not installed")

        # Create a dummy Excel file
        test_file = tmp_path / "test.xlsx"
        # Create a minimal xlsx using openpyxl if available
        try:
            import openpyxl
            wb = openpyxl.Workbook()
            wb.save(str(test_file))
        except ImportError:
            pytest.skip("openpyxl required for this test")

        initial_count = _count_processes("EXCEL.EXE")

        from office.excel.reader import ExcelComReader
        with ExcelComReader(test_file) as reader:
            pass

        final_count = _count_processes("EXCEL.EXE")

        # If Excel was already running (initial > 0), it should still be running
        if initial_count > 0:
            assert final_count >= initial_count, (
                f"Excel process count dropped from {initial_count} to {final_count}. "
                "ExcelComReader may have killed the user's session!"
            )

    def test_get_or_create_excel_reuses_existing(self):
        """get_or_create_excel should detect a running Excel instance."""
        from office.utils.com import com_initialized, is_available, get_win32com_client
        if not is_available():
            pytest.skip("pywin32 not installed")

        win32 = get_win32com_client()

        with com_initialized():
            # Create an Excel instance manually
            excel1 = win32.Dispatch("Excel.Application")
            excel1.Visible = False

            try:
                from office.utils.helpers import get_or_create_excel
                excel2, was_running = get_or_create_excel()

                # Should detect the running instance
                assert was_running is True
            finally:
                excel1.Quit()


@pytest.mark.integration
class TestOutlookLifecycleIntegration:
    """Integration tests verifying Outlook COM lifecycle (REQ-COM-08)."""

    def test_outlook_client_does_not_kill_existing_outlook(self):
        """
        If Outlook is already running, OutlookClient must NOT call Quit().
        """
        from office.utils.com import is_available
        if not is_available():
            pytest.skip("pywin32 not installed")

        initial_count = _count_processes("OUTLOOK.EXE")
        if initial_count == 0:
            pytest.skip("Outlook is not running — cannot test preservation")

        # We can't easily create a full OutlookClient without a valid account,
        # but we can test the helper directly
        from office.utils.com import com_initialized
        from office.utils.helpers import get_or_create_outlook, safe_quit_outlook

        with com_initialized():
            outlook, was_running = get_or_create_outlook()
            assert was_running is True, "Expected to attach to running Outlook"
            safe_quit_outlook(outlook, was_running)

        final_count = _count_processes("OUTLOOK.EXE")
        assert final_count >= initial_count, (
            f"Outlook process count dropped from {initial_count} to {final_count}. "
            "safe_quit_outlook may have killed the user's session!"
        )
