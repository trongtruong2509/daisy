"""
Unit tests for office/utils/helpers.py — Excel COM lifecycle helpers.

Covers:
- create_excel_background() — always creates new isolated process via Dispatch
- get_or_create_excel() — attaches to running instance or creates new
- safe_quit_excel() — only quits if tool created the instance
- REQ-COM-08 compliance

All tests mock COM at the office/utils/com level (no real Excel needed).
"""

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch, PropertyMock

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


class TestGetOrCreateExcel:
    """Tests for get_or_create_excel()."""

    def test_attaches_to_running_instance(self):
        """When Excel is running, GetObject succeeds → was_already_running=True."""
        mock_win32 = MagicMock()
        mock_excel = MagicMock()
        mock_win32.GetObject.return_value = mock_excel

        with patch("office.utils.helpers.ensure_com_available"), \
             patch("office.utils.helpers.get_win32com_client", return_value=mock_win32), \
             patch("office.utils.helpers.get_pythoncom"):
            from office.utils.helpers import get_or_create_excel
            app, was_running = get_or_create_excel()

        assert app is mock_excel
        assert was_running is True
        mock_win32.GetObject.assert_called_once_with(Class="Excel.Application")

    def test_creates_new_instance_when_not_running(self):
        """When Excel is not running, GetObject fails → Dispatch → was_already_running=False."""
        mock_win32 = MagicMock()
        mock_pythoncom = MagicMock()
        mock_excel = MagicMock()

        # GetObject raises com_error
        mock_pythoncom.com_error = type("com_error", (Exception,), {})
        mock_win32.GetObject.side_effect = mock_pythoncom.com_error()
        mock_win32.Dispatch.return_value = mock_excel

        with patch("office.utils.helpers.ensure_com_available"), \
             patch("office.utils.helpers.get_win32com_client", return_value=mock_win32), \
             patch("office.utils.helpers.get_pythoncom", return_value=mock_pythoncom):
            from office.utils.helpers import get_or_create_excel
            app, was_running = get_or_create_excel()

        assert app is mock_excel
        assert was_running is False
        mock_win32.Dispatch.assert_called_once_with("Excel.Application")


class TestSafeQuitExcel:
    """Tests for safe_quit_excel()."""

    def test_quits_when_tool_created(self):
        """Should call Quit() when was_already_running=False."""
        mock_excel = MagicMock()

        from office.utils.helpers import safe_quit_excel
        safe_quit_excel(mock_excel, was_already_running=False)

        mock_excel.Quit.assert_called_once()

    def test_no_quit_when_already_running(self):
        """Should NOT call Quit() when was_already_running=True."""
        mock_excel = MagicMock()

        from office.utils.helpers import safe_quit_excel
        safe_quit_excel(mock_excel, was_already_running=True)

        mock_excel.Quit.assert_not_called()

    def test_handles_none_app(self):
        """Should not crash when excel_app is None."""
        from office.utils.helpers import safe_quit_excel
        safe_quit_excel(None, was_already_running=False)  # Should not raise

    def test_handles_quit_exception(self):
        """Should swallow exceptions from Quit() gracefully."""
        mock_excel = MagicMock()
        mock_excel.Quit.side_effect = Exception("already closed")

        from office.utils.helpers import safe_quit_excel
        safe_quit_excel(mock_excel, was_already_running=False)  # Should not raise


class TestCreateExcelBackground:
    """Tests for create_excel_background() — the background processing helper."""

    def test_always_dispatches_new_instance(self):
        """create_excel_background() must call DispatchEx (not Dispatch) to bypass the ROT."""
        mock_win32 = MagicMock()
        mock_excel = MagicMock()
        mock_win32.DispatchEx.return_value = mock_excel

        with patch("office.utils.helpers.ensure_com_available"), \
             patch("office.utils.helpers.get_win32com_client", return_value=mock_win32):
            from office.utils.helpers import create_excel_background
            app, was_running = create_excel_background()

        assert app is mock_excel
        assert was_running is False  # always False — tool always owns it
        mock_win32.DispatchEx.assert_called_once_with("Excel.Application")
        mock_win32.Dispatch.assert_not_called()  # must never fall back to Dispatch

    def test_was_already_running_is_always_false(self):
        """Second return value must always be False regardless of environment."""
        mock_win32 = MagicMock()
        mock_win32.DispatchEx.return_value = MagicMock()

        with patch("office.utils.helpers.ensure_com_available"), \
             patch("office.utils.helpers.get_win32com_client", return_value=mock_win32):
            from office.utils.helpers import create_excel_background
            _, was_running = create_excel_background()

        assert was_running is False

    def test_safe_quit_always_quits_background_instance(self):
        """safe_quit_excel with was_already_running=False always calls Quit()."""
        mock_excel = MagicMock()
        from office.utils.helpers import safe_quit_excel
        safe_quit_excel(mock_excel, was_already_running=False)
        mock_excel.Quit.assert_called_once()
