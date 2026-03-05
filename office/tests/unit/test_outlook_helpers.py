"""
Unit tests for office/utils/helpers.py — Outlook COM lifecycle helpers.

Covers:
- get_or_create_outlook() — attaches to running instance or creates new
- safe_quit_outlook() — only quits if tool created the instance
- REQ-COM-08 compliance

All tests mock COM at the office/utils/com level (no real Outlook needed).
"""

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


class TestGetOrCreateOutlook:
    """Tests for get_or_create_outlook()."""

    def test_attaches_to_running_instance(self):
        """When Outlook is running, GetObject succeeds → was_already_running=True."""
        mock_win32 = MagicMock()
        mock_outlook = MagicMock()
        mock_win32.GetObject.return_value = mock_outlook

        with patch("office.utils.helpers.ensure_com_available"), \
             patch("office.utils.helpers.get_win32com_client", return_value=mock_win32), \
             patch("office.utils.helpers.get_pythoncom"):
            from office.utils.helpers import get_or_create_outlook
            app, was_running = get_or_create_outlook()

        assert app is mock_outlook
        assert was_running is True

    def test_creates_new_instance_when_not_running(self):
        """When Outlook is not running, GetObject fails → Dispatch."""
        mock_win32 = MagicMock()
        mock_pythoncom = MagicMock()
        mock_outlook = MagicMock()

        mock_pythoncom.com_error = type("com_error", (Exception,), {})
        mock_win32.GetObject.side_effect = mock_pythoncom.com_error()
        mock_win32.Dispatch.return_value = mock_outlook

        with patch("office.utils.helpers.ensure_com_available"), \
             patch("office.utils.helpers.get_win32com_client", return_value=mock_win32), \
             patch("office.utils.helpers.get_pythoncom", return_value=mock_pythoncom):
            from office.utils.helpers import get_or_create_outlook
            app, was_running = get_or_create_outlook()

        assert app is mock_outlook
        assert was_running is False


class TestSafeQuitOutlook:
    """Tests for safe_quit_outlook()."""

    def test_quits_when_tool_created(self):
        """Should call Quit() when was_already_running=False."""
        mock_outlook = MagicMock()

        from office.utils.helpers import safe_quit_outlook
        safe_quit_outlook(mock_outlook, was_already_running=False)

        mock_outlook.Quit.assert_called_once()

    def test_no_quit_when_already_running(self):
        """Should NOT call Quit() when was_already_running=True."""
        mock_outlook = MagicMock()

        from office.utils.helpers import safe_quit_outlook
        safe_quit_outlook(mock_outlook, was_already_running=True)

        mock_outlook.Quit.assert_not_called()

    def test_handles_none_app(self):
        """Should not crash when outlook_app is None."""
        from office.utils.helpers import safe_quit_outlook
        safe_quit_outlook(None, was_already_running=False)

    def test_handles_quit_exception(self):
        """Should swallow exceptions from Quit() gracefully."""
        mock_outlook = MagicMock()
        mock_outlook.Quit.side_effect = Exception("already closed")

        from office.utils.helpers import safe_quit_outlook
        safe_quit_outlook(mock_outlook, was_already_running=False)
