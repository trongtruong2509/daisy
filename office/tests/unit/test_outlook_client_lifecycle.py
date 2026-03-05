"""
Unit tests for refactored office/outlook/client.py — OutlookClient.

Verifies that the COM lifecycle follows REQ-COM-02 and REQ-COM-08:
- Uses com_initialized() context manager
- Uses get_or_create_outlook() / safe_quit_outlook()
- Does not call Quit() if Outlook was already running
"""

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


@pytest.fixture
def mock_outlook_com():
    """Set up mocks for the COM stack used by OutlookClient."""
    mock_com_ctx = MagicMock()
    mock_com_ctx.__enter__ = MagicMock(return_value=None)
    mock_com_ctx.__exit__ = MagicMock(return_value=False)

    mock_outlook = MagicMock()
    mock_namespace = MagicMock()
    mock_account = MagicMock()
    mock_account.SmtpAddress = "test@company.com"
    mock_outlook.GetNamespace.return_value = mock_namespace
    mock_namespace.Accounts.Count = 1
    mock_namespace.Accounts.Item.return_value = mock_account

    patches = {
        "com_init": patch(
            "office.outlook.client.com_initialized",
            return_value=mock_com_ctx,
        ),
        "ensure": patch("office.outlook.client.ensure_com_available"),
        "is_available": patch("office.outlook.client.is_available", return_value=True),
        "get_outlook": patch(
            "office.outlook.client.get_or_create_outlook",
            return_value=(mock_outlook, False),
        ),
        "safe_quit": patch("office.outlook.client.safe_quit_outlook"),
        "get_pywintypes": patch("office.outlook.client.get_pywintypes"),
    }

    started = {k: p.start() for k, p in patches.items()}
    yield {
        "com_ctx": mock_com_ctx,
        "outlook": mock_outlook,
        "namespace": mock_namespace,
        "account": mock_account,
        "patches": started,
    }
    for p in patches.values():
        p.stop()


class TestOutlookClientLifecycle:
    """Verify OutlookClient uses the proper COM lifecycle."""

    def test_connect_enters_com_context(self, mock_outlook_com):
        """connect() should enter com_initialized context."""
        from office.outlook.client import OutlookClient
        client = OutlookClient(account="test@company.com")
        client.connect()

        mock_outlook_com["com_ctx"].__enter__.assert_called_once()

    def test_connect_uses_get_or_create(self, mock_outlook_com):
        """connect() should use get_or_create_outlook."""
        from office.outlook.client import OutlookClient
        client = OutlookClient(account="test@company.com")
        client.connect()

        mock_outlook_com["patches"]["get_outlook"].assert_called_once()

    def test_disconnect_calls_safe_quit(self, mock_outlook_com):
        """disconnect() should call safe_quit_outlook."""
        from office.outlook.client import OutlookClient
        client = OutlookClient(account="test@company.com")
        client.connect()
        client.disconnect()

        mock_outlook_com["patches"]["safe_quit"].assert_called_once()

    def test_disconnect_exits_com_context(self, mock_outlook_com):
        """disconnect() should exit the com_initialized context."""
        from office.outlook.client import OutlookClient
        client = OutlookClient(account="test@company.com")
        client.connect()
        client.disconnect()

        mock_outlook_com["com_ctx"].__exit__.assert_called_once()

    def test_context_manager_full_lifecycle(self, mock_outlook_com):
        """with OutlookClient() should connect and disconnect properly."""
        from office.outlook.client import OutlookClient
        with OutlookClient(account="test@company.com") as client:
            mock_outlook_com["com_ctx"].__enter__.assert_called_once()
            assert client.is_connected

        mock_outlook_com["com_ctx"].__exit__.assert_called_once()

    def test_was_already_running_propagated(self):
        """safe_quit_outlook receives the correct was_already_running flag."""
        mock_com_ctx = MagicMock()
        mock_com_ctx.__enter__ = MagicMock(return_value=None)
        mock_com_ctx.__exit__ = MagicMock(return_value=False)

        mock_outlook = MagicMock()
        mock_ns = MagicMock()
        mock_acc = MagicMock()
        mock_acc.SmtpAddress = "test@company.com"
        mock_outlook.GetNamespace.return_value = mock_ns
        mock_ns.Accounts.Count = 1
        mock_ns.Accounts.Item.return_value = mock_acc

        with patch("office.outlook.client.com_initialized", return_value=mock_com_ctx), \
             patch("office.outlook.client.ensure_com_available"), \
             patch("office.outlook.client.is_available", return_value=True), \
             patch("office.outlook.client.get_or_create_outlook", return_value=(mock_outlook, True)), \
             patch("office.outlook.client.safe_quit_outlook") as mock_safe_quit, \
             patch("office.outlook.client.get_pywintypes"):

            from office.outlook.client import OutlookClient
            with OutlookClient(account="test@company.com"):
                pass

            mock_safe_quit.assert_called_once_with(mock_outlook, True)


class TestOutlookClientGetAvailableAccounts:
    """Tests for the static get_available_accounts method."""

    def test_returns_empty_when_unavailable(self):
        """Should return [] when COM is not available."""
        with patch("office.outlook.client.ensure_com_available"), \
             patch("office.outlook.client.is_available", return_value=False):
            from office.outlook.client import OutlookClient
            accounts = OutlookClient.get_available_accounts()
            assert accounts == []

    def test_uses_com_initialized_context(self):
        """Should use com_initialized() for the temporary COM connection."""
        mock_com_ctx = MagicMock()
        mock_com_ctx.__enter__ = MagicMock(return_value=None)
        mock_com_ctx.__exit__ = MagicMock(return_value=False)

        mock_outlook = MagicMock()
        mock_ns = MagicMock()
        mock_ns.Accounts.Count = 0
        mock_outlook.GetNamespace.return_value = mock_ns

        with patch("office.outlook.client.ensure_com_available"), \
             patch("office.outlook.client.is_available", return_value=True), \
             patch("office.outlook.client.com_initialized", return_value=mock_com_ctx), \
             patch("office.outlook.client.get_or_create_outlook", return_value=(mock_outlook, False)), \
             patch("office.outlook.client.safe_quit_outlook") as mock_quit:

            from office.outlook.client import OutlookClient
            OutlookClient.get_available_accounts()

            mock_com_ctx.__enter__.assert_called_once()
            mock_com_ctx.__exit__.assert_called_once()
            mock_quit.assert_called_once_with(mock_outlook, False)
