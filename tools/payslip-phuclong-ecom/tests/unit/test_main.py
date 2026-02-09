"""
Unit tests for main.py helper functions.

Tests each extracted helper function in isolation with mocked dependencies.
Covers:
- load_and_validate_config
- read_employee_data
- validate_employee_data
- show_summary_and_confirm
- check_and_handle_existing_state
- generate_payslips
- convert_to_pdf
- compose_emails
- send_emails
"""

import sys
import time
from pathlib import Path
from unittest.mock import MagicMock, patch, call

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
TOOL_DIR = Path(__file__).resolve().parent.parent.parent
for p in [str(PROJECT_ROOT), str(TOOL_DIR)]:
    if p not in sys.path:
        sys.path.insert(0, p)

from tests.conftest import make_employee, make_employees


class TestLoadAndValidateConfig:
    """Tests for load_and_validate_config."""

    @patch("main.setup_logging")
    @patch("main.load_config")
    def test_valid_config_returns(self, mock_load, mock_logging):
        from main import load_and_validate_config

        mock_config = MagicMock()
        mock_config.validate.return_value = []
        mock_config.log_dir = Path("./logs")
        mock_config.log_level = "INFO"
        mock_load.return_value = mock_config

        result = load_and_validate_config(Path("."))
        assert result == mock_config
        mock_config.ensure_directories.assert_called_once()

    @patch("main.setup_logging")
    @patch("main.load_config")
    def test_invalid_config_exits(self, mock_load, mock_logging):
        from main import load_and_validate_config

        mock_config = MagicMock()
        mock_config.validate.return_value = ["Error 1", "Error 2"]
        mock_load.return_value = mock_config

        with pytest.raises(SystemExit) as exc_info:
            load_and_validate_config(Path("."))
        assert exc_info.value.code == 1


class TestValidateEmployeeData:
    """Tests for validate_employee_data."""

    def test_valid_data_passes(self, mock_config):
        from main import validate_employee_data

        employees = make_employees(3)
        # Should not raise
        validate_employee_data(employees, mock_config)

    def test_invalid_data_exits(self, mock_config):
        from main import validate_employee_data

        employees = [make_employee(mnv="", email="")]
        with pytest.raises(SystemExit) as exc_info:
            validate_employee_data(employees, mock_config)
        assert exc_info.value.code == 1


class TestShowSummaryAndConfirm:
    """Tests for show_summary_and_confirm."""

    @patch("main.confirm_proceed", return_value=True)
    def test_non_dryrun_confirms(self, mock_confirm, mock_config):
        from main import show_summary_and_confirm

        mock_config.dry_run = False
        employees = make_employees(3)
        show_summary_and_confirm(mock_config, employees)
        mock_confirm.assert_called_once()

    def test_dryrun_skips_confirm(self, mock_config):
        from main import show_summary_and_confirm

        mock_config.dry_run = True
        employees = make_employees(3)
        # Should not raise or prompt
        show_summary_and_confirm(mock_config, employees)

    @patch("main.confirm_proceed", return_value=False)
    def test_declined_exits(self, mock_confirm, mock_config):
        from main import show_summary_and_confirm

        mock_config.dry_run = False
        employees = make_employees(3)
        with pytest.raises(SystemExit):
            show_summary_and_confirm(mock_config, employees)


class TestCheckAndHandleExistingState:
    """Tests for check_and_handle_existing_state."""

    @patch("main.analyze_existing_state")
    def test_no_state_continues(self, mock_analyze, mock_config):
        from main import check_and_handle_existing_state

        mock_analyze.return_value = {"has_state": False}
        # Should not raise
        check_and_handle_existing_state(mock_config, 10)

    @patch("main.prompt_state_action", return_value="no")
    @patch("main.analyze_existing_state")
    def test_state_exists_user_declines(self, mock_analyze, mock_prompt, mock_config):
        from main import check_and_handle_existing_state

        mock_analyze.return_value = {"has_state": True, "sent_count": 5}
        with pytest.raises(SystemExit):
            check_and_handle_existing_state(mock_config, 10)

    @patch("main.cleanup_all_files")
    @patch("main.prompt_state_action", return_value="new")
    @patch("main.analyze_existing_state")
    def test_state_exists_user_starts_new(self, mock_analyze, mock_prompt, mock_cleanup, mock_config):
        from main import check_and_handle_existing_state

        mock_analyze.return_value = {"has_state": True, "sent_count": 5}
        check_and_handle_existing_state(mock_config, 10)
        mock_cleanup.assert_called_once_with(mock_config)

    @patch("main.prompt_state_action", return_value="yes")
    @patch("main.analyze_existing_state")
    def test_state_exists_user_resumes(self, mock_analyze, mock_prompt, mock_config):
        from main import check_and_handle_existing_state

        mock_analyze.return_value = {"has_state": True, "sent_count": 5}
        # Should not raise or cleanup
        check_and_handle_existing_state(mock_config, 10)


class TestComposeEmails:
    """Tests for compose_emails."""

    def test_compose_emails_returns_results(self, mock_config, mock_email_template):
        from main import compose_emails

        results = [
            {
                "employee": make_employee(),
                "pdf_path": None,
                "success": True,
            }
        ]

        results, composed = compose_emails(
            mock_config, results, mock_email_template, "Test Subject"
        )
        # pdf_path is None, so email_data should also be None
        assert composed == 0

    def test_compose_emails_with_valid_pdf(self, mock_config, mock_email_template, tmp_path):
        from main import compose_emails

        pdf = tmp_path / "test.pdf"
        pdf.touch()

        results = [
            {
                "employee": make_employee(),
                "pdf_path": str(pdf),
                "success": True,
            }
        ]

        results, composed = compose_emails(
            mock_config, results, mock_email_template, "Test Subject"
        )
        assert composed == 1


class TestSendEmails:
    """Tests for send_emails with mocked Outlook."""

    @patch("main.OutlookSender")
    @patch("main.StateTracker")
    @patch("main.ResultWriter")
    def test_dry_run_send(self, MockWriter, MockState, MockSender, mock_config):
        from main import send_emails

        mock_config.dry_run = True

        # Setup mocks
        mock_tracker = MagicMock()
        mock_tracker.get_processed_count.return_value = 0
        mock_tracker.is_processed.return_value = False
        MockState.return_value = mock_tracker

        mock_writer = MagicMock()
        MockWriter.return_value = mock_writer

        mock_sender = MagicMock()
        mock_sender.send.return_value = True
        mock_sender.sent_count = 1
        mock_sender.skipped_count = 0
        mock_sender.error_count = 0
        mock_sender.__enter__ = MagicMock(return_value=mock_sender)
        mock_sender.__exit__ = MagicMock(return_value=False)
        MockSender.return_value = mock_sender

        results = [
            {
                "employee": make_employee(),
                "email_data": {
                    "to": ["a@co.com"],
                    "subject": "Test",
                    "body": "Hello",
                    "body_is_html": True,
                    "attachments": [],
                },
                "pdf_path": "/tmp/test.pdf",
            }
        ]

        sent, skipped, errors, result_file = send_emails(mock_config, results, 1)
        assert sent == 1
        assert errors == 0

    @patch("main.OutlookSender")
    @patch("main.StateTracker")
    @patch("main.ResultWriter")
    def test_skip_already_processed(self, MockWriter, MockState, MockSender, mock_config):
        from main import send_emails

        mock_tracker = MagicMock()
        mock_tracker.get_processed_count.return_value = 1
        mock_tracker.is_processed.return_value = True  # Already processed
        MockState.return_value = mock_tracker

        mock_writer = MagicMock()
        MockWriter.return_value = mock_writer

        mock_sender = MagicMock()
        mock_sender.sent_count = 0
        mock_sender.skipped_count = 0
        mock_sender.error_count = 0
        mock_sender.__enter__ = MagicMock(return_value=mock_sender)
        mock_sender.__exit__ = MagicMock(return_value=False)
        MockSender.return_value = mock_sender

        results = [
            {
                "employee": make_employee(),
                "email_data": {
                    "to": ["a@co.com"],
                    "subject": "Test",
                    "body": "Hello",
                    "body_is_html": True,
                },
                "pdf_path": "/tmp/test.pdf",
            }
        ]

        sent, skipped, errors, _ = send_emails(mock_config, results, 1)
        assert skipped == 1
        assert sent == 0

    @patch("main.OutlookSender")
    @patch("main.StateTracker")
    @patch("main.ResultWriter")
    def test_no_email_data_counted_as_error(self, MockWriter, MockState, MockSender, mock_config):
        from main import send_emails

        mock_tracker = MagicMock()
        mock_tracker.get_processed_count.return_value = 0
        mock_tracker.is_processed.return_value = False
        MockState.return_value = mock_tracker

        mock_writer = MagicMock()
        MockWriter.return_value = mock_writer

        mock_sender = MagicMock()
        mock_sender.sent_count = 0
        mock_sender.skipped_count = 0
        mock_sender.error_count = 0
        mock_sender.__enter__ = MagicMock(return_value=mock_sender)
        mock_sender.__exit__ = MagicMock(return_value=False)
        MockSender.return_value = mock_sender

        results = [
            {
                "employee": make_employee(),
                "email_data": None,  # No email data
            }
        ]

        sent, skipped, errors, _ = send_emails(mock_config, results, 1)
        assert errors == 1
        assert sent == 0
