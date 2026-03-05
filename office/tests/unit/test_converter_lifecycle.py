"""
Unit tests for refactored office/excel/converter.py — PdfConverter.

Verifies that the COM lifecycle follows REQ-COM-02 and REQ-COM-08:
- Uses com_initialized() context manager
- Uses get_or_create_excel() / safe_quit_excel()
- Does not call Quit() if Excel was already running
"""

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


@pytest.fixture
def mock_converter_com():
    """Set up mocks for the COM stack used by PdfConverter."""
    mock_com_ctx = MagicMock()
    mock_com_ctx.__enter__ = MagicMock(return_value=None)
    mock_com_ctx.__exit__ = MagicMock(return_value=False)

    mock_excel = MagicMock()

    patches = {
        "com_init": patch(
            "office.excel.converter.com_initialized",
            return_value=mock_com_ctx,
        ),
        "ensure": patch("office.excel.converter.ensure_com_available"),
        "get_excel": patch(
            "office.excel.converter.create_excel_background",
            return_value=(mock_excel, False),
        ),
        "safe_quit": patch("office.excel.converter.safe_quit_excel"),
    }

    started = {k: p.start() for k, p in patches.items()}
    yield {
        "com_ctx": mock_com_ctx,
        "excel": mock_excel,
        "patches": started,
    }
    for p in patches.values():
        p.stop()


class TestPdfConverterLifecycle:
    """Verify PdfConverter uses the proper COM lifecycle."""

    def test_init_excel_calls_com_initialized(self, mock_converter_com, tmp_path):
        """_init_excel should enter a com_initialized context."""
        from office.excel.converter import PdfConverter
        converter = PdfConverter(output_dir=tmp_path)
        converter._init_excel()

        mock_converter_com["com_ctx"].__enter__.assert_called_once()

    def test_init_excel_calls_get_or_create(self, mock_converter_com, tmp_path):
        """_init_excel should use create_excel_background."""
        from office.excel.converter import PdfConverter
        converter = PdfConverter(output_dir=tmp_path)
        converter._init_excel()

        mock_converter_com["patches"]["get_excel"].assert_called_once()

    def test_cleanup_calls_safe_quit(self, mock_converter_com, tmp_path):
        """_cleanup_excel should call safe_quit_excel."""
        from office.excel.converter import PdfConverter
        converter = PdfConverter(output_dir=tmp_path)
        converter._init_excel()
        converter._cleanup_excel()

        mock_converter_com["patches"]["safe_quit"].assert_called_once()

    def test_cleanup_exits_com_context(self, mock_converter_com, tmp_path):
        """_cleanup_excel should exit the com_initialized context."""
        from office.excel.converter import PdfConverter
        converter = PdfConverter(output_dir=tmp_path)
        converter._init_excel()
        converter._cleanup_excel()

        mock_converter_com["com_ctx"].__exit__.assert_called_once()

    def test_context_manager_full_lifecycle(self, mock_converter_com, tmp_path):
        """with PdfConverter() should init and cleanup COM."""
        from office.excel.converter import PdfConverter
        with PdfConverter(output_dir=tmp_path) as converter:
            mock_converter_com["com_ctx"].__enter__.assert_called_once()

        mock_converter_com["com_ctx"].__exit__.assert_called_once()
        mock_converter_com["patches"]["safe_quit"].assert_called_once()

    def test_was_already_running_propagated(self, tmp_path):
        """safe_quit_excel receives the correct was_already_running flag."""
        mock_com_ctx = MagicMock()
        mock_com_ctx.__enter__ = MagicMock(return_value=None)
        mock_com_ctx.__exit__ = MagicMock(return_value=False)
        mock_excel = MagicMock()

        with patch("office.excel.converter.com_initialized", return_value=mock_com_ctx), \
             patch("office.excel.converter.ensure_com_available"), \
             patch("office.excel.converter.create_excel_background", return_value=(mock_excel, False)), \
             patch("office.excel.converter.safe_quit_excel") as mock_safe_quit:

            from office.excel.converter import PdfConverter
            with PdfConverter(output_dir=tmp_path):
                pass

            # was_already_running=False means safe_quit always calls Quit()
            mock_safe_quit.assert_called_once_with(mock_excel, False)

    def test_double_init_is_noop(self, mock_converter_com, tmp_path):
        """Calling _init_excel twice should not create two COM instances."""
        from office.excel.converter import PdfConverter
        converter = PdfConverter(output_dir=tmp_path)
        converter._init_excel()
        converter._init_excel()

        # Should only be called once
        mock_converter_com["patches"]["get_excel"].assert_called_once()
