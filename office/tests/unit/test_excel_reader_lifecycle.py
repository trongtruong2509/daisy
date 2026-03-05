"""
Unit tests for refactored office/excel/reader.py — ExcelComReader.

Verifies that the COM lifecycle follows REQ-COM-02 and REQ-COM-08:
- Uses com_initialized() context manager
- Uses get_or_create_excel() / safe_quit_excel()
- Does not call Quit() if Excel was already running
"""

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch, call

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


@pytest.fixture
def mock_com_stack():
    """Set up mocks for the entire COM stack used by ExcelComReader."""
    mock_com_ctx = MagicMock()
    mock_com_ctx.__enter__ = MagicMock(return_value=None)
    mock_com_ctx.__exit__ = MagicMock(return_value=False)

    mock_excel = MagicMock()
    mock_wb = MagicMock()
    mock_excel.Workbooks.Open.return_value = mock_wb

    patches = {
        "com_init": patch(
            "office.excel.reader.com_initialized",
            return_value=mock_com_ctx,
        ),
        "ensure": patch("office.excel.reader.ensure_com_available"),
        "get_excel": patch(
            "office.excel.reader.create_excel_background",
            return_value=(mock_excel, False),
        ),
        "safe_quit": patch("office.excel.reader.safe_quit_excel"),
    }

    started = {k: p.start() for k, p in patches.items()}
    yield {
        "com_ctx": mock_com_ctx,
        "excel": mock_excel,
        "workbook": mock_wb,
        "patches": started,
    }
    for p in patches.values():
        p.stop()


class TestExcelComReaderLifecycle:
    """Verify ExcelComReader uses the proper COM lifecycle."""

    def test_open_calls_com_initialized(self, mock_com_stack, tmp_path):
        """open() should enter com_initialized context."""
        test_file = tmp_path / "test.xlsx"
        test_file.touch()

        from office.excel.reader import ExcelComReader
        reader = ExcelComReader(test_file)
        reader.open()

        mock_com_stack["com_ctx"].__enter__.assert_called_once()

    def test_open_calls_get_or_create_excel(self, mock_com_stack, tmp_path):
        """open() should use create_excel_background instead of raw Dispatch."""
        test_file = tmp_path / "test.xlsx"
        test_file.touch()

        from office.excel.reader import ExcelComReader
        reader = ExcelComReader(test_file)
        reader.open()

        mock_com_stack["patches"]["get_excel"].assert_called_once()

    def test_close_calls_safe_quit(self, mock_com_stack, tmp_path):
        """close() should call safe_quit_excel."""
        test_file = tmp_path / "test.xlsx"
        test_file.touch()

        from office.excel.reader import ExcelComReader
        reader = ExcelComReader(test_file)
        reader.open()
        reader.close()

        mock_com_stack["patches"]["safe_quit"].assert_called_once()

    def test_close_exits_com_context(self, mock_com_stack, tmp_path):
        """close() should exit the com_initialized context."""
        test_file = tmp_path / "test.xlsx"
        test_file.touch()

        from office.excel.reader import ExcelComReader
        reader = ExcelComReader(test_file)
        reader.open()
        reader.close()

        mock_com_stack["com_ctx"].__exit__.assert_called_once()

    def test_context_manager_lifecycle(self, mock_com_stack, tmp_path):
        """with ExcelComReader() should open and close properly."""
        test_file = tmp_path / "test.xlsx"
        test_file.touch()

        from office.excel.reader import ExcelComReader
        with ExcelComReader(test_file) as reader:
            mock_com_stack["com_ctx"].__enter__.assert_called_once()

        mock_com_stack["com_ctx"].__exit__.assert_called_once()
        mock_com_stack["patches"]["safe_quit"].assert_called_once()

    def test_was_already_running_flag_propagated(self, tmp_path):
        """safe_quit_excel receives the was_already_running flag from create_excel_background."""
        test_file = tmp_path / "test.xlsx"
        test_file.touch()

        mock_com_ctx = MagicMock()
        mock_com_ctx.__enter__ = MagicMock(return_value=None)
        mock_com_ctx.__exit__ = MagicMock(return_value=False)
        mock_excel = MagicMock()
        mock_excel.Workbooks.Open.return_value = MagicMock()

        with patch("office.excel.reader.com_initialized", return_value=mock_com_ctx), \
             patch("office.excel.reader.ensure_com_available"), \
             patch("office.excel.reader.create_excel_background", return_value=(mock_excel, False)), \
             patch("office.excel.reader.safe_quit_excel") as mock_safe_quit:

            from office.excel.reader import ExcelComReader
            with ExcelComReader(test_file):
                pass

            # create_excel_background always returns False — tool always owns the instance
            mock_safe_quit.assert_called_once_with(mock_excel, False)


class TestExcelComReaderFileNotFound:
    """Verify file validation still works after refactor."""

    def test_raises_file_not_found(self):
        """Should raise FileNotFoundError for non-existent file."""
        with patch("office.excel.reader.ensure_com_available"):
            from office.excel.reader import ExcelComReader
            with pytest.raises(FileNotFoundError):
                ExcelComReader(Path("nonexistent.xlsx"))
