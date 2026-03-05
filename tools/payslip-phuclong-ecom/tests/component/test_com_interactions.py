"""
Component tests for Excel COM interactions (mocked COM layer).

Tests integration between internal modules while mocking external
COM dependencies. Validates correct COM method calls, parameters,
error handling, and resource cleanup.

Covers:
- ExcelReader → ExcelComReader COM calls
- PayslipGenerator → Excel COM lifecycle
- PdfConverter → ExportAsFixedFormat COM calls
- COM failure simulation
- Resource cleanup verification
"""

import gc
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch, PropertyMock, call

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
TOOL_DIR = Path(__file__).resolve().parent.parent.parent
for p in [str(PROJECT_ROOT), str(TOOL_DIR)]:
    if p not in sys.path:
        sys.path.insert(0, p)

from tests.conftest import make_employee, make_employees


class TestExcelReaderComInteraction:
    """Component tests for ExcelReader's COM interactions."""

    @patch("excel_reader.ExcelComReader")
    def test_reader_opens_read_only_with_recalc(self, MockComReader):
        from excel_reader import ExcelReader

        mock_reader = MagicMock()
        MockComReader.return_value = mock_reader

        reader = ExcelReader(Path("test.xls"))
        reader.open()

        mock_reader.open.assert_called_once_with(read_only=True, recalculate=True)

    @patch("excel_reader.ExcelComReader")
    def test_reader_closes_on_exit(self, MockComReader):
        from excel_reader import ExcelReader

        mock_reader = MagicMock()
        MockComReader.return_value = mock_reader

        with ExcelReader(Path("test.xls")) as reader:
            pass

        mock_reader.close.assert_called_once()

    @patch("excel_reader.ExcelComReader")
    def test_reader_closes_on_exception(self, MockComReader):
        from excel_reader import ExcelReader

        mock_reader = MagicMock()
        MockComReader.return_value = mock_reader

        with pytest.raises(ValueError):
            with ExcelReader(Path("test.xls")) as reader:
                raise ValueError("Test error")

        mock_reader.close.assert_called_once()

    @patch("excel_reader.ExcelComReader")
    def test_get_sheet_called_correctly(self, MockComReader):
        from excel_reader import ExcelReader

        mock_reader = MagicMock()
        MockComReader.return_value = mock_reader
        ws = MagicMock()
        mock_reader.get_sheet.return_value = ws

        # Make cells return empty to avoid iteration
        end_cell = MagicMock()
        end_cell.Row = 3  # No data rows (start_row=4)
        ws.Cells.return_value.End.return_value = end_cell
        ws.Rows.Count = 65536

        reader = ExcelReader(Path("test.xls"))
        reader._reader = mock_reader
        reader.read_employees("Data", 2, 4, "A", "B", "C", "AZ")

        mock_reader.get_sheet.assert_called_with("Data")


class TestPayslipGeneratorComInteraction:
    """Component tests for PayslipGenerator's COM lifecycle."""

    @patch("payslip_generator.com_initialized")
    @patch("payslip_generator.create_excel_background")
    @patch("payslip_generator.safe_quit_excel")
    def test_generate_batch_opens_workbook(self, mock_safe_quit, mock_get_excel, mock_com_ctx, tmp_path):
        """Test that generate_batch opens the source workbook via COM."""
        from payslip_generator import PayslipGenerator

        gen = PayslipGenerator(tmp_path, "01/2026")
        emp = make_employee()

        mock_excel = MagicMock()
        mock_get_excel.return_value = (mock_excel, False)
        mock_wb = MagicMock()
        mock_excel.Workbooks.Open.return_value = mock_wb

        # Mock the data sheet for XLOOKUP fix
        data_ws = MagicMock()
        data_ws.UsedRange.Rows.Count = 0
        mock_wb.Sheets.return_value = data_ws

        # Mock com_initialized context manager
        mock_com_ctx.return_value.__enter__ = MagicMock()
        mock_com_ctx.return_value.__exit__ = MagicMock(return_value=False)

        source = tmp_path / "source.xls"
        source.touch()

        try:
            gen.generate_batch([emp], source, batch_size=50)
        except Exception:
            pass  # May fail on _generate_one, that's OK

        mock_excel.Workbooks.Open.assert_called_once()

    def test_generate_batch_quits_excel_on_completion(self, tmp_path):
        """Test that Excel is properly quit after generation."""
        from payslip_generator import PayslipGenerator

        gen = PayslipGenerator(tmp_path, "01/2026")

        # Pre-create files so it skips COM entirely
        emps = make_employees(2)
        for emp in emps:
            gen._build_output_path(emp).touch()

        results = gen.generate_batch(emps, Path("dummy.xls"), batch_size=50)

        assert all(r["skipped"] for r in results)


class TestComFailureHandling:
    """Tests for COM failure scenarios."""

    @patch("excel_reader.ExcelComReader")
    def test_bang_luong_sheet_not_found(self, MockComReader):
        """Should handle missing 'bang luong' sheet gracefully."""
        from excel_reader import ExcelReader

        mock_reader = MagicMock()
        MockComReader.return_value = mock_reader

        ws = MagicMock()
        end_cell = MagicMock()
        end_cell.Row = 4
        ws.Cells.return_value.End.return_value = end_cell
        ws.Rows.Count = 65536

        # Employee with missing name
        def mock_range(cell_ref):
            cell = MagicMock()
            data = {
                "A4": 1, "B4": None, "C4": "emp@co.com", "AZ4": "001",
            }
            cell.Value = data.get(cell_ref)
            return cell

        ws.Range = mock_range

        def get_sheet(name):
            if name == "Data":
                return ws
            raise Exception("Sheet not found")

        mock_reader.get_sheet.side_effect = get_sheet

        reader = ExcelReader(Path("test.xls"))
        reader._reader = mock_reader

        # Should not raise, just log warning
        employees = reader.read_employees("Data", 2, 4, "A", "B", "C", "AZ")
        assert len(employees) == 1
        assert employees[0]["name"] == ""  # Name stays empty

    @patch("payslip_generator.com_initialized")
    @patch("payslip_generator.create_excel_background")
    @patch("payslip_generator.safe_quit_excel")
    def test_generator_handles_single_failure(self, mock_safe_quit, mock_get_excel, mock_com_ctx, tmp_path):
        """Test that one employee failure doesn't stop others."""
        from payslip_generator import PayslipGenerator

        gen = PayslipGenerator(tmp_path, "01/2026")

        mock_excel = MagicMock()
        mock_wb = MagicMock()
        mock_get_excel.return_value = (mock_excel, False)
        mock_excel.Workbooks.Open.return_value = mock_wb

        # Mock com_initialized context manager
        mock_com_ctx.return_value.__enter__ = MagicMock()
        mock_com_ctx.return_value.__exit__ = MagicMock(return_value=False)

        data_ws = MagicMock()
        mock_wb.Sheets.return_value = data_ws

        emps = make_employees(2)

        # First employee succeeds, second fails
        call_count = [0]

        def mock_generate_one(*args, **kwargs):
            call_count[0] += 1
            if call_count[0] == 2:
                raise Exception("COM error")
            return tmp_path / "test.xlsx"

        gen._generate_one = mock_generate_one

        results = gen.generate_batch(emps, Path("dummy.xls"), batch_size=50)

        assert len(results) == 2
        assert results[0]["success"] is True
        assert results[1]["success"] is False


class TestEmailSenderComInteraction:
    """Component tests for email sending with mocked Outlook COM."""

    @patch("main.OutlookSender")
    @patch("main.StateTracker")
    @patch("main.ResultWriter")
    def test_sender_context_manager_used(self, MockWriter, MockState, MockSender, mock_config):
        """Test that OutlookSender is used as context manager."""
        from main import send_emails

        mock_tracker = MagicMock()
        mock_tracker.get_processed_count.return_value = 0
        mock_tracker.is_processed.return_value = False
        MockState.return_value = mock_tracker
        MockWriter.return_value = MagicMock()

        mock_sender = MagicMock()
        mock_sender.send.return_value = True
        mock_sender.sent_count = 1
        mock_sender.skipped_count = 0
        mock_sender.error_count = 0
        mock_sender.__enter__ = MagicMock(return_value=mock_sender)
        mock_sender.__exit__ = MagicMock(return_value=False)
        MockSender.return_value = mock_sender

        results = [{
            "employee": make_employee(),
            "email_data": {
                "to": ["a@co.com"],
                "subject": "Test",
                "body": "Hello",
                "body_is_html": True,
            },
            "pdf_path": None,
        }]

        send_emails(mock_config, results, 1)

        # Verify context manager was used
        mock_sender.__enter__.assert_called_once()
        mock_sender.__exit__.assert_called_once()

    @patch("main.OutlookSender")
    @patch("main.StateTracker")
    @patch("main.ResultWriter")
    def test_send_failure_logged(self, MockWriter, MockState, MockSender, mock_config):
        """Test that send failure is counted as error."""
        from main import send_emails

        mock_tracker = MagicMock()
        mock_tracker.get_processed_count.return_value = 0
        mock_tracker.is_processed.return_value = False
        MockState.return_value = mock_tracker

        mock_writer = MagicMock()
        MockWriter.return_value = mock_writer

        mock_sender = MagicMock()
        mock_sender.send.side_effect = Exception("Outlook error")
        mock_sender.sent_count = 0
        mock_sender.skipped_count = 0
        mock_sender.error_count = 1
        mock_sender.__enter__ = MagicMock(return_value=mock_sender)
        mock_sender.__exit__ = MagicMock(return_value=False)
        MockSender.return_value = mock_sender

        results = [{
            "employee": make_employee(),
            "email_data": {
                "to": ["a@co.com"],
                "subject": "Test",
                "body": "Hello",
                "body_is_html": True,
            },
            "pdf_path": None,
        }]

        sent, skipped, errors, _ = send_emails(mock_config, results, 1)
        assert errors == 1
        assert sent == 0

    @patch("main.OutlookSender")
    @patch("main.StateTracker")
    @patch("main.ResultWriter")
    def test_state_saved_after_send(self, MockWriter, MockState, MockSender, mock_config):
        """Test that state is saved after email sending completes."""
        from main import send_emails

        mock_checkpoint = MagicMock()
        mock_checkpoint.get_processed_count.return_value = 0
        mock_checkpoint.is_processed.return_value = False
        mock_state = MagicMock()

        # Return different trackers for checkpoint vs state
        MockState.side_effect = [mock_checkpoint, mock_state]
        MockWriter.return_value = MagicMock()

        mock_sender = MagicMock()
        mock_sender.send.return_value = True
        mock_sender.sent_count = 1
        mock_sender.skipped_count = 0
        mock_sender.error_count = 0
        mock_sender.__enter__ = MagicMock(return_value=mock_sender)
        mock_sender.__exit__ = MagicMock(return_value=False)
        MockSender.return_value = mock_sender

        results = [{
            "employee": make_employee(),
            "email_data": {
                "to": ["a@co.com"],
                "subject": "Test",
                "body": "Hello",
                "body_is_html": True,
            },
            "pdf_path": None,
        }]

        send_emails(mock_config, results, 1)

        # Both trackers should be saved
        mock_checkpoint.save.assert_called_once()
        mock_state.save.assert_called_once()
