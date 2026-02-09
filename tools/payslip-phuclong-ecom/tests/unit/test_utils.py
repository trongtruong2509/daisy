"""
Unit tests for utils.py — ResultWriter, progress_interval, state analysis, cleanup.

Covers:
- Progress interval calculation
- CSV ResultWriter (creation, append, header)
- State analysis with checkpoint/state/result files
- File cleanup functions
- confirm_proceed prompt
"""

import csv
import json
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
TOOL_DIR = Path(__file__).resolve().parent.parent.parent
for p in [str(PROJECT_ROOT), str(TOOL_DIR)]:
    if p not in sys.path:
        sys.path.insert(0, p)

from utils import (
    CSV_COLUMNS,
    ResultWriter,
    analyze_existing_state,
    cleanup_all_files,
    cleanup_output_files,
    cleanup_pdf,
    confirm_proceed,
    progress_interval,
)


class TestProgressInterval:
    """Tests for progress_interval calculation."""

    def test_small_count(self):
        assert progress_interval(5) == 1
        assert progress_interval(20) == 1

    def test_medium_count(self):
        assert progress_interval(30) == 5
        assert progress_interval(50) == 5

    def test_large_count(self):
        assert progress_interval(100) == 10
        assert progress_interval(200) == 10

    def test_very_large_count(self):
        assert progress_interval(300) == 25
        assert progress_interval(500) == 25

    def test_huge_count(self):
        assert progress_interval(1000) == 50
        assert progress_interval(2000) == 50


class TestResultWriter:
    """Tests for CSV ResultWriter."""

    def test_creates_file_with_header(self, tmp_path):
        output = tmp_path / "results.csv"
        writer = ResultWriter(output, "01/2026")

        assert output.exists()
        with open(output, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            header = next(reader)
        assert header == CSV_COLUMNS

    def test_append_row(self, tmp_path):
        output = tmp_path / "results.csv"
        writer = ResultWriter(output, "01/2026")
        writer.append("001", "Nguyen Van A", "a@co.com", "SUCCESS", "payslip.pdf", "")

        with open(output, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            next(reader)  # header
            row = next(reader)

        assert row[0] == "001"  # employee_id
        assert row[1] == "Nguyen Van A"  # employee_name
        assert row[2] == "a@co.com"  # email
        assert row[3] == "payslip.pdf"  # payslip_filename
        assert row[4] == "SUCCESS"  # status
        assert row[5]  # timestamp not empty
        assert row[6] == ""  # error_message

    def test_append_multiple_rows(self, tmp_path):
        output = tmp_path / "results.csv"
        writer = ResultWriter(output, "01/2026")
        writer.append("001", "A", "a@co.com", "SUCCESS", "a.pdf", "")
        writer.append("002", "B", "b@co.com", "FAILED", "b.pdf", "Error msg")
        writer.append("003", "C", "c@co.com", "DRY_RUN", "c.pdf", "")

        with open(output, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            rows = list(reader)

        assert len(rows) == 4  # header + 3 data rows

    def test_does_not_overwrite_existing(self, tmp_path):
        """If file exists, should not create new header."""
        output = tmp_path / "results.csv"
        writer1 = ResultWriter(output, "01/2026")
        writer1.append("001", "A", "a@co.com", "SUCCESS", "", "")

        writer2 = ResultWriter(output, "01/2026")
        writer2.append("002", "B", "b@co.com", "SUCCESS", "", "")

        with open(output, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            rows = list(reader)

        # Should have 1 header + 2 rows (not 2 headers)
        assert len(rows) == 3

    def test_creates_parent_directories(self, tmp_path):
        output = tmp_path / "nested" / "dir" / "results.csv"
        writer = ResultWriter(output, "01/2026")
        assert output.exists()


class TestAnalyzeExistingState:
    """Tests for analyze_existing_state."""

    def test_no_state_files(self, mock_config):
        result = analyze_existing_state(mock_config)
        assert result["has_state"] is False
        assert result["sent_count"] == 0

    def test_checkpoint_file_detected(self, mock_config):
        checkpoint = mock_config.state_dir / "payslip_checkpoint_send_012026_state.json"
        checkpoint.write_text(json.dumps({"total_processed": 10, "processed_ids": []}))

        result = analyze_existing_state(mock_config)
        assert result["has_state"] is True
        assert result["sent_count"] == 10

    def test_state_file_detected(self, mock_config):
        state_file = mock_config.state_dir / "payslip_send_012026_state.json"
        state_file.write_text(json.dumps({"processed_ids": ["001", "002"]}))

        result = analyze_existing_state(mock_config)
        assert result["has_state"] is True
        assert result["state_file"] == state_file

    def test_result_file_detected(self, mock_config):
        result_file = mock_config.output_dir / "sent_results_012026.csv"
        with open(result_file, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(CSV_COLUMNS)
            writer.writerow(["001", "A", "a@co.com", "a.pdf", "SUCCESS", "2026-01-01", ""])

        result = analyze_existing_state(mock_config)
        assert result["has_state"] is True
        assert result["total_in_results"] == 1


class TestCleanupFunctions:
    """Tests for cleanup functions."""

    def test_cleanup_output_files(self, mock_config):
        # Create dummy files
        (mock_config.output_dir / "test.pdf").touch()
        (mock_config.output_dir / "test.csv").touch()
        (mock_config.output_dir / "test.txt").touch()

        cleanup_output_files(mock_config)

        assert not (mock_config.output_dir / "test.pdf").exists()
        assert not (mock_config.output_dir / "test.csv").exists()
        assert not (mock_config.output_dir / "test.txt").exists()

    def test_cleanup_all_files(self, mock_config):
        # Create state files
        cp = mock_config.state_dir / "payslip_checkpoint_send_012026_state.json"
        cp.write_text("{}")
        st = mock_config.state_dir / "payslip_send_012026_state.json"
        st.write_text("{}")
        (mock_config.output_dir / "test.pdf").touch()

        cleanup_all_files(mock_config)

        assert not cp.exists()
        assert not st.exists()
        assert not (mock_config.output_dir / "test.pdf").exists()

    def test_cleanup_pdf_existing(self, tmp_path):
        pdf = tmp_path / "test.pdf"
        pdf.touch()
        cleanup_pdf(pdf)
        assert not pdf.exists()

    def test_cleanup_pdf_nonexistent(self, tmp_path):
        """Should not raise for missing file."""
        pdf = tmp_path / "nonexistent.pdf"
        cleanup_pdf(pdf)  # No exception

    def test_cleanup_pdf_none(self):
        """Should not raise for None path."""
        cleanup_pdf(None)  # No exception


class TestConfirmProceed:
    """Tests for confirm_proceed user prompt."""

    @patch("builtins.input", return_value="yes")
    def test_yes_returns_true(self, mock_input):
        assert confirm_proceed() is True

    @patch("builtins.input", return_value="y")
    def test_y_returns_true(self, mock_input):
        assert confirm_proceed() is True

    @patch("builtins.input", return_value="no")
    def test_no_returns_false(self, mock_input):
        assert confirm_proceed() is False

    @patch("builtins.input", return_value="n")
    def test_n_returns_false(self, mock_input):
        assert confirm_proceed() is False

    @patch("builtins.input", side_effect=["maybe", "yes"])
    def test_invalid_then_valid(self, mock_input):
        assert confirm_proceed() is True
