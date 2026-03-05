"""
Unit tests for check_reuse_existing_pdfs and build_results_from_existing_pdfs.

Covers the PDF-reuse feature in main.py that lets users skip regeneration
when PDFs from a previous run already exist in the output directory.
"""

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
TOOL_DIR = Path(__file__).resolve().parent.parent.parent
for p in [str(PROJECT_ROOT), str(TOOL_DIR)]:
    if p not in sys.path:
        sys.path.insert(0, p)

from tests.conftest import make_employee, make_employees


class TestCheckReuseExistingPdfs:
    """Tests for check_reuse_existing_pdfs."""

    def test_no_output_dir_returns_false(self, tmp_path):
        from main import check_reuse_existing_pdfs

        config = MagicMock()
        config.output_dir = tmp_path / "nonexistent"
        # output_dir does not exist
        result = check_reuse_existing_pdfs(config, make_employees(2))
        assert result is False

    def test_no_pdfs_returns_false(self, tmp_path):
        from main import check_reuse_existing_pdfs

        config = MagicMock()
        config.output_dir = tmp_path
        # Directory exists but no PDFs
        result = check_reuse_existing_pdfs(config, make_employees(2))
        assert result is False

    def test_unmatched_pdfs_returns_false(self, tmp_path):
        """PDFs exist but don't match any employee names."""
        from main import check_reuse_existing_pdfs

        config = MagicMock()
        config.output_dir = tmp_path
        config.pdf_filename_pattern = "TBKQ_{name}_{mmyyyy}"
        config.date_mmyyyy = "012026"

        # Create PDFs that don't match employee names
        (tmp_path / "TBKQ_UnknownPerson_012026.pdf").touch()

        emps = [make_employee(name="Alice"), make_employee(mnv="002", name="Bob")]
        result = check_reuse_existing_pdfs(config, emps)
        assert result is False

    @patch("builtins.input", return_value="yes")
    def test_matched_pdfs_prompts_user_yes(self, mock_input, tmp_path):
        """User chooses to reuse when matching PDFs found."""
        from main import check_reuse_existing_pdfs

        config = MagicMock()
        config.output_dir = tmp_path
        config.pdf_filename_pattern = "TBKQ_{name}_{mmyyyy}"
        config.date_mmyyyy = "012026"

        # Create a matching PDF
        (tmp_path / "TBKQ_Alice_012026.pdf").touch()

        emps = [make_employee(name="Alice")]
        result = check_reuse_existing_pdfs(config, emps)
        assert result is True

    @patch("builtins.input", return_value="no")
    def test_matched_pdfs_prompts_user_no(self, mock_input, tmp_path):
        """User declines to reuse when matching PDFs found."""
        from main import check_reuse_existing_pdfs

        config = MagicMock()
        config.output_dir = tmp_path
        config.pdf_filename_pattern = "TBKQ_{name}_{mmyyyy}"
        config.date_mmyyyy = "012026"

        (tmp_path / "TBKQ_Alice_012026.pdf").touch()

        emps = [make_employee(name="Alice")]
        result = check_reuse_existing_pdfs(config, emps)
        assert result is False

    @patch("builtins.input", return_value="y")
    def test_partial_match_counts_correctly(self, mock_input, tmp_path):
        """Only some employees have matching PDFs."""
        from main import check_reuse_existing_pdfs

        config = MagicMock()
        config.output_dir = tmp_path
        config.pdf_filename_pattern = "TBKQ_{name}_{mmyyyy}"
        config.date_mmyyyy = "012026"

        # Only Alice has a PDF
        (tmp_path / "TBKQ_Alice_012026.pdf").touch()

        emps = [
            make_employee(mnv="001", name="Alice"),
            make_employee(mnv="002", name="Bob"),
        ]
        result = check_reuse_existing_pdfs(config, emps)
        assert result is True  # User said yes


class TestBuildResultsFromExistingPdfs:
    """Tests for build_results_from_existing_pdfs."""

    def test_all_pdfs_exist(self, tmp_path):
        from main import build_results_from_existing_pdfs

        config = MagicMock()
        config.output_dir = tmp_path
        config.pdf_filename_pattern = "TBKQ_{name}_{mmyyyy}"
        config.date_mmyyyy = "012026"

        # Create matching PDFs
        (tmp_path / "TBKQ_Alice_012026.pdf").touch()
        (tmp_path / "TBKQ_Bob_012026.pdf").touch()

        emps = [
            make_employee(mnv="001", name="Alice"),
            make_employee(mnv="002", name="Bob"),
        ]

        results = build_results_from_existing_pdfs(config, emps)
        assert len(results) == 2
        assert all(r["success"] for r in results)
        assert all(r.get("pdf_skipped") for r in results)
        for r in results:
            assert r["pdf_path"] is not None
            assert r["pdf_path"].endswith(".pdf")

    def test_no_pdfs_exist(self, tmp_path):
        from main import build_results_from_existing_pdfs

        config = MagicMock()
        config.output_dir = tmp_path
        config.pdf_filename_pattern = "TBKQ_{name}_{mmyyyy}"
        config.date_mmyyyy = "012026"

        emps = [make_employee(name="Alice")]
        results = build_results_from_existing_pdfs(config, emps)
        assert len(results) == 1
        assert results[0]["success"] is False
        assert results[0]["pdf_path"] is None

    def test_partial_pdfs(self, tmp_path):
        from main import build_results_from_existing_pdfs

        config = MagicMock()
        config.output_dir = tmp_path
        config.pdf_filename_pattern = "TBKQ_{name}_{mmyyyy}"
        config.date_mmyyyy = "012026"

        (tmp_path / "TBKQ_Alice_012026.pdf").touch()
        # Bob pdf does NOT exist

        emps = [
            make_employee(mnv="001", name="Alice"),
            make_employee(mnv="002", name="Bob"),
        ]

        results = build_results_from_existing_pdfs(config, emps)
        assert results[0]["success"] is True
        assert results[0]["pdf_path"] is not None
        assert results[1]["success"] is False
        assert results[1]["pdf_path"] is None

    def test_xlsx_included_when_exists(self, tmp_path):
        from main import build_results_from_existing_pdfs

        config = MagicMock()
        config.output_dir = tmp_path
        config.pdf_filename_pattern = "TBKQ_{name}_{mmyyyy}"
        config.date_mmyyyy = "012026"

        (tmp_path / "TBKQ_Alice_012026.pdf").touch()
        (tmp_path / "TBKQ_Alice_012026.xlsx").touch()

        emps = [make_employee(name="Alice")]
        results = build_results_from_existing_pdfs(config, emps)
        assert results[0]["xlsx_path"] is not None

    def test_xlsx_none_when_missing(self, tmp_path):
        from main import build_results_from_existing_pdfs

        config = MagicMock()
        config.output_dir = tmp_path
        config.pdf_filename_pattern = "TBKQ_{name}_{mmyyyy}"
        config.date_mmyyyy = "012026"

        (tmp_path / "TBKQ_Alice_012026.pdf").touch()
        # No xlsx file

        emps = [make_employee(name="Alice")]
        results = build_results_from_existing_pdfs(config, emps)
        assert results[0]["xlsx_path"] is None
