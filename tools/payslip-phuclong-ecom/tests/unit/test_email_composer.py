"""
Unit tests for email_composer.py — EmailComposer.

Covers:
- Date replacement in template cells
- Date replacement in email subject
- HTML body composition with paragraph spacing
- Email composition for single employee
- Batch email composition
- Missing PDF handling
- Bold formatting for A5 cell
"""

import sys
from pathlib import Path
from unittest.mock import MagicMock

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
TOOL_DIR = Path(__file__).resolve().parent.parent.parent
for p in [str(PROJECT_ROOT), str(TOOL_DIR)]:
    if p not in sys.path:
        sys.path.insert(0, p)

from email_composer import EmailComposer
from tests.conftest import make_employee


class TestDateReplacement:
    """Tests for date replacement in template cells and subject."""

    def test_date_replaces_in_template_cell(self):
        template = {
            "A1": "Dear Employee,",
            "A3": "Your payslip for tháng 01/2025 is attached.",
        }
        composer = EmailComposer(template, "Subject", "02/2026", "A3")
        assert "02/2026" in composer.template_cells["A3"]
        assert "01/2025" not in composer.template_cells["A3"]

    def test_date_replaces_in_subject(self):
        template = {"A1": "Hello"}
        composer = EmailComposer(
            template,
            "Phiếu lương tháng 01/2025",
            "03/2026",
            "A1",
        )
        assert "03/2026" in composer.subject
        assert "01/2025" not in composer.subject

    def test_no_date_replacement_when_empty(self):
        template = {"A1": "No date here"}
        composer = EmailComposer(template, "Subject", "", "A1")
        assert composer.template_cells["A1"] == "No date here"

    def test_no_crash_when_date_cell_not_in_template(self):
        template = {"A1": "Hello"}
        composer = EmailComposer(template, "Subject", "01/2026", "A99")
        assert composer.template_cells["A1"] == "Hello"


class TestHtmlBodyComposition:
    """Tests for HTML body generation."""

    def test_basic_html_body(self):
        template = {
            "A1": "Line 1",
            "A3": "Line 2",
        }
        composer = EmailComposer(template, "Subject", "01/2026", "A3")
        body = composer.compose_html_body()
        assert "Line 1" in body
        assert "<br />" in body

    def test_empty_template_returns_empty(self):
        composer = EmailComposer({}, "Subject", "01/2026", "A1")
        assert composer.compose_html_body() == ""

    def test_paragraph_spacing_for_row_gaps(self):
        """Skipped rows should create double line breaks."""
        template = {
            "A1": "Paragraph 1",
            "A5": "Paragraph 2",  # Gap of 4 rows
        }
        composer = EmailComposer(template, "Subject", "", "")
        body = composer.compose_html_body()
        # Row gap > 1 should produce an empty part (double <br />)
        assert "<br /><br />" in body

    def test_consecutive_rows_single_break(self):
        """Consecutive rows should have single line break."""
        template = {
            "A1": "Line 1",
            "A2": "Line 2",
        }
        composer = EmailComposer(template, "Subject", "", "")
        body = composer.compose_html_body()
        parts = body.split("<br />")
        # Should not have empty parts (which would mean double break)
        assert "" not in parts or parts.count("") == 0

    def test_bold_formatting_a5(self):
        """A5 cell should be wrapped in <strong> tags."""
        template = {
            "A1": "Hello",
            "A5": "Password hint",
        }
        composer = EmailComposer(template, "Subject", "", "")
        body = composer.compose_html_body()
        assert "<strong>Password hint</strong>" in body

    def test_empty_cell_values_skipped(self):
        template = {
            "A1": "Hello",
            "A3": "",
            "A5": "End",
        }
        composer = EmailComposer(template, "Subject", "", "")
        body = composer.compose_html_body()
        assert "Hello" in body
        assert "End" in body


class TestComposeEmail:
    """Tests for single email composition."""

    def test_compose_valid_email(self, tmp_path):
        pdf = tmp_path / "test.pdf"
        pdf.touch()

        template = {"A1": "Hello"}
        composer = EmailComposer(template, "Test Subject", "01/2026", "A1")

        emp = make_employee()
        result = composer.compose_email(emp, pdf)

        assert result is not None
        assert result["to"] == ["a@company.com"]
        assert result["subject"] == "Test Subject"
        assert result["body_is_html"] is True
        assert len(result["attachments"]) == 1

    def test_compose_no_email_returns_none(self, tmp_path):
        pdf = tmp_path / "test.pdf"
        pdf.touch()

        template = {"A1": "Hello"}
        composer = EmailComposer(template, "Subject", "01/2026", "A1")

        emp = make_employee(email="")
        result = composer.compose_email(emp, pdf)
        assert result is None


class TestComposeBatch:
    """Tests for batch email composition."""

    def test_compose_batch_all_valid(self, tmp_path):
        pdfs = []
        items = []
        for i in range(3):
            pdf = tmp_path / f"test_{i}.pdf"
            pdf.touch()
            pdfs.append(pdf)
            items.append({
                "employee": make_employee(
                    row=i + 4, mnv=f"{i:03d}",
                    email=f"emp{i}@co.com", password=f"{i:03d}",
                ),
                "pdf_path": str(pdf),
            })

        template = {"A1": "Hello"}
        composer = EmailComposer(template, "Subject", "01/2026", "A1")
        results = composer.compose_batch(items)

        composed = sum(1 for r in results if r.get("email_data"))
        assert composed == 3

    def test_compose_batch_missing_pdf(self, tmp_path):
        items = [{
            "employee": make_employee(),
            "pdf_path": str(tmp_path / "missing.pdf"),
        }]

        template = {"A1": "Hello"}
        composer = EmailComposer(template, "Subject", "01/2026", "A1")
        results = composer.compose_batch(items)

        assert results[0].get("email_data") is None

    def test_compose_batch_none_pdf(self):
        items = [{
            "employee": make_employee(),
            "pdf_path": None,
        }]

        template = {"A1": "Hello"}
        composer = EmailComposer(template, "Subject", "01/2026", "A1")
        results = composer.compose_batch(items)

        assert results[0].get("email_data") is None
