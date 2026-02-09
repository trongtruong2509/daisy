"""Unit tests for email_composer module."""

import sys
from pathlib import Path

import pytest

TOOL_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(TOOL_DIR))

from email_composer import EmailComposer


class TestEmailComposer:
    """Tests for EmailComposer."""

    def test_compose_html_body(self, email_template):
        composer = EmailComposer(
            template_cells=email_template,
            subject="Test Subject",
            date_str="01/2026",
        )
        html = composer.compose_html_body()
        assert "<br />" in html
        assert "Kính gửi Anh/Chị," in html
        assert "ADECCO VIỆT NAM" in html

    def test_date_replacement_in_body(self, email_template):
        composer = EmailComposer(
            template_cells=email_template,
            subject="Test",
            date_str="01/2026",
            date_cell="A3",
        )
        html = composer.compose_html_body()
        assert "01/2026" in html
        assert "11/2025" not in html

    def test_date_replacement_in_subject(self, email_template):
        composer = EmailComposer(
            template_cells=email_template,
            subject="THÔNG BÁO PHIẾU LƯƠNG KỲ LƯƠNG THÁNG 11/2025",
            date_str="01/2026",
        )
        assert "01/2026" in composer.subject
        assert "11/2025" not in composer.subject

    def test_subject_lowercase_thang(self, email_template):
        composer = EmailComposer(
            template_cells=email_template,
            subject="Phiếu lương tháng 09/2025",
            date_str="03/2026",
        )
        assert "03/2026" in composer.subject
        assert "09/2025" not in composer.subject

    def test_a5_bold(self, email_template):
        composer = EmailComposer(
            template_cells=email_template,
            subject="Test",
            date_str="01/2026",
        )
        html = composer.compose_html_body()
        assert "<strong>" in html
        assert "Mật khẩu" in html

    def test_cell_order_preserved(self, email_template):
        composer = EmailComposer(
            template_cells=email_template,
            subject="Test",
            date_str="01/2026",
        )
        html = composer.compose_html_body()
        # A1 should come before A3
        idx_a1 = html.find("Kính gửi")
        idx_a3 = html.find("Công ty gửi")
        assert idx_a1 < idx_a3

    def test_compose_email(self, email_template, sample_employee, tmp_path):
        # Create a dummy PDF
        dummy_pdf = tmp_path / "test.pdf"
        dummy_pdf.write_text("dummy")

        composer = EmailComposer(
            template_cells=email_template,
            subject="Payslip 01/2026",
            date_str="01/2026",
        )
        result = composer.compose_email(sample_employee, dummy_pdf)
        assert result is not None
        assert result["to"] == ["test@example.com"]
        assert result["subject"] == "Payslip 01/2026"
        assert result["body_is_html"] is True
        assert len(result["attachments"]) == 1

    def test_compose_email_no_email(self, email_template, sample_employee, tmp_path):
        sample_employee["email"] = ""
        dummy_pdf = tmp_path / "test.pdf"
        dummy_pdf.write_text("dummy")

        composer = EmailComposer(
            template_cells=email_template,
            subject="Test",
            date_str="01/2026",
        )
        result = composer.compose_email(sample_employee, dummy_pdf)
        assert result is None

    def test_compose_batch(self, email_template, sample_employees, tmp_path):
        # Create dummy PDFs
        items = []
        for emp in sample_employees:
            pdf = tmp_path / f"{emp['mnv']}.pdf"
            pdf.write_text("dummy")
            items.append({"employee": emp, "pdf_path": pdf})

        composer = EmailComposer(
            template_cells=email_template,
            subject="Test",
            date_str="01/2026",
        )
        results = composer.compose_batch(items)
        assert all(r.get("email_data") is not None for r in results)
