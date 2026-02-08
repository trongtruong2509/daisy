"""
Email composer for payslip distribution.

Builds personalized HTML emails from the bodymail template,
matching the VBA format for consistency.
"""

import logging
import re
from pathlib import Path
from typing import Any, Dict, List

logger = logging.getLogger(__name__)


class EmailComposer:
    """
    Composes payslip emails from bodymail template.

    Reads cell values from the bodymail sheet and constructs HTML body
    matching the original VBA email format.
    """

    def __init__(
        self,
        template_cells: Dict[str, str],
        subject: str,
        date_str: str,
        date_cell: str = "A3",
    ):
        """
        Args:
            template_cells: Dict of cell_ref -> text value from bodymail sheet.
            subject: Email subject (from TBKQ G1 or .env override).
            date_str: Payroll date (MM/YYYY format).
            date_cell: Cell reference that contains the date placeholder.
        """
        self.template_cells = dict(template_cells)
        self.subject = subject
        self.date_str = date_str
        self.date_cell = date_cell

        # Replace date placeholder in the template
        self._apply_date_replacement()

    def _apply_date_replacement(self):
        """Replace date references in the email template with actual date."""
        if not self.date_str or not self.date_cell:
            return

        cell_value = self.template_cells.get(self.date_cell, "")
        if cell_value:
            # Replace patterns like "tháng 11/2025" or "tháng XX/XXXX"
            updated = re.sub(
                r"tháng\s+\d{1,2}/\d{4}",
                f"tháng {self.date_str}",
                cell_value,
            )
            # Also replace standalone MM/YYYY patterns
            updated = re.sub(
                r"\d{2}/\d{4}",
                self.date_str,
                updated,
            )
            self.template_cells[self.date_cell] = updated

        # Also update subject if it contains date reference
        if self.subject:
            # Case-insensitive replacement for THÁNG/tháng
            def _replace_date(m):
                prefix = m.group(0).split()[0]  # Keep original case
                return f"{prefix} {self.date_str}"

            self.subject = re.sub(
                r"(?:tháng|THÁNG)\s+\d{1,2}/\d{4}",
                _replace_date,
                self.subject,
            )

    def compose_html_body(self) -> str:
        """
        Build HTML email body from template cells.

        Matches the VBA HTMLBody format:
        - Each cell value becomes a line separated by <br />
        - Cell A5 is wrapped in <strong> tags
        - Cells are joined with <br /> separators

        Returns:
            HTML body string.
        """
        # Maintain the original cell order from config (not alphabetical)
        body_cells = list(self.template_cells.keys())

        parts = []
        for cell in body_cells:
            value = self.template_cells.get(cell, "").strip()
            if not value:
                # Empty cells still add a line break (matching VBA behavior)
                parts.append("")
                continue

            # A5 is typically the password instruction — make it bold
            if cell == "A5":
                parts.append(f"<strong>{value}</strong>")
            else:
                parts.append(value)

        # Join with <br /> matching VBA format
        html_body = "<br />".join(parts)
        return html_body

    def compose_email(
        self,
        employee: Dict[str, Any],
        pdf_path: Path,
    ) -> Dict[str, Any]:
        """
        Compose a complete email for an employee.

        Args:
            employee: Employee data dict with 'email', 'name', etc.
            pdf_path: Path to the password-protected PDF attachment.

        Returns:
            Dict with 'to', 'subject', 'body', 'body_is_html',
            'attachments' keys ready for NewEmail construction.
        """
        email_addr = employee.get("email", "")
        name = employee.get("name", "")

        if not email_addr:
            logger.warning(f"No email for employee: {name}")
            return None

        html_body = self.compose_html_body()

        return {
            "to": [email_addr],
            "subject": self.subject,
            "body": html_body,
            "body_is_html": True,
            "attachments": [Path(pdf_path)] if pdf_path else [],
        }

    def compose_batch(
        self,
        items: List[Dict[str, Any]],
    ) -> List[Dict[str, Any]]:
        """
        Compose emails for all employees.

        Args:
            items: List of dicts with 'employee' and 'pdf_path' keys.

        Returns:
            Updated list with 'email_data' added to each dict.
        """
        for item in items:
            emp = item.get("employee", {})
            pdf_path = item.get("pdf_path")

            if pdf_path and Path(pdf_path).exists():
                email_data = self.compose_email(emp, pdf_path)
                item["email_data"] = email_data
            else:
                item["email_data"] = None
                name = emp.get("name", "N/A")
                logger.warning(f"No PDF for {name}, email not composed")

        composed = sum(1 for item in items if item.get("email_data"))
        logger.info(
            f"Emails composed: {composed}/{len(items)}"
        )
        return items
