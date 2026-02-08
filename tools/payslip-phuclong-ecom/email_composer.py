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
        self.template_cells = dict(template_cells)
        self.subject = subject
        self.date_str = date_str
        self.date_cell = date_cell
        self._apply_date_replacement()

    def _apply_date_replacement(self):
        """Replace date references in the email template with actual date."""
        if not self.date_str or not self.date_cell:
            return

        cell_value = self.template_cells.get(self.date_cell, "")
        if cell_value:
            updated = re.sub(
                r"tháng\s+\d{1,2}/\d{4}",
                f"tháng {self.date_str}",
                cell_value,
            )
            updated = re.sub(r"\d{2}/\d{4}", self.date_str, updated)
            self.template_cells[self.date_cell] = updated

        if self.subject:
            def _replace_date(m):
                prefix = m.group(0).split()[0]
                return f"{prefix} {self.date_str}"

            self.subject = re.sub(
                r"(?:tháng|THÁNG)\s+\d{1,2}/\d{4}",
                _replace_date,
                self.subject,
            )

    def compose_html_body(self) -> str:
        """Build HTML email body from template cells matching VBA format."""
        parts = []
        for cell in self.template_cells:
            value = self.template_cells.get(cell, "").strip()
            if not value:
                parts.append("")
                continue
            if cell == "A5":
                parts.append(f"<strong>{value}</strong>")
            else:
                parts.append(value)

        return "<br />".join(parts)

    def compose_email(
        self,
        employee: Dict[str, Any],
        pdf_path: Path,
    ) -> Dict[str, Any]:
        """Compose a complete email for an employee."""
        email_addr = employee.get("email", "")
        name = employee.get("name", "")

        if not email_addr:
            logger.warning(f"No email for employee: {name}")
            return None

        return {
            "to": [email_addr],
            "subject": self.subject,
            "body": self.compose_html_body(),
            "body_is_html": True,
            "attachments": [Path(pdf_path)] if pdf_path else [],
        }

    def compose_batch(self, items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Compose emails for all employees."""
        for item in items:
            emp = item.get("employee", {})
            pdf_path = item.get("pdf_path")

            if pdf_path and Path(pdf_path).exists():
                item["email_data"] = self.compose_email(emp, pdf_path)
            else:
                item["email_data"] = None
                logger.warning(f"No PDF for {emp.get('name', 'N/A')}, email not composed")

        composed = sum(1 for item in items if item.get("email_data"))
        logger.info(f"Emails composed: {composed}/{len(items)}")
        return items
