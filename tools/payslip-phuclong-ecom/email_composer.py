"""
Email composer for payslip distribution.

Builds personalized HTML emails from the bodymail template,
matching the VBA format for consistency.
"""

import re
from pathlib import Path
from typing import Any, Dict, List

from core.logger import get_logger

logger = get_logger(__name__)


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
        """Build HTML email body from template cells with proper paragraph spacing.

        Uses row gaps between cells to determine paragraph spacing:
        - Consecutive rows (gap=1): single line break <br />
        - Skipped rows (gap>1): double line break <br /><br /> for paragraph spacing
        """
        import re as _re

        def _cell_row(cell_ref: str) -> int:
            """Extract row number from cell reference like 'A5' -> 5."""
            match = _re.match(r"[A-Z]+(\d+)", cell_ref, _re.IGNORECASE)
            return int(match.group(1)) if match else 0

        # Sort cells by row number to maintain correct order
        sorted_cells = sorted(self.template_cells.keys(), key=_cell_row)
        if not sorted_cells:
            return ""

        html_parts = []
        prev_row = None

        for cell in sorted_cells:
            value = self.template_cells.get(cell, "").strip()
            row_num = _cell_row(cell)

            # Add paragraph spacing for row gaps > 1
            if prev_row is not None and row_num - prev_row > 1:
                html_parts.append("")  # Creates <br /><br /> when joined

            if not value:
                prev_row = row_num
                continue

            # Apply bold formatting for password hint cell
            if cell.upper() == "A5":
                html_parts.append(f"<strong>{value}</strong>")
            else:
                html_parts.append(value)

            prev_row = row_num

        return "<br />".join(html_parts)

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
