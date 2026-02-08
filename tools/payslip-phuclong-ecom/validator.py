"""
Data validation for payslip processing.

Validates employee data before payslip generation or email sending.
Follows fail-fast approach: if any issues are found, processing terminates.
"""

import logging
import re
from collections import Counter
from typing import Any, Dict, List, Tuple

logger = logging.getLogger(__name__)

EMAIL_PATTERN = re.compile(
    r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$"
)


class ValidationError:
    """A single validation error."""

    def __init__(self, row: int, field: str, message: str):
        self.row = row
        self.field = field
        self.message = message

    def __str__(self):
        return f"Row {self.row}: [{self.field}] {self.message}"


class DataValidator:
    """
    Validates employee data for payslip processing.

    Checks:
    - Required fields present (MNV, Name, Email, Password)
    - Email format validation
    - Duplicate email detection
    - Empty/missing data detection
    """

    def __init__(self, employees: List[Dict[str, Any]]):
        self.employees = employees
        self.errors: List[ValidationError] = []
        self.warnings: List[ValidationError] = []

    def validate_all(self) -> Tuple[List[ValidationError], List[ValidationError]]:
        """Run all validations. Returns (errors, warnings)."""
        self.errors = []
        self.warnings = []

        if not self.employees:
            self.errors.append(ValidationError(0, "data", "No employee data found"))
            return self.errors, self.warnings

        self._validate_required_fields()
        self._validate_email_format()
        self._validate_duplicate_emails()
        self._validate_passwords()

        if self.errors:
            logger.error(f"Validation failed with {len(self.errors)} error(s)")
            for err in self.errors:
                logger.error(f"  {err}")
        if self.warnings:
            logger.warning(f"Validation has {len(self.warnings)} warning(s)")

        if not self.errors:
            logger.info(f"Validation passed: {len(self.employees)} employees OK")

        return self.errors, self.warnings

    def _validate_required_fields(self):
        """Check that required fields are non-empty."""
        for emp in self.employees:
            row = emp.get("row", 0)
            if not emp.get("mnv"):
                self.errors.append(
                    ValidationError(row, "MNV", "Employee ID (MNV) is empty")
                )
            if not emp.get("name"):
                self.warnings.append(
                    ValidationError(row, "Name", "Employee name is empty")
                )
            if not emp.get("email"):
                self.errors.append(
                    ValidationError(row, "Email", "Email address is empty")
                )

    def _validate_email_format(self):
        """Validate email format with regex."""
        for emp in self.employees:
            row = emp.get("row", 0)
            email = emp.get("email", "")
            if email and not EMAIL_PATTERN.match(email):
                self.errors.append(
                    ValidationError(row, "Email", f"Invalid email format: '{email}'")
                )

    def _validate_duplicate_emails(self):
        """Check for duplicate email addresses."""
        email_counts = Counter(
            emp.get("email", "").lower()
            for emp in self.employees
            if emp.get("email")
        )
        for email, count in email_counts.items():
            if count > 1:
                rows = [
                    emp.get("row", 0)
                    for emp in self.employees
                    if emp.get("email", "").lower() == email
                ]
                self.errors.append(
                    ValidationError(
                        rows[0], "Email",
                        f"Duplicate email '{email}' found in rows: {rows}",
                    )
                )

    def _validate_passwords(self):
        """Check that passwords are present."""
        for emp in self.employees:
            row = emp.get("row", 0)
            if not emp.get("password", ""):
                self.errors.append(
                    ValidationError(row, "Password", "Password is empty")
                )

    @property
    def is_valid(self) -> bool:
        return len(self.errors) == 0
