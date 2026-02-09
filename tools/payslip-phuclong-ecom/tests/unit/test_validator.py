"""
Unit tests for validator.py — DataValidator and ValidationError.

Covers:
- Required field validation (MNV, Name, Email, Password)
- Email format validation
- Duplicate email detection
- Allow duplicate emails override
- Empty employee list
- All-or-nothing validation policy
"""

import sys
from pathlib import Path

import pytest

# Ensure imports work
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
TOOL_DIR = Path(__file__).resolve().parent.parent.parent
for p in [str(PROJECT_ROOT), str(TOOL_DIR)]:
    if p not in sys.path:
        sys.path.insert(0, p)

from validator import DataValidator, ValidationError
from tests.conftest import make_employee, make_employees


class TestValidationError:
    """Tests for the ValidationError class."""

    def test_str_format(self):
        err = ValidationError(row=5, field="Email", message="Invalid format")
        assert str(err) == "Row 5: [Email] Invalid format"

    def test_attributes(self):
        err = ValidationError(row=10, field="MNV", message="Empty")
        assert err.row == 10
        assert err.field == "MNV"
        assert err.message == "Empty"


class TestDataValidatorRequiredFields:
    """Tests for required field validation."""

    def test_valid_employees_pass(self):
        employees = make_employees(3)
        v = DataValidator(employees)
        errors, warnings = v.validate_all()
        assert len(errors) == 0
        assert v.is_valid

    def test_missing_mnv_is_error(self):
        emp = make_employee(mnv="")
        v = DataValidator([emp])
        errors, _ = v.validate_all()
        assert any(e.field == "MNV" for e in errors)
        assert not v.is_valid

    def test_missing_email_is_error(self):
        emp = make_employee(email="")
        v = DataValidator([emp])
        errors, _ = v.validate_all()
        assert any(e.field == "Email" for e in errors)

    def test_missing_name_is_warning(self):
        emp = make_employee(name="")
        v = DataValidator([emp])
        errors, warnings = v.validate_all()
        assert len(errors) == 0  # Name missing is not fatal
        assert any(w.field == "Name" for w in warnings)

    def test_missing_password_is_error(self):
        emp = make_employee(password="")
        v = DataValidator([emp])
        errors, _ = v.validate_all()
        assert any(e.field == "Password" for e in errors)


class TestDataValidatorEmailFormat:
    """Tests for email format validation."""

    @pytest.mark.parametrize(
        "email",
        [
            "user@company.com",
            "first.last@domain.co.th",
            "user+tag@sub.domain.com",
            "a@b.cd",
        ],
    )
    def test_valid_emails(self, email):
        emp = make_employee(email=email)
        v = DataValidator([emp])
        errors, _ = v.validate_all()
        email_format_errors = [
            e for e in errors if e.field == "Email" and "Invalid email" in e.message
        ]
        assert len(email_format_errors) == 0

    @pytest.mark.parametrize(
        "email",
        [
            "not-an-email",
            "@no-local.com",
            "user@",
            "user@.com",
            "user name@domain.com",
        ],
    )
    def test_invalid_emails(self, email):
        emp = make_employee(email=email)
        v = DataValidator([emp])
        errors, _ = v.validate_all()
        email_errors = [
            e for e in errors if e.field == "Email" and "Invalid email" in e.message
        ]
        assert len(email_errors) == 1


class TestDataValidatorDuplicates:
    """Tests for duplicate email detection."""

    def test_duplicate_emails_detected(self):
        emps = [
            make_employee(row=4, mnv="001", email="same@co.com", password="001"),
            make_employee(row=5, mnv="002", email="same@co.com", password="002"),
        ]
        v = DataValidator(emps)
        errors, _ = v.validate_all()
        assert any("Duplicate" in e.message for e in errors)
        assert not v.is_valid

    def test_duplicate_emails_case_insensitive(self):
        emps = [
            make_employee(row=4, mnv="001", email="User@Co.com", password="001"),
            make_employee(row=5, mnv="002", email="user@co.com", password="002"),
        ]
        v = DataValidator(emps)
        errors, _ = v.validate_all()
        assert any("Duplicate" in e.message or "duplicate" in e.message for e in errors)

    def test_allow_duplicate_emails_config(self):
        emps = [
            make_employee(row=4, mnv="001", email="same@co.com", password="001"),
            make_employee(row=5, mnv="002", email="same@co.com", password="002"),
        ]
        v = DataValidator(emps, allow_duplicate_emails=True)
        errors, warnings = v.validate_all()
        # No errors for duplicates when allowed
        dup_errors = [e for e in errors if "Duplicate" in e.message or "duplicate" in e.message]
        assert len(dup_errors) == 0
        # But should still have warnings
        dup_warnings = [w for w in warnings if "Duplicate" in w.message or "duplicate" in w.message]
        assert len(dup_warnings) > 0

    def test_no_duplicate_when_unique(self):
        emps = make_employees(5)
        v = DataValidator(emps)
        errors, _ = v.validate_all()
        dup_errors = [e for e in errors if "Duplicate" in e.message or "duplicate" in e.message]
        assert len(dup_errors) == 0


class TestDataValidatorEdgeCases:
    """Edge case tests."""

    def test_empty_employee_list(self):
        v = DataValidator([])
        errors, _ = v.validate_all()
        assert len(errors) == 1
        assert "No employee data" in errors[0].message
        assert not v.is_valid

    def test_is_valid_property_true(self):
        emps = make_employees(2)
        v = DataValidator(emps)
        v.validate_all()
        assert v.is_valid is True

    def test_is_valid_property_false(self):
        emp = make_employee(mnv="", email="")
        v = DataValidator([emp])
        v.validate_all()
        assert v.is_valid is False

    def test_multiple_errors_accumulated(self):
        """Multiple validation errors should all be collected (fail-fast per FR-09)."""
        emps = [
            make_employee(row=4, mnv="", email="bad", password=""),
            make_employee(row=5, mnv="002", email="", password="002"),
        ]
        v = DataValidator(emps)
        errors, _ = v.validate_all()
        # Should have: missing MNV (row 4), invalid email (row 4),
        #              missing password (row 4), missing email (row 5)
        assert len(errors) >= 4
