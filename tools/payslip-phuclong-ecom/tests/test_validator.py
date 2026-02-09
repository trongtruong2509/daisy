"""Unit tests for validator module."""

import sys
from pathlib import Path

import pytest

TOOL_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(TOOL_DIR))

from validator import DataValidator, ValidationError


class TestDataValidator:
    """Tests for DataValidator."""

    def test_valid_employees(self, sample_employees):
        validator = DataValidator(sample_employees)
        errors, warnings = validator.validate_all()
        assert len(errors) == 0

    def test_empty_list(self):
        validator = DataValidator([])
        errors, _ = validator.validate_all()
        assert len(errors) == 1
        assert "No employee data" in str(errors[0])

    def test_missing_mnv(self, sample_employee):
        sample_employee["mnv"] = ""
        validator = DataValidator([sample_employee])
        errors, _ = validator.validate_all()
        assert any("MNV" in str(e) for e in errors)

    def test_missing_email(self, sample_employee):
        sample_employee["email"] = ""
        validator = DataValidator([sample_employee])
        errors, _ = validator.validate_all()
        assert any("Email" in str(e) and "empty" in str(e) for e in errors)

    def test_invalid_email_format(self, sample_employee):
        sample_employee["email"] = "not-an-email"
        validator = DataValidator([sample_employee])
        errors, _ = validator.validate_all()
        assert any("Invalid email" in str(e) for e in errors)

    def test_valid_email_formats(self):
        emails = [
            "user@example.com",
            "user.name@company.co.vn",
            "user+tag@domain.org",
            "a@b.cd",
        ]
        for email in emails:
            emp = {
                "row": 1, "mnv": "123", "name": "Test",
                "email": email, "password": "123",
            }
            validator = DataValidator([emp])
            errors, _ = validator.validate_all()
            assert not any("Invalid email" in str(e) for e in errors), f"{email} should be valid"

    def test_duplicate_emails(self, sample_employees):
        sample_employees[1]["email"] = sample_employees[0]["email"]
        validator = DataValidator(sample_employees)
        errors, _ = validator.validate_all()
        assert any("Duplicate" in str(e) for e in errors)

    def test_duplicate_emails_case_insensitive(self, sample_employee):
        emp2 = dict(sample_employee)
        emp2["row"] = 5
        emp2["mnv"] = "999"
        emp2["email"] = sample_employee["email"].upper()
        validator = DataValidator([sample_employee, emp2])
        errors, _ = validator.validate_all()
        assert any("Duplicate" in str(e) for e in errors)

    def test_missing_password(self, sample_employee):
        sample_employee["password"] = ""
        validator = DataValidator([sample_employee])
        errors, _ = validator.validate_all()
        assert any("Password" in str(e) for e in errors)

    def test_empty_name_is_warning(self, sample_employee):
        sample_employee["name"] = ""
        validator = DataValidator([sample_employee])
        errors, warnings = validator.validate_all()
        assert len(errors) == 0
        assert any("Name" in str(w) for w in warnings)

    def test_is_valid_property(self, sample_employee):
        validator = DataValidator([sample_employee])
        validator.validate_all()
        assert validator.is_valid

    def test_not_valid_property(self, sample_employee):
        sample_employee["email"] = ""
        validator = DataValidator([sample_employee])
        validator.validate_all()
        assert not validator.is_valid

    def test_error_summary(self, sample_employee):
        sample_employee["email"] = "bad"
        sample_employee["password"] = ""
        validator = DataValidator([sample_employee])
        validator.validate_all()
        summary = validator.get_error_summary()
        assert "error" in summary.lower()
        assert "2" in summary  # 2 errors


class TestValidationError:
    """Tests for ValidationError."""

    def test_str_format(self):
        err = ValidationError(4, "Email", "is invalid")
        assert "Row 4" in str(err)
        assert "Email" in str(err)
        assert "is invalid" in str(err)

    def test_repr(self):
        err = ValidationError(5, "Name", "missing")
        r = repr(err)
        assert "ValidationError" in r
        assert "5" in r
