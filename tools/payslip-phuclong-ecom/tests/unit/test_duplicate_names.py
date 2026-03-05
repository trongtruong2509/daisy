"""
Unit tests for duplicate name handling in PayslipGenerator.

Covers:
- _build_name_suffix_map() — suffix assignment for duplicate names
- _build_output_path() with name_suffix parameter
- generate_batch integration with suffix_map
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

from payslip_generator import PayslipGenerator
from tests.conftest import make_employee, make_employees


class TestBuildNameSuffixMap:
    """Tests for _build_name_suffix_map static method."""

    def test_unique_names_no_suffix(self):
        emps = [
            make_employee(mnv="001", name="Alice"),
            make_employee(mnv="002", name="Bob"),
            make_employee(mnv="003", name="Charlie"),
        ]
        result = PayslipGenerator._build_name_suffix_map(emps)
        assert result == {"001": "", "002": "", "003": ""}

    def test_duplicate_names_get_suffix(self):
        emps = [
            make_employee(mnv="001", name="Alice"),
            make_employee(mnv="002", name="Alice"),
        ]
        result = PayslipGenerator._build_name_suffix_map(emps)
        assert result["001"] == "_1"
        assert result["002"] == "_2"

    def test_three_duplicates(self):
        emps = [
            make_employee(mnv="001", name="Alice"),
            make_employee(mnv="002", name="Alice"),
            make_employee(mnv="003", name="Alice"),
        ]
        result = PayslipGenerator._build_name_suffix_map(emps)
        assert result["001"] == "_1"
        assert result["002"] == "_2"
        assert result["003"] == "_3"

    def test_mixed_unique_and_duplicate(self):
        emps = [
            make_employee(mnv="001", name="Alice"),
            make_employee(mnv="002", name="Bob"),
            make_employee(mnv="003", name="Alice"),
            make_employee(mnv="004", name="Charlie"),
        ]
        result = PayslipGenerator._build_name_suffix_map(emps)
        assert result["001"] == "_1"  # Alice first
        assert result["002"] == ""     # Bob unique
        assert result["003"] == "_2"  # Alice second
        assert result["004"] == ""     # Charlie unique

    def test_empty_list(self):
        result = PayslipGenerator._build_name_suffix_map([])
        assert result == {}

    def test_empty_name_no_suffix(self):
        """Employees with empty names are not counted as duplicates."""
        emps = [
            make_employee(mnv="001", name=""),
            make_employee(mnv="002", name=""),
        ]
        result = PayslipGenerator._build_name_suffix_map(emps)
        assert result["001"] == ""
        assert result["002"] == ""

    def test_whitespace_name_stripped(self):
        """Names with leading/trailing whitespace are normalized."""
        emps = [
            make_employee(mnv="001", name="  Alice  "),
            make_employee(mnv="002", name="Alice"),
        ]
        result = PayslipGenerator._build_name_suffix_map(emps)
        assert result["001"] == "_1"
        assert result["002"] == "_2"

    def test_multiple_duplicate_groups(self):
        """Multiple groups of duplicates handled independently."""
        emps = [
            make_employee(mnv="001", name="Alice"),
            make_employee(mnv="002", name="Bob"),
            make_employee(mnv="003", name="Alice"),
            make_employee(mnv="004", name="Bob"),
            make_employee(mnv="005", name="Charlie"),
        ]
        result = PayslipGenerator._build_name_suffix_map(emps)
        assert result["001"] == "_1"  # Alice 1
        assert result["002"] == "_1"  # Bob 1
        assert result["003"] == "_2"  # Alice 2
        assert result["004"] == "_2"  # Bob 2
        assert result["005"] == ""     # Charlie unique


class TestBuildOutputPathWithSuffix:
    """Tests for _build_output_path with name_suffix."""

    def test_no_suffix(self, tmp_path):
        gen = PayslipGenerator(tmp_path, "01/2026")
        emp = make_employee(name="Alice")
        path = gen._build_output_path(emp, name_suffix="")
        assert path.name == "TBKQ_Alice_012026.xlsx"

    def test_with_suffix(self, tmp_path):
        gen = PayslipGenerator(tmp_path, "01/2026")
        emp = make_employee(name="Alice")
        path = gen._build_output_path(emp, name_suffix="_1")
        assert path.name == "TBKQ_Alice_1_012026.xlsx"

    def test_suffix_2(self, tmp_path):
        gen = PayslipGenerator(tmp_path, "01/2026")
        emp = make_employee(name="Alice")
        path = gen._build_output_path(emp, name_suffix="_2")
        assert path.name == "TBKQ_Alice_2_012026.xlsx"

    def test_default_no_suffix(self, tmp_path):
        """Default name_suffix parameter is empty string."""
        gen = PayslipGenerator(tmp_path, "01/2026")
        emp = make_employee(name="Test")
        # Without passing name_suffix => default ""
        path = gen._build_output_path(emp)
        assert path.name == "TBKQ_Test_012026.xlsx"

    def test_suffix_with_special_chars_in_name(self, tmp_path):
        gen = PayslipGenerator(tmp_path, "01/2026")
        emp = make_employee(name='Bad/Name*')
        path = gen._build_output_path(emp, name_suffix="_1")
        # Special chars replaced, then suffix appended
        assert "_1_" in path.name
        assert "/" not in path.name
        assert "*" not in path.name


class TestGenerateBatchWithDuplicateNames:
    """Tests for generate_batch skipping logic with duplicate name suffixes."""

    def test_skip_with_suffix_files_exist(self, tmp_path):
        """Skip logic should use suffixed filenames when checking existence."""
        gen = PayslipGenerator(tmp_path, "01/2026")
        emps = [
            make_employee(mnv="001", name="Alice"),
            make_employee(mnv="002", name="Alice"),
        ]

        # Pre-create output files WITH suffixes
        path1 = tmp_path / "TBKQ_Alice_1_012026.xlsx"
        path2 = tmp_path / "TBKQ_Alice_2_012026.xlsx"
        path1.touch()
        path2.touch()

        results = gen.generate_batch(
            employees=emps,
            source_xls=Path("dummy.xls"),
            batch_size=50,
        )

        assert len(results) == 2
        assert all(r["skipped"] for r in results)
        assert all(r["success"] for r in results)

    def test_skip_with_mixed_unique_and_duplicate(self, tmp_path):
        """Mix of unique and duplicate names — all skipped when files exist."""
        gen = PayslipGenerator(tmp_path, "01/2026")
        emps = [
            make_employee(mnv="001", name="Alice"),
            make_employee(mnv="002", name="Bob"),
            make_employee(mnv="003", name="Alice"),
        ]

        # Create output files: Alice_1, Bob (no suffix), Alice_2
        (tmp_path / "TBKQ_Alice_1_012026.xlsx").touch()
        (tmp_path / "TBKQ_Bob_012026.xlsx").touch()
        (tmp_path / "TBKQ_Alice_2_012026.xlsx").touch()

        results = gen.generate_batch(
            employees=emps,
            source_xls=Path("dummy.xls"),
            batch_size=50,
        )

        assert len(results) == 3
        assert all(r["skipped"] for r in results)
