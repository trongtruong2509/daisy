"""
Unit tests for office/utils/com.py — centralised COM bootstrapping.

Covers:
- is_available() reflects actual environment
- ensure_com_available() raises ImportError when pywin32 missing
- com_initialized() context manager calls CoInitialize/CoUninitialize
- com_initialized() handles nested usage safely
- get_pythoncom/get_win32com_client/get_pywintypes return modules
- REQ-COM-06 compliance: no forbidden imports outside office/utils/com.py

Requirements tested: REQ-COM-01, REQ-COM-06.
"""

import importlib
import re
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

# Ensure project root is importable
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


class TestIsAvailable:
    """Tests for is_available()."""

    def test_returns_bool(self):
        from office.utils.com import is_available
        result = is_available()
        assert isinstance(result, bool)

    def test_true_when_pywin32_installed(self):
        """On a Windows dev machine with pywin32, should return True."""
        from office.utils.com import is_available
        # This test will pass on dev machines and skip on CI without pywin32
        if not is_available():
            pytest.skip("pywin32 not installed — expected in CI")
        assert is_available() is True


class TestEnsureComAvailable:
    """Tests for ensure_com_available()."""

    def test_raises_when_not_available(self):
        """Should raise ImportError with actionable message when HAS_COM is False."""
        import office.utils.com as com_mod
        original = com_mod.HAS_COM
        try:
            com_mod.HAS_COM = False
            with pytest.raises(ImportError, match="pywin32"):
                com_mod.ensure_com_available()
        finally:
            com_mod.HAS_COM = original

    def test_no_raise_when_available(self):
        """Should not raise when pywin32 is present."""
        from office.utils.com import is_available, ensure_com_available
        if not is_available():
            pytest.skip("pywin32 not installed")
        # Should not raise
        ensure_com_available()


class TestComInitialized:
    """Tests for the com_initialized() context manager."""

    def test_raises_import_error_when_unavailable(self):
        """Should raise ImportError if pywin32 missing."""
        import office.utils.com as com_mod
        original = com_mod.HAS_COM
        try:
            com_mod.HAS_COM = False
            with pytest.raises(ImportError):
                with com_mod.com_initialized():
                    pass
        finally:
            com_mod.HAS_COM = original

    def test_calls_coinitialize_and_couninitialize(self):
        """Should call CoInitialize on entry and CoUninitialize on exit."""
        from office.utils.com import is_available
        if not is_available():
            pytest.skip("pywin32 not installed")

        import office.utils.com as com_mod
        mock_pythoncom = MagicMock()
        original = com_mod._pythoncom
        try:
            com_mod._pythoncom = mock_pythoncom
            with com_mod.com_initialized():
                mock_pythoncom.CoInitialize.assert_called_once()
                mock_pythoncom.CoUninitialize.assert_not_called()
            mock_pythoncom.CoUninitialize.assert_called_once()
        finally:
            com_mod._pythoncom = original

    def test_couninitialize_called_even_on_exception(self):
        """CoUninitialize must be called even if the body raises."""
        from office.utils.com import is_available
        if not is_available():
            pytest.skip("pywin32 not installed")

        import office.utils.com as com_mod
        mock_pythoncom = MagicMock()
        original = com_mod._pythoncom
        try:
            com_mod._pythoncom = mock_pythoncom
            with pytest.raises(ValueError):
                with com_mod.com_initialized():
                    raise ValueError("test error")
            mock_pythoncom.CoUninitialize.assert_called_once()
        finally:
            com_mod._pythoncom = original

    def test_nested_contexts_safe(self):
        """Nested com_initialized() should not crash (COM is ref-counted)."""
        from office.utils.com import is_available
        if not is_available():
            pytest.skip("pywin32 not installed")

        import office.utils.com as com_mod
        mock_pythoncom = MagicMock()
        original = com_mod._pythoncom
        try:
            com_mod._pythoncom = mock_pythoncom
            with com_mod.com_initialized():
                with com_mod.com_initialized():
                    pass
            assert mock_pythoncom.CoInitialize.call_count == 2
            assert mock_pythoncom.CoUninitialize.call_count == 2
        finally:
            com_mod._pythoncom = original


class TestInternalAccessors:
    """Tests for get_pythoncom/get_win32com_client/get_pywintypes."""

    def test_get_pythoncom_returns_module(self):
        from office.utils.com import is_available, get_pythoncom
        if not is_available():
            pytest.skip("pywin32 not installed")
        mod = get_pythoncom()
        assert hasattr(mod, "CoInitialize")

    def test_get_win32com_client_returns_module(self):
        from office.utils.com import is_available, get_win32com_client
        if not is_available():
            pytest.skip("pywin32 not installed")
        mod = get_win32com_client()
        assert hasattr(mod, "Dispatch")

    def test_get_pywintypes_returns_module(self):
        from office.utils.com import is_available, get_pywintypes
        if not is_available():
            pytest.skip("pywin32 not installed")
        mod = get_pywintypes()
        assert hasattr(mod, "com_error")


class TestReqCom06Compliance:
    """
    REQ-COM-06: HAS_PYTHONCOM / HAS_WIN32COM guards forbidden outside office/utils/com.py.
    
    Scans source files to ensure no forbidden patterns exist.
    """

    # Directories to check (relative to project root)
    CHECK_DIRS = ["office/excel", "office/outlook", "tools"]
    # Forbidden patterns
    FORBIDDEN_PATTERNS = [
        r"import\s+pythoncom",
        r"import\s+win32com",
        r"import\s+pywintypes",
        r"from\s+win32com",
        r"HAS_PYTHONCOM",
        r"HAS_WIN32COM",
    ]

    def _get_python_files(self):
        """Yield all .py files in CHECK_DIRS."""
        root = Path(__file__).resolve().parent.parent.parent.parent
        for dir_name in self.CHECK_DIRS:
            dir_path = root / dir_name
            if dir_path.exists():
                yield from dir_path.rglob("*.py")

    def test_no_forbidden_imports_in_office_and_tools(self):
        """No file outside office/com.py should import pythoncom/win32com directly."""
        violations = []
        for py_file in self._get_python_files():
            # Skip __pycache__
            if "__pycache__" in str(py_file):
                continue
            # Skip test files — tests may need to mock things
            if "tests" in str(py_file).replace("\\", "/"):
                continue

            content = py_file.read_text(encoding="utf-8", errors="ignore")
            for pattern in self.FORBIDDEN_PATTERNS:
                matches = re.findall(pattern, content)
                if matches:
                    rel = py_file.relative_to(
                        Path(__file__).resolve().parent.parent.parent.parent
                    )
                    violations.append(f"{rel}: {pattern} ({len(matches)} match(es))")

        assert violations == [], (
            "Forbidden COM imports found outside office/com.py:\n"
            + "\n".join(f"  - {v}" for v in violations)
        )
