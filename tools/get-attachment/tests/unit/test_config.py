"""
Unit tests for config.py.

Covers:
- GetAttachmentConfig defaults and post-init normalisation
- GetAttachmentConfig.validate() — valid and invalid states for both dates
- GetAttachmentConfig.start_date_parsed / end_date_parsed — parsing and defaults
- GetAttachmentConfig.date_range_display — formatted display string
- _validate_date_ddmmyyyy — valid/invalid date strings
- load_config — .env reading and interactive prompt mocking
"""

import os
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

# Ensure both project root and tool dir are importable
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
TOOL_DIR = Path(__file__).resolve().parent.parent.parent
for _p in [str(PROJECT_ROOT), str(TOOL_DIR)]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

from config import (
    GetAttachmentConfig,
    _validate_date_ddmmyyyy,
    load_config,
)


# ── _validate_date_ddmmyyyy ──────────────────────────────────────

class TestValidateDateDdMmYyyy:
    """Tests for the DD/MM/YYYY date validator."""

    def test_valid_date(self):
        ok, msg = _validate_date_ddmmyyyy("02/03/2026")
        assert ok is True
        assert msg == ""

    def test_valid_date_start_of_year(self):
        ok, _ = _validate_date_ddmmyyyy("01/01/2000")
        assert ok is True

    def test_valid_date_last_day_of_month(self):
        ok, _ = _validate_date_ddmmyyyy("31/01/2026")
        assert ok is True

    def test_invalid_format_no_separators(self):
        ok, msg = _validate_date_ddmmyyyy("02032026")
        assert ok is False
        assert "DD/MM/YYYY" in msg

    def test_invalid_format_wrong_order(self):
        # MM/DD/YYYY format
        ok, msg = _validate_date_ddmmyyyy("03/02/2026")
        # This still passes DD/MM/YYYY format check; 03 Feb 2026 is valid
        assert ok is True  # 03/02/2026 is valid as 3 Feb 2026

    def test_invalid_day(self):
        # 31 February does not exist
        ok, msg = _validate_date_ddmmyyyy("31/02/2026")
        assert ok is False
        assert "Invalid date" in msg or "day is out of range" in msg.lower()

    def test_invalid_month(self):
        ok, msg = _validate_date_ddmmyyyy("01/13/2026")
        assert ok is False

    def test_whitespace_stripped(self):
        ok, _ = _validate_date_ddmmyyyy("  02/03/2026  ")
        assert ok is True


# ── GetAttachmentConfig defaults ─────────────────────────────────

class TestGetAttachmentConfigDefaults:
    """Verify default field values."""

    def test_empty_strings(self):
        cfg = GetAttachmentConfig()
        assert cfg.outlook_account == ""
        assert cfg.start_date == ""
        assert cfg.end_date == ""

    def test_empty_keywords(self):
        cfg = GetAttachmentConfig()
        assert cfg.subject_keywords == []

    def test_default_log_level_uppercased(self):
        cfg = GetAttachmentConfig(log_level="info")
        assert cfg.log_level == "INFO"

    def test_paths_are_path_objects(self):
        cfg = GetAttachmentConfig()
        assert isinstance(cfg.attachment_save_path, Path)
        assert isinstance(cfg.log_dir, Path)

    def test_string_paths_normalised(self):
        cfg = GetAttachmentConfig(
            attachment_save_path="D:\\Downloads\\attachments",
            log_dir="./logs",
        )
        assert isinstance(cfg.attachment_save_path, Path)
        assert isinstance(cfg.log_dir, Path)


# ── GetAttachmentConfig.validate() ──────────────────────────────

class TestGetAttachmentConfigValidation:
    """Tests for GetAttachmentConfig.validate()."""

    def test_valid_config(self):
        cfg = GetAttachmentConfig(
            outlook_account="user@company.com",
            start_date="02/03/2026",
        )
        assert cfg.validate() == []

    def test_missing_account(self):
        cfg = GetAttachmentConfig(start_date="02/03/2026")
        errors = cfg.validate()
        assert any("OUTLOOK_ACCOUNT" in e for e in errors)

    def test_missing_date(self):
        cfg = GetAttachmentConfig(outlook_account="user@company.com")
        errors = cfg.validate()
        assert any("START_DATE" in e for e in errors)

    def test_invalid_date_format(self):
        cfg = GetAttachmentConfig(
            outlook_account="user@company.com",
            start_date="2026-03-02",  # ISO format — wrong
        )
        errors = cfg.validate()
        assert any("START_DATE" in e for e in errors)

    def test_multiple_errors(self):
        cfg = GetAttachmentConfig()
        errors = cfg.validate()
        assert len(errors) >= 2  # missing account AND missing date

    def test_valid_end_date(self):
        cfg = GetAttachmentConfig(
            outlook_account="user@company.com",
            start_date="01/03/2026",
            end_date="05/03/2026",
        )
        assert cfg.validate() == []

    def test_invalid_end_date_format(self):
        cfg = GetAttachmentConfig(
            outlook_account="user@company.com",
            start_date="01/03/2026",
            end_date="2026-03-05",  # ISO format — wrong
        )
        errors = cfg.validate()
        assert any("END_DATE" in e for e in errors)

    def test_end_date_before_start_date(self):
        cfg = GetAttachmentConfig(
            outlook_account="user@company.com",
            start_date="05/03/2026",
            end_date="01/03/2026",  # earlier than start
        )
        errors = cfg.validate()
        assert any("must not be before" in e for e in errors)

    def test_end_date_same_as_start_date_is_valid(self):
        cfg = GetAttachmentConfig(
            outlook_account="user@company.com",
            start_date="02/03/2026",
            end_date="02/03/2026",
        )
        assert cfg.validate() == []

    def test_empty_end_date_is_valid(self):
        """Empty end_date means today; should not produce a validation error."""
        cfg = GetAttachmentConfig(
            outlook_account="user@company.com",
            start_date="01/03/2026",
            end_date="",
        )
        assert cfg.validate() == []


# ── GetAttachmentConfig.target_date_parsed ───────────────────────

class TestStartDateParsed:
    """Tests for the start_date_parsed property."""

    def test_valid_date_returns_date_object(self):
        from datetime import date
        cfg = GetAttachmentConfig(
            outlook_account="a@b.com",
            start_date="02/03/2026",
        )
        result = cfg.start_date_parsed
        assert result == date(2026, 3, 2)

    def test_empty_date_returns_none(self):
        cfg = GetAttachmentConfig()
        assert cfg.start_date_parsed is None

    def test_invalid_date_returns_none(self):
        cfg = GetAttachmentConfig(start_date="not-a-date")
        assert cfg.start_date_parsed is None


# ── GetAttachmentConfig.end_date_parsed ──────────────────────────

class TestEndDateParsed:
    """Tests for the end_date_parsed property."""

    def test_empty_end_date_returns_today(self):
        from datetime import date
        cfg = GetAttachmentConfig(end_date="")
        assert cfg.end_date_parsed == date.today()

    def test_valid_end_date_returns_date_object(self):
        from datetime import date
        cfg = GetAttachmentConfig(end_date="05/03/2026")
        assert cfg.end_date_parsed == date(2026, 3, 5)

    def test_invalid_end_date_falls_back_to_today(self):
        from datetime import date
        cfg = GetAttachmentConfig(end_date="not-a-date")
        assert cfg.end_date_parsed == date.today()


# ── GetAttachmentConfig.date_range_display ───────────────────────

class TestDateRangeDisplay:
    """Tests for the date_range_display property."""

    def test_with_explicit_end_date(self):
        cfg = GetAttachmentConfig(start_date="01/03/2026", end_date="05/03/2026")
        assert cfg.date_range_display == "01/03/2026 \u2192 05/03/2026"

    def test_empty_end_date_shows_today(self):
        from datetime import date
        cfg = GetAttachmentConfig(start_date="01/03/2026", end_date="")
        today = date.today().strftime("%d/%m/%Y")
        assert cfg.date_range_display == f"01/03/2026 \u2192 {today}"


# ── load_config ──────────────────────────────────────────────────

class TestLoadConfig:
    """Tests for load_config() with mocked env and prompts."""

    def test_all_values_from_env(self, tmp_path, monkeypatch):
        """When all values are in the env, no prompts should be called."""
        env_content = (
            "OUTLOOK_ACCOUNT=test@co.com\n"
            "START_DATE=01/03/2026\n"
            "END_DATE=05/03/2026\n"
            "SUBJECT_KEYWORDS=invoice,report\n"
            "ATTACHMENT_SAVE_PATH=D:\\Downloads\\att\n"
        )
        local_env = tmp_path / ".env"
        local_env.write_text(env_content)

        # Patch ConfigManager.prompt_for_value to fail if called unexpectedly
        with patch("config.ConfigManager.prompt_for_value", side_effect=AssertionError("should not prompt")) as _mock_prompt, \
             patch("config._prompt_for_outlook_account", side_effect=AssertionError("should not prompt")) as _mock_acc, \
             patch("builtins.input", side_effect=AssertionError("should not prompt")):
            cfg = load_config(tool_dir=tmp_path)

        assert cfg.outlook_account == "test@co.com"
        assert cfg.start_date == "01/03/2026"
        assert cfg.end_date == "05/03/2026"
        assert cfg.subject_keywords == ["invoice", "report"]
        assert cfg.attachment_save_path == Path("D:\\Downloads\\att")

    def test_missing_account_prompts(self, tmp_path, monkeypatch):
        """Missing OUTLOOK_ACCOUNT triggers the account-selection prompt."""
        # Clear any leftover env vars from previous tests in this process
        monkeypatch.delenv("OUTLOOK_ACCOUNT", raising=False)
        monkeypatch.delenv("SUBJECT_KEYWORDS", raising=False)
        monkeypatch.setenv("START_DATE", "01/03/2026")
        monkeypatch.setenv("ATTACHMENT_SAVE_PATH", r"D:\Downloads\att")

        # Write an empty .env so load_env finds the tool dir's env file
        (tmp_path / ".env").write_text("")

        with patch("config._prompt_for_outlook_account", return_value="prompted@co.com"), \
             patch("builtins.input", return_value=""):  # empty keywords
            cfg = load_config(tool_dir=tmp_path)

        assert cfg.outlook_account == "prompted@co.com"

    def test_missing_date_prompts(self, tmp_path, monkeypatch):
        """Missing START_DATE triggers the date prompt."""
        monkeypatch.delenv("START_DATE", raising=False)
        monkeypatch.delenv("SUBJECT_KEYWORDS", raising=False)
        monkeypatch.setenv("OUTLOOK_ACCOUNT", "test@co.com")
        monkeypatch.setenv("ATTACHMENT_SAVE_PATH", r"D:\Downloads\att")

        (tmp_path / ".env").write_text("")

        with patch("config.ConfigManager.prompt_for_value", return_value="05/03/2026"), \
             patch("builtins.input", return_value=""):  # empty keywords
            cfg = load_config(tool_dir=tmp_path)

        assert cfg.start_date == "05/03/2026"

    def test_empty_keywords_accepted(self, tmp_path):
        """An empty SUBJECT_KEYWORDS value results in an empty list."""
        env_content = (
            "OUTLOOK_ACCOUNT=test@co.com\n"
            "START_DATE=01/03/2026\n"
            "SUBJECT_KEYWORDS=\n"
            "ATTACHMENT_SAVE_PATH=D:\\Downloads\\att\n"
        )
        local_env = tmp_path / ".env"
        local_env.write_text(env_content)

        with patch("builtins.input", return_value=""):
            cfg = load_config(tool_dir=tmp_path)

        assert cfg.subject_keywords == []
