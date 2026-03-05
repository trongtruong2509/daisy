"""
Unit tests for attachment_downloader.py.

Covers:
- _sanitise_for_filename — unsafe character replacement
- _candidate_filenames   — deduplication name generation
- save_attachment_with_dedup — file saving and dedup logic (COM mocked)
- AttachmentDownloader._matches_keywords — OR keyword logic
- AttachmentDownloader._process_email   — email processing flow
- AttachmentDownloader.run              — full run with mocked Outlook
"""

import sys
from datetime import datetime, date
from pathlib import Path
from unittest.mock import MagicMock, patch, call

import pytest

# Ensure both project root and tool dir are importable
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
TOOL_DIR = Path(__file__).resolve().parent.parent.parent
for _p in [str(PROJECT_ROOT), str(TOOL_DIR)]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

from attachment_downloader import (
    _sanitise_for_filename,
    _candidate_filenames,
    save_attachment_with_dedup,
    AttachmentDownloader,
    DownloadResult,
)
from config import GetAttachmentConfig


# ── Fixtures ─────────────────────────────────────────────────────

def make_config(**kwargs) -> GetAttachmentConfig:
    defaults = dict(
        outlook_account="test@co.com",
        start_date="02/03/2026",
        subject_keywords=[],
        attachment_save_path=Path("/tmp/attachments"),
    )
    defaults.update(kwargs)
    return GetAttachmentConfig(**defaults)


def make_attachment(filename: str, save_side_effect=None) -> MagicMock:
    """Return a mock Attachment whose _com_attachment.SaveAsFile can be controlled."""
    att = MagicMock()
    att.filename = filename
    att.size = 100
    att.is_inline = False  # Regular file attachment by default
    if save_side_effect:
        att._com_attachment.SaveAsFile.side_effect = save_side_effect
    return att


def make_email(
    subject: str = "Test subject",
    sender: str = "sender@example.com",
    received: datetime = None,
    attachments: list = None,
) -> MagicMock:
    """Return a mock Email object."""
    email = MagicMock()
    email.subject = subject
    email.sender_address = sender
    email.received_time = received or datetime(2026, 3, 2, 10, 0, 0)
    email.attachments = attachments or []
    email.has_attachments = bool(email.attachments)
    return email


# ── _sanitise_for_filename ───────────────────────────────────────

class TestSanitiseForFilename:
    """Tests for the filename sanitisation helper."""

    def test_email_at_and_dots_replaced(self):
        result = _sanitise_for_filename("user@company.com")
        assert "@" not in result
        assert "." not in result

    def test_spaces_replaced(self):
        result = _sanitise_for_filename("John Doe")
        assert " " not in result

    def test_windows_unsafe_chars_replaced(self):
        result = _sanitise_for_filename(r'a\b/c:d*e?f"g<h>i|j')
        for ch in r'\/:*?"<>|':
            assert ch not in result

    def test_leading_trailing_underscores_stripped(self):
        result = _sanitise_for_filename("@user@")
        assert not result.startswith("_")
        assert not result.endswith("_")

    def test_consecutive_underscores_collapsed(self):
        result = _sanitise_for_filename("a@@b")
        assert "__" not in result

    def test_plain_name_unchanged(self):
        result = _sanitise_for_filename("invoice")
        assert result == "invoice"


# ── _candidate_filenames ─────────────────────────────────────────

class TestCandidateFilenames:
    """Tests for deduplication candidate generation."""

    def test_returns_two_candidates(self):
        candidates = _candidate_filenames("report.xlsx", "sender@co.com", None)
        assert len(candidates) == 2

    def test_first_candidate_has_sender_suffix(self):
        candidates = _candidate_filenames("report.xlsx", "a@b.com", None)
        assert candidates[0].endswith(".xlsx")
        assert "_a_b_com" in candidates[0]

    def test_second_candidate_has_timestamp_when_provided(self):
        dt = datetime(2026, 3, 2, 10, 5, 30)
        candidates = _candidate_filenames("doc.pdf", "a@b.com", dt)
        assert "20260302_100530" in candidates[1]

    def test_second_candidate_unknown_time_when_none(self):
        candidates = _candidate_filenames("doc.pdf", "a@b.com", None)
        assert "unknown_time" in candidates[1]

    def test_extension_preserved_in_all_candidates(self):
        candidates = _candidate_filenames("data.csv", "x@y.com", None)
        for c in candidates:
            assert c.endswith(".csv")

    def test_no_extension_file(self):
        candidates = _candidate_filenames("makefile", "x@y.com", None)
        # stem is 'makefile', suffix is ''
        for c in candidates:
            assert "makefile" in c


# ── save_attachment_with_dedup ───────────────────────────────────

class TestSaveAttachmentWithDedup:
    """Tests for save_attachment_with_dedup (file system + COM mocked)."""

    RECEIVED = datetime(2026, 3, 2, 10, 0, 0)
    SENDER = "sender@example.com"

    def test_saves_original_when_no_conflict(self, tmp_path):
        att = make_attachment("report.xlsx")
        path, status = save_attachment_with_dedup(att, tmp_path, self.SENDER, self.RECEIVED)
        assert status == "saved"
        assert path == tmp_path / "report.xlsx"
        att._com_attachment.SaveAsFile.assert_called_once_with(str(tmp_path / "report.xlsx"))

    def test_renames_with_sender_on_conflict(self, tmp_path):
        # Pre-create the original file
        (tmp_path / "report.xlsx").touch()
        att = make_attachment("report.xlsx")
        path, status = save_attachment_with_dedup(att, tmp_path, self.SENDER, self.RECEIVED)
        assert status == "renamed_sender"
        assert "sender_example_com" in path.name

    def test_renames_with_timestamp_on_double_conflict(self, tmp_path):
        # Pre-create original + sender variant
        (tmp_path / "report.xlsx").touch()
        safe = _sanitise_for_filename(self.SENDER)
        (tmp_path / f"report_{safe}.xlsx").touch()
        att = make_attachment("report.xlsx")
        path, status = save_attachment_with_dedup(att, tmp_path, self.SENDER, self.RECEIVED)
        assert status == "renamed_timestamp"
        assert "20260302" in path.name

    def test_renames_with_counter_on_triple_conflict(self, tmp_path):
        (tmp_path / "report.xlsx").touch()
        safe = _sanitise_for_filename(self.SENDER)
        ts = self.RECEIVED.strftime("%Y%m%d_%H%M%S")
        (tmp_path / f"report_{safe}.xlsx").touch()
        (tmp_path / f"report_{safe}_{ts}.xlsx").touch()
        att = make_attachment("report.xlsx")
        path, status = save_attachment_with_dedup(att, tmp_path, self.SENDER, self.RECEIVED)
        assert status == "renamed_counter"
        assert path.name.endswith(".xlsx")

    def test_raises_if_com_reference_missing(self, tmp_path):
        att = MagicMock()
        att.filename = "x.xlsx"
        att._com_attachment = None
        with pytest.raises(ValueError, match="COM reference"):
            save_attachment_with_dedup(att, tmp_path, self.SENDER, self.RECEIVED)

    def test_creates_save_dir_if_missing(self, tmp_path):
        nested = tmp_path / "a" / "b" / "c"
        att = make_attachment("f.txt")
        save_attachment_with_dedup(att, nested, self.SENDER, self.RECEIVED)
        assert nested.is_dir()


# ── AttachmentDownloader._matches_keywords ───────────────────────

class TestMatchesKeywords:
    """Tests for the keyword filter logic in AttachmentDownloader."""

    def _downloader(self, keywords):
        cfg = make_config(subject_keywords=keywords)
        return AttachmentDownloader(cfg)

    def test_no_keywords_always_matches(self):
        dl = self._downloader([])
        email = make_email(subject="Random subject")
        assert dl._matches_keywords(email) is True

    def test_single_keyword_match(self):
        dl = self._downloader(["invoice"])
        email = make_email(subject="January Invoice 2026")
        assert dl._matches_keywords(email) is True

    def test_single_keyword_no_match(self):
        dl = self._downloader(["invoice"])
        email = make_email(subject="Meeting notes")
        assert dl._matches_keywords(email) is False

    def test_multiple_keywords_or_logic(self):
        dl = self._downloader(["invoice", "report"])
        email_invoice = make_email(subject="Monthly invoice")
        email_report = make_email(subject="Sales report Q1")
        email_other = make_email(subject="Team lunch")
        assert dl._matches_keywords(email_invoice) is True
        assert dl._matches_keywords(email_report) is True
        assert dl._matches_keywords(email_other) is False

    def test_keyword_matching_is_case_insensitive(self):
        dl = self._downloader(["INVOICE"])
        email = make_email(subject="monthly invoice")
        assert dl._matches_keywords(email) is True

    def test_partial_keyword_match(self):
        dl = self._downloader(["inv"])
        email = make_email(subject="invoice")
        assert dl._matches_keywords(email) is True


# ── AttachmentDownloader._process_email ─────────────────────────

class TestProcessEmail:
    """Tests for the per-email processing method."""

    START_DATE = "02/03/2026"

    def _downloader(self, keywords=None, save_dir=None, tmp_path=None):
        if save_dir is None:
            save_dir = tmp_path or Path("/tmp/test_attachments")
        cfg = make_config(
            subject_keywords=keywords or [],
            start_date=self.START_DATE,
            attachment_save_path=save_dir,
        )
        return AttachmentDownloader(cfg)

    def test_email_skipped_when_keywords_dont_match(self, tmp_path):
        dl = self._downloader(keywords=["invoice"], tmp_path=tmp_path)
        result = DownloadResult(emails_found=1)
        email = make_email(subject="Meeting notes")
        dl._process_email(email, result)
        assert result.emails_matched == 0
        assert result.emails_with_attachments == 0

    def test_email_with_no_attachments_counted_as_matched(self, tmp_path):
        dl = self._downloader(tmp_path=tmp_path)
        result = DownloadResult(emails_found=1)
        email = make_email(subject="No attachments here", attachments=[])
        email.has_attachments = False
        dl._process_email(email, result)
        assert result.emails_matched == 1
        assert result.emails_with_attachments == 0
        assert result.attachments_saved == 0

    def test_attachment_saved_on_match(self, tmp_path):
        dl = self._downloader(tmp_path=tmp_path)
        result = DownloadResult(emails_found=1)
        att = make_attachment("data.csv")
        email = make_email(subject="Report", attachments=[att])
        email.has_attachments = True

        with patch("attachment_downloader._retry_save",
                   return_value=(tmp_path / "data.csv", "saved")):
            dl._process_email(email, result)

        assert result.emails_with_attachments == 1
        assert result.attachments_saved == 1
        assert result.attachments_failed == 0

    def test_failed_attachment_recorded_in_result(self, tmp_path):
        dl = self._downloader(tmp_path=tmp_path)
        result = DownloadResult(emails_found=1)
        att = make_attachment("bad.pdf")
        email = make_email(subject="Error email", attachments=[att])
        email.has_attachments = True

        with patch("attachment_downloader._retry_save", side_effect=OSError("disk full")):
            dl._process_email(email, result)

        assert result.attachments_failed == 1
        assert result.attachments_saved == 0
        assert len(result.errors) == 1
        assert "bad.pdf" in result.errors[0]


# ── AttachmentDownloader.run ─────────────────────────────────────

class TestAttachmentDownloaderRun:
    """Integration-level unit tests for the full run() method (Outlook mocked)."""

    START_DATE = "02/03/2026"

    def _downloader(self, keywords=None, tmp_path=None):
        save_dir = tmp_path or Path("/tmp/run_test")
        cfg = make_config(
            subject_keywords=keywords or [],
            start_date=self.START_DATE,
            attachment_save_path=save_dir,
        )
        return AttachmentDownloader(cfg)

    def test_no_emails_returns_empty_result(self, tmp_path):
        dl = self._downloader(tmp_path=tmp_path)

        mock_client = MagicMock()
        mock_client.__enter__ = MagicMock(return_value=mock_client)
        mock_client.__exit__ = MagicMock(return_value=False)
        mock_client.get_inbox.return_value = MagicMock()
        mock_client.get_emails_from_folder.return_value = []

        with patch("attachment_downloader.OutlookReader", return_value=mock_client):
            result = dl.run()

        assert result.emails_found == 0
        assert result.attachments_saved == 0

    def test_emails_with_attachments_are_saved(self, tmp_path):
        dl = self._downloader(tmp_path=tmp_path)

        att = make_attachment("report.xlsx")
        email = make_email(subject="Monthly report", attachments=[att])
        email.has_attachments = True

        mock_client = MagicMock()
        mock_client.__enter__ = MagicMock(return_value=mock_client)
        mock_client.__exit__ = MagicMock(return_value=False)
        mock_client.get_inbox.return_value = MagicMock()
        mock_client.get_emails_from_folder.return_value = [email]

        saved_path = tmp_path / "report.xlsx"

        with patch("attachment_downloader.OutlookReader", return_value=mock_client), \
             patch("attachment_downloader._retry_save", return_value=(saved_path, "saved")):
            result = dl.run()

        assert result.emails_found == 1
        assert result.emails_matched == 1
        assert result.emails_with_attachments == 1
        assert result.attachments_saved == 1
        assert saved_path in result.saved_files

    def test_keyword_filtered_emails_excluded(self, tmp_path):
        dl = self._downloader(keywords=["invoice"], tmp_path=tmp_path)

        att = make_attachment("notes.docx")
        email = make_email(subject="Meeting notes", attachments=[att])
        email.has_attachments = True

        mock_client = MagicMock()
        mock_client.__enter__ = MagicMock(return_value=mock_client)
        mock_client.__exit__ = MagicMock(return_value=False)
        mock_client.get_inbox.return_value = MagicMock()
        mock_client.get_emails_from_folder.return_value = [email]

        with patch("attachment_downloader.OutlookReader", return_value=mock_client):
            result = dl.run()

        assert result.emails_found == 1
        assert result.emails_matched == 0
        assert result.attachments_saved == 0

    def test_invalid_date_raises_value_error(self, tmp_path):
        cfg = make_config(start_date="not-a-date", attachment_save_path=tmp_path)
        dl = AttachmentDownloader(cfg)
        with pytest.raises(ValueError, match="Cannot parse start_date"):
            dl.run()
