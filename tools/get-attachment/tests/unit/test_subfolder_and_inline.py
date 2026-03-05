"""
Unit tests for get-attachment updates:
- Subfolder support (OUTLOOK_FOLDER config field)
- Inline/embedded image filtering
- Folder navigation in AttachmentDownloader.run()
"""

import sys
from datetime import datetime
from pathlib import Path
from unittest.mock import MagicMock, patch, PropertyMock

import pytest

# Ensure both project root and tool dir are importable
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
TOOL_DIR = Path(__file__).resolve().parent.parent.parent
for _p in [str(PROJECT_ROOT), str(TOOL_DIR)]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

from attachment_downloader import (
    AttachmentDownloader,
    DownloadResult,
)
from config import GetAttachmentConfig
from office.outlook.models import (
    Attachment,
    ATTACH_BY_VALUE,
    ATTACH_OLE,
)


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


def make_attachment(
    filename: str,
    attachment_type: int = ATTACH_BY_VALUE,
    content_id: str = "",
) -> MagicMock:
    """Return a mock Attachment with inline detection support."""
    att = MagicMock(spec=Attachment)
    att.filename = filename
    att.size = 100
    att.content_type = ""
    att.attachment_type = attachment_type
    att.content_id = content_id

    # Set up is_inline property to use actual logic
    att.is_inline = bool(content_id) or (attachment_type == ATTACH_OLE)
    att._com_attachment = MagicMock()
    return att


def make_email(
    subject: str = "Test subject",
    sender: str = "sender@example.com",
    received: datetime = None,
    attachments: list = None,
) -> MagicMock:
    email = MagicMock()
    email.subject = subject
    email.sender_address = sender
    email.received_time = received or datetime(2026, 3, 2, 10, 0, 0)
    email.attachments = attachments or []
    email.has_attachments = bool(email.attachments)
    return email


# ── Config: OUTLOOK_FOLDER ───────────────────────────────────────

class TestOutlookFolderConfig:
    """Tests for the OUTLOOK_FOLDER configuration field."""

    def test_default_outlook_folder_is_empty(self):
        cfg = GetAttachmentConfig()
        assert cfg.outlook_folder == ""

    def test_outlook_folder_set_in_constructor(self):
        cfg = GetAttachmentConfig(
            outlook_account="test@co.com",
            start_date="01/03/2026",
            outlook_folder="Inbox/Subfolder1",
        )
        assert cfg.outlook_folder == "Inbox/Subfolder1"

    def test_outlook_folder_preserved_through_validation(self):
        cfg = GetAttachmentConfig(
            outlook_account="test@co.com",
            start_date="01/03/2026",
            outlook_folder="Inbox/Reports/2026",
        )
        errors = cfg.validate()
        assert errors == []
        assert cfg.outlook_folder == "Inbox/Reports/2026"

    def test_load_config_reads_outlook_folder_from_env(self, tmp_path, monkeypatch):
        """OUTLOOK_FOLDER from .env is read into config."""
        env_content = (
            "OUTLOOK_ACCOUNT=test@co.com\n"
            "START_DATE=01/03/2026\n"
            "ATTACHMENT_SAVE_PATH=D:\\Downloads\\att\n"
            "OUTLOOK_FOLDER=Inbox/Reports\n"
        )
        local_env = tmp_path / ".env"
        local_env.write_text(env_content)

        # Mock the folder prompt (user confirms current value)
        with patch("builtins.input", return_value=""):
            from config import load_config
            cfg = load_config(tool_dir=tmp_path)

        assert cfg.outlook_folder == "Inbox/Reports"


# ── Attachment.is_inline ─────────────────────────────────────────

class TestAttachmentIsInline:
    """Tests for the Attachment.is_inline property."""

    def test_regular_file_attachment_not_inline(self):
        att = Attachment(
            filename="report.xlsx",
            size=1024,
            attachment_type=ATTACH_BY_VALUE,
            content_id="",
        )
        assert att.is_inline is False

    def test_attachment_with_content_id_is_inline(self):
        att = Attachment(
            filename="logo.png",
            size=5000,
            attachment_type=ATTACH_BY_VALUE,
            content_id="image001.png@01D6A9F0.12345678",
        )
        assert att.is_inline is True

    def test_ole_attachment_is_inline(self):
        att = Attachment(
            filename="embedded.png",
            size=2000,
            attachment_type=ATTACH_OLE,
            content_id="",
        )
        assert att.is_inline is True

    def test_regular_png_without_cid_not_inline(self):
        """A regular .png attachment (not embedded) should NOT be filtered."""
        att = Attachment(
            filename="screenshot.png",
            size=50000,
            attachment_type=ATTACH_BY_VALUE,
            content_id="",
        )
        assert att.is_inline is False


# ── Downloader: inline image filtering ───────────────────────────

class TestInlineImageFiltering:
    """Tests that inline/embedded images are skipped during download."""

    START_DATE = "02/03/2026"

    def _downloader(self, tmp_path):
        cfg = make_config(
            start_date=self.START_DATE,
            attachment_save_path=tmp_path,
        )
        return AttachmentDownloader(cfg)

    def test_inline_attachment_skipped(self, tmp_path):
        dl = self._downloader(tmp_path)
        result = DownloadResult(emails_found=1)

        # One inline image + one regular attachment
        inline_att = make_attachment("logo.png", content_id="cid:logo@header")
        regular_att = make_attachment("report.xlsx")

        email = make_email(
            subject="Report",
            attachments=[inline_att, regular_att],
        )
        email.has_attachments = True

        with patch("attachment_downloader._retry_save",
                   return_value=(tmp_path / "report.xlsx", "saved")):
            dl._process_email(email, result)

        assert result.emails_matched == 1
        assert result.attachments_saved == 1  # Only regular attachment

    def test_all_inline_attachments_skipped(self, tmp_path):
        dl = self._downloader(tmp_path)
        result = DownloadResult(emails_found=1)

        inline1 = make_attachment("header.png", content_id="cid:header@img")
        inline2 = make_attachment("footer.png", content_id="cid:footer@img")

        email = make_email(
            subject="Newsletter",
            attachments=[inline1, inline2],
        )
        email.has_attachments = True

        dl._process_email(email, result)

        assert result.emails_matched == 1
        assert result.emails_with_attachments == 1
        assert result.attachments_saved == 0  # All skipped

    def test_ole_attachment_skipped(self, tmp_path):
        dl = self._downloader(tmp_path)
        result = DownloadResult(emails_found=1)

        ole_att = make_attachment("embedded.png", attachment_type=ATTACH_OLE)
        regular_att = make_attachment("data.csv")

        email = make_email(
            subject="Data",
            attachments=[ole_att, regular_att],
        )
        email.has_attachments = True

        with patch("attachment_downloader._retry_save",
                   return_value=(tmp_path / "data.csv", "saved")):
            dl._process_email(email, result)

        assert result.attachments_saved == 1


# ── Downloader.run(): subfolder navigation ───────────────────────

class TestDownloaderSubfolderNavigation:
    """Tests for folder navigation in run() method."""

    START_DATE = "02/03/2026"

    def test_run_uses_inbox_when_folder_empty(self, tmp_path):
        """When outlook_folder is empty, should use default Inbox."""
        cfg = make_config(
            outlook_folder="",
            attachment_save_path=tmp_path,
        )
        dl = AttachmentDownloader(cfg)

        mock_client = MagicMock()
        mock_client.__enter__ = MagicMock(return_value=mock_client)
        mock_client.__exit__ = MagicMock(return_value=False)
        mock_client.get_inbox.return_value = MagicMock()
        mock_client.get_emails_from_folder.return_value = []

        with patch("attachment_downloader.OutlookReader", return_value=mock_client):
            result = dl.run()

        mock_client.get_inbox.assert_called_once()
        mock_client.get_folder_by_path.assert_not_called()

    def test_run_uses_inbox_when_folder_is_inbox(self, tmp_path):
        """When outlook_folder is 'Inbox', should use default Inbox."""
        cfg = make_config(
            outlook_folder="Inbox",
            attachment_save_path=tmp_path,
        )
        dl = AttachmentDownloader(cfg)

        mock_client = MagicMock()
        mock_client.__enter__ = MagicMock(return_value=mock_client)
        mock_client.__exit__ = MagicMock(return_value=False)
        mock_client.get_inbox.return_value = MagicMock()
        mock_client.get_emails_from_folder.return_value = []

        with patch("attachment_downloader.OutlookReader", return_value=mock_client):
            result = dl.run()

        mock_client.get_inbox.assert_called_once()

    def test_run_navigates_to_subfolder(self, tmp_path):
        """When outlook_folder is set, should navigate to that folder."""
        cfg = make_config(
            outlook_folder="Inbox/Reports/2026",
            attachment_save_path=tmp_path,
        )
        dl = AttachmentDownloader(cfg)

        mock_folder = MagicMock()
        mock_client = MagicMock()
        mock_client.__enter__ = MagicMock(return_value=mock_client)
        mock_client.__exit__ = MagicMock(return_value=False)
        mock_client.get_folder_by_path.return_value = mock_folder
        mock_client.get_emails_from_folder.return_value = []

        with patch("attachment_downloader.OutlookReader", return_value=mock_client):
            result = dl.run()

        mock_client.get_folder_by_path.assert_called_once_with("Inbox/Reports/2026")
        mock_client.get_emails_from_folder.assert_called_once()
        # Verify the folder passed to get_emails_from_folder is the resolved one
        args, kwargs = mock_client.get_emails_from_folder.call_args
        assert args[0] == mock_folder

    def test_run_handles_folder_not_found(self, tmp_path):
        """When folder doesn't exist, should return empty result."""
        from office.outlook.exceptions import OutlookFolderNotFoundError

        cfg = make_config(
            outlook_folder="Inbox/NonExistent",
            attachment_save_path=tmp_path,
        )
        dl = AttachmentDownloader(cfg)

        mock_client = MagicMock()
        mock_client.__enter__ = MagicMock(return_value=mock_client)
        mock_client.__exit__ = MagicMock(return_value=False)
        mock_client.get_folder_by_path.side_effect = OutlookFolderNotFoundError(
            "Inbox/NonExistent", "test@co.com"
        )

        with patch("attachment_downloader.OutlookReader", return_value=mock_client):
            result = dl.run()

        assert result.emails_found == 0
        assert result.attachments_saved == 0


class TestPromptClearOutputDir:
    """Tests for _prompt_clear_output_dir in main.py."""

    def test_no_action_when_dir_missing(self, tmp_path):
        from main import _prompt_clear_output_dir

        missing = tmp_path / "nonexistent"
        # Should not raise, no prompt needed
        _prompt_clear_output_dir(missing)  # no exception == pass

    def test_no_action_when_dir_empty(self, tmp_path):
        from main import _prompt_clear_output_dir

        empty_dir = tmp_path / "empty"
        empty_dir.mkdir()
        # No files → no prompt, no error
        _prompt_clear_output_dir(empty_dir)  # no exception == pass

    @patch("builtins.input", return_value="yes")
    def test_yes_deletes_all_files(self, mock_input, tmp_path):
        from main import _prompt_clear_output_dir

        (tmp_path / "file1.pdf").touch()
        (tmp_path / "file2.xlsx").touch()
        (tmp_path / "file3.txt").touch()

        _prompt_clear_output_dir(tmp_path)

        remaining = list(tmp_path.iterdir())
        assert remaining == []

    @patch("builtins.input", return_value="y")
    def test_y_shorthand_deletes_all_files(self, mock_input, tmp_path):
        from main import _prompt_clear_output_dir

        (tmp_path / "a.pdf").touch()
        (tmp_path / "b.pdf").touch()

        _prompt_clear_output_dir(tmp_path)

        assert list(tmp_path.iterdir()) == []

    @patch("builtins.input", return_value="no")
    def test_no_keeps_all_files(self, mock_input, tmp_path):
        from main import _prompt_clear_output_dir

        (tmp_path / "keep1.pdf").touch()
        (tmp_path / "keep2.xlsx").touch()

        _prompt_clear_output_dir(tmp_path)

        assert len(list(tmp_path.iterdir())) == 2

    @patch("builtins.input", return_value="n")
    def test_n_shorthand_keeps_all_files(self, mock_input, tmp_path):
        from main import _prompt_clear_output_dir

        (tmp_path / "keep.pdf").touch()

        _prompt_clear_output_dir(tmp_path)

        assert len(list(tmp_path.iterdir())) == 1

    @patch("builtins.input", side_effect=["maybe", "yes"])
    def test_invalid_input_retries_then_deletes(self, mock_input, tmp_path):
        from main import _prompt_clear_output_dir

        (tmp_path / "file.pdf").touch()

        _prompt_clear_output_dir(tmp_path)

        assert list(tmp_path.iterdir()) == []

    @patch("builtins.input", return_value="yes")
    def test_subdirectories_not_deleted(self, mock_input, tmp_path):
        """Only files are deleted; subdirectories are preserved."""
        from main import _prompt_clear_output_dir

        (tmp_path / "file.pdf").touch()
        subdir = tmp_path / "subdir"
        subdir.mkdir()
        (subdir / "nested.pdf").touch()

        _prompt_clear_output_dir(tmp_path)

        # Top-level file gone, subdir intact
        assert not (tmp_path / "file.pdf").exists()
        assert subdir.exists()
