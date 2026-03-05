"""
Core attachment download logic for the Get Attachment tool.

Searches the Outlook Inbox for emails matching a specific date and (optional)
subject keyword criteria, then saves all attachments to a local directory.

Filename deduplication strategy (applied when the same filename already exists
in the target directory):
  1. ``{stem}_{safe_sender}{suffix}``      — append sanitised sender address
  2. ``{stem}_{safe_sender}_{timestamp}{suffix}`` — append sender + received time
  3. ``{stem}_{safe_sender}_{timestamp}_{n}{suffix}`` — add counter as last resort
"""

import logging
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.console import cprint
from core.retry import retry_operation, RetryConfig
from office.outlook.reader import OutlookReader
from office.outlook.models import Email, EmailFilter, Attachment

logger = logging.getLogger(__name__)

# Retry config for COM attachment save operations
_SAVE_RETRY = RetryConfig(max_attempts=3, base_delay=1.0, max_delay=10.0)


# ── Result container ─────────────────────────────────────────────

@dataclass
class DownloadResult:
    """Aggregated result of one download run."""

    emails_found: int = 0
    """Total emails on the target date (before keyword filtering)."""

    emails_matched: int = 0
    """Emails that passed the keyword filter."""

    emails_with_attachments: int = 0
    """Matched emails that had at least one attachment."""

    attachments_saved: int = 0
    """Number of attachments successfully written to disk."""

    attachments_failed: int = 0
    """Number of attachments that could not be saved due to errors."""

    saved_files: List[Path] = field(default_factory=list)
    """Absolute paths of every file that was saved."""

    errors: List[str] = field(default_factory=list)
    """Human-readable error messages for failed attachments."""


# ── Filename deduplication ───────────────────────────────────────

def _sanitise_for_filename(value: str) -> str:
    """Replace characters that are unsafe in Windows filenames with underscores."""
    unsafe = r'\/:*?"<>|@. '
    result = value
    for ch in unsafe:
        result = result.replace(ch, "_")
    # Collapse consecutive underscores and strip leading/trailing ones
    while "__" in result:
        result = result.replace("__", "_")
    return result.strip("_")


def _candidate_filenames(
    original: str,
    sender_address: str,
    received_time: Optional[datetime],
) -> List[str]:
    """
    Return ordered list of filename candidates for deduplication.

    Args:
        original:       Original attachment filename (e.g. ``report.xlsx``).
        sender_address: Sender's email address.
        received_time:  Email received timestamp (used for the third candidate).

    Returns:
        List of candidate filenames, from most preferred to least preferred.
        The original name is *not* included; callers should try it first.
    """
    stem = Path(original).stem
    suffix = Path(original).suffix
    safe_sender = _sanitise_for_filename(sender_address) or "unknown_sender"

    ts = received_time.strftime("%Y%m%d_%H%M%S") if received_time else "unknown_time"

    return [
        f"{stem}_{safe_sender}{suffix}",
        f"{stem}_{safe_sender}_{ts}{suffix}",
    ]


def save_attachment_with_dedup(
    attachment: Attachment,
    save_dir: Path,
    sender_address: str,
    received_time: Optional[datetime],
) -> Tuple[Path, str]:
    """
    Save a single attachment to *save_dir*, avoiding filename collisions.

    Deduplication order:
      1. ``<original>``
      2. ``<stem>_<safe_sender><suffix>``
      3. ``<stem>_<safe_sender>_<YYYYMMDD_HHMMSS><suffix>``
      4. ``<stem>_<safe_sender>_<YYYYMMDD_HHMMSS>_<n><suffix>``  (counter)

    Args:
        attachment:     :class:`~office.outlook.models.Attachment` to save.
        save_dir:       Target directory (created if it does not exist).
        sender_address: Sender email address used for dedup naming.
        received_time:  Email received datetime used for dedup naming.

    Returns:
        A ``(saved_path, status)`` tuple, where *status* is one of
        ``"saved"``, ``"renamed_sender"``, ``"renamed_timestamp"``, or
        ``"renamed_counter"``.

    Raises:
        ValueError: If the attachment COM reference is not available.
        OSError: If the file cannot be written to disk.
    """
    if attachment._com_attachment is None:
        raise ValueError(
            f"Cannot save '{attachment.filename}': COM reference is not available."
        )

    save_dir = Path(save_dir)
    save_dir.mkdir(parents=True, exist_ok=True)

    def _write(path: Path) -> None:
        """Invoke the COM SaveAsFile call (wrapped for retries)."""
        attachment._com_attachment.SaveAsFile(str(path))

    # 1. Try original filename first
    original_path = save_dir / attachment.filename
    if not original_path.exists():
        _write(original_path)
        return original_path, "saved"

    # 2 & 3. Try sender-based and timestamp-based candidates
    candidates = _candidate_filenames(attachment.filename, sender_address, received_time)
    status_labels = ["renamed_sender", "renamed_timestamp"]

    for candidate, label in zip(candidates, status_labels):
        candidate_path = save_dir / candidate
        if not candidate_path.exists():
            _write(candidate_path)
            return candidate_path, label

    # 4. Counter fallback: base name is the timestamp variant
    base_ts_name = candidates[-1]   # e.g. report_user_example_com_20260302_143000.xlsx
    stem = Path(base_ts_name).stem
    suffix = Path(base_ts_name).suffix
    counter = 1
    while True:
        fallback = save_dir / f"{stem}_{counter}{suffix}"
        if not fallback.exists():
            _write(fallback)
            return fallback, "renamed_counter"
        counter += 1


# ── Main downloader ──────────────────────────────────────────────

class AttachmentDownloader:
    """
    Downloads email attachments from the Outlook Inbox.

    Filters emails by a specific calendar date and optional subject keywords,
    then saves all attachments to a flat local directory.

    Usage::

        from config import load_config
        from attachment_downloader import AttachmentDownloader

        config = load_config()
        downloader = AttachmentDownloader(config)
        result = downloader.run()
    """

    def __init__(self, config) -> None:
        """
        Args:
            config: A :class:`~config.GetAttachmentConfig` instance.
        """
        self.config = config

    # ── Public API ────────────────────────────────────

    def run(self) -> DownloadResult:
        """
        Execute the full download workflow.

        Connects to Outlook, queries the Inbox for emails from
        :attr:`~config.GetAttachmentConfig.start_date` through
        :attr:`~config.GetAttachmentConfig.end_date_parsed` (inclusive),
        applies keyword filtering, and saves every attachment.

        Returns:
            :class:`DownloadResult` with counts and saved file paths.

        Raises:
            ValueError: If ``start_date`` cannot be parsed.
        """
        result = DownloadResult()

        start_date = self.config.start_date_parsed
        if start_date is None:
            raise ValueError(
                f"Cannot parse start_date: '{self.config.start_date}'. "
                "Expected DD/MM/YYYY."
            )

        end_date = self.config.end_date_parsed  # never None; defaults to today

        # Build date-range filter spanning from start of start_date
        # to end of end_date (inclusive)
        received_after = datetime.combine(start_date, datetime.min.time())
        received_before = datetime.combine(end_date, datetime.max.time())

        email_filter = EmailFilter(
            received_after=received_after,
            received_before=received_before,
            limit=10000,
        )

        cprint(f"Connecting to Outlook account: {self.config.outlook_account}", level="INFO")

        with OutlookReader(account=self.config.outlook_account) as client:
            cprint("Connected to Outlook", level="SUCCESS")
            cprint(
                f"Searching Inbox for emails from {self.config.start_date} "
                f"to {self.config.date_range_display.split(' \u2192 ')[-1]}...",
                level="INFO",
            )

            emails = client.get_inbox_emails(filter=email_filter)
            result.emails_found = len(emails)
            cprint(
                f"Found {len(emails)} email(s) in range {self.config.date_range_display}",
                level="INFO",
            )

            for email in emails:
                self._process_email(email, result)

        return result

    # ── Internal helpers ──────────────────────────────

    def _matches_keywords(self, email: Email) -> bool:
        """
        Return ``True`` if the email subject matches the configured keywords.

        Matching uses OR logic: the email passes if its subject contains
        *any* of the configured keywords (case-insensitive).  If no
        keywords are configured (empty list), every email is accepted.
        """
        keywords = self.config.subject_keywords
        if not keywords:
            return True
        subject_lower = email.subject.lower()
        return any(kw.lower() in subject_lower for kw in keywords)

    def _process_email(self, email: Email, result: DownloadResult) -> None:
        """
        Process a single email: keyword-filter then save all attachments.

        Errors for individual attachments are captured in *result* and do
        not stop processing of the remaining attachments or emails.
        """
        if not self._matches_keywords(email):
            logger.debug(
                "Skipped (keyword mismatch): subject='%s' sender='%s'",
                email.subject,
                email.sender_address,
            )
            return

        result.emails_matched += 1

        if not email.has_attachments:
            logger.info(
                "No attachments: subject='%s' sender='%s'",
                email.subject,
                email.sender_address,
            )
            return

        result.emails_with_attachments += 1

        received_str = (
            email.received_time.strftime("%Y-%m-%d %H:%M:%S")
            if email.received_time
            else "unknown"
        )
        cprint(
            f"[{received_str}] {email.sender_address}: {email.subject} "
            f"({len(email.attachments)} attachment(s))",
            level="PROGRESS",
        )
        logger.info(
            "Processing email: subject='%s' sender='%s' received='%s' attachments=%d",
            email.subject,
            email.sender_address,
            received_str,
            len(email.attachments),
        )

        for attachment in email.attachments:
            self._save_one_attachment(attachment, email, result)

    def _save_one_attachment(
        self,
        attachment: Attachment,
        email: Email,
        result: DownloadResult,
    ) -> None:
        """
        Save a single attachment with COM retry handling.

        On success: increments ``result.attachments_saved`` and appends to
        ``result.saved_files``.
        On failure: increments ``result.attachments_failed`` and appends an
        error message to ``result.errors``.
        """
        try:
            saved_path, status = _retry_save(
                attachment=attachment,
                save_dir=self.config.attachment_save_path,
                sender_address=email.sender_address,
                received_time=email.received_time,
            )
            result.attachments_saved += 1
            result.saved_files.append(saved_path)

            _STATUS_NOTE = {
                "saved": "",
                "renamed_sender": " (renamed: added sender suffix)",
                "renamed_timestamp": " (renamed: added sender+timestamp suffix)",
                "renamed_counter": " (renamed: added counter suffix)",
            }
            note = _STATUS_NOTE.get(status, "")
            cprint(f"  Saved: {saved_path.name}{note}", level="SUCCESS", indent=2)
            logger.info("Saved attachment: path='%s' status=%s", saved_path, status)

        except Exception as exc:
            result.attachments_failed += 1
            err_msg = (
                f"Failed to save '{attachment.filename}' "
                f"from email '{email.subject}': {exc}"
            )
            result.errors.append(err_msg)
            cprint(f"  Failed: {attachment.filename} — {exc}", level="ERROR", indent=2)
            logger.error(err_msg, exc_info=True)


# ── Retry wrapper ────────────────────────────────────────────────

@retry_operation(_SAVE_RETRY)
def _retry_save(
    attachment: Attachment,
    save_dir: Path,
    sender_address: str,
    received_time: Optional[datetime],
) -> Tuple[Path, str]:
    """
    Thin retry wrapper around :func:`save_attachment_with_dedup`.

    Decorated with :func:`~core.retry.retry_operation` so transient
    COM errors cause automatic retries with exponential back-off.
    """
    return save_attachment_with_dedup(attachment, save_dir, sender_address, received_time)
