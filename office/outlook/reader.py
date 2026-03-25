"""
Outlook email reader — read-only mailbox operations.

Extends OutlookClient with folder navigation and email retrieval.
Not intended for mutation operations (sending, saving, etc.).

Usage:
    from office.outlook.reader import OutlookReader

    with OutlookReader(account="your.email@company.com") as reader:
        emails = reader.get_inbox_emails()
        inbox = reader.get_inbox()
        for email in reader.iterate_emails(inbox):
            print(email.subject)
"""

import logging
from datetime import datetime
from pathlib import Path
from typing import Generator, List, Optional

from office.outlook.client import OutlookClient
from office.outlook.exceptions import OutlookFolderNotFoundError
from office.outlook.models import (
    Attachment,
    Email,
    EmailFilter,
    FolderInfo,
    Importance,
)

logger = logging.getLogger(__name__)

# Outlook default folder type constants (olDefaultFolders)
FOLDER_TYPE_INBOX = 6
FOLDER_TYPE_SENT = 5
FOLDER_TYPE_DRAFTS = 16
FOLDER_TYPE_DELETED = 3
FOLDER_TYPE_OUTBOX = 4


class OutlookReader(OutlookClient):
    """
    Read-only Outlook client for folder navigation and email retrieval.

    Extends the base OutlookClient connection lifecycle with methods for
    reading emails, iterating folders, and downloading attachments.

    Use as a context manager:

        with OutlookReader(account="your.email@company.com") as reader:
            emails = reader.get_inbox_emails()
    """

    def connect(self) -> None:
        """
        Connect to Outlook and resolve the account root folder.

        Calls the base connection logic (with retry) then resolves the
        account's delivery store root folder for folder tree navigation.
        """
        super().connect()
        self._account_folder = self._account_obj.DeliveryStore.GetRootFolder()
        logger.debug(
            f"[READER] Account root folder resolved: "
            f"{getattr(self._account_folder, 'FolderPath', '?')}"
        )

    # ── Folder navigation ────────────────────────────────────────

    def _get_default_folder(self, folder_type: int) -> object:
        """
        Get a default folder by its Outlook folder type constant.

        Args:
            folder_type: Outlook olDefaultFolders constant
                         (e.g. FOLDER_TYPE_INBOX).

        Returns:
            The folder COM object.
        """
        return self._namespace.GetDefaultFolder(folder_type)

    def get_inbox(self) -> object:
        """Get the Inbox default folder."""
        return self._get_default_folder(FOLDER_TYPE_INBOX)

    def get_sent_folder(self) -> object:
        """Get the Sent Items default folder."""
        return self._get_default_folder(FOLDER_TYPE_SENT)

    def get_folder_by_path(self, path: str) -> object:
        """
        Get a subfolder by path relative to the account root.

        Args:
            path: Slash- or backslash-separated folder path
                  (e.g. "Inbox/Subfolder" or "Inbox\\Subfolder").

        Returns:
            The folder COM object.

        Raises:
            OutlookFolderNotFoundError: If any segment in the path is not found.
        """
        self._ensure_connected()

        path = path.replace("\\", "/")
        parts = [p for p in path.split("/") if p]

        if not parts:
            raise OutlookFolderNotFoundError(path, self.account)

        current = self._account_folder

        for part in parts:
            found = False
            for folder in current.Folders:
                if folder.Name.lower() == part.lower():
                    current = folder
                    found = True
                    break

            if not found:
                raise OutlookFolderNotFoundError(path, self.account)

        return current

    def get_folder_info(self, folder: object) -> FolderInfo:
        """
        Get metadata about a folder.

        Uses safe attribute access to avoid crashes on folders with
        restricted access or missing properties.

        Args:
            folder: Folder COM object.

        Returns:
            FolderInfo dataclass.
        """
        items = getattr(folder, "Items", None)
        return FolderInfo(
            name=getattr(folder, "Name", ""),
            path=getattr(folder, "FolderPath", ""),
            item_count=getattr(items, "Count", 0) if items is not None else 0,
            unread_count=getattr(folder, "UnReadItemCount", 0),
        )

    def list_folders(self, parent: Optional[object] = None) -> List[FolderInfo]:
        """
        List all direct sub-folders under a parent.

        Args:
            parent: Parent folder COM object (uses account root if None).

        Returns:
            List of FolderInfo objects.
        """
        self._ensure_connected()

        if parent is None:
            parent = self._account_folder

        folders = []
        for folder in parent.Folders:
            try:
                folders.append(self.get_folder_info(folder))
            except Exception as e:
                logger.debug(f"Error getting folder info: {e}")

        return folders

    # ── Email retrieval ──────────────────────────────────────────

    def get_inbox_emails(
        self,
        filter: Optional[EmailFilter] = None,
    ) -> List[Email]:
        """
        Get emails from the Inbox with optional filtering.

        Args:
            filter: Optional filter criteria.

        Returns:
            List of Email objects.
        """
        return self.get_emails_from_folder(self.get_inbox(), filter)

    def get_emails_from_folder(
        self,
        folder: object,
        filter: Optional[EmailFilter] = None,
    ) -> List[Email]:
        """
        Get emails from a specific folder with optional filtering.

        Applies Outlook server-side restrictions where possible and falls
        back to client-side filtering for more complex criteria.

        Args:
            folder: Folder COM object.
            filter: Optional filter criteria.

        Returns:
            List of Email objects (newest first, up to filter.limit).
        """
        self._ensure_connected()

        if filter is None:
            filter = EmailFilter()

        emails: List[Email] = []
        items = folder.Items
        items.Sort("[ReceivedTime]", True)

        outlook_filter = filter.to_outlook_filter()
        if outlook_filter:
            logger.debug(f"Applying Outlook filter: {outlook_filter}")
            try:
                items = items.Restrict(outlook_filter)
            except Exception as e:
                logger.warning(
                    f"Outlook filter failed, using client-side filtering: {e}"
                )

        count = 0
        for item in items:
            if count >= filter.limit:
                break

            try:
                if not hasattr(item, "MessageClass"):
                    continue
                if not item.MessageClass.startswith("IPM.Note"):
                    continue

                email = self._item_to_email(item)

                if filter.matches(email):
                    emails.append(email)
                    count += 1

            except Exception as e:
                logger.debug(f"Error processing email item: {e}")
                continue

        logger.debug(f"Retrieved {len(emails)} emails from folder")
        return emails

    def iterate_emails(
        self,
        folder: object,
        filter: Optional[EmailFilter] = None,
    ) -> Generator[Email, None, None]:
        """
        Iterate through emails one at a time (memory-efficient).

        Applies the same server-side + client-side filtering as
        get_emails_from_folder, but yields items individually to avoid
        building a large list in memory.

        Args:
            folder: Folder COM object.
            filter: Optional filter criteria.

        Yields:
            Email objects one at a time, newest first.
        """
        self._ensure_connected()

        if filter is None:
            filter = EmailFilter(limit=10000)

        items = folder.Items
        items.Sort("[ReceivedTime]", True)

        outlook_filter = filter.to_outlook_filter()
        if outlook_filter:
            try:
                items = items.Restrict(outlook_filter)
            except Exception as e:
                logger.warning(f"Outlook filter failed: {e}")

        count = 0
        for item in items:
            if count >= filter.limit:
                break

            try:
                if not hasattr(item, "MessageClass"):
                    continue
                if not item.MessageClass.startswith("IPM.Note"):
                    continue

                email = self._item_to_email(item)

                if filter.matches(email):
                    yield email
                    count += 1

            except Exception as e:
                logger.debug(f"Error processing email: {e}")
                continue

    def get_email_by_id(self, entry_id: str) -> Optional[Email]:
        """
        Get a specific email by its Outlook Entry ID.

        Args:
            entry_id: Outlook Entry ID of the email.

        Returns:
            Email object, or None if not found.
        """
        self._ensure_connected()

        try:
            item = self._namespace.GetItemFromID(entry_id)
            return self._item_to_email(item)
        except Exception as e:
            logger.debug(f"Email not found: {entry_id}: {e}")
            return None

    # ── Attachment operations ────────────────────────────────────

    def save_attachments(self, email: Email, directory: Path) -> List[Path]:
        """
        Save all attachments from an email to a directory.

        Args:
            email: Email containing attachments.
            directory: Destination directory (created if missing).

        Returns:
            List of paths to the saved attachment files.
        """
        saved: List[Path] = []
        for attachment in email.attachments:
            try:
                path = attachment.save(directory)
                saved.append(path)
                logger.debug(f"Saved attachment: {path}")
            except Exception as e:
                logger.warning(
                    f"Failed to save attachment {attachment.filename}: {e}"
                )
        return saved

    # ── Internal helpers ─────────────────────────────────────────

    def _item_to_email(self, item: object) -> Email:
        """
        Convert an Outlook MailItem COM object to an Email dataclass.

        Args:
            item: Outlook MailItem COM object.

        Returns:
            Email dataclass.
        """
        sender_address = ""
        sender_name = ""

        try:
            sender_name = getattr(item, "SenderName", "")

            if hasattr(item, "SenderEmailAddress"):
                sender_address = item.SenderEmailAddress

            # Resolve Exchange internal address format (/O=...)
            if sender_address.startswith("/O=") or "@" not in sender_address:
                try:
                    sender = item.Sender
                    if sender and hasattr(sender, "GetExchangeUser"):
                        exchange_user = sender.GetExchangeUser()
                        if exchange_user:
                            sender_address = exchange_user.PrimarySmtpAddress
                except Exception:
                    pass

        except Exception as e:
            logger.debug(f"Error getting sender info: {e}")

        recipients: List[str] = []
        try:
            for i in range(1, item.Recipients.Count + 1):
                recip = item.Recipients.Item(i)
                recip_addr = getattr(recip, "Address", "")
                if recip_addr:
                    recipients.append(recip_addr)
        except Exception as e:
            logger.debug(f"Error getting recipients: {e}")

        attachments: List[Attachment] = []
        try:
            for i in range(1, item.Attachments.Count + 1):
                att = item.Attachments.Item(i)

                # Read Content-ID for inline image detection
                content_id = ""
                try:
                    PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
                    cid = att.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID)
                    if cid:
                        content_id = str(cid)
                except Exception:
                    pass

                # Read PR_ATTACH_FLAGS for ATT_MHTML_REF inline detection
                attach_flags = 0
                try:
                    PR_ATTACH_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x37140003"
                    flags = att.PropertyAccessor.GetProperty(PR_ATTACH_FLAGS)
                    if flags is not None:
                        attach_flags = int(flags)
                except Exception:
                    pass

                attachments.append(Attachment(
                    filename=getattr(att, "FileName", f"attachment_{i}"),
                    size=getattr(att, "Size", 0),
                    content_type=getattr(att, "ContentType", ""),
                    attachment_type=getattr(att, "Type", 1),
                    content_id=content_id,
                    attach_flags=attach_flags,
                    _com_attachment=att,
                ))
        except Exception as e:
            logger.debug(f"Error getting attachments: {e}")

        categories: List[str] = []
        try:
            cat_str = getattr(item, "Categories", "")
            if cat_str:
                categories = [c.strip() for c in cat_str.split(",")]
        except Exception:
            pass

        importance = Importance.NORMAL
        try:
            imp_value = getattr(item, "Importance", 1)
            importance = Importance(imp_value)
        except Exception:
            pass

        received_time = None
        sent_time = None
        try:
            rt = getattr(item, "ReceivedTime", None)
            if rt:
                received_time = datetime(
                    rt.year, rt.month, rt.day,
                    rt.hour, rt.minute, rt.second,
                )
            st = getattr(item, "SentOn", None)
            if st:
                sent_time = datetime(
                    st.year, st.month, st.day,
                    st.hour, st.minute, st.second,
                )
        except Exception as e:
            logger.debug(f"Error parsing timestamps: {e}")

        return Email(
            entry_id=getattr(item, "EntryID", ""),
            message_id=self._get_message_id(item),
            subject=getattr(item, "Subject", ""),
            sender_address=sender_address,
            sender_name=sender_name,
            recipients=recipients,
            received_time=received_time,
            sent_time=sent_time,
            body_text=getattr(item, "Body", ""),
            body_html=getattr(item, "HTMLBody", ""),
            is_read=not getattr(item, "UnRead", False),
            importance=importance,
            attachments=attachments,
            categories=categories,
            conversation_id=getattr(item, "ConversationID", ""),
            _com_item=item,
        )

    def _get_message_id(self, item: object) -> str:
        """
        Extract the Internet Message-ID header from an email.

        Args:
            item: Outlook MailItem COM object.

        Returns:
            Message-ID string, or empty string if unavailable.
        """
        try:
            PR_INTERNET_MESSAGE_ID = (
                "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
            )
            prop_accessor = item.PropertyAccessor
            message_id = prop_accessor.GetProperty(PR_INTERNET_MESSAGE_ID)
            return message_id or ""
        except Exception:
            return ""
