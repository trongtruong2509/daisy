"""
Data models for Outlook operations.

These classes provide clean, typed representations of Outlook objects
without exposing COM complexities to calling code.
"""

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import List, Optional


class Importance(Enum):
    """Email importance levels."""
    LOW = 0
    NORMAL = 1
    HIGH = 2


@dataclass
class AccountInfo:
    """
    Information about an Outlook account.
    
    Attributes:
        smtp_address: The SMTP email address of the account.
        display_name: Display name of the account.
        account_type: Type of account (e.g., "Exchange", "IMAP").
    """
    smtp_address: str
    display_name: str
    account_type: str = ""
    
    def __str__(self) -> str:
        return f"{self.display_name} <{self.smtp_address}>"


@dataclass
class FolderInfo:
    """
    Information about an Outlook folder.
    
    Attributes:
        name: Folder name.
        path: Full path to the folder.
        item_count: Number of items in the folder.
        unread_count: Number of unread items.
        folder_type: Type of folder (Inbox, Sent, etc.).
    """
    name: str
    path: str
    item_count: int = 0
    unread_count: int = 0
    folder_type: str = ""
    
    def __str__(self) -> str:
        return self.path


@dataclass
class Attachment:
    """
    Represents an email attachment.
    
    Attributes:
        filename: Name of the attachment file.
        size: Size in bytes.
        content_type: MIME type of the attachment.
        _com_attachment: Internal reference to COM object (not for external use).
    """
    filename: str
    size: int
    content_type: str = ""
    _com_attachment: object = field(default=None, repr=False, compare=False)
    
    def save(self, directory: Path) -> Path:
        """
        Save the attachment to a directory.
        
        Args:
            directory: Directory to save the attachment to.
            
        Returns:
            Path to the saved file.
            
        Raises:
            ValueError: If the attachment cannot be saved.
        """
        if self._com_attachment is None:
            raise ValueError("Cannot save attachment: COM reference not available")
        
        directory = Path(directory)
        directory.mkdir(parents=True, exist_ok=True)
        
        save_path = directory / self.filename
        
        # Handle duplicate filenames
        counter = 1
        while save_path.exists():
            stem = Path(self.filename).stem
            suffix = Path(self.filename).suffix
            save_path = directory / f"{stem}_{counter}{suffix}"
            counter += 1
        
        self._com_attachment.SaveAsFile(str(save_path))
        return save_path


@dataclass
class Email:
    """
    Represents an email message.
    
    This is a clean, typed representation of an Outlook MailItem.
    All COM interaction is abstracted away.
    
    Attributes:
        entry_id: Unique Outlook identifier for the email.
        message_id: Internet Message-ID header (for duplicate detection).
        subject: Email subject line.
        sender_address: Sender's email address.
        sender_name: Sender's display name.
        recipients: List of recipient email addresses.
        received_time: When the email was received.
        sent_time: When the email was sent.
        body_text: Plain text body.
        body_html: HTML body (if available).
        is_read: Whether the email has been read.
        importance: Email importance level.
        attachments: List of attachments.
        categories: List of categories/tags.
        conversation_id: Conversation thread ID.
        _com_item: Internal reference to COM object (not for external use).
    """
    entry_id: str
    message_id: str
    subject: str
    sender_address: str
    sender_name: str
    recipients: List[str]
    received_time: Optional[datetime]
    sent_time: Optional[datetime]
    body_text: str
    body_html: str
    is_read: bool
    importance: Importance
    attachments: List[Attachment] = field(default_factory=list)
    categories: List[str] = field(default_factory=list)
    conversation_id: str = ""
    _com_item: object = field(default=None, repr=False, compare=False)
    
    def __str__(self) -> str:
        date_str = self.received_time.strftime("%Y-%m-%d %H:%M") if self.received_time else "Unknown"
        return f"[{date_str}] {self.sender_name}: {self.subject}"
    
    def save_as_msg(self, path: Path) -> Path:
        """
        Save the email as a .msg file.
        
        Args:
            path: Path to save the .msg file (including filename).
            
        Returns:
            Path to the saved file.
            
        Raises:
            ValueError: If the email cannot be saved.
        """
        if self._com_item is None:
            raise ValueError("Cannot save email: COM reference not available")
        
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        
        # Outlook SaveAs format: 3 = olMsg
        self._com_item.SaveAs(str(path), 3)
        return path
    
    def mark_as_read(self) -> None:
        """Mark the email as read."""
        if self._com_item is None:
            raise ValueError("Cannot modify email: COM reference not available")
        
        self._com_item.UnRead = False
        self._com_item.Save()
        self.is_read = True
    
    def mark_as_unread(self) -> None:
        """Mark the email as unread."""
        if self._com_item is None:
            raise ValueError("Cannot modify email: COM reference not available")
        
        self._com_item.UnRead = True
        self._com_item.Save()
        self.is_read = False
    
    @property
    def has_attachments(self) -> bool:
        """Check if the email has attachments."""
        return len(self.attachments) > 0
    
    @property
    def unique_id(self) -> str:
        """
        Get a unique identifier for duplicate detection.
        
        Prefers Message-ID, falls back to Entry-ID.
        """
        return self.message_id if self.message_id else self.entry_id


@dataclass
class EmailFilter:
    """
    Filter criteria for email queries.
    
    All criteria are optional and combined with AND logic.
    
    Attributes:
        unread_only: Only return unread emails.
        sender_contains: Sender address must contain this string (case-insensitive).
        subject_contains: Subject must contain this string (case-insensitive).
        received_after: Only emails received after this datetime.
        received_before: Only emails received before this datetime.
        has_attachments: Only emails with attachments.
        limit: Maximum number of emails to return.
        categories: Only emails with these categories.
    """
    unread_only: bool = False
    sender_contains: Optional[str] = None
    subject_contains: Optional[str] = None
    received_after: Optional[datetime] = None
    received_before: Optional[datetime] = None
    has_attachments: Optional[bool] = None
    limit: int = 100
    categories: Optional[List[str]] = None
    
    def __post_init__(self):
        """Validate filter parameters."""
        if self.limit < 1:
            self.limit = 1
        if self.limit > 10000:
            self.limit = 10000
    
    def matches(self, email: Email) -> bool:
        """
        Check if an email matches this filter.
        
        Used for client-side filtering when server-side is not possible.
        
        Args:
            email: Email to check.
            
        Returns:
            True if the email matches all criteria.
        """
        if self.unread_only and email.is_read:
            return False
        
        if self.sender_contains:
            if self.sender_contains.lower() not in email.sender_address.lower():
                return False
        
        if self.subject_contains:
            if self.subject_contains.lower() not in email.subject.lower():
                return False
        
        if self.received_after and email.received_time:
            if email.received_time < self.received_after:
                return False
        
        if self.received_before and email.received_time:
            if email.received_time > self.received_before:
                return False
        
        if self.has_attachments is not None:
            if email.has_attachments != self.has_attachments:
                return False
        
        if self.categories:
            if not any(cat in email.categories for cat in self.categories):
                return False
        
        return True
    
    def to_outlook_filter(self) -> Optional[str]:
        """
        Convert to Outlook restriction filter string.
        
        Returns:
            Filter string for Outlook Items.Restrict(), or None for no filter.
        """
        conditions = []
        
        if self.unread_only:
            conditions.append("[UnRead] = True")
        
        if self.received_after:
            date_str = self.received_after.strftime("%m/%d/%Y %H:%M %p")
            conditions.append(f"[ReceivedTime] >= '{date_str}'")
        
        if self.received_before:
            date_str = self.received_before.strftime("%m/%d/%Y %H:%M %p")
            conditions.append(f"[ReceivedTime] <= '{date_str}'")
        
        # Note: Subject and Sender filters require DASL syntax for reliable results
        # We'll handle those client-side for simplicity and reliability
        
        if not conditions:
            return None
        
        return " AND ".join(conditions)


@dataclass
class NewEmail:
    """
    Data for creating a new email to send.
    
    Attributes:
        to: List of recipient email addresses.
        subject: Email subject.
        body: Email body (plain text or HTML).
        body_is_html: Whether body is HTML formatted.
        cc: List of CC recipients.
        bcc: List of BCC recipients.
        attachments: List of file paths to attach.
        importance: Email importance level.
    """
    to: List[str]
    subject: str
    body: str
    body_is_html: bool = False
    cc: List[str] = field(default_factory=list)
    bcc: List[str] = field(default_factory=list)
    attachments: List[Path] = field(default_factory=list)
    importance: Importance = Importance.NORMAL
    
    def __post_init__(self):
        """Validate and normalize email data."""
        # Ensure to is a list
        if isinstance(self.to, str):
            self.to = [self.to]
        
        # Ensure cc and bcc are lists
        if isinstance(self.cc, str):
            self.cc = [self.cc]
        if isinstance(self.bcc, str):
            self.bcc = [self.bcc]
        
        # Ensure attachments are Path objects
        self.attachments = [Path(a) if isinstance(a, str) else a for a in self.attachments]
    
    def validate(self) -> List[str]:
        """
        Validate the email data.
        
        Returns:
            List of validation error messages (empty if valid).
        """
        errors = []
        
        if not self.to:
            errors.append("At least one recipient is required")
        
        for addr in self.to:
            if "@" not in addr:
                errors.append(f"Invalid recipient address: {addr}")
        
        for addr in self.cc:
            if "@" not in addr:
                errors.append(f"Invalid CC address: {addr}")
        
        for addr in self.bcc:
            if "@" not in addr:
                errors.append(f"Invalid BCC address: {addr}")
        
        if not self.subject:
            errors.append("Subject is required")
        
        for attachment in self.attachments:
            if not attachment.exists():
                errors.append(f"Attachment not found: {attachment}")
        
        return errors
    
    def is_valid(self) -> bool:
        """Check if the email data is valid."""
        return len(self.validate()) == 0
    
    @property
    def all_recipients(self) -> List[str]:
        """Get all recipients (to + cc + bcc)."""
        return self.to + self.cc + self.bcc
    
    def get_content_hash(self) -> str:
        """
        Get a hash for duplicate detection.
        
        Based on recipients, subject, and body content.
        """
        import hashlib
        content = f"{','.join(sorted(self.all_recipients))}|{self.subject}|{self.body}"
        return hashlib.sha256(content.encode()).hexdigest()
