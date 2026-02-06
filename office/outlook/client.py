"""
Outlook client for reading emails.

Provides a clean interface to Outlook Desktop COM automation.
All COM operations are isolated in this module.

IMPORTANT COM CAVEATS:
- Outlook must be running for COM to work
- COM objects must be accessed from the same thread that created them
- Some operations may fail if Outlook is "busy" (retry logic helps)
- Large mailboxes can be slow - use filters and limits

Usage:
    from office.outlook import OutlookClient, EmailFilter
    
    with OutlookClient(account="your.email@company.com") as client:
        emails = client.get_inbox_emails(
            filter=EmailFilter(unread_only=True, limit=50)
        )
        for email in emails:
            print(f"{email.subject} from {email.sender_name}")
"""

import logging
from datetime import datetime
from pathlib import Path
from typing import Generator, List, Optional

# COM imports - will fail on non-Windows systems
try:
    import win32com.client
    import pywintypes
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

from core.retry import retry_operation, RetryConfig
from office.outlook.exceptions import (
    OutlookConnectionError,
    OutlookAccountNotFoundError,
    OutlookFolderNotFoundError,
    OutlookError,
    OutlookItemError,
)
from office.outlook.models import (
    Email,
    EmailFilter,
    Attachment,
    FolderInfo,
    AccountInfo,
    Importance,
)

logger = logging.getLogger(__name__)

# Outlook folder type constants
FOLDER_TYPE_INBOX = 6
FOLDER_TYPE_SENT = 5
FOLDER_TYPE_DRAFTS = 16
FOLDER_TYPE_DELETED = 3
FOLDER_TYPE_OUTBOX = 4


class OutlookClient:
    """
    Client for reading emails from Outlook Desktop.
    
    Provides a clean interface hiding all COM complexity.
    Supports multiple accounts in one Outlook profile.
    
    Use as a context manager for proper resource cleanup:
    
        with OutlookClient(account="user@example.com") as client:
            emails = client.get_inbox_emails()
    
    Attributes:
        account: SMTP address of the account to use.
        retry_config: Configuration for retry behavior.
    """
    
    def __init__(
        self,
        account: str,
        retry_config: Optional[RetryConfig] = None
    ):
        """
        Initialize Outlook client.
        
        Args:
            account: SMTP address of the Outlook account to use.
            retry_config: Retry configuration (uses sensible defaults if None).
            
        Raises:
            ImportError: If win32com is not available (non-Windows).
        """
        if not HAS_WIN32COM:
            raise ImportError(
                "win32com is required for Outlook operations. "
                "Install pywin32: pip install pywin32"
            )
        
        self.account = account
        self.retry_config = retry_config or RetryConfig(
            max_attempts=3,
            base_delay=2.0,
            max_delay=30.0
        )
        
        self._outlook: Optional[object] = None
        self._namespace: Optional[object] = None
        self._account_folder: Optional[object] = None
        self._connected = False
    
    def __enter__(self) -> "OutlookClient":
        """Connect to Outlook on context entry."""
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb) -> bool:
        """Clean up on context exit."""
        self.disconnect()
        return False
    
    @retry_operation()
    def connect(self) -> None:
        """
        Connect to Outlook and select the specified account.
        
        Raises:
            OutlookConnectionError: If unable to connect to Outlook.
            OutlookAccountNotFoundError: If the specified account is not found.
        """
        if self._connected:
            return
        
        logger.debug("Connecting to Outlook...")
        
        try:
            # Get Outlook application
            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook.GetNamespace("MAPI")
            
            logger.debug("Connected to Outlook application")
        
        except pywintypes.com_error as e:
            raise OutlookConnectionError(
                f"Failed to connect to Outlook: {e}. Is Outlook running?"
            )
        
        # Find the specified account
        self._account_folder = self._find_account_folder()
        self._connected = True
        
        logger.info(f"Connected to Outlook account: {self.account}")
    
    def _find_account_folder(self) -> object:
        """
        Find the root folder for the specified account.
        
        Returns:
            The account's root folder COM object.
            
        Raises:
            OutlookAccountNotFoundError: If the account is not found.
        """
        available_accounts = []
        
        # Search through all stores (each account has its own store)
        for store in self._namespace.Stores:
            try:
                # Try to get the account's SMTP address
                if hasattr(store, "DisplayName"):
                    available_accounts.append(store.DisplayName)
                
                root_folder = store.GetRootFolder()
                
                # Check if this is the target account
                # Try multiple approaches to match the account
                
                # Approach 1: Check store display name
                if self.account.lower() in store.DisplayName.lower():
                    logger.debug(f"Found account by store name: {store.DisplayName}")
                    return root_folder
                
                # Approach 2: Check folders for account SMTP address
                for folder in root_folder.Folders:
                    folder_name = getattr(folder, "Name", "")
                    if self.account.lower() in folder_name.lower():
                        logger.debug(f"Found account by folder name: {folder_name}")
                        return root_folder
            
            except Exception as e:
                logger.debug(f"Error checking store: {e}")
                continue
        
        # Also try the Accounts collection
        try:
            for i in range(1, self._namespace.Accounts.Count + 1):
                acc = self._namespace.Accounts.Item(i)
                smtp = getattr(acc, "SmtpAddress", "")
                available_accounts.append(smtp)
                
                if smtp.lower() == self.account.lower():
                    # Find the delivery store for this account
                    delivery_store = acc.DeliveryStore
                    if delivery_store:
                        logger.debug(f"Found account by SMTP address: {smtp}")
                        return delivery_store.GetRootFolder()
        
        except Exception as e:
            logger.debug(f"Error checking Accounts collection: {e}")
        
        raise OutlookAccountNotFoundError(self.account, available_accounts)
    
    def disconnect(self) -> None:
        """Disconnect from Outlook and clean up resources."""
        self._outlook = None
        self._namespace = None
        self._account_folder = None
        self._connected = False
        logger.debug("Disconnected from Outlook")
    
    @property
    def is_connected(self) -> bool:
        """Check if connected to Outlook."""
        return self._connected
    
    def get_available_accounts(self) -> List[AccountInfo]:
        """
        Get list of available accounts in the Outlook profile.
        
        Returns:
            List of AccountInfo objects.
        """
        self._ensure_connected()
        accounts = []
        
        try:
            for i in range(1, self._namespace.Accounts.Count + 1):
                acc = self._namespace.Accounts.Item(i)
                accounts.append(AccountInfo(
                    smtp_address=getattr(acc, "SmtpAddress", ""),
                    display_name=getattr(acc, "DisplayName", ""),
                    account_type=getattr(acc, "AccountType", ""),
                ))
        except Exception as e:
            logger.warning(f"Error getting accounts: {e}")
        
        return accounts
    
    def _ensure_connected(self) -> None:
        """Ensure the client is connected."""
        if not self._connected:
            self.connect()
    
    def _get_default_folder(self, folder_type: int) -> object:
        """
        Get a default folder (Inbox, Sent, etc.) for the account.
        
        Args:
            folder_type: Outlook folder type constant.
            
        Returns:
            The folder COM object.
        """
        self._ensure_connected()
        
        # Try to get the folder from the account's store
        for folder in self._account_folder.Folders:
            folder_name = getattr(folder, "Name", "").lower()
            
            if folder_type == FOLDER_TYPE_INBOX and folder_name in ("inbox", "posteingang", "boîte de réception"):
                return folder
            if folder_type == FOLDER_TYPE_SENT and "sent" in folder_name:
                return folder
            if folder_type == FOLDER_TYPE_DRAFTS and "draft" in folder_name:
                return folder
            if folder_type == FOLDER_TYPE_DELETED and ("deleted" in folder_name or "trash" in folder_name):
                return folder
        
        # Fallback: use namespace default folder
        return self._namespace.GetDefaultFolder(folder_type)
    
    def get_inbox(self) -> object:
        """Get the Inbox folder."""
        return self._get_default_folder(FOLDER_TYPE_INBOX)
    
    def get_sent_folder(self) -> object:
        """Get the Sent Items folder."""
        return self._get_default_folder(FOLDER_TYPE_SENT)
    
    def get_folder_by_path(self, path: str) -> object:
        """
        Get a folder by its path.
        
        Args:
            path: Folder path (e.g., "Inbox/Subfolder" or "Inbox\\Subfolder").
            
        Returns:
            The folder COM object.
            
        Raises:
            OutlookFolderNotFoundError: If the folder is not found.
        """
        self._ensure_connected()
        
        # Normalize path separators
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
        Get information about a folder.
        
        Args:
            folder: Folder COM object.
            
        Returns:
            FolderInfo object.
        """
        return FolderInfo(
            name=getattr(folder, "Name", ""),
            path=getattr(folder, "FolderPath", ""),
            item_count=getattr(folder, "Items", object()).Count if hasattr(folder, "Items") else 0,
            unread_count=getattr(folder, "UnReadItemCount", 0),
        )
    
    def list_folders(self, parent: Optional[object] = None) -> List[FolderInfo]:
        """
        List all folders under a parent.
        
        Args:
            parent: Parent folder (uses account root if None).
            
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
    
    def get_inbox_emails(
        self,
        filter: Optional[EmailFilter] = None
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
        filter: Optional[EmailFilter] = None
    ) -> List[Email]:
        """
        Get emails from a specific folder with optional filtering.
        
        Args:
            folder: Folder COM object.
            filter: Optional filter criteria.
            
        Returns:
            List of Email objects.
        """
        self._ensure_connected()
        
        if filter is None:
            filter = EmailFilter()
        
        emails = []
        items = folder.Items
        
        # Sort by received time (newest first)
        items.Sort("[ReceivedTime]", True)
        
        # Apply Outlook-side filter if possible
        outlook_filter = filter.to_outlook_filter()
        if outlook_filter:
            logger.debug(f"Applying Outlook filter: {outlook_filter}")
            try:
                items = items.Restrict(outlook_filter)
            except Exception as e:
                logger.warning(f"Outlook filter failed, using client-side filtering: {e}")
        
        count = 0
        for item in items:
            if count >= filter.limit:
                break
            
            try:
                # Check if it's a mail item (not a meeting request, etc.)
                if not hasattr(item, "MessageClass"):
                    continue
                if not item.MessageClass.startswith("IPM.Note"):
                    continue
                
                email = self._item_to_email(item)
                
                # Apply client-side filtering
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
        batch_size: int = 50
    ) -> Generator[Email, None, None]:
        """
        Iterate through emails in batches (memory-efficient).
        
        Args:
            folder: Folder COM object.
            filter: Optional filter criteria.
            batch_size: Number of emails to fetch at a time.
            
        Yields:
            Email objects one at a time.
        """
        self._ensure_connected()
        
        if filter is None:
            filter = EmailFilter(limit=10000)  # High limit for iteration
        
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
        Get a specific email by its Entry ID.
        
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
    
    def _item_to_email(self, item: object) -> Email:
        """
        Convert an Outlook MailItem to an Email object.
        
        Args:
            item: Outlook MailItem COM object.
            
        Returns:
            Email object.
        """
        # Get sender information
        sender_address = ""
        sender_name = ""
        
        try:
            sender_name = getattr(item, "SenderName", "")
            
            # Try to get sender email address
            if hasattr(item, "SenderEmailAddress"):
                sender_address = item.SenderEmailAddress
            
            # If exchange format, try to resolve
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
        
        # Get recipients
        recipients = []
        try:
            for i in range(1, item.Recipients.Count + 1):
                recip = item.Recipients.Item(i)
                recip_addr = getattr(recip, "Address", "")
                if recip_addr:
                    recipients.append(recip_addr)
        except Exception as e:
            logger.debug(f"Error getting recipients: {e}")
        
        # Get attachments
        attachments = []
        try:
            for i in range(1, item.Attachments.Count + 1):
                att = item.Attachments.Item(i)
                attachments.append(Attachment(
                    filename=getattr(att, "FileName", f"attachment_{i}"),
                    size=getattr(att, "Size", 0),
                    content_type=getattr(att, "ContentType", ""),
                    _com_attachment=att,
                ))
        except Exception as e:
            logger.debug(f"Error getting attachments: {e}")
        
        # Get categories
        categories = []
        try:
            cat_str = getattr(item, "Categories", "")
            if cat_str:
                categories = [c.strip() for c in cat_str.split(",")]
        except Exception:
            pass
        
        # Get importance
        importance = Importance.NORMAL
        try:
            imp_value = getattr(item, "Importance", 1)
            importance = Importance(imp_value)
        except Exception:
            pass
        
        # Parse timestamps
        received_time = None
        sent_time = None
        try:
            rt = getattr(item, "ReceivedTime", None)
            if rt:
                received_time = datetime(
                    rt.year, rt.month, rt.day,
                    rt.hour, rt.minute, rt.second
                )
            
            st = getattr(item, "SentOn", None)
            if st:
                sent_time = datetime(
                    st.year, st.month, st.day,
                    st.hour, st.minute, st.second
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
        Extract the Message-ID header from an email.
        
        Args:
            item: Outlook MailItem COM object.
            
        Returns:
            Message-ID string, or empty string if not found.
        """
        try:
            # Property tag for PR_INTERNET_MESSAGE_ID
            PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
            
            prop_accessor = item.PropertyAccessor
            message_id = prop_accessor.GetProperty(PR_INTERNET_MESSAGE_ID)
            return message_id or ""
        except Exception:
            return ""
    
    def save_email_as_msg(self, email: Email, path: Path) -> Path:
        """
        Save an email as a .msg file.
        
        Args:
            email: Email to save.
            path: Destination path (including filename).
            
        Returns:
            Path to the saved file.
        """
        return email.save_as_msg(path)
    
    def save_attachments(self, email: Email, directory: Path) -> List[Path]:
        """
        Save all attachments from an email.
        
        Args:
            email: Email with attachments.
            directory: Directory to save attachments to.
            
        Returns:
            List of paths to saved files.
        """
        saved = []
        for attachment in email.attachments:
            try:
                path = attachment.save(directory)
                saved.append(path)
                logger.debug(f"Saved attachment: {path}")
            except Exception as e:
                logger.warning(f"Failed to save attachment {attachment.filename}: {e}")
        
        return saved
