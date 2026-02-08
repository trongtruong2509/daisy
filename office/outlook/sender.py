"""
Outlook email sender with safety controls.

Provides safe email sending with:
- Dry-run mode support
- Duplicate send protection
- Retry logic
- Comprehensive logging

IMPORTANT SAFETY NOTES:
- Always test with DRY_RUN=true first
- All send operations are logged for audit
- Used ContentHashTracker to prevent duplicate sends
- Validate recipients before sending

Usage:
    from office.outlook import OutlookSender
    from office.outlook.models import NewEmail
    from core.state import ContentHashTracker
    
    tracker = ContentHashTracker(state_dir, "email_send")
    
    with OutlookSender(
        account="your.email@company.com",
        dry_run=True,
        state_tracker=tracker
    ) as sender:
        email = NewEmail(
            to=["recipient@example.com"],
            subject="Test",
            body="Hello world"
        )
        sender.send(email)
"""

import logging
from pathlib import Path
from typing import List, Optional

# COM imports
try:
    import win32com.client
    import pywintypes
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False
    pythoncom = None

from core.retry import retry_operation, RetryConfig
from core.state import ContentHashTracker, StateTracker
from office.outlook.exceptions import (
    OutlookConnectionError,
    OutlookAccountNotFoundError,
    OutlookSendError,
    DryRunBlockedError,
    OutlookError,
)
from office.outlook.models import NewEmail, Importance

logger = logging.getLogger(__name__)


class OutlookSender:
    """
    Outlook email sender with safety controls.
    
    Provides:
    - Dry-run mode: logs what would be sent without actually sending
    - Duplicate prevention: uses state tracker to prevent repeat sends
    - Retry logic: handles transient COM errors
    - Audit logging: all operations are logged
    
    Use as a context manager:
    
        with OutlookSender(account="user@example.com", dry_run=True) as sender:
            sender.send(email)
    
    Attributes:
        account: SMTP address of the sending account.
        dry_run: If True, no emails are actually sent.
        state_tracker: Optional tracker for duplicate prevention.
    """
    
    def __init__(
        self,
        account: str,
        dry_run: bool = True,
        state_tracker: Optional[StateTracker] = None,
        retry_config: Optional[RetryConfig] = None
    ):
        """
        Initialize the sender.
        
        Args:
            account: SMTP address of the account to send from.
            dry_run: If True, don't actually send emails (safe mode).
            state_tracker: Optional state tracker for duplicate prevention.
            retry_config: Retry configuration for COM operations.
        """
        if not HAS_WIN32COM:
            raise ImportError(
                "win32com is required for Outlook operations. "
                "Install pywin32: pip install pywin32"
            )
        
        self.account = account
        self.dry_run = dry_run
        self.state_tracker = state_tracker
        self.retry_config = retry_config or RetryConfig(
            max_attempts=3,
            base_delay=2.0
        )
        
        self._outlook: Optional[object] = None
        self._namespace: Optional[object] = None
        self._account_obj: Optional[object] = None
        self._connected = False
        
        # Statistics
        self.sent_count = 0
        self.skipped_count = 0
        self.error_count = 0
        
        if dry_run:
            logger.warning("DRY-RUN MODE: No emails will actually be sent")
    
    def __enter__(self) -> "OutlookSender":
        """Connect on context entry."""
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb) -> bool:
        """Clean up on context exit."""
        self.disconnect()
        
        # Save state tracker if present
        if self.state_tracker:
            self.state_tracker.save()
        
        # Log summary
        logger.info(
            f"Send session complete - Sent: {self.sent_count}, "
            f"Skipped: {self.skipped_count}, Errors: {self.error_count}"
        )
        
        return False
    
    @retry_operation()
    def connect(self) -> None:
        """
        Connect to Outlook.
        
        Raises:
            OutlookConnectionError: If unable to connect.
            OutlookAccountNotFoundError: If account not found.
        """
        if self._connected:
            return
        
        logger.info(f"[ACCOUNT-LOOKUP] Looking for account: {self.account}")
        logger.debug("Connecting to Outlook...")
        
        try:
            # Initialize COM for this thread
            if pythoncom:
                pythoncom.CoInitialize()
            
            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook.GetNamespace("MAPI")
        except pywintypes.com_error as e:
            raise OutlookConnectionError(
                f"Failed to connect to Outlook: {e}"
            )
        
        # Find the account
        self._account_obj = self._find_account()
        found_smtp = getattr(self._account_obj, "SmtpAddress", "UNKNOWN")
        self._connected = True
        
        logger.info(f"[ACCOUNT-FOUND] ✓ Successfully connected with account: {found_smtp}")
    
    def _find_account(self) -> object:
        """
        Find the specified account in Outlook.
        
        Returns:
            Account COM object.
            
        Raises:
            OutlookAccountNotFoundError: If account not found.
        """
        available = []
        
        try:
            total_accounts = self._namespace.Accounts.Count
            logger.info(f"[ACCOUNT-SEARCH] Total accounts in Outlook: {total_accounts}")
            
            for i in range(1, total_accounts + 1):
                acc = self._namespace.Accounts.Item(i)
                smtp = getattr(acc, "SmtpAddress", "")
                available.append(smtp)
                
                match = "✓ MATCH" if smtp.lower() == self.account.lower() else ""
                logger.info(f"[ACCOUNT-SEARCH] Account [{i}] = '{smtp}' {match}")
                
                if smtp.lower() == self.account.lower():
                    logger.info(f"[ACCOUNT-SELECTED] ✓ Found requested account: {smtp}")
                    return acc
            
            # No match found
            logger.error(
                f"[ACCOUNT-ERROR] Account '{self.account}' not found in Outlook.\n"
                f"[ACCOUNT-ERROR] Available accounts: {', '.join(available)}"
            )
        except Exception as e:
            logger.error(f"[ACCOUNT-ERROR] Error finding account: {e}", exc_info=True)
        
        raise OutlookAccountNotFoundError(self.account, available)
    
    def disconnect(self) -> None:
        """Disconnect from Outlook."""
        self._outlook = None
        self._namespace = None
        self._account_obj = None
        self._connected = False
        logger.debug("Disconnected from Outlook")
    
    def _ensure_connected(self) -> None:
        """Ensure connected to Outlook."""
        if not self._connected:
            self.connect()
    
    def is_duplicate(self, email: NewEmail) -> bool:
        """
        Check if an email would be a duplicate send.
        
        Args:
            email: Email to check.
            
        Returns:
            True if this email was already sent.
        """
        if not self.state_tracker:
            return False
        
        content_hash = email.get_content_hash()
        return self.state_tracker.is_processed(content_hash)
    
    def send(self, email: NewEmail, skip_duplicate_check: bool = False) -> bool:
        """
        Send an email.
        
        Args:
            email: Email to send.
            skip_duplicate_check: If True, skip duplicate detection.
            
        Returns:
            True if sent successfully (or would be sent in dry-run).
            False if skipped (duplicate or validation failed).
            
        Raises:
            OutlookSendError: If sending fails.
        """
        self._ensure_connected()
        
        # Validate
        validation_errors = email.validate()
        if validation_errors:
            logger.error(f"Email validation failed: {validation_errors}")
            self.error_count += 1
            return False
        
        recipients_str = ", ".join(email.to[:3])
        if len(email.to) > 3:
            recipients_str += f" (+{len(email.to) - 3} more)"
        
        # Check for duplicates
        if not skip_duplicate_check and self.is_duplicate(email):
            logger.warning(
                f"[SKIP] Duplicate email detected - To: {recipients_str}, "
                f"Subject: {email.subject}"
            )
            self.skipped_count += 1
            return False
        
        # Dry-run mode
        if self.dry_run:
            logger.info(
                f"[DRY-RUN] Would send email - To: {recipients_str}, "
                f"Subject: {email.subject}"
            )
            
            # Do NOT track dry-run emails in state — they weren't actually sent
            # Tracking them would cause false "duplicate" detection on real runs
            
            self.sent_count += 1
            return True
        
        # Actually send
        try:
            self._do_send(email)
            
            logger.info(
                f"[SENT] Email sent - To: {recipients_str}, "
                f"Subject: {email.subject}"
            )
            
            # Track as sent
            if self.state_tracker:
                self.state_tracker.mark_processed(
                    email.get_content_hash(),
                    metadata={
                        "to": email.to,
                        "subject": email.subject,
                        "dry_run": False,
                    }
                )
            
            self.sent_count += 1
            return True
        
        except Exception as e:
            logger.error(
                f"[ERROR] Failed to send email - To: {recipients_str}, "
                f"Subject: {email.subject}, Error: {e}"
            )
            self.error_count += 1
            raise OutlookSendError(
                recipients_str,
                email.subject,
                str(e)
            )
    
    @retry_operation()
    def _do_send(self, email: NewEmail) -> None:
        """
        Actually send the email via COM.
        
        Args:
            email: Email to send.
        """
        # Log which account is sending this email
        account_smtp = getattr(self._account_obj, "SmtpAddress", "UNKNOWN")
        logger.debug(f"[SEND] Using account '{account_smtp}' to send to {email.to}")
        
        # Verify account object is valid
        try:
            account_email = self._account_obj.SmtpAddress
            logger.debug(f"[SEND-METHOD] Account object valid: {account_email}")
        except Exception as e:
            logger.error(f"[SEND-METHOD] Account object invalid: {e}")
            raise
        
        # Create mail item (try account's Outbox first)
        try:
            store = self._account_obj.DeliveryStore
            outbox = store.GetDefaultFolder(4)  # 4 = olFolderOutbox
            mail = outbox.Items.Add(0)  # 0 = olMailItem
            logger.debug(f"[SEND-METHOD] Created mail item in account's Outbox folder")
        except Exception as e:
            logger.warning(f"[SEND-METHOD] Could not create mail in account store: {e}")
            logger.info(f"[SEND-METHOD] Falling back to CreateItem")
            mail = self._outlook.CreateItem(0)  # 0 = olMailItem
        
        # Set recipients FIRST (some COM properties require recipients to be set first)
        for recipient in email.to:
            mail.Recipients.Add(recipient)
        
        for recipient in email.cc:
            recip = mail.Recipients.Add(recipient)
            recip.Type = 2  # CC
        
        for recipient in email.bcc:
            recip = mail.Recipients.Add(recipient)
            recip.Type = 3  # BCC
        
        # Resolve all recipients
        mail.Recipients.ResolveAll()
        
        # NOW set SendUsingAccount AFTER recipients are resolved
        try:
            mail.SendUsingAccount = self._account_obj
            logger.debug(f"[SEND-METHOD] Set SendUsingAccount property (after recipient setup)")
        except Exception as e:
            logger.error(f"[SEND-METHOD] Failed to set SendUsingAccount: {e}")
        
        # Verify which account will be used
        try:
            if mail.SendUsingAccount is not None:
                actual_account = mail.SendUsingAccount.SmtpAddress
                logger.info(f"[SEND-VERIFY] Mail item's SendUsingAccount = '{actual_account}'")
            else:
                logger.debug(f"[SEND-VERIFY] Mail item's SendUsingAccount is None")
        except Exception as e:
            logger.debug(f"[SEND-VERIFY] Could not verify SendUsingAccount: {e}")
        
        # Set content
        mail.Subject = email.subject
        
        if email.body_is_html:
            mail.HTMLBody = email.body
        else:
            mail.Body = email.body
        
        # Set importance
        mail.Importance = email.importance.value
        
        # Add attachments
        for attachment_path in email.attachments:
            mail.Attachments.Add(str(attachment_path))
        
        # Send
        mail.Send()
    
    def send_batch(
        self,
        emails: List[NewEmail],
        continue_on_error: bool = True
    ) -> tuple[int, int, int]:
        """
        Send multiple emails.
        
        Args:
            emails: List of emails to send.
            continue_on_error: If True, continue sending after errors.
            
        Returns:
            Tuple of (sent_count, skipped_count, error_count) for this batch.
        """
        batch_sent = 0
        batch_skipped = 0
        batch_errors = 0
        
        for i, email in enumerate(emails, 1):
            logger.debug(f"Processing email {i}/{len(emails)}")
            
            try:
                result = self.send(email)
                if result:
                    batch_sent += 1
                else:
                    batch_skipped += 1
            
            except OutlookSendError as e:
                batch_errors += 1
                if not continue_on_error:
                    raise
        
        logger.info(
            f"Batch complete - Sent: {batch_sent}, "
            f"Skipped: {batch_skipped}, Errors: {batch_errors}"
        )
        
        return batch_sent, batch_skipped, batch_errors
    
    def create_draft(self, email: NewEmail) -> bool:
        """
        Create a draft email (not sent).
        
        Useful for review before sending.
        
        Args:
            email: Email to create as draft.
            
        Returns:
            True if draft created successfully.
        """
        self._ensure_connected()
        
        if self.dry_run:
            logger.info(
                f"[DRY-RUN] Would create draft - Subject: {email.subject}"
            )
            return True
        
        try:
            mail = self._outlook.CreateItem(0)
            mail.SendUsingAccount = self._account_obj
            
            for recipient in email.to:
                mail.Recipients.Add(recipient)
            
            for recipient in email.cc:
                recip = mail.Recipients.Add(recipient)
                recip.Type = 2
            
            mail.Subject = email.subject
            
            if email.body_is_html:
                mail.HTMLBody = email.body
            else:
                mail.Body = email.body
            
            for attachment_path in email.attachments:
                mail.Attachments.Add(str(attachment_path))
            
            mail.Save()  # Save as draft
            
            logger.info(f"Draft created - Subject: {email.subject}")
            return True
        
        except Exception as e:
            logger.error(f"Failed to create draft: {e}")
            return False
