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
from typing import List, Optional, Tuple

from core.retry import retry_operation, RetryConfig
from core.state import StateTracker
from office.outlook.client import OutlookClient
from office.outlook.exceptions import (
    OutlookSendError)
from office.outlook.models import NewEmail

logger = logging.getLogger(__name__)


class OutlookSender(OutlookClient):
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
        super().__init__(account, retry_config)
        self.dry_run = dry_run
        self.state_tracker = state_tracker

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
        super().__exit__(exc_type, exc_val, exc_tb)

        # Save state tracker if present
        if self.state_tracker:
            self.state_tracker.save()
        
        # Log summary
        logger.info(
            f"Send session complete - Sent: {self.sent_count}, "
            f"Skipped: {self.skipped_count}, Errors: {self.error_count}"
        )
        
        return False
    
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
    
    def _log_item_location(self, mail: object, step: str) -> None:
        """
        Debug-log the mail item's current parent folder path.

        Helps pinpoint which COM operation triggers an account re-bind on Exchange.
        Guarded by isEnabledFor(DEBUG) so COM accessors are never invoked in
        production when debug logging is off.
        """
        if not logger.isEnabledFor(logging.DEBUG):
            return
        try:
            parent = getattr(mail, "Parent", None)
            path = parent.FolderPath if parent else "(no parent yet)"
            logger.debug(f"[SEND-LOCATION:{step}] parent='{path}'")
        except Exception as e:
            logger.debug(f"[SEND-LOCATION:{step}] could not read parent: {e}")

    @retry_operation()
    def _do_send(self, email: NewEmail) -> None:
        """
        Actually send the email via COM.

        Account binding strategy (account-type-aware):

        Exchange (AccountType == 0):
            Proven failures on Exchange shared/delegate mailboxes:
              - DeliveryStore.Outbox.Items.Add(0): item silently lands in
                default account's Outbox (confirmed via SEND-LOCATION logs).
              - CreateItem + SendUsingAccount (set early): Exchange ignores it;
                email sends from default account.
            Working approach:
              1. CreateItem(0) — in-memory item, no store association.
              2. Set ALL content (recipients, subject, body, attachments).
              3. Set SendUsingAccount LAST (right before Send) — some Exchange
                 configs reset it if content is modified after.
              4. Set SentOnBehalfOfName — the standard Exchange mechanism for
                 shared/functional mailboxes. With "Send As" permission the
                 recipient sees VN.Phuclong as the sender; with "Send on Behalf"
                 permission they see "Xuan.Truong on behalf of VN.Phuclong".

        IMAP/Gmail (AccountType != 0):
            SendUsingAccount silently reverts to None for IMAP accounts.
            The only reliable routing signal is Outbox folder ownership:
                account.DeliveryStore.GetDefaultFolder(4).Items.Add(0)

        Neither strategy calls ResolveAll() — on Exchange it triggers a MAPI
        transport re-bind that moves the item to the default account's Outbox.
        """
        account_smtp = getattr(self._account_obj, "SmtpAddress", "UNKNOWN")
        logger.debug(f"[SEND] Using account '{account_smtp}' to send to {email.to}")

        # Verify the account COM object is still valid
        try:
            _ = self._account_obj.SmtpAddress
        except Exception as e:
            logger.error(f"[SEND-METHOD] Account object invalid: {e}")
            raise

        # ── Account-type-aware item creation ──────────────────────────────
        #
        # OlAccountType: 0=olExchange, 1=olImap, 2=olPop3, 3=olHttp, 4=olEas
        #
        # Exchange: NOTHING store-based works for non-default accounts.
        #   - DeliveryStore Outbox Items.Add(0) → item in default Outbox
        #   - SendUsingAccount set early → ignored after content changes
        #   Fix: CreateItem(0), set ALL content, then set both
        #   SendUsingAccount + SentOnBehalfOfName right before Send().
        #
        # IMAP/Gmail: SendUsingAccount silently reverts to None.
        #   Fix: DeliveryStore Outbox folder ownership is the routing signal.
        #
        mail = None
        expected_outbox_path = None
        used_delivery_store = False

        account_type = getattr(self._account_obj, "AccountType", None)
        is_exchange = (account_type == 0)  # 0 = olExchange
        logger.debug(f"[SEND-STEP-1] AccountType={account_type} ({'Exchange' if is_exchange else 'IMAP/other'})")

        if is_exchange:
            # Exchange: CreateItem — set account later (after all content)
            mail = self._outlook.CreateItem(0)  # 0 = olMailItem
            logger.debug(f"[SEND-STEP-1] ✓ Exchange: CreateItem(0) — account binding deferred to after content")
        else:
            # IMAP/Gmail: DeliveryStore Outbox folder ownership
            try:
                store = self._account_obj.DeliveryStore
                outbox = store.GetDefaultFolder(4)  # 4 = olFolderOutbox
                expected_outbox_path = getattr(outbox, "FolderPath", None)
                mail = outbox.Items.Add(0)           # 0 = olMailItem
                used_delivery_store = True
                logger.debug(f"[SEND-STEP-1] ✓ IMAP: Created item in account outbox: {expected_outbox_path}")
            except Exception as e:
                logger.warning(
                    f"[SEND-STEP-1] IMAP: DeliveryStore/Outbox unavailable ({e}), "
                    f"falling back to CreateItem — routing may fail"
                )
                mail = self._outlook.CreateItem(0)
                # IMAP fallback — best effort
                try:
                    mail.SendUsingAccount = self._account_obj
                    logger.debug(f"[SEND-STEP-1] IMAP fallback: Set SendUsingAccount: {account_smtp}")
                except Exception as e2:
                    logger.warning(f"[SEND-STEP-1] IMAP fallback: SendUsingAccount failed: {e2}")

        self._log_item_location(mail, "after-create")

        # ── Recipients (NO ResolveAll — triggers MAPI rebind on Exchange) ─
        for recipient in email.to:
            mail.Recipients.Add(recipient)
        for recipient in email.cc:
            recip = mail.Recipients.Add(recipient)
            recip.Type = 2  # CC
        for recipient in email.bcc:
            recip = mail.Recipients.Add(recipient)
            recip.Type = 3  # BCC

        self._log_item_location(mail, "after-recipients")

        # ── Content ───────────────────────────────────────────────────────
        mail.Subject = email.subject

        if email.body_is_html:
            mail.HTMLBody = email.body
        else:
            mail.Body = email.body

        mail.Importance = email.importance.value

        for attachment_path in email.attachments:
            mail.Attachments.Add(str(attachment_path))

        self._log_item_location(mail, "after-content")

        # ── Exchange: bind account LAST (after all content is set) ────────
        #
        # On Exchange shared/delegate mailboxes, SendUsingAccount set early
        # gets ignored (email sends from default account). Setting it LAST,
        # combined with SentOnBehalfOfName, is the reliable approach.
        #
        if is_exchange:
            # SendUsingAccount — set as late as possible
            try:
                mail.SendUsingAccount = self._account_obj
                # Verify it stuck by reading back
                readback = getattr(mail, "SendUsingAccount", None)
                readback_smtp = getattr(readback, "SmtpAddress", "None") if readback else "None"
                logger.debug(f"[SEND-STEP-2] Exchange: SendUsingAccount={account_smtp}, readback={readback_smtp}")
            except Exception as e:
                logger.warning(f"[SEND-STEP-2] Exchange: SendUsingAccount failed: {e}")

            # SentOnBehalfOfName — the standard Exchange mechanism for
            # shared/functional mailboxes. This directly sets the "From" field.
            # With "Send As" permission → appears as VN.Phuclong (no "on behalf of").
            # With "Send on Behalf" permission → appears as "X on behalf of Y".
            try:
                mail.SentOnBehalfOfName = account_smtp
                logger.debug(f"[SEND-STEP-2] Exchange: SentOnBehalfOfName={account_smtp}")
            except Exception as e:
                logger.debug(f"[SEND-STEP-2] Exchange: SentOnBehalfOfName failed (OK if not delegate): {e}")

            self._log_item_location(mail, "after-account-bind")
        else:
            logger.debug(f"[SEND-STEP-2] IMAP: account binding via {'DeliveryStore Outbox' if used_delivery_store else 'fallback'}")

        # ── IMAP verification (Exchange items live in default Outbox, skip) ─
        if expected_outbox_path:
            try:
                parent = getattr(mail, "Parent", None)
                parent_path = parent.FolderPath if parent else None
                if parent_path:
                    if parent_path == expected_outbox_path:
                        logger.debug(f"[SEND-VERIFY] ✓ Item still in correct outbox: {parent_path}")
                    else:
                        logger.warning(
                            f"[SEND-VERIFY] ✗ Item moved! "
                            f"Expected='{expected_outbox_path}' → Current='{parent_path}' "
                            f"— email will send from wrong account"
                        )
                else:
                    logger.debug("[SEND-VERIFY] Parent path unavailable (unsaved item) — skipping check")
            except Exception as e:
                logger.debug(f"[SEND-VERIFY] Could not read parent folder: {e}")

        # ── Send ──────────────────────────────────────────────────────────
        logger.debug(f"[SEND-STEP-3] Calling mail.Send() for account '{account_smtp}'")
        mail.Send()
        logger.debug(f"[SEND-STEP-3] mail.Send() returned successfully")
    
    def send_batch(
        self,
        emails: List[NewEmail],
        continue_on_error: bool = True
    ) -> Tuple[int, int, int]:
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
            account_smtp = getattr(self._account_obj, "SmtpAddress", "UNKNOWN")

            # Account-type-aware: same strategy as _do_send.
            # Exchange: CreateItem, set content, then bind account LAST.
            # IMAP: DeliveryStore Drafts folder ownership.
            mail = None
            used_delivery_store = False
            account_type = getattr(self._account_obj, "AccountType", None)
            is_exchange = (account_type == 0)  # 0=olExchange

            if is_exchange:
                mail = self._outlook.CreateItem(0)
                logger.debug(f"[DRAFT] Exchange: CreateItem(0) — account binding deferred")
            else:
                try:
                    store = self._account_obj.DeliveryStore
                    drafts = store.GetDefaultFolder(16)  # 16 = olFolderDrafts
                    mail = drafts.Items.Add(0)            # 0 = olMailItem
                    used_delivery_store = True
                    logger.debug(
                        f"[DRAFT] IMAP: Created item in account drafts: "
                        f"{getattr(drafts, 'FolderPath', '?')}"
                    )
                except Exception as e:
                    logger.warning(f"[DRAFT] IMAP: DeliveryStore/Drafts unavailable ({e}), falling back to CreateItem")
                    mail = self._outlook.CreateItem(0)

            # IMAP fallback only
            if not is_exchange and not used_delivery_store:
                try:
                    mail.SendUsingAccount = self._account_obj
                except Exception:
                    pass

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

            # Exchange: bind account LAST (after all content), same as _do_send
            if is_exchange:
                try:
                    mail.SendUsingAccount = self._account_obj
                except Exception:
                    pass
                try:
                    mail.SentOnBehalfOfName = account_smtp
                except Exception:
                    pass

            mail.Save()  # Save as draft
            
            logger.info(f"Draft created - Subject: {email.subject}")
            return True
        
        except Exception as e:
            logger.error(f"Failed to create draft: {e}")
            return False
