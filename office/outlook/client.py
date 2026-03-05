"""
Outlook base client — COM connection lifecycle and account lookup.

Owns the COM connection, account resolution, and shared utilities.
Not intended for direct use by tools; use OutlookReader or OutlookSender.

IMPORTANT COM CAVEATS:
- Outlook must be running for COM to work
- COM objects must be accessed from the same thread that created them
- Some operations may fail if Outlook is "busy" (retry logic helps)

Usage:
    # For reading emails, use OutlookReader
    from office.outlook.reader import OutlookReader

    with OutlookReader(account="your.email@company.com") as reader:
        emails = reader.get_inbox_emails()

    # For sending emails, use OutlookSender
    from office.outlook.sender import OutlookSender

    with OutlookSender(account="your.email@company.com", dry_run=True) as sender:
        sender.send(email)

    # To list accounts without an instance
    accounts = OutlookClient.get_available_accounts()
"""

import logging
from typing import List, Optional

from office.utils.com import (
    com_initialized,
    ensure_com_available,
    is_available,
    get_pywintypes,
)
from office.utils.helpers import get_or_create_outlook, safe_quit_outlook
from core.retry import retry_operation, RetryConfig
from office.outlook.exceptions import (
    OutlookConnectionError,
    OutlookAccountNotFoundError,
)
from office.outlook.models import AccountInfo

logger = logging.getLogger(__name__)


class OutlookClient:
    """
    Base class owning the Outlook COM connection lifecycle.

    Handles COM initialization, account lookup via SMTP address, and cleanup.
    Subclass this \u2014 do not use directly in tools.

    Attributes:
        account: SMTP address of the account to connect to.
        retry_config: Configuration for retry behaviour on COM errors.
    """

    def __init__(
        self,
        account: str,
        retry_config: Optional[RetryConfig] = None,
    ) -> None:
        """
        Initialize the client.

        Args:
            account: SMTP address of the Outlook account to use.
            retry_config: Retry configuration (uses sensible defaults if None).

        Raises:
            ImportError: If win32com is not available (non-Windows).
        """
        ensure_com_available()

        self.account = account
        self.retry_config = retry_config or RetryConfig(
            max_attempts=3,
            base_delay=2.0,
            max_delay=30.0,
        )

        self._outlook: Optional[object] = None
        self._namespace: Optional[object] = None
        self._account_obj: Optional[object] = None
        self._connected = False
        self._was_already_running = False
        self._com_ctx = None

    # ── Context manager ──────────────────────────────────────────

    def __enter__(self) -> "OutlookClient":
        """Connect to Outlook on context entry."""
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> bool:
        """Disconnect and clean up on context exit."""
        self.disconnect()
        return False

    # ── Public API ───────────────────────────────────────────────

    @property
    def is_connected(self) -> bool:
        """Return True if the client is currently connected."""
        return self._connected

    @retry_operation()
    def connect(self) -> None:
        """
        Connect to Outlook and resolve the specified account.

        Calls CoInitialize() for the current thread before dispatching,
        then locates the account object via SMTP address.

        Raises:
            OutlookConnectionError: If the COM dispatch fails.
            OutlookAccountNotFoundError: If the account is not in the profile.
        """
        if self._connected:
            return

        logger.debug("[ACCOUNT-CONNECT] Initialising COM and connecting to Outlook...")

        try:
            self._com_ctx = com_initialized()
            self._com_ctx.__enter__()

            pywintypes = get_pywintypes()
            self._outlook, self._was_already_running = get_or_create_outlook()
            self._namespace = self._outlook.GetNamespace("MAPI")
        except Exception as e:
            if hasattr(e, '__class__') and 'com_error' in type(e).__name__:
                raise OutlookConnectionError(
                    f"Failed to connect to Outlook: {e}. Is Outlook running?"
                )
            raise

        self._account_obj = self._find_account()
        self._connected = True

        found_smtp = getattr(self._account_obj, "SmtpAddress", self.account)
        logger.info(f"[ACCOUNT-FOUND] Connected to Outlook account: {found_smtp}")

    def disconnect(self) -> None:
        """Disconnect from Outlook and release COM resources."""
        safe_quit_outlook(self._outlook, self._was_already_running)
        self._outlook = None
        self._namespace = None
        self._account_obj = None
        self._connected = False

        if self._com_ctx:
            self._com_ctx.__exit__(None, None, None)
            self._com_ctx = None

        logger.debug("[ACCOUNT-CONNECT] Disconnected from Outlook")

    @staticmethod
    def get_available_accounts() -> List[str]:
        """
        Return SMTP addresses of all accounts configured in the Outlook profile.

        Opens a temporary COM connection (no instance required). Always wraps
        COM initialization in CoInitialize/CoUninitialize so it is safe to
        call from any thread or at startup before a context manager is used.

        Returns:
            List of email addresses (SMTP, deduplicated). Returns empty list
            if Outlook is unavailable.
        """
        if not is_available():
            logger.debug("win32com not available \u2014 cannot retrieve Outlook accounts")
            return []

        try:
            with com_initialized():
                outlook, was_running = get_or_create_outlook()
                try:
                    ns = outlook.GetNamespace("MAPI")

                    accounts: List[str] = []
                    seen: set = set()

                    for i in range(1, ns.Accounts.Count + 1):
                        acc = ns.Accounts.Item(i)
                        smtp = getattr(acc, "SmtpAddress", "") or ""
                        if smtp and smtp not in seen:
                            accounts.append(smtp)
                            seen.add(smtp)

                    logger.debug(f"[ACCOUNT-SEARCH] Found {len(accounts)} Outlook account(s)")
                    return accounts
                finally:
                    safe_quit_outlook(outlook, was_running)

        except Exception as e:
            logger.warning(f"Could not retrieve Outlook accounts: {e}")
            return []

    # ── Internal helpers ─────────────────────────────────────────

    def _ensure_connected(self) -> None:
        """Auto-connect if not already connected."""
        if not self._connected:
            self.connect()

    def _find_account(self) -> object:
        """
        Find the Outlook Account COM object by SMTP address.

        Uses SMTP-only matching \u2014 no display-name or folder-name fallbacks.

        Returns:
            The Account COM object.

        Raises:
            OutlookAccountNotFoundError: If the account is not found.
        """
        available: List[str] = []

        logger.info(f"[ACCOUNT-SEARCH] Looking for account: {self.account}")

        try:
            total = self._namespace.Accounts.Count
            logger.info(f"[ACCOUNT-SEARCH] Total accounts in Outlook: {total}")

            for i in range(1, total + 1):
                acc = self._namespace.Accounts.Item(i)
                smtp = getattr(acc, "SmtpAddress", "") or ""
                available.append(smtp)

                match_marker = "\u2713 MATCH" if smtp.lower() == self.account.lower() else ""
                logger.info(f"[ACCOUNT-SEARCH] Account [{i}] = '{smtp}' {match_marker}")

                if smtp.lower() == self.account.lower():
                    logger.info(f"[ACCOUNT-SELECTED] Found account: {smtp}")
                    return acc

        except Exception as e:
            logger.error(f"[ACCOUNT-ERROR] Error searching accounts: {e}", exc_info=True)

        logger.error(
            f"[ACCOUNT-ERROR] Account '{self.account}' not found. "
            f"Available: {', '.join(available)}"
        )
        raise OutlookAccountNotFoundError(self.account, available)
    
