"""
Outlook-specific module for Office Automation Foundation.

Provides abstraction over Outlook Desktop COM automation.
"""

from office.outlook.client import OutlookClient, get_outlook_accounts
from office.outlook.models import (
    Email,
    EmailFilter,
    Attachment,
    FolderInfo,
    AccountInfo,
)
from office.outlook.sender import OutlookSender
from office.outlook.exceptions import (
    OutlookError,
    OutlookConnectionError,
    OutlookAccountNotFoundError,
    OutlookFolderNotFoundError,
    OutlookSendError,
)

__all__ = [
    "OutlookClient",
    "OutlookSender",
    "get_outlook_accounts",
    "Email",
    "EmailFilter",
    "Attachment",
    "FolderInfo",
    "AccountInfo",
    "OutlookError",
    "OutlookConnectionError",
    "OutlookAccountNotFoundError",
    "OutlookFolderNotFoundError",
    "OutlookSendError",
]
