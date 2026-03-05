"""
Outlook-specific module for Office Automation Foundation.

Provides abstraction over Outlook Desktop COM automation.
"""

from office.outlook.client import OutlookClient
from office.outlook.reader import OutlookReader
from office.outlook.models import (
    Email,
    EmailFilter,
    Attachment,
    FolderInfo,
    AccountInfo,
    NewEmail,
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
    "OutlookReader",
    "OutlookSender",
    "Email",
    "EmailFilter",
    "Attachment",
    "FolderInfo",
    "AccountInfo",
    "NewEmail",
    "OutlookError",
    "OutlookConnectionError",
    "OutlookAccountNotFoundError",
    "OutlookFolderNotFoundError",
    "OutlookSendError",
]
