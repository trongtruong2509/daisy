"""
Outlook abstraction layer for Office Automation Foundation.

This package provides a clean interface to Outlook Desktop via COM,
hiding complexity from calling code.

Key features:
- Support for multiple accounts in one Outlook profile
- Folder navigation and email filtering
- Read email content (text, HTML, attachments)
- Send emails with safety controls
- All COM operations isolated here

Usage:
    from office.outlook import OutlookReader, OutlookSender, EmailFilter

    with OutlookReader(account="your.email@company.com") as reader:
        emails = reader.get_inbox_emails(
            filter=EmailFilter(unread_only=True, limit=100)
        )
        for email in emails:
            print(email.subject)

    # To list configured accounts without an instance:
    accounts = OutlookClient.get_available_accounts()
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
