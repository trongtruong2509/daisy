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
    from office.outlook import OutlookClient, EmailFilter
    
    with OutlookClient(account="your.email@company.com") as client:
        emails = client.get_inbox_emails(
            filter=EmailFilter(unread_only=True, limit=100)
        )
        for email in emails:
            print(email.subject)
"""

from office.outlook.client import OutlookClient
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
