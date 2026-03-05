"""
Custom exceptions for Outlook operations.

All Outlook-specific errors inherit from OutlookError for easy catching.
"""


class OutlookError(Exception):
    """Base exception for all Outlook-related errors."""
    pass


class OutlookConnectionError(OutlookError):
    """Raised when unable to connect to Outlook application."""
    
    def __init__(self, message: str = "Failed to connect to Outlook"):
        super().__init__(message)


class OutlookAccountNotFoundError(OutlookError):
    """Raised when the specified account is not found in Outlook profile."""
    
    def __init__(self, account: str, available_accounts: list[str] = None):
        self.account = account
        self.available_accounts = available_accounts or []
        
        message = f"Account '{account}' not found in Outlook profile"
        if self.available_accounts:
            message += f". Available accounts: {', '.join(self.available_accounts)}"
        
        super().__init__(message)


class OutlookFolderNotFoundError(OutlookError):
    """Raised when the specified folder is not found."""
    
    def __init__(self, folder_path: str, account: str = None):
        self.folder_path = folder_path
        self.account = account
        
        message = f"Folder '{folder_path}' not found"
        if account:
            message += f" in account '{account}'"
        
        super().__init__(message)


class OutlookSendError(OutlookError):
    """Raised when sending an email fails."""
    
    def __init__(self, recipient: str, subject: str, reason: str = None):
        self.recipient = recipient
        self.subject = subject
        self.reason = reason
        
        message = f"Failed to send email to '{recipient}' with subject '{subject}'"
        if reason:
            message += f": {reason}"
        
        super().__init__(message)


class OutlookItemError(OutlookError):
    """Raised when there's an error accessing or processing an email item."""
    
    def __init__(self, item_id: str = None, reason: str = None):
        self.item_id = item_id
        self.reason = reason
        
        message = "Error processing email item"
        if item_id:
            message += f" (ID: {item_id})"
        if reason:
            message += f": {reason}"
        
        super().__init__(message)


