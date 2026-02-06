#!/usr/bin/env python
"""
Example script demonstrating email sending with safety controls.

This script shows how to:
- Send emails with dry-run mode
- Use state tracking to prevent duplicate sends
- Handle failures gracefully

IMPORTANT:
- Always test with DRY_RUN=true first
- Review logs before enabling actual sends
- This is a DEMONSTRATION SCRIPT

Usage:
    python scripts/example_send_email.py
"""

import sys
from pathlib import Path

# Add project root to path (for development)
project_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(project_root))

from core.config import load_config
from core.logger import setup_logging, get_logger
from core.state import ContentHashTracker
from office.outlook import OutlookSender
from office.outlook.models import NewEmail
from office.outlook.exceptions import OutlookError


def main():
    """Main entry point for the send example."""
    
    # Load configuration
    config = load_config()
    
    # Validate
    errors = config.validate()
    if errors:
        print("Configuration errors:")
        for error in errors:
            print(f"  - {error}")
        return 1
    
    # Create directories
    config.ensure_directories()
    
    # Set up logging
    log_file = setup_logging(
        log_dir=config.log_dir,
        level=config.log_level,
        run_name="example_send"
    )
    
    logger = get_logger(__name__)
    
    logger.info("=" * 60)
    logger.info("Office Automation Foundation - Send Example")
    logger.info("=" * 60)
    logger.info(f"Sending account: {config.outlook_account}")
    logger.info(f"Dry-run mode: {config.dry_run}")
    
    if not config.dry_run:
        logger.warning("LIVE MODE: Emails will actually be sent!")
    
    # Initialize state tracker for duplicate prevention
    state_tracker = ContentHashTracker(
        state_dir=config.state_dir,
        state_name="example_send"
    )
    
    try:
        with OutlookSender(
            account=config.outlook_account,
            dry_run=config.dry_run,
            state_tracker=state_tracker
        ) as sender:
            
            # Example: Create a test email
            # In a real script, you would build emails from data
            test_email = NewEmail(
                to=["test@example.com"],  # Replace with actual recipient
                subject="Test Email from Office Automation Foundation",
                body="""Hello,

This is a test email sent from the Office Automation Foundation.

If you received this email, the automation is working correctly.

Best regards,
Automation System
""",
                body_is_html=False,
            )
            
            # Validate before sending
            validation_errors = test_email.validate()
            if validation_errors:
                logger.error(f"Email validation failed: {validation_errors}")
                return 1
            
            # Check for duplicate
            if sender.is_duplicate(test_email):
                logger.warning("This email has already been sent (duplicate detected)")
                logger.info("To resend, clear the state file or use skip_duplicate_check=True")
            else:
                # Send the email
                logger.info("Sending test email...")
                result = sender.send(test_email)
                
                if result:
                    logger.info("Email sent successfully (or would be sent in dry-run)")
                else:
                    logger.warning("Email was not sent (validation failed or duplicate)")
            
            # Show statistics
            logger.info("=" * 60)
            logger.info(f"Session statistics:")
            logger.info(f"  Sent: {sender.sent_count}")
            logger.info(f"  Skipped: {sender.skipped_count}")
            logger.info(f"  Errors: {sender.error_count}")
    
    except OutlookError as e:
        logger.error(f"Outlook error: {e}")
        return 1
    
    except Exception as e:
        logger.error(f"Unexpected error: {e}", exc_info=True)
        return 1
    
    logger.info("Script completed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
