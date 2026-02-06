#!/usr/bin/env python
"""
Example script demonstrating the Office Automation Foundation.

This script shows how to:
- Load configuration
- Set up logging
- Connect to Outlook
- Read emails with filtering
- Save emails and attachments
- Use state tracking

This is a DEMONSTRATION SCRIPT, not a production tool.
Modify and extend for your specific needs.

Usage:
    python scripts/example_read_emails.py

Requirements:
    - .env file configured with OUTLOOK_ACCOUNT
    - Outlook Desktop running
    - Windows environment
"""

import sys
from pathlib import Path

# Add project root to path (for development)
project_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(project_root))

from core.config import load_config
from core.logger import setup_logging, get_logger, ProgressLogger
from core.state import StateTracker
from office.outlook import OutlookClient, EmailFilter
from office.outlook.exceptions import OutlookError


def main():
    """Main entry point for the example script."""
    
    # Load configuration
    config = load_config()
    
    # Validate configuration
    errors = config.validate()
    if errors:
        print("Configuration errors:")
        for error in errors:
            print(f"  - {error}")
        print("\nPlease check your .env file.")
        return 1
    
    # Create directories
    config.ensure_directories()
    
    # Set up logging
    log_file = setup_logging(
        log_dir=config.log_dir,
        level=config.log_level,
        run_name="example_read"
    )
    
    logger = get_logger(__name__)
    
    # Log startup info
    logger.info("=" * 60)
    logger.info("Office Automation Foundation - Example Script")
    logger.info("=" * 60)
    logger.info(f"Account: {config.outlook_account}")
    logger.info(f"Dry-run mode: {config.dry_run}")
    logger.info(f"Batch size: {config.batch_size}")
    logger.info(f"Log file: {log_file}")
    
    if config.dry_run:
        logger.warning("DRY-RUN MODE: No mutations will be made")
    
    # Initialize state tracker (for demonstrating duplicate detection)
    state_tracker = StateTracker(
        state_dir=config.state_dir,
        state_name="example_read"
    )
    
    logger.info(f"Previously processed emails: {state_tracker.get_processed_count()}")
    
    try:
        # Connect to Outlook
        logger.info("Connecting to Outlook...")
        
        with OutlookClient(account=config.outlook_account) as client:
            # Show available accounts
            accounts = client.get_available_accounts()
            logger.info(f"Available accounts: {len(accounts)}")
            for acc in accounts:
                logger.debug(f"  - {acc.smtp_address}")
            
            # Get inbox information
            inbox = client.get_inbox()
            inbox_info = client.get_folder_info(inbox)
            logger.info(f"Inbox: {inbox_info.item_count} emails, {inbox_info.unread_count} unread")
            
            # Create filter for recent emails
            filter = EmailFilter(
                unread_only=False,  # Get all emails, not just unread
                limit=config.batch_size
            )
            
            # Read emails
            logger.info(f"Reading up to {filter.limit} emails...")
            emails = client.get_inbox_emails(filter=filter)
            
            logger.info(f"Retrieved {len(emails)} emails")
            
            # Process emails (demonstration)
            progress = ProgressLogger(
                total=len(emails),
                logger=logger,
                operation="Processing emails",
                log_every=10
            )
            
            new_emails = 0
            skipped_emails = 0
            
            for i, email in enumerate(emails):
                # Check if already processed
                if state_tracker.is_processed(email.unique_id):
                    logger.debug(f"Skipping already processed: {email.subject}")
                    skipped_emails += 1
                    progress.update(i + 1)
                    continue
                
                # Log email details
                logger.debug(
                    f"Email: {email.subject} | From: {email.sender_name} | "
                    f"Date: {email.received_time}"
                )
                
                # Here you would add your business logic:
                # - Parse the email body
                # - Extract data
                # - Take action
                
                # Example: Save emails with attachments
                if email.has_attachments:
                    logger.info(f"Email with attachments: {email.subject}")
                    for att in email.attachments:
                        logger.debug(f"  Attachment: {att.filename} ({att.size} bytes)")
                    
                    # In a real script, you might save attachments:
                    # if not config.dry_run:
                    #     client.save_attachments(email, config.output_dir / "attachments")
                
                # Mark as processed
                state_tracker.mark_processed(
                    email.unique_id,
                    metadata={
                        "subject": email.subject,
                        "sender": email.sender_address,
                    }
                )
                new_emails += 1
                progress.update(i + 1)
            
            progress.complete()
            
            # Save state
            state_tracker.save()
            
            # Summary
            logger.info("=" * 60)
            logger.info("Summary")
            logger.info("=" * 60)
            logger.info(f"Total emails retrieved: {len(emails)}")
            logger.info(f"New emails processed: {new_emails}")
            logger.info(f"Skipped (already processed): {skipped_emails}")
            logger.info(f"Total in state tracker: {state_tracker.get_processed_count()}")
    
    except OutlookError as e:
        logger.error(f"Outlook error: {e}")
        return 1
    
    except KeyboardInterrupt:
        logger.warning("Operation cancelled by user")
        state_tracker.save()  # Save progress
        return 130
    
    except Exception as e:
        logger.error(f"Unexpected error: {e}", exc_info=True)
        return 1
    
    logger.info("Script completed successfully")
    return 0


if __name__ == "__main__":
    sys.exit(main())
