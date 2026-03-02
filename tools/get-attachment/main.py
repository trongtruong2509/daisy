"""
Main entry point for the Get Attachment tool.

Connects to Outlook, searches the Inbox for emails on a specific date
(optionally filtered by subject keywords), and saves all attachments to
a local directory.

Workflow:
  1. Load configuration (from .env or interactive prompts)
  2. Validate configuration — exit on error
  3. Set up logging
  4. Show a run summary and confirm with the user
  5. Connect to Outlook, search, and download attachments
  6. Print a final summary

Usage:
    cd tools/get-attachment
    python main.py
"""

import sys
from pathlib import Path

# Ensure project root is importable regardless of working directory
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.console import cprint, cprint_banner, cprint_summary_box, cprint_summary_box_lite
from core.logger import setup_logging, get_logger

from config import load_config
from attachment_downloader import AttachmentDownloader, DownloadResult


logger = get_logger(__name__)

TOOL_DIR = Path(__file__).resolve().parent


# ── Helpers ──────────────────────────────────────────────────────

def _confirm_proceed() -> bool:
    """Ask the user for a yes/no confirmation before starting the download."""
    cprint("Proceed with download? (yes/no)", level="WARNING")
    while True:
        answer = input("\u2192 ").strip().lower()
        if answer in ("yes", "y"):
            return True
        if answer in ("no", "n"):
            return False
        print("  Please enter 'yes' or 'no'.")


def _show_config_summary(config) -> None:
    """Print a compact configuration summary before proceeding."""
    keywords_display = (
        ", ".join(config.subject_keywords) if config.subject_keywords else "(none — all emails)"
    )
    save_short = str(config.attachment_save_path)
    end_display = config.end_date if config.end_date else f"{config.end_date_parsed.strftime('%d/%m/%Y')} (today)"

    cprint_summary_box_lite(
        "Get Attachment \u2014 Run Summary",
        {
            "Outlook account": config.outlook_account,
            "Start date": config.start_date,
            "End date": end_display,
            "Subject keywords": keywords_display,
            "Save path": save_short,
            "Log directory": str(config.log_dir),
        },
    )


def _show_final_summary(result: DownloadResult, config) -> None:
    """Print the final download summary."""
    end_display = config.end_date if config.end_date else config.end_date_parsed.strftime("%d/%m/%Y")
    summary = {
        "Date range": f"{config.start_date} \u2192 {end_display}",
        "Emails found": result.emails_found,
        "Emails matched": result.emails_matched,
        "With attachments": result.emails_with_attachments,
        "Attachments saved": result.attachments_saved,
        "Attachments failed": result.attachments_failed,
    }

    footer = ""
    if result.attachments_failed:
        footer = f"WARNING: {result.attachments_failed} attachment(s) could not be saved. Check log for details."

    cprint_summary_box("Download Complete", summary, footer=footer)

    if result.saved_files:
        print()
        cprint("Saved files:", level="INFO")
        for path in result.saved_files:
            print(f"    {path}")

    if result.errors:
        print()
        cprint("Errors:", level="ERROR")
        for err in result.errors:
            print(f"    {err}")


# ── Main ─────────────────────────────────────────────────────────

def main() -> None:
    """Run the Get Attachment tool."""

    cprint_banner(
        "Get Attachment",
        "Download email attachments from Outlook Inbox",
    )

    # 1. Load and validate configuration
    config = load_config(tool_dir=TOOL_DIR)

    errors = config.validate()
    if errors:
        print()
        cprint("Configuration errors:", level="ERROR")
        for err in errors:
            print(f"      {err}")
        print()
        print("  Please fix the .env file or re-enter the values and try again.")
        sys.exit(1)

    # 2. Set up file logging
    config.ensure_directories()
    setup_logging(
        log_dir=config.log_dir,
        level=config.log_level,
        run_name="get_attachment",
    )
    logger.info(
        "Starting get-attachment: account=%s start=%s end=%s keywords=%s save_path=%s",
        config.outlook_account,
        config.start_date,
        config.end_date or "today",
        config.subject_keywords,
        config.attachment_save_path,
    )

    # 3. Summary + confirmation
    print()
    _show_config_summary(config)
    print()

    if not _confirm_proceed():
        cprint("Aborted by user.", level="INFO")
        logger.info("Run aborted by user at confirmation prompt.")
        sys.exit(0)

    print()

    # 4. Download attachments
    try:
        downloader = AttachmentDownloader(config)
        result = downloader.run()
    except Exception as exc:
        cprint(f"Fatal error: {exc}", level="ERROR")
        logger.exception("Fatal error during attachment download")
        sys.exit(1)

    # 5. Final summary
    print()
    _show_final_summary(result, config)

    logger.info(
        "Run complete: emails_found=%d emails_matched=%d saved=%d failed=%d",
        result.emails_found,
        result.emails_matched,
        result.attachments_saved,
        result.attachments_failed,
    )

    if result.attachments_failed:
        sys.exit(2)  # Non-zero exit so callers know something went wrong


if __name__ == "__main__":
    main()
