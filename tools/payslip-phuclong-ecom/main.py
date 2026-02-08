"""
Main orchestrator for Payslip generation and distribution (Excel COM variant).

Uses Excel COM for ALL formula evaluation, guaranteeing correct
values regardless of formula complexity (VLOOKUP, XLOOKUP, etc.).

Workflow:
1. Load configuration
2. Read employee metadata from Excel (MNV, Name, Email, Password)
3. Validate all data (fail-fast)
4. Generate payslip Excel files via COM (set B3=MNV, copy, paste values)
5. Convert to password-protected PDFs
6. Compose and send emails via Outlook
7. Report summary

Usage:
    cd tools/payslip-phuclong-ecom
    python main.py
"""

import gc
import logging
import sys
import time
from pathlib import Path

# Add project root to path for foundation imports
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.logger import setup_logging, get_logger
from core.state import StateTracker
from office.outlook.models import NewEmail
from office.outlook.sender import OutlookSender

from config import load_config
from excel_reader import ExcelReader
from validator import DataValidator
from payslip_generator import PayslipGenerator
from pdf_converter import PdfConverter
from email_composer import EmailComposer


logger = get_logger(__name__)


def print_banner():
    """Print tool banner."""
    print("\n" + "=" * 60)
    print("  Payslip Generator & Distributor - Phuc Long (Excel COM)")
    print("  Powered by Daisy Foundation")
    print("=" * 60)


def print_pre_summary(config, employee_count: int):
    """Print pre-execution summary."""
    print("\n--- Pre-Execution Summary ---")
    print(f"  Excel file      : {config.excel_path}")
    print(f"  Payroll date     : {config.date}")
    print(f"  Total employees  : {employee_count}")
    print(f"  Outlook account  : {config.outlook_account}")
    print(f"  Dry run          : {config.dry_run}")
    print(f"  PDF password     : {'Enabled' if config.pdf_password_enabled else 'Disabled'}")
    print(f"  Output directory : {config.output_dir}")
    print(f"  Log directory    : {config.log_dir}")
    print("-----------------------------\n")


def print_post_summary(stats: dict):
    """Print post-execution summary."""
    print("\n--- Post-Execution Summary ---")
    print(f"  Total employees  : {stats.get('total', 0)}")
    print(f"  Payslips generated: {stats.get('generated', 0)}")
    print(f"  PDFs converted   : {stats.get('converted', 0)}")
    print(f"  Emails sent      : {stats.get('sent', 0)}")
    print(f"  Emails skipped   : {stats.get('skipped', 0)}")
    print(f"  Errors           : {stats.get('errors', 0)}")
    elapsed = stats.get("elapsed", 0)
    print(f"  Time elapsed     : {elapsed:.1f}s ({elapsed/60:.1f}m)")
    print("------------------------------\n")


def confirm_proceed() -> bool:
    """Ask user for confirmation before sending."""
    while True:
        answer = input("Proceed with payslip generation and email sending? (yes/no): ").strip().lower()
        if answer in ("yes", "y"):
            return True
        if answer in ("no", "n"):
            return False
        print("Please enter 'yes' or 'no'.")


def main():
    """Main entry point."""
    import pythoncom
    pythoncom.CoInitialize()

    start_time = time.time()
    print_banner()

    # ─── 1. Load Configuration ───
    print("Loading configuration...")
    tool_dir = Path(__file__).resolve().parent
    config = load_config(tool_dir=tool_dir)

    config_errors = config.validate()
    if config_errors:
        print("\nConfiguration errors:")
        for err in config_errors:
            print(f"  ERROR: {err}")
        print("\nPlease fix .env file and try again.")
        sys.exit(1)

    config.ensure_directories()

    setup_logging(
        log_dir=config.log_dir,
        level=config.log_level,
        run_name="payslip",
    )
    logger.info("Configuration loaded successfully")
    logger.info(f"Excel path: {config.excel_path}")
    logger.info(f"Date: {config.date}")

    # ─── 2. Read Employee Metadata ───
    print("Reading employee data from Excel...")
    try:
        with ExcelReader(config.excel_path) as reader:
            employees = reader.read_employees(
                data_sheet=config.data_sheet,
                header_row=config.data_header_row,
                start_row=config.data_start_row,
                col_mnv=config.col_mnv,
                col_name=config.col_name,
                col_email=config.col_email,
                col_password=config.col_password,
            )

            email_template = reader.read_email_template(
                sheet_name=config.email_body_sheet,
                body_cells=config.email_body_cells,
                date_cell=config.email_date_cell,
            )

            if config.email_subject:
                subject = config.email_subject
            else:
                subject = reader.read_email_subject(
                    sheet_name=config.template_sheet,
                    subject_cell=config.email_subject_cell,
                )

    except Exception as e:
        logger.error(f"Failed to read Excel file: {e}")
        print(f"\nERROR: Failed to read Excel file: {e}")
        sys.exit(1)

    # Allow ExcelReader's COM to fully release before generator starts
    gc.collect()
    time.sleep(1)

    if not employees:
        logger.error("No employee data found")
        print("\nERROR: No employee data found in the Excel file.")
        sys.exit(1)

    print(f"  Found {len(employees)} employees")
    logger.info(f"Found {len(employees)} employees")

    # ─── 3. Validate Data ───
    print("Validating employee data...")
    validator = DataValidator(employees)
    errors, warnings = validator.validate_all()

    if errors:
        print(f"\nValidation FAILED with {len(errors)} error(s):")
        for err in errors:
            print(f"  ERROR: {err}")
        print("\nPlease fix the data and try again.")
        sys.exit(1)

    if warnings:
        print(f"  {len(warnings)} warning(s) found (non-blocking)")

    print(f"  Validation passed: {len(employees)} employees OK")

    # ─── Pre-Execution Summary & Confirmation ───
    print_pre_summary(config, len(employees))

    if not config.dry_run:
        if not confirm_proceed():
            print("Aborted by user.")
            sys.exit(0)
    else:
        print("[DRY-RUN MODE] Simulating — no emails will be sent\n")

    # ─── 4. Generate Payslip Excel Files via COM ───
    print("Generating payslip Excel files (via Excel COM)...")
    generator = PayslipGenerator(
        output_dir=config.output_dir,
        date_str=config.date,
        filename_pattern=config.pdf_filename_pattern,
    )

    try:
        results = generator.generate_batch(
            employees=employees,
            source_xls=config.excel_path,
            template_sheet=config.template_sheet,
            data_sheet=config.data_sheet,
            col_mnv=config.col_mnv,
        )
    except Exception as e:
        logger.error(f"Payslip generation failed: {e}")
        print(f"\nERROR: Payslip generation failed: {e}")
        sys.exit(1)

    generated = sum(1 for r in results if r["success"])
    print(f"  Generated {generated}/{len(employees)} payslips")

    # Allow Excel COM to fully release before starting PDF converter
    gc.collect()
    time.sleep(2)

    # ─── 5. Convert to Password-Protected PDFs ───
    print("Converting payslips to PDF...")
    try:
        with PdfConverter(
            output_dir=config.output_dir,
            password_enabled=config.pdf_password_enabled,
            strip_leading_zeros=config.pdf_password_strip_zeros,
        ) as converter:
            results = converter.convert_batch(
                [r for r in results if r["success"]]
            )
    except Exception as e:
        logger.error(f"PDF conversion failed: {e}")
        print(f"\nERROR: PDF conversion failed: {e}")

    converted = sum(1 for r in results if r.get("pdf_path"))
    print(f"  Converted {converted}/{generated} PDFs")

    # ─── 6. Compose and Send Emails ───
    print("Composing emails...")
    composer = EmailComposer(
        template_cells=email_template,
        subject=subject,
        date_str=config.date,
        date_cell=config.email_date_cell,
    )
    results = composer.compose_batch(results)
    composed = sum(1 for r in results if r.get("email_data"))
    print(f"  Composed {composed} emails")

    print("Sending emails via Outlook...")
    state_tracker = StateTracker(
        state_dir=config.state_dir,
        state_name=f"payslip_send_{config.date_mmyyyy}",
    )

    sent_count = 0
    skipped_count = 0
    error_count = 0

    try:
        with OutlookSender(
            account=config.outlook_account,
            dry_run=config.dry_run,
            state_tracker=state_tracker,
        ) as sender:
            for i, item in enumerate(results, 1):
                email_data = item.get("email_data")
                emp = item.get("employee", {})
                name = emp.get("name", "N/A")

                if not email_data:
                    logger.warning(f"[{i}/{composed}] No email data for {name}")
                    error_count += 1
                    continue

                email = NewEmail(
                    to=email_data["to"],
                    subject=email_data["subject"],
                    body=email_data["body"],
                    body_is_html=email_data["body_is_html"],
                    attachments=email_data.get("attachments", []),
                )

                try:
                    result = sender.send(email)
                    if result:
                        sent_count += 1
                        logger.info(f"[{i}/{composed}] Sent to {name}")
                    else:
                        skipped_count += 1
                        logger.info(f"[{i}/{composed}] Skipped {name} (duplicate)")
                except Exception as e:
                    error_count += 1
                    logger.error(f"[{i}/{composed}] Failed for {name}: {e}")

            print(
                f"  Sent: {sender.sent_count}, "
                f"Skipped: {sender.skipped_count}, "
                f"Errors: {sender.error_count}"
            )

    except ImportError:
        logger.error("win32com not available - cannot send emails")
        print("\nERROR: Outlook COM not available.")
        if config.dry_run:
            print("[DRY-RUN] Would have sent emails. Skipping Outlook.")
            sent_count = composed
    except Exception as e:
        logger.error(f"Email sending failed: {e}")
        print(f"\nERROR: Email sending failed: {e}")

    # ─── 7. Post-Execution Summary ───
    elapsed = time.time() - start_time
    stats = {
        "total": len(employees),
        "generated": generated,
        "converted": converted,
        "sent": sent_count,
        "skipped": skipped_count,
        "errors": error_count,
        "elapsed": elapsed,
    }
    print_post_summary(stats)
    logger.info(f"Payslip processing complete: {stats}")

    state_tracker.save()

    # Release COM apartment
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass

    if error_count > 0:
        print(f"WARNING: {error_count} error(s) occurred. Check logs for details.")
        sys.exit(1)

    print("Done!")
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
