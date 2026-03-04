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
import sys
import time
from pathlib import Path

# Add project root to path for foundation imports
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.console import cprint, cprint_banner, cprint_summary_box, cprint_summary_box_lite
from core.logger import setup_logging, get_logger
from core.state import StateTracker
from office.outlook.models import NewEmail
from office.outlook.sender import OutlookSender

from config import load_config
from excel_reader import ExcelReader
from validator import DataValidator
from payslip_generator import PayslipGenerator
from office.excel.converter import PdfConverter
from email_composer import EmailComposer
from utils import *


logger = get_logger(__name__)


def load_and_validate_config(tool_dir: Path):
    """Load and validate configuration."""
    config = load_config(tool_dir=tool_dir)

    config_errors = config.validate()
    if config_errors:
        print()
        cprint("Configuration errors:", level="ERROR")
        for err in config_errors:
            print(f"      {err}")
        print()
        print("  Please fix .env file and try again.")
        sys.exit(1)

    config.ensure_directories()
    setup_logging(
        log_dir=config.log_dir,
        level=config.log_level,
        run_name="payslip",
    )
    cprint("Configuration loaded", level="SUCCESS")
    return config


def read_employee_data(config):
    """Read employee data and email template from Excel."""
    cprint("Input & Validation", level="PHASE")
    print()
    cprint("Reading employee data...", level="INFO")
    
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
        cprint(f"Failed to read Excel file: {e}", level="ERROR")
        sys.exit(1)

    # Allow ExcelReader's COM to fully release
    gc.collect()
    time.sleep(1)

    if not employees:
        cprint("No employee data found in the Excel file", level="ERROR")
        sys.exit(1)

    cprint(f"Found {len(employees)} employee" + ("s" if len(employees) != 1 else ""), level="SUCCESS")
    return employees, email_template, subject


def check_and_handle_existing_state(config, total_employees: int):
    """Check for existing state and prompt user for action if found."""
    state_info = analyze_existing_state(config)
    if state_info["has_state"]:
        action = prompt_state_action(config, state_info, total_employees)
        if action == "no":
            cprint("Exited by user.", level="INFO")
            sys.exit(0)
        elif action == "new":
            cleanup_all_files(config)
            print()
            cprint("Starting fresh with new payroll", level="INFO")
        else:  # action == "yes"
            cprint("Resuming from last checkpoint.", level="INFO")


def validate_employee_data(employees, config):
    """Validate employee data and handle errors/warnings."""
    print()
    cprint("Validating data...", level="INFO")
    validator = DataValidator(
        employees,
        allow_duplicate_emails=config.allow_duplicate_emails,
    )
    errors, warnings = validator.validate_all()

    if errors:
        print()
        cprint(f"Validation FAILED with {len(errors)} error(s):", level="ERROR")
        for err in errors:
            print(f"      {err}")
        print()
        print("  Please fix the data and try again.")
        sys.exit(1)

    if warnings:
        cprint(f"{len(warnings)} warning(s) found (non-blocking)", level="WARNING")
        for w in warnings[:5]:
            print(f"      {w}")
        if len(warnings) > 5:
            print(f"      ... and {len(warnings) - 5} more (see log file)")

    cprint("Validation passed", level="SUCCESS")


def show_summary_and_confirm(config, employees):
    """Show configuration summary and get user confirmation."""
    excel_short = config.excel_path.name if hasattr(config.excel_path, 'name') else str(config.excel_path).split('\\')[-1]
    output_short = config.output_dir.name if hasattr(config.output_dir, 'name') else 'output/'
    cprint_summary_box_lite(
        "Configuration Summary",
        {
            "Excel file": excel_short,
            "Payroll date": config.date,
            "Employees": len(employees),
            "Outlook account": config.outlook_account,
            "Dry run": "Yes" if config.dry_run else "No",
            "PDF password": "Enabled" if config.pdf_password_enabled else "Disabled",
            "Keep PDFs": "Yes" if config.keep_pdf_payslips else "No",
            "Output directory": output_short,
        },
    )

    if not config.dry_run:
        if not confirm_proceed():
            cprint("Aborted by user.", level="ERROR")
            sys.exit(0)
    else:
        cprint("Dry Run Mode", level="WARNING")
        cprint("Simulating - no emails will be sent\n", level="INFO")


def generate_payslips(config, employees):
    """Generate payslip Excel files via COM."""
    cprint("Generating payslips", level="PHASE")
    generator = PayslipGenerator(
        output_dir=config.output_dir,
        date_str=config.date,
        filename_pattern=config.pdf_filename_pattern,
    )

    interval = progress_interval(len(employees))

    def gen_progress(current, total, name, skipped=False):
        action = "Skipped (exists)" if skipped else "Done"
        level = "WARNING" if skipped else "SUCCESS"
        if current == 1 or current == total or current % interval == 0:
            cprint(f"[{current}/{total}] {action} {name}", level=level, indent=2)

    gen_start = time.time()
    try:
        results = generator.generate_batch(
            employees=employees,
            source_xls=config.excel_path,
            batch_size=config.batch_size,
            template_sheet=config.template_sheet,
            data_sheet=config.data_sheet,
            col_mnv=config.col_mnv,
            progress_callback=gen_progress
        )
    except Exception as e:
        cprint(f"Payslip generation failed: {e}", level="ERROR")
        sys.exit(1)

    gen_elapsed = time.time() - gen_start
    generated = sum(1 for r in results if r["success"] and not r.get("skipped"))
    gen_skipped = sum(1 for r in results if r.get("skipped"))
    gen_failed = sum(1 for r in results if not r["success"])
    print()
    cprint(f"Result: Generated {generated} | Skipped {gen_skipped} | Failed {gen_failed} ({gen_elapsed:.1f}s)", level="INFO")

    # Allow Excel COM to fully release
    gc.collect()
    time.sleep(2)

    return results


def convert_to_pdf(config, results):
    """Convert successful Excel payslips to password-protected PDFs."""
    cprint("Converting to PDF", level="PHASE")
    successful_items = [r for r in results if r["success"]]
    interval = progress_interval(len(successful_items))

    def pdf_progress(current, total, name, skipped=False):
        action = "Skipped (exists)" if skipped else "Done"
        level = "WARNING" if skipped else "SUCCESS"
        if current == 1 or current == total or current % interval == 0:
            cprint(f"[{current}/{total}] {action} {name}", level=level, indent=2)

    pdf_start = time.time()
    try:
        with PdfConverter(
            output_dir=config.output_dir,
            password_enabled=config.pdf_password_enabled,
            strip_leading_zeros=config.pdf_password_strip_zeros,
        ) as converter:
            results = converter.convert_batch(
                successful_items,
                progress_callback=pdf_progress,
            )
    except Exception as e:
        cprint(f"PDF conversion failed: {e}", level="ERROR")

    pdf_elapsed = time.time() - pdf_start
    converted = sum(1 for r in results if r.get("pdf_path"))
    pdf_skipped = sum(1 for r in results if r.get("pdf_skipped"))
    pdf_failed = len(successful_items) - converted
    print()
    cprint(f"Result: Converted {converted - pdf_skipped} | Skipped {pdf_skipped} | Failed {pdf_failed} ({pdf_elapsed:.1f}s)", level="INFO")

    return results


def compose_emails(config, results, email_template, subject):
    """Compose email messages for all results."""
    cprint("Sending emails", level="PHASE")
    cprint("Composing emails…", level="INFO", indent=2)
    composer = EmailComposer(
        template_cells=email_template,
        subject=subject,
        date_str=config.date,
        date_cell=config.email_date_cell,
    )
    results = composer.compose_batch(results)
    composed = sum(1 for r in results if r.get("email_data"))
    cprint("Sending via Outlook…", level="INFO", indent=2)
    print()
    return results, composed


def send_emails(config, results, composed):
    """Send emails via Outlook COM and track results."""
    # Initialize trackers
    run_mode = "dryrun" if config.dry_run else "send"
    checkpoint = StateTracker(
        state_dir=config.state_dir,
        state_name=f"payslip_checkpoint_{run_mode}_{config.date_mmyyyy}",
        auto_save=True,
        auto_save_interval=1,
    )

    state_tracker = StateTracker(
        state_dir=config.state_dir,
        state_name=f"payslip_send_{config.date_mmyyyy}",
    )

    result_file = config.output_dir / f"{SENT_RESULT_FILE_PREFIX}{config.date_mmyyyy}.csv"
    result_writer = ResultWriter(result_file, config.date)

    resumed_count = checkpoint.get_processed_count()
    if resumed_count > 0:
        cprint(f"Resuming: {resumed_count} employees already processed", level="INFO")

    sent_count = 0
    skipped_count = 0
    error_count = 0
    interval = progress_interval(composed)

    try:
        with OutlookSender(
            account=config.outlook_account,
            dry_run=config.dry_run,
            state_tracker=state_tracker,
        ) as sender:
            for i, item in enumerate(results, 1):
                email_data = item.get("email_data")
                emp = item.get("employee", {})
                mnv = emp.get("mnv", "")
                name = emp.get("name", "N/A")
                email_addr = emp.get("email", "")

                if not email_data:
                    logger.warning(f"No email data for {name} (MNV: {mnv})")
                    error_count += 1
                    result_writer.append(mnv, name, email_addr, "NO_EMAIL_DATA", "", "No email data composed")
                    continue

                # Check checkpoint for resume support
                if checkpoint.is_processed(mnv):
                    skipped_count += 1
                    continue

                email = NewEmail(
                    to=email_data["to"],
                    subject=email_data["subject"],
                    body=email_data["body"],
                    body_is_html=email_data["body_is_html"],
                    attachments=email_data.get("attachments", []),
                )

                try:
                    result = sender.send(
                        email,
                        skip_duplicate_check=config.allow_duplicate_emails,
                    )
                    if result:
                        sent_count += 1
                        status = "DRY_RUN" if config.dry_run else "SUCCESS"
                        pdf_path = item.get("pdf_path", "")
                        payslip_fn = Path(pdf_path).name if pdf_path else ""
                        checkpoint.mark_processed(mnv, metadata={
                            "name": name,
                            "email": email_addr,
                            "status": status,
                        })
                        result_writer.append(mnv, name, email_addr, status, payslip_fn, "")
                        if i == 1 or i == composed or i % interval == 0:
                            cprint(f"[{i}/{composed}] {name}", level="SUCCESS", indent=2)
                        # Cleanup PDF after successful send if not keeping
                        if not config.keep_pdf_payslips and pdf_path:
                            cleanup_pdf(pdf_path)
                    else:
                        skipped_count += 1
                        checkpoint.mark_processed(mnv, metadata={
                            "name": name,
                            "email": email_addr,
                            "status": "SKIPPED_DUPLICATE",
                        })
                        result_writer.append(mnv, name, email_addr, "SKIPPED_DUPLICATE", "", "Duplicate email")
                        logger.info(f"[{i}/{composed}] Skipped {name} (duplicate)")
                except Exception as e:
                    error_count += 1
                    result_writer.append(mnv, name, email_addr, "FAILED", "", str(e))
                    logger.error(f"[{i}/{composed}] Failed for {name}: {e}")

            print()
            cprint(
                f"Sender stats: Sent {sender.sent_count} | "
                f"Skipped {sender.skipped_count} | "
                f"Errors {sender.error_count}",
                level="INFO",
            )

    except ImportError:
        cprint("Outlook COM not available.", level="ERROR")
        if config.dry_run:
            cprint("[DRY-RUN] Would have sent emails. Skipping Outlook.", level="INFO")
            sent_count = composed
    except Exception as e:
        cprint(f"Email sending failed: {e}", level="ERROR")

    # Save all state
    checkpoint.save()
    state_tracker.save()

    return sent_count, skipped_count, error_count, result_file


def main():
    """Main entry point."""
    import pythoncom
    pythoncom.CoInitialize()

    start_time = time.time()
    cprint_banner(
        "Payslip Generator & Distributor - Phuc Long",
        "Excel COM | Outlook | Daisy Platform",
    )

    # Load Configuration
    tool_dir = Path(__file__).resolve().parent
    config = load_and_validate_config(tool_dir)

    # Read Employee Data
    employees, email_template, subject = read_employee_data(config)

    # Check Existing State
    check_and_handle_existing_state(config, len(employees))

    # Validate Data
    validate_employee_data(employees, config)

    # Show Summary & Get Confirmation
    show_summary_and_confirm(config, employees)
    cleanup_output_files(config)

    # Generate Payslips
    results = generate_payslips(config, employees)
    generated = sum(1 for r in results if r["success"] and not r.get("skipped"))
    gen_skipped = sum(1 for r in results if r.get("skipped"))

    # Convert to PDF
    results = convert_to_pdf(config, results)
    converted = sum(1 for r in results if r.get("pdf_path"))
    pdf_skipped = sum(1 for r in results if r.get("pdf_skipped"))

    # Compose & Send Emails
    results, composed = compose_emails(config, results, email_template, subject)
    sent_count, skipped_count, error_count, result_file = send_emails(config, results, composed)

    # Post-Execution Summary
    elapsed = time.time() - start_time
    cprint_summary_box(
        "FINAL SUMMARY",
        {
            "Total employees": len(employees),
            "Generated": generated,
            "Gen skipped": gen_skipped,
            "Converted to PDF": converted,
            "PDF skipped": pdf_skipped,
            "Emails sent": sent_count,
            "Emails skipped": skipped_count,
            "Errors": error_count,
            "Elapsed": f"{elapsed:.1f}s",
            "Result file": str(result_file),
        },
    )

    # Release COM apartment
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass

    if error_count > 0:
        cprint(f"WARNING: {error_count} error(s) occurred. Check logs for details.", level="WARNING")
        sys.exit(1)

    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
