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
from datetime import datetime
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


SENT_RESULT_FILE_PREFIX="01_sent_results_"

# ── Progress Utilities ──────────────────────────────────────────

def _progress_interval(total: int) -> int:
    """Determine print interval based on total item count."""
    if total <= 20:
        return 1
    elif total <= 50:
        return 5
    elif total <= 200:
        return 10
    elif total <= 500:
        return 25
    return 50


def print_banner():
    """Print tool banner."""
    print("\n" + "=" * 60)
    print("  Payslip Generator & Distributor - Phuc Long (Excel COM)")
    print("  Powered by Daisy Platform")
    print("=" * 60)


def print_section(title: str):
    """Print a section header with separator."""
    print(f"\n{'─' * 55}")
    print(f"  {title}")
    print(f"{'─' * 55}")

def print_section_lite(title: str):
    """Print a lightweight section header."""
    print(f"\n-- {title} --\n")

def print_with_color(text: str, color_code: int = 92):
    """Print text with ANSI color codes."""
    print(f"\033[{color_code}m{text}\033[0m")


def print_pre_summary(config, employee_count: int):
    """Print pre-execution summary."""
    print("\n--- Configuration Summary ---")
    print(f"  Excel file         : {config.excel_path}")
    print(f"  Payroll date       : {config.date}")
    print(f"  Total employees    : {employee_count}")
    print(f"  Outlook account    : {config.outlook_account}")
    print(f"  Dry run            : {config.dry_run}")
    print(f"  PDF password       : {'Enabled' if config.pdf_password_enabled else 'Disabled'}")
    print(f"  Duplicate emails   : {'Allowed' if config.allow_duplicate_emails else 'Not allowed'}")
    print(f"  Output directory   : {config.output_dir}")
    print("-----------------------------\n")
    
    # Log to file for debugging account issues
    logger.info(f"=== CONFIGURATION ===")
    logger.info(f"Outlook account: {config.outlook_account}")
    logger.info(f"Dry run: {config.dry_run}")
    logger.info(f"Total employees: {employee_count}")


def print_post_summary(stats: dict):
    """Print post-execution summary."""
    print("\n" + "=" * 55)
    print("  FINAL SUMMARY")
    print("=" * 55)
    print(f"  Total employees    : {stats.get('total', 0)}")
    print(f"  Payslips generated : {stats.get('generated', 0)} (skipped: {stats.get('gen_skipped', 0)})")
    print(f"  PDFs converted     : {stats.get('converted', 0)} (skipped: {stats.get('pdf_skipped', 0)})")
    print(f"  Emails sent        : {stats.get('sent', 0)}")
    print(f"  Emails skipped     : {stats.get('skipped', 0)}")
    print(f"  Errors             : {stats.get('errors', 0)}")
    elapsed = stats.get("elapsed", 0)
    print(f"  Time elapsed       : {elapsed:.1f}s ({elapsed/60:.1f}m)")
    if stats.get("result_file"):
        print(f"  Results file       : {stats['result_file']}")
    print("=" * 55 + "\n")


def confirm_proceed() -> bool:
    """Ask user for confirmation before sending."""
    while True:
        answer = input("Proceed with payslip generation and email sending? (yes/no): ").strip().lower()
        if answer in ("yes", "y"):
            return True
        if answer in ("no", "n"):
            return False
        print("Please enter 'yes' or 'no'.")


# ── Result Writer ───────────────────────────────────────────────

class ResultWriter:
    """
    Writes per-employee processing results to a plain text file.

    This file can be used to track which employees have been processed
    and to update the original Excel input file accordingly.
    """

    def __init__(self, output_path: Path, date_str: str):
        self.output_path = Path(output_path)
        if not self.output_path.exists():
            with open(self.output_path, "w", encoding="utf-8") as f:
                f.write(f"# Payslip Distribution Results - {date_str}\n")
                f.write(f"# Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"# Format: MNV | Name | Email | Status | Timestamp\n")
                f.write(f"{'─' * 80}\n")

    def append(self, mnv: str, name: str, email: str, status: str):
        """Append a result entry."""
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(self.output_path, "a", encoding="utf-8") as f:
            f.write(f"{mnv} | {name} | {email} | {status} | {ts}\n")


# ── State Analysis & Cleanup ────────────────────────────────────

def analyze_existing_state(config) -> dict:
    """
    Analyze existing state files for the current payroll date.
    
    Returns dict with:
    - has_state: bool
    - checkpoint_file: Path or None
    - state_file: Path or None
    - result_file: Path or None
    - sent_count: int
    - total_in_results: int
    """
    result = {
        "has_state": False,
        "checkpoint_file": None,
        "state_file": None,
        "result_file": None,
        "sent_count": 0,
        "total_in_results": 0,
    }
    
    # Check checkpoint files (both dry-run and production)
    checkpoint_dryrun = config.state_dir / f"payslip_checkpoint_dryrun_{config.date_mmyyyy}_state.json"
    checkpoint_send = config.state_dir / f"payslip_checkpoint_send_{config.date_mmyyyy}_state.json"
    state_file = config.state_dir / f"payslip_send_{config.date_mmyyyy}_state.json"
    result_file = config.output_dir / f"{SENT_RESULT_FILE_PREFIX}{config.date_mmyyyy}.txt"
    
    if checkpoint_dryrun.exists():
        result["checkpoint_file"] = checkpoint_dryrun
        result["has_state"] = True
        try:
            import json
            with open(checkpoint_dryrun, "r", encoding="utf-8") as f:
                data = json.load(f)
                result["sent_count"] += data.get("total_processed", 0)
        except Exception:
            pass
    
    if checkpoint_send.exists():
        result["checkpoint_file"] = checkpoint_send
        result["has_state"] = True
        try:
            import json
            with open(checkpoint_send, "r", encoding="utf-8") as f:
                data = json.load(f)
                result["sent_count"] += data.get("total_processed", 0)
        except Exception:
            pass
    
    if state_file.exists():
        result["state_file"] = state_file
        result["has_state"] = True
    
    if result_file.exists():
        result["result_file"] = result_file
        result["has_state"] = True
        try:
            with open(result_file, "r", encoding="utf-8") as f:
                lines = f.readlines()
                # Count non-comment lines
                result["total_in_results"] = sum(
                    1 for line in lines 
                    if line.strip() and not line.strip().startswith("#") 
                    and not line.strip().startswith("─")
                )
        except Exception:
            pass
    
    return result


def prompt_state_action(config, state_info: dict, total_employees: int) -> str:
    """
    Prompt user for action when existing state is found.
    
    Returns: "yes", "no", or "new"
    """
    print("\n" + "-" * 60)
    print(f"  EXISTING PAYROLL FOR {config.date} DETECTED")
    print("-" * 60)
    print(f"  Total employees in data: {total_employees}")
    print(f"  Already processed: {state_info['sent_count']}")
    print(f"  Remaining: {total_employees - state_info['sent_count']}")
    
    # if state_info["checkpoint_file"]:
    #     print(f"  Checkpoint file: {state_info['checkpoint_file'].name}")
    # if state_info["state_file"]:
    #     print(f"  State file: {state_info['state_file'].name}")
    # if state_info["result_file"]:
    #     print(f"  Result file: {state_info['result_file'].name} ({state_info['total_in_results']} entries)")
    
    # print("\n")
    print_with_color("\n  You should decide how to continue:", 96)
    print("    yes - Continue from last checkpoint (resume processing)")
    print("    no  - Exit the tool without making any changes")
    print(f"    new - Clean the existing state and start the payroll for {config.date} again")
    print("-" * 60)
    
    while True:
        answer = input("\nYour choice (yes/no/new): ").strip().lower()
        if answer in ("yes", "y"):
            return "yes"
        elif answer in ("no", "n"):
            return "no"
        elif answer == "new":
            return "new"
        print("Please enter 'yes', 'no', or 'new'.")


def cleanup_output_files(config):
    """
    Delete all output files for the current payroll date.
    """
    print("\n  Cleaning up output files...")
    files_deleted = []
    
    # # Delete result file
    # result_file = config.output_dir / f"{SENT_RESULT_FILE_PREFIX}{config.date_mmyyyy}.txt"
    # if result_file.exists():
    #     result_file.unlink()
    #     files_deleted.append(result_file.name)
    
    # Delete output PDF, TXT, and CSV files
    if config.output_dir.exists():
        patterns = ("*.pdf", "*.txt", "*.csv")

        for pattern in patterns:
            for file in config.output_dir.glob(pattern):
                if file.is_file():
                    file.unlink()
                    files_deleted.append(file.name)

    

def cleanup_all_files(config):
    """
    Delete all state and output files for the current payroll date.
    """
    print("\n  Cleaning up state files...")
    files_deleted = []
    
    # Delete checkpoint files
    for mode in ["dryrun", "send"]:
        checkpoint = config.state_dir / f"payslip_checkpoint_{mode}_{config.date_mmyyyy}_state.json"
        if checkpoint.exists():
            checkpoint.unlink()
            files_deleted.append(checkpoint.name)
    
    # Delete state file
    state_file = config.state_dir / f"payslip_send_{config.date_mmyyyy}_state.json"
    if state_file.exists():
        state_file.unlink()
        files_deleted.append(state_file.name)
    
    cleanup_output_files(config)


def main():
    """Main entry point."""
    import pythoncom
    pythoncom.CoInitialize()

    start_time = time.time()
    print_banner()

    # ─── 1. Load Configuration ───
    print_section_lite("Loading Configuration")
    tool_dir = Path(__file__).resolve().parent
    config = load_config(tool_dir=tool_dir)

    config_errors = config.validate()
    if config_errors:
        print("\nConfiguration errors:")
        for err in config_errors:
            print(f"    ERROR: {err}")
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
    print("  Configuration loaded OK")

    # ─── 2. Read Employee Metadata ───
    print_section_lite("Reading Employee Data")
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
        print(f"\n  ERROR: Failed to read Excel file: {e}")
        sys.exit(1)

    # Allow ExcelReader's COM to fully release before generator starts
    gc.collect()
    time.sleep(1)

    if not employees:
        logger.error("No employee data found")
        print("\n  ERROR: No employee data found in the Excel file.")
        sys.exit(1)

    print(f"  Found {len(employees)} employees")
    logger.info(f"Found {len(employees)} employees")

    # ─── 2.5. Check for Existing State ───
    state_info = analyze_existing_state(config)
    if state_info["has_state"]:
        action = prompt_state_action(config, state_info, len(employees))
        if action == "no":
            print("\nExited by user. No changes made.")
            sys.exit(0)
        elif action == "new":
            cleanup_all_files(config)
            print("  Starting fresh with clean state.\n")
        else:  # action == "yes"
            print("\n  Resuming from last checkpoint.\n")

    # ─── 3. Validate Data ───
    print_section_lite("Validating Data")
    validator = DataValidator(
        employees,
        allow_duplicate_emails=config.allow_duplicate_emails,
    )
    errors, warnings = validator.validate_all()

    if errors:
        print(f"\n  Validation FAILED with {len(errors)} error(s):")
        for err in errors:
            print(f"    ERROR: {err}")
        print("\n  Please fix the data and try again.")
        sys.exit(1)

    if warnings:
        print(f"  {len(warnings)} warning(s) found (non-blocking)")
        for w in warnings[:5]:
            print(f"    WARNING: {w}")
        if len(warnings) > 5:
            print(f"    ... and {len(warnings) - 5} more (see log file)")

    print(f"  Validation passed: {len(employees)} employees OK")

    # ─── Pre-Execution Summary & Confirmation ───
    print_pre_summary(config, len(employees))

    if not config.dry_run:
        if not confirm_proceed():
            print("Aborted by user.")
            sys.exit(0)
    else:
        print("  [DRY-RUN MODE] Simulating — no emails will be sent\n")

    # Clean up output files for new session
    cleanup_output_files(config)

    # ─── 4. Generate Payslip Excel Files via COM ───
    print_section_lite("Generating Payslips")
    generator = PayslipGenerator(
        output_dir=config.output_dir,
        date_str=config.date,
        filename_pattern=config.pdf_filename_pattern,
    )

    interval = _progress_interval(len(employees))

    def gen_progress(current, total, name, skipped=False):
        action = "Skipped (exists)" if skipped else "Generated"
        if current == 1 or current == total or current % interval == 0:
            print(f"  [{current}/{total}] {action}: {name}")

    gen_start = time.time()
    try:
        results = generator.generate_batch(
            employees=employees,
            source_xls=config.excel_path,
            template_sheet=config.template_sheet,
            data_sheet=config.data_sheet,
            col_mnv=config.col_mnv,
            progress_callback=gen_progress,
        )
    except Exception as e:
        logger.error(f"Payslip generation failed: {e}")
        print(f"\n  ERROR: Payslip generation failed: {e}")
        sys.exit(1)

    gen_elapsed = time.time() - gen_start
    generated = sum(1 for r in results if r["success"] and not r.get("skipped"))
    gen_skipped = sum(1 for r in results if r.get("skipped"))
    gen_failed = sum(1 for r in results if not r["success"])
    print(f"\n  Result: Generated {generated}, Skipped {gen_skipped}, Failed {gen_failed} ({gen_elapsed:.1f}s)")

    # Allow Excel COM to fully release before starting PDF converter
    gc.collect()
    time.sleep(2)

    # ─── 5. Convert to Password-Protected PDFs ───
    print_section_lite("Converting to PDF")
    successful_items = [r for r in results if r["success"]]

    def pdf_progress(current, total, name, skipped=False):
        action = "Skipped (exists)" if skipped else "Converted"
        if current == 1 or current == total or current % interval == 0:
            print(f"  [{current}/{total}] {action}: {name}")

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
        logger.error(f"PDF conversion failed: {e}")
        print(f"\n  ERROR: PDF conversion failed: {e}")

    pdf_elapsed = time.time() - pdf_start
    converted = sum(1 for r in results if r.get("pdf_path"))
    pdf_skipped = sum(1 for r in results if r.get("pdf_skipped"))
    pdf_failed = len(successful_items) - converted
    print(f"\n  Result: Converted {converted - pdf_skipped}, Skipped {pdf_skipped}, Failed {pdf_failed} ({pdf_elapsed:.1f}s)")

    # ─── 6. Compose and Send Emails ───
    print_section_lite("Sending Emails")
    print("  Composing emails...")
    composer = EmailComposer(
        template_cells=email_template,
        subject=subject,
        date_str=config.date,
        date_cell=config.email_date_cell,
    )
    results = composer.compose_batch(results)
    composed = sum(1 for r in results if r.get("email_data"))
    print(f"  Composed {composed} emails")

    # Initialize MNV-based checkpoint tracker (for resume support)
    run_mode = "dryrun" if config.dry_run else "send"
    checkpoint = StateTracker(
        state_dir=config.state_dir,
        state_name=f"payslip_checkpoint_{run_mode}_{config.date_mmyyyy}",
        auto_save=True,
        auto_save_interval=1,  # Save after every email for crash safety
    )

    # Initialize content hash tracker (for Outlook duplicate prevention)
    state_tracker = StateTracker(
        state_dir=config.state_dir,
        state_name=f"payslip_send_{config.date_mmyyyy}",
    )

    # Initialize result writer
    result_file = config.output_dir / f"{SENT_RESULT_FILE_PREFIX}{config.date_mmyyyy}.txt"
    result_writer = ResultWriter(result_file, config.date)

    resumed_count = checkpoint.get_processed_count()
    if resumed_count > 0:
        print(f"  Resuming: {resumed_count} employees already processed in previous run")

    print("  Sending emails via Outlook...")
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
                mnv = emp.get("mnv", "")
                name = emp.get("name", "N/A")
                email_addr = emp.get("email", "")

                if not email_data:
                    logger.warning(f"No email data for {name} (MNV: {mnv})")
                    error_count += 1
                    result_writer.append(mnv, name, email_addr, "NO_EMAIL_DATA")
                    continue

                # Check MNV-based checkpoint (resume support)
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
                        status = "DRY_RUN" if config.dry_run else "SENT"
                        checkpoint.mark_processed(mnv, metadata={
                            "name": name,
                            "email": email_addr,
                            "status": status,
                        })
                        result_writer.append(mnv, name, email_addr, status)
                        if i == 1 or i == composed or i % interval == 0:
                            print(f"  [{i}/{composed}] Sent: {name}")
                    else:
                        skipped_count += 1
                        checkpoint.mark_processed(mnv, metadata={
                            "name": name,
                            "email": email_addr,
                            "status": "SKIPPED_DUPLICATE",
                        })
                        result_writer.append(mnv, name, email_addr, "SKIPPED_DUPLICATE")
                        logger.info(f"[{i}/{composed}] Skipped {name} (duplicate)")
                except Exception as e:
                    error_count += 1
                    result_writer.append(mnv, name, email_addr, f"FAILED: {e}")
                    logger.error(f"[{i}/{composed}] Failed for {name}: {e}")

            print(
                f"\n  Sender stats - Sent: {sender.sent_count}, "
                f"Skipped: {sender.skipped_count}, "
                f"Errors: {sender.error_count}"
            )

    except ImportError:
        logger.error("win32com not available - cannot send emails")
        print("\n  ERROR: Outlook COM not available.")
        if config.dry_run:
            print("  [DRY-RUN] Would have sent emails. Skipping Outlook.")
            sent_count = composed
    except Exception as e:
        logger.error(f"Email sending failed: {e}")
        print(f"\n  ERROR: Email sending failed: {e}")

    # ─── 7. Post-Execution Summary ───
    elapsed = time.time() - start_time
    stats = {
        "total": len(employees),
        "generated": generated,
        "gen_skipped": gen_skipped,
        "converted": converted,
        "pdf_skipped": pdf_skipped,
        "sent": sent_count,
        "skipped": skipped_count,
        "errors": error_count,
        "elapsed": elapsed,
        "result_file": result_file,
    }
    print_post_summary(stats)
    logger.info(f"Payslip processing complete: {stats}")

    # Save all state
    checkpoint.save()
    state_tracker.save()

    # Release COM apartment
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass

    if error_count > 0:
        print(f"  WARNING: {error_count} error(s) occurred. Check logs for details.")
        sys.exit(1)

    print("  Done!")
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
