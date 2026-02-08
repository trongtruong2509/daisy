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
RESET = "\033[0m"

# ── Console Output Style ────────────────────────────────────────
# Colors
PHASE = "\033[96m"    # Bright cyan for phase headers
OK = "\033[92m"       # Bright green for success
WARN = "\033[93m"     # Yellow for warnings
ERROR = "\033[91m"    # Bright red for errors
INFO = "\033[37m"     # White for info
BOX = "\033[96m"      # Bright cyan for boxes
RESET = "\033[0m"

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


def print_with_color(text: str, color_code: int = 92):
    """Print text with ANSI color codes."""
    print(f"\033[{color_code}m{text}{RESET}")

def print_banner():
    """Print tool banner."""
    print(f"\n{BOX}╔════════════════════════════════════════════════════════════════╗{RESET}")
    print(f"{BOX}║  Payslip Generator & Distributor – Phuc Long                   ║{RESET}")
    print(f"{BOX}║  Excel COM • Outlook • Daisy Platform                          ║{RESET}")
    print(f"{BOX}╚════════════════════════════════════════════════════════════════╝{RESET}")

def print_phase(title: str):
    """Print a phase header."""
    print(f"\n{PHASE}▶  {title}{RESET}")

def print_success(text: str, indent: int = 0):
    """Print success message with checkmark."""
    spaces = " " * indent
    print(f"{spaces}{OK}✓ {text}{RESET}")

def print_info(text: str, indent: int = 0):
    """Print info message."""
    spaces = " " * indent
    print(f"{spaces}{text}")

def print_error_msg(text: str, indent: int = 0):
    """Print error message."""
    spaces = " " * indent
    print(f"{spaces}{ERROR}✗ {text}{RESET}")

def print_warning_msg(text: str, indent: int = 0):
    """Print warning message."""
    spaces = " " * indent
    print(f"{spaces}{WARN}⚠ {text}{RESET}")

def print_pre_summary(config, employee_count: int):
    """Print pre-execution summary."""
    print_phase("Configuration Summary")
    print()
    excel_short = config.excel_path.name if hasattr(config.excel_path, 'name') else str(config.excel_path).split('\\')[-1]
    print(f"  Excel file       : {excel_short}")
    print(f"  Payroll date     : {config.date}")
    print(f"  Employees        : {employee_count}")
    print(f"  Outlook account  : {config.outlook_account}")
    print(f"  Dry run          : {'Yes' if config.dry_run else 'No'}")
    print(f"  PDF password     : {'Enabled' if config.pdf_password_enabled else 'Disabled'}")
    output_short = config.output_dir.name if hasattr(config.output_dir, 'name') else 'output/'
    print(f"  Output directory : {output_short}")
    
    # Log to file for debugging account issues
    logger.info(f"=== CONFIGURATION ===")
    logger.info(f"Outlook account: {config.outlook_account}")
    logger.info(f"Dry run: {config.dry_run}")
    logger.info(f"Total employees: {employee_count}")


def print_post_summary(stats: dict):
    """Print post-execution summary."""
    print(f"\n{OK}================================================================")
    print(f"  FINAL SUMMARY                                                 ")
    print(f"================================================================")
    print()
    print(f"  Employees           : {stats.get('total', 0)}")
    gen_skipped = stats.get('gen_skipped', 0)
    print(f"  Payslips generated  : {stats.get('generated', 0)}" + (f" (skipped: {gen_skipped})" if gen_skipped > 0 else ""))
    pdf_skipped = stats.get('pdf_skipped', 0)
    print(f"  PDFs converted      : {stats.get('converted', 0)}" + (f" (skipped: {pdf_skipped})" if pdf_skipped > 0 else ""))
    print(f"  Emails sent         : {stats.get('sent', 0)}{RESET}")
    errors = stats.get('errors', 0)
    if errors > 0:
        print(f"  {ERROR}Errors              : {errors}{RESET}")
    else:
        print(f"{OK}  Errors              : 0{RESET}")
    print()
    elapsed = stats.get("elapsed", 0)
    print(f"{OK}  Time elapsed        : {elapsed:.1f}s{RESET}")
    if stats.get("result_file"):
        result_short = stats['result_file'].name if hasattr(stats['result_file'], 'name') else str(stats['result_file']).split('\\')[-1]
        print(f"{OK}  Results file        : {result_short}{RESET}")
    print()
    if errors == 0:
        print(f"{OK}✔ Done{RESET}")
    else:
        print(f"{WARN}⚠ Done with errors{RESET}")
    print()


def confirm_proceed() -> bool:
    """Ask user for confirmation before sending."""
    print_phase("Confirmation")
    print()
    while True:
        answer = input("  Proceed with payslip generation and email sending? (yes/no): ").strip().lower()
        if answer in ("yes", "y"):
            return True
        if answer in ("no", "n"):
            return False
        print("  Please enter 'yes' or 'no'.")


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
    print()
    print(f"{WARN}▶ Existing Payroll Detected{RESET}")
    print()
    print(f"  Payroll for {config.date} has partial state:")
    print(f"    • Total employees    : {total_employees}")
    print(f"    • Already processed  : {state_info['sent_count']}")
    print(f"    • Remaining          : {total_employees - state_info['sent_count']}")
    print()
    print("  How to continue:")
    print("    yes - Continue from last checkpoint (resume)")
    print("    no  - Exit without changes")
    print(f"    new - Clean state and start fresh")
    
    while True:
        answer = input(f"\n  Your choice (yes/no/new): ").strip().lower()
        if answer in ("yes", "y"):
            return "yes"
        elif answer in ("no", "n"):
            return "no"
        elif answer == "new":
            return "new"
        print("  Please enter 'yes', 'no', or 'new'.")


def cleanup_output_files(config):
    """
    Delete all output files for the current payroll date.
    """
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
    tool_dir = Path(__file__).resolve().parent
    config = load_config(tool_dir=tool_dir)

    config_errors = config.validate()
    if config_errors:
        print()
        print_error_msg("Configuration errors:")
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
    logger.info("Configuration loaded successfully")
    logger.info(f"Excel path: {config.excel_path}")
    logger.info(f"Date: {config.date}")
    print_success("Configuration loaded")

    # ─── 2. Read Employee Metadata ───
    print_phase("Input & Validation")
    print()
    print_info("Reading employee data…")
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
        print_error_msg(f"Failed to read Excel file: {e}")
        sys.exit(1)

    # Allow ExcelReader's COM to fully release before generator starts
    gc.collect()
    time.sleep(1)

    if not employees:
        logger.error("No employee data found")
        print_error_msg("No employee data found in the Excel file")
        sys.exit(1)

    print_success(f"Found {len(employees)} employee" + ("s" if len(employees) != 1 else ""))
    logger.info(f"Found {len(employees)} employees")

    # ─── 2.5. Check for Existing State ───
    state_info = analyze_existing_state(config)
    if state_info["has_state"]:
        action = prompt_state_action(config, state_info, len(employees))
        if action == "no":
            print_info("Exited by user.")
            sys.exit(0)
        elif action == "new":
            cleanup_all_files(config)
            print_info("Starting fresh with clean state.")
        else:  # action == "yes"
            print_info("Resuming from last checkpoint.")

    # ─── 3. Validate Data ───
    print()
    print_info("Validating data…")
    validator = DataValidator(
        employees,
        allow_duplicate_emails=config.allow_duplicate_emails,
    )
    errors, warnings = validator.validate_all()

    if errors:
        print()
        print_error_msg(f"Validation FAILED with {len(errors)} error(s):")
        for err in errors:
            print(f"      {err}")
        print()
        print("  Please fix the data and try again.")
        sys.exit(1)

    if warnings:
        print_warning_msg(f"{len(warnings)} warning(s) found (non-blocking)")
        for w in warnings[:5]:
            print(f"      {w}")
        if len(warnings) > 5:
            print(f"      ... and {len(warnings) - 5} more (see log file)")

    print_success("Validation passed")

    # ─── Pre-Execution Summary & Confirmation ───
    print_pre_summary(config, len(employees))

    if not config.dry_run:
        if not confirm_proceed():
            print(f"\n{ERROR}Aborted by user.{RESET}")
            sys.exit(0)
    else:
        print_phase("Dry Run Mode")
        print_info("Simulating — no emails will be sent\n")

    # Clean up output files for new session
    print_info("Cleaning up output files…")
    cleanup_output_files(config)

    # ─── 4. Generate Payslip Excel Files via COM ───
    print_phase("Generating payslips")
    generator = PayslipGenerator(
        output_dir=config.output_dir,
        date_str=config.date,
        filename_pattern=config.pdf_filename_pattern,
    )

    interval = _progress_interval(len(employees))

    def gen_progress(current, total, name, skipped=False):
        action = "Skipped (exists)" if skipped else "✓"
        if current == 1 or current == total or current % interval == 0:
            print(f"  [{current}/{total}] {OK if not skipped else WARN}{action}{RESET} {name}")

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
        print_error_msg(f"Payslip generation failed: {e}")
        sys.exit(1)

    gen_elapsed = time.time() - gen_start
    generated = sum(1 for r in results if r["success"] and not r.get("skipped"))
    gen_skipped = sum(1 for r in results if r.get("skipped"))
    gen_failed = sum(1 for r in results if not r["success"])
    print()
    print_info(f"Result: Generated {generated} | Skipped {gen_skipped} | Failed {gen_failed} ({gen_elapsed:.1f}s)")

    # Allow Excel COM to fully release before starting PDF converter
    gc.collect()
    time.sleep(2)

    # ─── 5. Convert to Password-Protected PDFs ───
    print_phase("Converting to PDF")
    successful_items = [r for r in results if r["success"]]

    def pdf_progress(current, total, name, skipped=False):
        action = "Skipped (exists)" if skipped else "✓"
        if current == 1 or current == total or current % interval == 0:
            print(f"  [{current}/{total}] {OK if not skipped else WARN}{action}{RESET} {name}")

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
        print_error_msg(f"PDF conversion failed: {e}")

    pdf_elapsed = time.time() - pdf_start
    converted = sum(1 for r in results if r.get("pdf_path"))
    pdf_skipped = sum(1 for r in results if r.get("pdf_skipped"))
    pdf_failed = len(successful_items) - converted
    print()
    print_info(f"Result: Converted {converted - pdf_skipped} | Skipped {pdf_skipped} | Failed {pdf_failed} ({pdf_elapsed:.1f}s)")

    # ─── 6. Compose and Send Emails ───
    print_phase("Sending emails")
    print_info("Composing emails…")
    composer = EmailComposer(
        template_cells=email_template,
        subject=subject,
        date_str=config.date,
        date_cell=config.email_date_cell,
    )
    results = composer.compose_batch(results)
    composed = sum(1 for r in results if r.get("email_data"))
    print_info(f"Sending via Outlook…")
    print()

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
        print_info(f"Resuming: {resumed_count} employees already processed")

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
                            print(f"  [{i}/{composed}] {OK}✓{RESET} {name}")
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

            print()
            print_info(
                f"Sender stats: Sent {sender.sent_count} | "
                f"Skipped {sender.skipped_count} | "
                f"Errors {sender.error_count}"
            )

    except ImportError:
        logger.error("win32com not available - cannot send emails")
        print_error_msg("Outlook COM not available.")
        if config.dry_run:
            print_info("[DRY-RUN] Would have sent emails. Skipping Outlook.")
            sent_count = composed
    except Exception as e:
        logger.error(f"Email sending failed: {e}")
        print_error_msg(f"Email sending failed: {e}")

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
        print(f"\n{WARN}⚠ WARNING: {error_count} error(s) occurred. Check logs for details.{RESET}")
        sys.exit(1)

    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
