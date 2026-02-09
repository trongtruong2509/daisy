import csv
import sys
from datetime import datetime
from pathlib import Path

# Add project root to path for foundation imports
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.console import cprint
from core.logger import get_logger

logger = get_logger(__name__)

SENT_RESULT_FILE_PREFIX = "sent_results_"

# ── Progress Utilities ──────────────────────────────────────────

def progress_interval(total: int) -> int:
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


# ── CSV Result Writer ───────────────────────────────────────────

CSV_COLUMNS = [
    "employee_id",
    "employee_name",
    "email_address",
    "payslip_filename",
    "sent_status",
    "timestamp",
    "error_message",
]


class ResultWriter:
    """
    Writes per-employee processing results to a CSV file.

    Supports append mode for resume/checkpoint compatibility.
    """

    def __init__(self, output_path: Path, date_str: str):
        self.output_path = Path(output_path)
        # Write header only if file doesn't exist yet
        if not self.output_path.exists():
            self.output_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.output_path, "w", encoding="utf-8", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(CSV_COLUMNS)

    def append(
        self,
        mnv: str,
        name: str,
        email: str,
        status: str,
        payslip_filename: str = "",
        error_message: str = "",
    ):
        """Append a result row to the CSV file."""
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(self.output_path, "a", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow([mnv, name, email, payslip_filename, status, ts, error_message])


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
    result_file = config.output_dir / f"{SENT_RESULT_FILE_PREFIX}{config.date_mmyyyy}.csv"
    
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
                reader = csv.reader(f)
                next(reader, None)  # Skip header
                result["total_in_results"] = sum(1 for _ in reader)
        except Exception:
            pass
    
    return result


def prompt_state_action(config, state_info: dict, total_employees: int) -> str:
    """
    Prompt user for action when existing state is found.
    
    Returns: "yes", "no", or "new"
    """
    print()
    cprint("Existing Payroll Detected", level="WARNING")
    print()
    cprint(f"Payroll for {config.date} has partial state:", level="INFO", indent=2)
    cprint(f"  Total employees    : {total_employees}", level="PRE_SUMMARY", indent=2)
    cprint(f"  Already processed  : {state_info['sent_count']}", level="PRE_SUMMARY", indent=2)
    cprint(f"  Remaining          : {total_employees - state_info['sent_count']}", level="PRE_SUMMARY", indent=2)
    print()
    print("  How to continue:")
    print("    yes - Continue from last checkpoint (resume)")
    print("    no  - Exit without changes")
    print("    new - Clean state and start fresh")
    
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
    """Delete all output files for the current payroll date."""
    if config.output_dir.exists():
        patterns = ("*.pdf", "*.txt", "*.csv")
        for pattern in patterns:
            for file in config.output_dir.glob(pattern):
                if file.is_file():
                    file.unlink()

    
def cleanup_all_files(config):
    """Delete all state and output files for the current payroll date."""
    # Delete checkpoint files
    for mode in ["dryrun", "send"]:
        checkpoint = config.state_dir / f"payslip_checkpoint_{mode}_{config.date_mmyyyy}_state.json"
        if checkpoint.exists():
            checkpoint.unlink()

    # Delete state file
    state_file = config.state_dir / f"payslip_send_{config.date_mmyyyy}_state.json"
    if state_file.exists():
        state_file.unlink()

    cleanup_output_files(config)


def cleanup_pdf(pdf_path: Path) -> None:
    """Delete a PDF file after successful email send."""
    try:
        if pdf_path and Path(pdf_path).exists():
            Path(pdf_path).unlink()
            logger.debug(f"Cleaned up PDF: {pdf_path}")
    except Exception as e:
        logger.warning(f"Failed to delete PDF {pdf_path}: {e}")

def confirm_proceed() -> bool:
    """Ask user for confirmation before sending."""
    cprint("Confirmation", level="PHASE")
    while True:
        answer = input("  Proceed with payslip generation and email sending? (yes/no): ").strip().lower()
        if answer in ("yes", "y"):
            return True
        if answer in ("no", "n"):
            return False
        print("  Please enter 'yes' or 'no'.")