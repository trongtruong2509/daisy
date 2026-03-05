"""
Configuration management for the Get Attachment tool.

Loads settings from local .env (tool-specific) with fallback to global .env.
Prompts the user for any required values not present in the .env files.
Values entered interactively are NOT saved back to the .env file.
"""

import re
import sys
from dataclasses import dataclass, field
from datetime import datetime, date
from pathlib import Path
from typing import List, Optional, Tuple

_TODAY_STR = lambda: date.today().strftime("%d/%m/%Y")

# Add project root to path for foundation imports
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.config_manager import ConfigManager
from core.console import cprint
from office.outlook import OutlookClient


# ── Validators ───────────────────────────────────────────────────

def _validate_date_ddmmyyyy(value: str) -> Tuple[bool, str]:
    """Validate a date string in DD/MM/YYYY format."""
    value = value.strip()
    if not re.match(r"^\d{2}/\d{2}/\d{4}$", value):
        return False, "Expected format DD/MM/YYYY (e.g., 02/03/2026)"
    try:
        datetime.strptime(value, "%d/%m/%Y")
    except ValueError as exc:
        return False, f"Invalid date value: {exc}"
    return True, ""


# ── Account selection ────────────────────────────────────────────

def _prompt_for_outlook_account() -> str:
    """
    Prompt the user to choose an Outlook account from those configured in Outlook.

    Falls back to manual email entry if no accounts are detected.

    Returns:
        Selected email address string.
    """
    accounts = OutlookClient.get_available_accounts()

    print()
    if not accounts:
        cprint("No Outlook accounts detected. Enter email manually.", level="WARNING")
        return ConfigManager.prompt_for_value(
            "OUTLOOK_ACCOUNT",
            "Enter your Outlook email address",
            "user@company.com",
            validator=ConfigManager.validate_email,
        )

    cprint("Choose Outlook account", level="WARNING")
    for i, account in enumerate(accounts, 1):
        print(f"  [{i}] {account}")

    while True:
        try:
            choice = input("\u2192 Select: ").strip()
            if not choice:
                print("  Selection cannot be empty. Please try again.")
                continue
            choice_num = int(choice)
            if 1 <= choice_num <= len(accounts):
                return accounts[choice_num - 1]
            cprint(f"Invalid selection. Please enter a number between 1 and {len(accounts)}.", level="ERROR")
        except ValueError:
            cprint("Please enter a valid number.", level="ERROR")


# ── Config dataclass ─────────────────────────────────────────────

@dataclass
class GetAttachmentConfig:
    """Configuration for the Get Attachment tool."""

    # Outlook account to use
    outlook_account: str = ""

    # Outlook folder path (e.g. "Inbox/SubFolder1/SubFolder2")
    # Empty string means default Inbox
    outlook_folder: str = ""

    # Search criteria — date range from start_date to end_date (inclusive)
    # Dates in DD/MM/YYYY format
    start_date: str = ""
    # End date for the search range; empty string means today
    end_date: str = ""
    # Keywords to filter email subjects (OR logic); empty list means no filter
    subject_keywords: List[str] = field(default_factory=list)

    # Output
    attachment_save_path: Path = field(default_factory=lambda: Path("./attachments"))

    # Logging
    log_dir: Path = field(default_factory=lambda: Path("./logs"))
    log_level: str = "INFO"

    def __post_init__(self) -> None:
        if isinstance(self.attachment_save_path, str):
            self.attachment_save_path = Path(self.attachment_save_path)
        if isinstance(self.log_dir, str):
            self.log_dir = Path(self.log_dir)
        self.log_level = self.log_level.upper()

    # ── Derived properties ────────────────────────────

    @property
    def start_date_parsed(self) -> Optional[date]:
        """Return start_date as a :class:`datetime.date` object, or ``None`` on failure."""
        if not self.start_date:
            return None
        try:
            return datetime.strptime(self.start_date.strip(), "%d/%m/%Y").date()
        except ValueError:
            return None

    @property
    def end_date_parsed(self) -> date:
        """
        Return end_date as a :class:`datetime.date` object.

        Returns today's date when ``end_date`` is empty, so callers never
        receive ``None`` — the end of the search range is always defined.
        """
        if not self.end_date:
            return date.today()
        try:
            return datetime.strptime(self.end_date.strip(), "%d/%m/%Y").date()
        except ValueError:
            return date.today()

    @property
    def date_range_display(self) -> str:
        """Human-readable date range string, e.g. '01/03/2026 → 02/03/2026'."""
        end = self.end_date if self.end_date else _TODAY_STR()
        return f"{self.start_date} \u2192 {end}"

    # ── Validation ────────────────────────────────────

    def validate(self) -> List[str]:
        """Return a list of error messages; empty list means config is valid."""
        errors: List[str] = []
        if not self.outlook_account:
            errors.append("OUTLOOK_ACCOUNT is required")
        if not self.start_date:
            errors.append("START_DATE is required (format: DD/MM/YYYY)")
        elif self.start_date_parsed is None:
            errors.append(
                f"START_DATE format invalid: '{self.start_date}'. Expected DD/MM/YYYY"
            )
        if self.end_date:
            if not self._end_date_is_parseable():
                errors.append(
                    f"END_DATE format invalid: '{self.end_date}'. Expected DD/MM/YYYY"
                )
            elif self.start_date_parsed is not None:
                end = datetime.strptime(self.end_date.strip(), "%d/%m/%Y").date()
                if end < self.start_date_parsed:
                    errors.append(
                        f"END_DATE '{self.end_date}' must not be before "
                        f"START_DATE '{self.start_date}'"
                    )
        return errors

    def _end_date_is_parseable(self) -> bool:
        """Return True when end_date can be parsed as DD/MM/YYYY."""
        try:
            datetime.strptime(self.end_date.strip(), "%d/%m/%Y")
            return True
        except ValueError:
            return False

    def ensure_directories(self) -> None:
        """Create output directories if they do not already exist."""
        self.attachment_save_path.mkdir(parents=True, exist_ok=True)
        self.log_dir.mkdir(parents=True, exist_ok=True)


# ── Loader ───────────────────────────────────────────────────────

def load_config(tool_dir: Optional[Path] = None) -> "GetAttachmentConfig":
    """
    Load configuration from .env files and prompt for any missing required values.

    Priority: local .env (tool dir) > global .env (project root) > defaults.

    Prompted values are used for the current run only; they are NOT saved back
    to any .env file.

    Args:
        tool_dir: Directory containing the tool's .env file. Defaults to the
                  directory containing this module.

    Returns:
        Fully populated :class:`GetAttachmentConfig`.
    """
    if tool_dir is None:
        tool_dir = Path(__file__).resolve().parent
    else:
        tool_dir = Path(tool_dir).resolve()

    global_env = PROJECT_ROOT / ".env"
    local_env = tool_dir / ".env"

    mgr = ConfigManager()
    mgr.load_env([global_env, local_env])

    # --- Read values from env ---
    outlook_account = mgr.get("OUTLOOK_ACCOUNT", "")
    start_date = mgr.get("START_DATE", "")
    end_date_raw = mgr.get("END_DATE", "")

    keywords_raw = mgr.get("SUBJECT_KEYWORDS", "")
    subject_keywords: List[str] = (
        [k.strip() for k in keywords_raw.split(",") if k.strip()]
        if keywords_raw
        else []
    )

    save_path_raw = mgr.get("ATTACHMENT_SAVE_PATH", "")
    attachment_save_path = (
        Path(save_path_raw) if save_path_raw else tool_dir / "attachments"
    )

    outlook_folder = mgr.get("OUTLOOK_FOLDER", "")

    config = GetAttachmentConfig(
        outlook_account=outlook_account,
        outlook_folder=outlook_folder,
        start_date=start_date,
        end_date=end_date_raw,
        subject_keywords=subject_keywords,
        attachment_save_path=attachment_save_path,
        log_dir=tool_dir / Path(mgr.get("LOG_DIR", "./logs")),
        log_level=mgr.get("LOG_LEVEL", "INFO"),
    )

    # --- Interactive prompts for missing values ---

    needs_prompt = (
        not config.outlook_account
        or not config.start_date
        or not end_date_raw
        or not save_path_raw
    )
    if needs_prompt:
        cprint("Configuration", level="PHASE")
        print()

    if not config.outlook_account:
        config.outlook_account = _prompt_for_outlook_account()

    if not config.start_date:
        config.start_date = ConfigManager.prompt_for_value(
            "START_DATE",
            "Enter the start date for email search (DD/MM/YYYY)",
            _TODAY_STR(),
            validator=_validate_date_ddmmyyyy,
        )

    # END_DATE — optional; empty means today
    if not end_date_raw:
        today_str = _TODAY_STR()
        print()
        cprint(
            f"Enter end date for email search (DD/MM/YYYY) — leave empty to use today ({today_str})",
            level="WARNING",
        )
        while True:
            raw = input(f"END_DATE [{today_str}]: ").strip()
            if not raw:
                # Empty → stay as "" so end_date_parsed returns today
                break
            ok, msg = _validate_date_ddmmyyyy(raw)
            if ok:
                config.end_date = raw
                break
            cprint(f"Invalid date: {msg}", level="ERROR")

    # OUTLOOK_FOLDER — always prompt so user can change each run
    # Show current value (from .env) and allow user to confirm or re-enter
    print()
    current_folder = config.outlook_folder or "Inbox"
    cprint(
        "Enter the Outlook folder path to read attachments from "
        "(e.g. Inbox/SubFolder1/SubFolder2).",
        level="WARNING",
    )
    cprint(
        f"Leave blank to use the current value: {current_folder}",
        level="INFO",
    )
    folder_input = input("OUTLOOK_FOLDER: ").strip()
    if folder_input:
        config.outlook_folder = folder_input
        # Save to .env for next time
        ConfigManager.save_to_env(local_env, "OUTLOOK_FOLDER", folder_input)
    elif not config.outlook_folder:
        config.outlook_folder = "Inbox"

    # Subject keywords are optional; prompt only if not set in env
    if not keywords_raw:
        print()
        cprint(
            "Enter subject keywords to filter emails (comma-separated, leave empty for all emails)",
            level="WARNING",
        )
        raw_input = input("SUBJECT_KEYWORDS: ").strip()
        if raw_input:
            config.subject_keywords = [k.strip() for k in raw_input.split(",") if k.strip()]

    if not save_path_raw:
        save_path_str = ConfigManager.prompt_for_value(
            "ATTACHMENT_SAVE_PATH",
            "Enter the directory path where attachments will be saved",
            r"D:\Downloads\attachments",
        )
        config.attachment_save_path = Path(save_path_str)

    return config
