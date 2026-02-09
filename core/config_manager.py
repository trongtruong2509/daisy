"""
Reusable configuration management for Office Automation Foundation.

Provides standardized .env loading, typed getters, interactive prompting,
and built-in validators. Tools use this module to eliminate duplicated
config utilities.

Usage:
    from core.config_manager import ConfigManager

    mgr = ConfigManager()
    mgr.load_env([global_env, local_env])

    date = mgr.get("DATE", default="")
    dry_run = mgr.get_bool("DRY_RUN", default=True)
    excel_path = mgr.get_path("EXCEL_PATH", base_dir=tool_dir)

    value = mgr.prompt_for_value(
        "DATE", "Set Payroll Date", "01/2026",
        validator=ConfigManager.validate_date,
    )
"""

import os
import re
from pathlib import Path
from typing import Any, Callable, List, Optional, Tuple

from dotenv import load_dotenv, set_key

from core.console import cprint


class ConfigManager:
    """
    Base configuration manager with .env loading, prompting, and validation.

    Handles the common patterns shared across all tools:
    - Loading .env files with priority ordering
    - Typed environment variable access
    - Interactive prompting with colored output and validation
    - Persisting values back to .env files
    """

    def __init__(self):
        self._loaded_files: List[Path] = []

    # ── .env Loading ────────────────────────────────────────

    def load_env(self, env_files: List[Path]) -> None:
        """
        Load .env files in order. Later files override earlier ones.

        Args:
            env_files: List of .env file paths, ordered from lowest
                       to highest priority (e.g., [global, local]).
        """
        for env_file in env_files:
            env_path = Path(env_file)
            if env_path.exists():
                # First file: don't override existing env vars
                # Subsequent files: override (higher priority)
                override = len(self._loaded_files) > 0
                load_dotenv(env_path, override=override)
                self._loaded_files.append(env_path)

    # ── Typed Getters ───────────────────────────────────────

    @staticmethod
    def get(key: str, default: str = "") -> str:
        """Get environment variable as string."""
        return os.getenv(key, default)

    @staticmethod
    def get_bool(key: str, default: bool = False) -> bool:
        """Get environment variable as boolean."""
        value = os.getenv(key)
        if value is None:
            return default
        return value.strip().lower() in ("true", "1", "yes", "on")

    @staticmethod
    def get_int(key: str, default: int = 0) -> int:
        """Get environment variable as integer."""
        value = os.getenv(key)
        if value is None or value.strip() == "":
            return default
        try:
            return int(value)
        except ValueError:
            return default

    @staticmethod
    def get_path(
        key: str,
        default: str = "",
        base_dir: Optional[Path] = None,
    ) -> Path:
        """
        Get environment variable as Path.

        Relative paths are resolved against base_dir if provided.
        """
        value = os.getenv(key, default)
        if not value:
            return Path(default) if default else Path("")
        path = Path(value)
        if base_dir and not path.is_absolute():
            path = Path(base_dir) / path
        return path

    @staticmethod
    def get_list(
        key: str,
        default: Optional[List[str]] = None,
        separator: str = ",",
    ) -> List[str]:
        """
        Get environment variable as a list of strings.

        Splits by separator and strips whitespace from each item.
        """
        value = os.getenv(key)
        if value is None:
            return default if default is not None else []
        return [item.strip() for item in value.split(separator) if item.strip()]

    # ── Interactive Prompting ───────────────────────────────

    @staticmethod
    def prompt_for_value(
        key: str,
        description: str,
        example: str = "",
        validator: Optional[Callable[[str], Tuple[bool, str]]] = None,
    ) -> str:
        """
        Prompt user for a configuration value with optional validation.

        Uses cprint() for colored output. Loops until a valid value is entered.

        Args:
            key: Config key name (shown as prompt label).
            description: User-facing description of what to enter.
            example: Example value to display.
            validator: Optional callable(value) -> (is_valid, error_msg).

        Returns:
            The validated user input string.
        """
        cprint(description, level="WARNING")
        if example:
            print(f"Example: {example}")
        while True:
            value = input(f"{key}: ").strip()
            if not value:
                print("Value cannot be empty. Please try again.")
                continue
            if validator:
                is_valid, error_msg = validator(value)
                if not is_valid:
                    cprint(f"Invalid input: {error_msg}", level="ERROR")
                    continue
            return value

    @staticmethod
    def save_to_env(env_file: Path, key: str, value: str) -> None:
        """
        Persist a key-value pair to a .env file.

        Creates the file if it doesn't exist.
        """
        try:
            env_path = Path(env_file)
            if not env_path.exists():
                env_path.touch()
            set_key(str(env_path), key, value)
            print(f"  Saved {key} to .env file")
        except Exception as e:
            print(f"  Warning: Could not save to .env: {e}")

    # ── Built-in Validators ─────────────────────────────────

    @staticmethod
    def validate_date(value: str) -> Tuple[bool, str]:
        """Validate MM/YYYY date format with valid month (01-12)."""
        if not re.match(r"^\d{2}/\d{4}$", value):
            return False, "Expected format: MM/YYYY (e.g., 01/2026)"
        try:
            month = int(value[:2])
            if month < 1 or month > 12:
                return False, f"Month must be 01-12, got {month}"
        except ValueError:
            return False, "Invalid month value"
        return True, ""

    @staticmethod
    def validate_email(value: str) -> Tuple[bool, str]:
        """Validate email address format."""
        pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
        if not re.match(pattern, value):
            return False, "Invalid email format (e.g., user@company.com)"
        return True, ""

    @staticmethod
    def validate_file_path(
        value: str,
        base_dir: Optional[Path] = None,
    ) -> Tuple[bool, str]:
        """
        Validate that a file path exists.

        Relative paths are resolved against base_dir if provided.
        """
        file_path = Path(value)
        if base_dir and not file_path.is_absolute():
            file_path = Path(base_dir) / file_path
        if not file_path.exists():
            return False, f"File not found: {file_path}"
        if not file_path.is_file():
            return False, f"Path is not a file: {file_path}"
        return True, ""
