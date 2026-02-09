"""
Console output module for Office Automation Foundation.

Provides standardized, colored console output with automatic logging
at the custom CONSOLE level for file traceability.

All console output goes through cprint() for consistency.
Each call also logs the message via logger.console() so it appears
in log files alongside regular log entries.

Usage:
    from core.console import cprint

    cprint("Starting process", level="PHASE")
    cprint("File loaded", level="SUCCESS")
    cprint("Something went wrong", level="ERROR")
    cprint("Heads up", level="WARNING", indent=2)
"""

import logging

from core.logger import CONSOLE

logger = logging.getLogger("core.console")

# ── ANSI Color Codes ────────────────────────────────────────────
CC_PHASE = "\033[96m"     # Bright cyan for phase headers
CC_OK = "\033[92m"        # Bright green for success
CC_WARN = "\033[93m"      # Yellow for warnings
CC_ERROR = "\033[91m"     # Bright red for errors
CC_INFO = "\033[37m"      # White for info
CC_BOX = "\033[96m"       # Bright cyan for boxes/banners
CC_RESET = "\033[0m"

# Box-drawing characters
_BOX_H = "═"
_BOX_TL = "╔"
_BOX_TR = "╗"
_BOX_BL = "╚"
_BOX_BR = "╝"
_BOX_V = "║"
_BOX_WIDTH = 64


def cprint(message: str, level: str = "INFO", indent: int = 0) -> None:
    """
    Print message to console with color formatting and log to file.

    Args:
        message: Text to display.
        level: Output style. One of:
            INFO, BANNER, PHASE, SUCCESS, ERROR, WARNING,
            PRE_SUMMARY, SUMMARY, PROGRESS
        indent: Number of leading spaces.
    """
    spaces = " " * indent
    level = level.upper()

    if level == "BANNER":
        _print_banner(message)
    elif level == "PHASE":
        _print_phase(message, spaces)
    elif level == "SUCCESS":
        _print_success(message, spaces)
    elif level == "ERROR":
        _print_error(message, spaces)
    elif level == "WARNING":
        _print_warning(message, spaces)
    elif level == "SUMMARY":
        _print_summary_line(message, spaces)
    elif level == "PRE_SUMMARY":
        _print_pre_summary_line(message, spaces)
    elif level == "PROGRESS":
        _print_progress(message, spaces)
    else:  # INFO and fallback
        _print_info(message, spaces)


# ── Internal Formatters ─────────────────────────────────────────

def _print_info(message: str, spaces: str) -> None:
    print(f"{spaces}{CC_INFO}{message}{CC_RESET}")
    logger.log(CONSOLE, message)


def _print_phase(title: str, spaces: str) -> None:
    print(f"\n{spaces}{CC_PHASE}▶ {title}{CC_RESET}")
    logger.log(CONSOLE, f"▶ {title}")


def _print_success(message: str, spaces: str) -> None:
    print(f"{spaces}{CC_OK}✓ {message}{CC_RESET}")
    logger.log(CONSOLE, f"✓ {message}")


def _print_error(message: str, spaces: str) -> None:
    print(f"{spaces}{CC_ERROR}✗ {message}{CC_RESET}")
    logger.log(CONSOLE, f"✗ {message}")


def _print_warning(message: str, spaces: str) -> None:
    print(f"{spaces}{CC_WARN}⚠ {message}{CC_RESET}")
    logger.log(CONSOLE, f"⚠ {message}")


def _print_progress(message: str, spaces: str) -> None:
    print(f"{spaces}{CC_INFO}{message}{CC_RESET}")
    logger.log(CONSOLE, message)


def _print_summary_line(message: str, spaces: str) -> None:
    print(f"{spaces}{CC_OK}{message}{CC_RESET}")
    logger.log(CONSOLE, message)


def _print_pre_summary_line(message: str, spaces: str) -> None:
    print(f"{spaces}{message}")
    logger.log(CONSOLE, message)


def _print_banner(message: str) -> None:
    """Print a boxed banner header."""
    lines = message.strip().split("\n")
    width = max(len(line) for line in lines)
    box_w = max(width + 4, _BOX_WIDTH)
    inner = box_w - 2  # inside the box walls

    print(f"\n{CC_BOX}{_BOX_TL}{_BOX_H * inner}{_BOX_TR}{CC_RESET}")
    for line in lines:
        padded = f"  {line}".ljust(inner)
        print(f"{CC_BOX}{_BOX_V}{padded}{_BOX_V}{CC_RESET}")
    print(f"{CC_BOX}{_BOX_BL}{_BOX_H * inner}{_BOX_BR}{CC_RESET}")

    logger.log(CONSOLE, f"[BANNER] {message}")


# ── High-Level Helpers ──────────────────────────────────────────

def cprint_banner(title: str, subtitle: str = "") -> None:
    """Print a tool banner with title and optional subtitle."""
    text = title
    if subtitle:
        text += f"\n{subtitle}"
    cprint(text, level="BANNER")


def cprint_summary_box(title: str, items: dict, footer: str = "") -> None:
    """
    Print a summary box with key-value pairs.

    Args:
        title: Box header text.
        items: Dict of label -> value to display.
        footer: Optional footer message.
    """
    bar = "=" * _BOX_WIDTH
    cprint(f"\n{bar}", level="SUMMARY")
    cprint(f"  {title}", level="SUMMARY")
    cprint(bar, level="SUMMARY")
    print()
    for label, value in items.items():
        cprint(f"  {label:<22}: {value}", level="SUMMARY", indent=0)
    print()
    if footer:
        cprint(footer, level="SUMMARY")
    cprint(bar, level="SUMMARY")


def cprint_summary_box_lite(title: str, items: dict, footer: str = "") -> None:
    """
    Print a summary box with key-value pairs.

    Args:
        title: Box header text.
        items: Dict of label -> value to display.
        footer: Optional footer message.
    """
    bar = "-" * _BOX_WIDTH
    cprint(f"\n{bar}", level="PRE_SUMMARY")
    cprint(f"  {title}", level="PRE_SUMMARY")
    cprint(bar, level="PRE_SUMMARY")
    for label, value in items.items():
        cprint(f"  {label:<22}: {value}", level="PRE_SUMMARY", indent=0)
    if footer:
        cprint(footer, level="PRE_SUMMARY")
    cprint(bar, level="PRE_SUMMARY")


# def confirm_proceed(prompt_text: str = "Proceed? (yes/no)") -> bool:
#     """
#     Ask user for yes/no confirmation.

#     Args:
#         prompt_text: The question to display.

#     Returns:
#         True if user confirms, False otherwise.
#     """
#     cprint(prompt_text, level="INFO")
#     # print()
#     while True:
#         answer = input("  → ").strip().lower()
#         if answer in ("yes", "y"):
#             return True
#         if answer in ("no", "n"):
#             return False
#         print("  Please enter 'yes' or 'no'.")
