"""
Core utilities for the Office Automation Foundation.

This package provides foundational components:
- Configuration loading and validation
- Logging (file + console) with custom CONSOLE level
- Console output with standardized formatting
- Retry mechanisms with exponential backoff
- State tracking for duplicate prevention
"""

from core.config import Config, load_config
from core.config_manager import ConfigManager
from core.console import cprint, cprint_banner, cprint_summary_box
from core.logger import get_logger, setup_logging, CONSOLE
from core.retry import retry_operation, RetryConfig
from core.state import StateTracker

__all__ = [
    "Config",
    "ConfigManager",
    "load_config",
    "cprint",
    "cprint_banner",
    "cprint_summary_box",
    "confirm_proceed",
    "CONSOLE",
    "get_logger",
    "setup_logging",
    "retry_operation",
    "RetryConfig",
    "StateTracker",
]
