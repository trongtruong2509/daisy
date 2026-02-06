"""
Core utilities for the Office Automation Foundation.

This package provides foundational components:
- Configuration loading and validation
- Logging (file + console)
- Retry mechanisms with exponential backoff
- State tracking for duplicate prevention
"""

from core.config import Config, load_config
from core.logger import get_logger, setup_logging
from core.retry import retry_operation, RetryConfig
from core.state import StateTracker

__all__ = [
    "Config",
    "load_config",
    "get_logger",
    "setup_logging",
    "retry_operation",
    "RetryConfig",
    "StateTracker",
]
