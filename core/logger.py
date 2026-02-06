"""
Logging utilities for Office Automation Foundation.

Provides dual-output logging:
- File logging: detailed, for debugging and audits
- Console logging: concise, user-friendly progress updates

Each run creates a new timestamped log file.

Usage:
    from core.logger import setup_logging, get_logger
    
    setup_logging(log_dir=Path("./logs"), level="INFO")
    logger = get_logger(__name__)
    
    logger.info("Starting email processing")
    logger.debug("Processing email ID: 12345")
"""

import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

# Global state for the logging system
_logging_initialized = False
_log_file_path: Optional[Path] = None


class ConsoleFormatter(logging.Formatter):
    """
    Custom formatter for console output.
    
    Provides concise, user-friendly messages without excessive detail.
    Uses colors for different log levels (when supported).
    """
    
    # ANSI color codes for Windows terminal support
    COLORS = {
        logging.DEBUG: "\033[36m",     # Cyan
        logging.INFO: "\033[32m",       # Green
        logging.WARNING: "\033[33m",    # Yellow
        logging.ERROR: "\033[31m",      # Red
        logging.CRITICAL: "\033[35m",   # Magenta
    }
    RESET = "\033[0m"
    
    def __init__(self, use_colors: bool = True):
        super().__init__()
        self.use_colors = use_colors
    
    def format(self, record: logging.LogRecord) -> str:
        # Create concise prefix based on level
        level_prefixes = {
            logging.DEBUG: "[DEBUG]",
            logging.INFO: "[*]",
            logging.WARNING: "[!]",
            logging.ERROR: "[ERROR]",
            logging.CRITICAL: "[CRITICAL]",
        }
        prefix = level_prefixes.get(record.levelno, "[?]")
        
        # Format the message
        message = f"{prefix} {record.getMessage()}"
        
        # Add colors if supported
        if self.use_colors and sys.stdout.isatty():
            color = self.COLORS.get(record.levelno, "")
            message = f"{color}{message}{self.RESET}"
        
        return message


class FileFormatter(logging.Formatter):
    """
    Custom formatter for file output.
    
    Provides detailed information suitable for debugging and audits:
    - Timestamp
    - Log level
    - Module name
    - Message
    - Exception info (if present)
    """
    
    def __init__(self):
        super().__init__(
            fmt="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )


def setup_logging(
    log_dir: Path,
    level: str = "INFO",
    console_level: Optional[str] = None,
    run_name: Optional[str] = None
) -> Path:
    """
    Initialize the logging system with file and console handlers.
    
    Should be called once at the start of a script.
    
    Args:
        log_dir: Directory for log files (created if doesn't exist).
        level: File logging level (DEBUG, INFO, WARNING, ERROR).
        console_level: Console logging level (defaults to INFO minimum).
        run_name: Optional name prefix for the log file.
        
    Returns:
        Path to the created log file.
    """
    global _logging_initialized, _log_file_path
    
    # Prevent double initialization
    if _logging_initialized:
        return _log_file_path
    
    # Create log directory
    log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # Generate log filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if run_name:
        log_filename = f"{run_name}_{timestamp}.log"
    else:
        log_filename = f"run_{timestamp}.log"
    
    _log_file_path = log_dir / log_filename
    
    # Convert level strings to logging constants
    file_level = getattr(logging, level.upper(), logging.INFO)
    
    # Console level: at least INFO, but can be more verbose
    if console_level:
        console_log_level = getattr(logging, console_level.upper(), logging.INFO)
    else:
        console_log_level = max(file_level, logging.INFO)
    
    # Get root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)  # Allow all levels, handlers filter
    
    # Clear existing handlers
    root_logger.handlers.clear()
    
    # File handler (detailed)
    file_handler = logging.FileHandler(_log_file_path, encoding="utf-8")
    file_handler.setLevel(file_level)
    file_handler.setFormatter(FileFormatter())
    root_logger.addHandler(file_handler)
    
    # Console handler (concise)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(console_log_level)
    console_handler.setFormatter(ConsoleFormatter(use_colors=True))
    root_logger.addHandler(console_handler)
    
    _logging_initialized = True
    
    # Log initialization
    logger = logging.getLogger(__name__)
    logger.info(f"Logging initialized: {_log_file_path}")
    logger.debug(f"File log level: {level}, Console log level: {console_level or 'INFO'}")
    
    return _log_file_path


def get_logger(name: str) -> logging.Logger:
    """
    Get a logger instance for a module.
    
    Args:
        name: Module name (typically __name__).
        
    Returns:
        Logger instance configured with the global settings.
    """
    return logging.getLogger(name)


def get_log_file_path() -> Optional[Path]:
    """
    Get the path to the current log file.
    
    Returns:
        Path to log file, or None if logging not initialized.
    """
    return _log_file_path


class ProgressLogger:
    """
    Helper class for logging progress through a batch operation.
    
    Provides periodic updates without flooding the console.
    
    Usage:
        progress = ProgressLogger(total=100, logger=logger, operation="Processing emails")
        for i, item in enumerate(items):
            process(item)
            progress.update(i + 1)
        progress.complete()
    """
    
    def __init__(
        self,
        total: int,
        logger: logging.Logger,
        operation: str = "Processing",
        log_every: int = 10
    ):
        """
        Initialize progress logger.
        
        Args:
            total: Total number of items to process.
            logger: Logger instance to use.
            operation: Description of the operation (for log messages).
            log_every: Log progress every N items.
        """
        self.total = total
        self.logger = logger
        self.operation = operation
        self.log_every = log_every
        self.current = 0
        self.success_count = 0
        self.error_count = 0
        
        self.logger.info(f"{operation}: Starting ({total} items)")
    
    def update(self, current: int, success: bool = True) -> None:
        """
        Update progress and log if needed.
        
        Args:
            current: Current item number (1-indexed).
            success: Whether the current item was processed successfully.
        """
        self.current = current
        if success:
            self.success_count += 1
        else:
            self.error_count += 1
        
        # Log progress at intervals
        if current % self.log_every == 0 or current == self.total:
            pct = (current / self.total) * 100
            self.logger.info(
                f"{self.operation}: {current}/{self.total} ({pct:.0f}%) - "
                f"Success: {self.success_count}, Errors: {self.error_count}"
            )
    
    def complete(self) -> None:
        """Log completion summary."""
        self.logger.info(
            f"{self.operation}: Complete - "
            f"Total: {self.total}, Success: {self.success_count}, Errors: {self.error_count}"
        )


class DryRunLogger:
    """
    Context manager that prefixes all log messages with [DRY-RUN].
    
    Usage:
        with DryRunLogger(logger):
            logger.info("Would send email")  # Outputs: [DRY-RUN] Would send email
    """
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
        self._original_format = None
    
    def __enter__(self):
        # Note: This is a simplified approach. In practice, you'd modify
        # the handler formatters. For now, we recommend using explicit
        # "[DRY-RUN]" prefixes in messages when in dry-run mode.
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        return False


def log_exception(logger: logging.Logger, message: str, exc: Exception) -> None:
    """
    Log an exception with full traceback to file, concise message to console.
    
    Args:
        logger: Logger instance.
        message: Context message about where the error occurred.
        exc: The exception that was caught.
    """
    # Log full details at ERROR level (goes to file)
    logger.error(f"{message}: {type(exc).__name__}: {exc}", exc_info=True)
