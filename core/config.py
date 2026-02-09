"""
Configuration management for Office Automation Foundation.

Loads configuration from .env files and provides a typed Config object.
Supports environment-specific overrides and validation.

Usage:
    from core.config import load_config
    
    config = load_config()  # Loads from .env in project root
    print(config.outlook_account)
    print(config.dry_run)
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from core.config_manager import ConfigManager


@dataclass
class Config:
    """
    Configuration container for the automation framework.
    
    All settings are loaded from environment variables (via .env file).
    Provides sensible defaults for optional settings.
    
    Attributes:
        outlook_account: SMTP address of the Outlook account to use.
        dry_run: If True, no mutations are made to Outlook (read-only mode).
        batch_size: Number of emails to process per batch (default: 50).
        retry_count: Number of retry attempts for Outlook operations (default: 3).
        retry_delay_seconds: Base delay between retries in seconds (default: 2).
        log_dir: Directory for log files (default: ./logs).
        output_dir: Directory for output files (default: ./output).
        state_dir: Directory for state tracking files (default: ./state).
        log_level: Logging level (DEBUG, INFO, WARNING, ERROR) (default: INFO).
    """
    
    # Required settings (no defaults)
    outlook_account: str = ""
    
    # Safety settings
    dry_run: bool = True  # Default to safe mode
    
    # Batch processing
    batch_size: int = 50
    
    # Retry configuration
    retry_count: int = 3
    retry_delay_seconds: int = 2
    
    # Paths
    log_dir: Path = field(default_factory=lambda: Path("./logs"))
    output_dir: Path = field(default_factory=lambda: Path("./output"))
    state_dir: Path = field(default_factory=lambda: Path("./state"))
    
    # Logging
    log_level: str = "INFO"
    
    def __post_init__(self):
        """Validate and normalize configuration values."""
        # Ensure paths are Path objects
        if isinstance(self.log_dir, str):
            self.log_dir = Path(self.log_dir)
        if isinstance(self.output_dir, str):
            self.output_dir = Path(self.output_dir)
        if isinstance(self.state_dir, str):
            self.state_dir = Path(self.state_dir)
        
        # Normalize log level
        self.log_level = self.log_level.upper()
        
        # Validate batch size
        if self.batch_size < 1:
            self.batch_size = 50
        if self.batch_size > 1000:
            self.batch_size = 1000
        
        # Validate retry count
        if self.retry_count < 0:
            self.retry_count = 0
        if self.retry_count > 10:
            self.retry_count = 10

    def validate(self) -> list[str]:
        """
        Validate the configuration and return a list of errors.
        
        Returns:
            List of error messages. Empty list if valid.
        """
        errors = []
        
        if not self.outlook_account:
            errors.append("OUTLOOK_ACCOUNT is required but not set")
        elif "@" not in self.outlook_account:
            errors.append(f"OUTLOOK_ACCOUNT '{self.outlook_account}' does not appear to be a valid email address")
        
        if self.log_level not in ("DEBUG", "INFO", "WARNING", "ERROR"):
            errors.append(f"Invalid LOG_LEVEL: {self.log_level}")
        
        return errors

    def ensure_directories(self) -> None:
        """Create necessary directories if they don't exist."""
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.state_dir.mkdir(parents=True, exist_ok=True)

    def is_valid(self) -> bool:
        """Check if configuration is valid."""
        return len(self.validate()) == 0


def load_config(env_file: Optional[Path] = None) -> Config:
    """
    Load configuration from .env file and environment variables.
    
    Environment variables take precedence over .env file values.
    
    Args:
        env_file: Path to .env file. If None, uses .env in project root.
        
    Returns:
        Config object with loaded values.
        
    Raises:
        FileNotFoundError: If specified env_file doesn't exist.
    """
    # Determine .env file location
    if env_file is None:
        project_root = Path(__file__).resolve().parent.parent
        env_file = project_root / ".env"
    
    # Load .env file via ConfigManager
    mgr = ConfigManager()
    mgr.load_env([env_file])
    
    # Build config using ConfigManager typed getters
    config = Config(
        outlook_account=mgr.get("OUTLOOK_ACCOUNT", ""),
        dry_run=mgr.get_bool("DRY_RUN", True),
        batch_size=mgr.get_int("BATCH_SIZE", 50),
        retry_count=mgr.get_int("RETRY_COUNT", 3),
        retry_delay_seconds=mgr.get_int("RETRY_DELAY_SECONDS", 2),
        log_dir=Path(mgr.get("LOG_DIR", "./logs")),
        output_dir=Path(mgr.get("OUTPUT_DIR", "./output")),
        state_dir=Path(mgr.get("STATE_DIR", "./state")),
        log_level=mgr.get("LOG_LEVEL", "INFO"),
    )
    
    return config


def get_config_template() -> str:
    """
    Return a template .env file content with all available settings.
    
    Useful for generating .env.example files.
    """
    return '''# Office Automation Foundation Configuration
# Copy this file to .env and adjust values as needed.

# =============================================================================
# OUTLOOK SETTINGS (Required)
# =============================================================================

# SMTP address of the Outlook account to use
# This must match one of the accounts configured in your Outlook profile
OUTLOOK_ACCOUNT=your.email@company.com

# =============================================================================
# SAFETY SETTINGS
# =============================================================================

# Dry-run mode: when true, no mutations are made to Outlook
# Logs and reports are still generated
# ALWAYS test with DRY_RUN=true first!
DRY_RUN=true

# =============================================================================
# BATCH PROCESSING
# =============================================================================

# Number of emails to process per batch (1-1000)
BATCH_SIZE=50

# =============================================================================
# RETRY CONFIGURATION
# =============================================================================

# Number of retry attempts for failed Outlook operations (0-10)
RETRY_COUNT=3

# Base delay between retries in seconds (uses exponential backoff)
RETRY_DELAY_SECONDS=2

# =============================================================================
# PATHS
# =============================================================================

# Directory for log files (created automatically)
LOG_DIR=./logs

# Directory for output files (created automatically)
OUTPUT_DIR=./output

# Directory for state tracking files (created automatically)
STATE_DIR=./state

# =============================================================================
# LOGGING
# =============================================================================

# Logging level: DEBUG, INFO, WARNING, ERROR
LOG_LEVEL=INFO
'''
