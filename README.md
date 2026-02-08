# Daisy Automation Platform

A **production-grade Python foundation** for automating office work on Windows, with a primary focus on **Outlook email automation** and integrated automation tools.

This project provides reusable modules for building reliable, safe email automation tools and a master launcher for managing multiple automation tools from a single shared virtual environment.

## Features

### Core Foundation
- **Outlook Email Operations**: Read, filter, save, and send emails via Outlook Desktop
- **Multi-Account Support**: Work with multiple accounts in one Outlook profile
- **Safety First**: Dry-run mode, duplicate prevention, comprehensive logging
- **Extensible Parsing**: Foundation for extracting data from email content
- **Resilient Design**: Retry logic, state tracking, crash recovery

### Master Launcher System
- **Single Shared Virtual Environment**: One setup for all tools
- **Interactive Menu**: Easy access to all tools and scripts
- **Direct Invocation**: Run specific tools from command line
- **Auto-Discovery**: Automatically detects tools in `tools/` directory
- **Simple for Non-Technical Users**: Double-click `.bat` files to run

### Available Tools
- **payslip-phuclong**: Automated payslip generation and email distribution

## Target Environment

- **OS**: Windows 10/11
- **Python**: 3.9+
- **Email Client**: Outlook Desktop (not web/O365 API)
- **Permissions**: No admin rights required
- **Distribution**: Git clone + Python scripts (no executables)

## Quick Start

### 1. Clone the Repository

```bash
git clone <repository-url>
cd daisy
```

### 2. Run Setup

```cmd
setup.bat
```

This will:
- Create a Python virtual environment
- Install dependencies
- Create `.Tools

**Interactive menu** (recommended for new users):
```cmd
run.bat
```

**Direct tool invocation**:
```cmd
run.bat payslip-phuclong
```

**Run example scripts**:
```cmd
run.bat example_read_emails
```

**Get help**:
```cmd
run.bat --help
```ini
OUTLOOK_ACCOUNT=your.email@company.com
DRY_RUN=true
```

### 4. Run Example Script

Make sure Outlook Desktop is running, then:

```cmd
run.bat example_read_emails
```

## Project Structure

```
daisy/
├── venv/                    # Virtual environment (gitignored)
├── core/                    # Core utilities
│   ├── __init__.py
│   ├── config.py            # Configuration management
│   ├── logger.py            # Logging (file + console)
│   ├── retry.py             # Retry with exponential backoff
│   └── state.py             # State tracking for duplicates
├── office/
│   └── outlook/             # Outlook abstraction layer
│       ├── __init__.py
│       ├── client.py        # Email reading client
│       ├── sender.py        # Email sending with safety
│       ├── models.py        # Data models (Email, Filter, etc.)
│       └── exceptions.py    # Custom exceptions
├── parsing/                 # Email content parsing
│   ├── __init__.py
│   ├── base.py              # Parser interface
│   ├── text.py              # Plain text parsing
│   └── html.py              # HTML parsing
├── tools/                   # Automation tools
│   └── payslip-phuclong/    # Payslip generation tool
│       ├── main.py          # Tool entry point
│       ├── run.bat          # Tool wrapper
│       ├── .env.example     # Tool configuration template
│       └── README.md        # Tool documentation
├── scripts/                 # Utility scripts
│   ├── example_read_emails.py
│   └── example_send_email.py
├── docs/                    # Documentation
│   └── venv-strategy.md     # Virtual environment strategy
├── .env.example             # Root configuration template
├── requirements.txt         # Shared Python dependencies
├── setup.bat                # Master setup (creates venv)
├── run.bat                  # Master launcher
└── README.md                # This file
```

## Configuration

All configuration is via `.env` file. Copy `.env.example` to `.env` and adjust:

| Variable | Description | Default |
|----------|-------------|---------|
| `OUTLOOK_ACCOUNT` | Email address to use | (required) |
| `DRY_RUN` | Safety mode - no mutations | `true` |
| `BATCH_SIZE` | Emails per batch | `50` |
| `RETRY_COUNT` | Retries for failed operations | `3` |
| `RETRY_DELAY_SECONDS` | Base retry delay | `2` |
| `LOG_DIR` | Log file directory | `./logs` |
| `OUTPUT_DIR` | Output file directory | `./output` |
| `STATE_DIR` | State tracking directory | `./state` |
| `LOG_LEVEL` | DEBUG, INFO, WARNING, ERROR | `INFO` |

## Master Launcher and Setup

### Setup Script (`setup.bat`)

The setup script creates the virtual environment and installs all dependencies. It supports several options:

**Standard setup** (first time):
```cmd
setup.bat
```

**Recreate virtual environment** (if corrupted):
```cmd
setup.bat --recreate
```

**Upgrade installed packages**:
```cmd
setup.bat --upgrade
```

The setup script will:
1. Check Python 3.14+ is installed
2. Create or update virtual environment
3. Install all dependencies from `requirements.txt`
4. Validate core modules
5. Create `.env` from template if needed

### Master Launcher (`run.bat`)

The master launcher provides three ways to run tools:

**1. Interactive Menu** (easiest for non-technical users):
```cmd
run.bat
```
This shows a numbered menu of all available tools and scripts.

**2. Direct Tool Invocation** (fastest for regular users):
```cmd
run.bat payslip-phuclong
run.bat example_read_emails
```

**3. Tool Wrapper** (from tool directory):
```cmd
cd tools\payslip-phuclong
run.bat
```

The launcher automatically:
- Activates the virtual environment
- Discovers available tools in `tools/` directory
- Discovers available scripts in `scripts/` directory
- Handles errors gracefully
- Deactivates virtual environment on exit

## Safety Guarantees

### Dry-Run Mode

When `DRY_RUN=true` (default):

- No emails are sent
- No emails are modified
- Logs show what *would* happen
- State tracking still works (for testing)

**Always test with dry-run before going live!**

### Duplicate Prevention

The state tracking system prevents:

- Sending the same email twice
- Processing the same email repeatedly

State is persisted to JSON files and survives restarts.

### Comprehensive Logging

Every run creates a timestamped log file in `logs/`:

- Console: Concise progress updates
- File: Detailed audit trail
- All operations are logged for traceability

### Retry Logic

Outlook COM operations are wrapped with:

- Configurable retry attempts
- Exponential backoff
- Transient error detection
- Clear logging of failures

## Building Your Own Tools

The foundation is designed for building custom automation tools.

### Basic Pattern

```python
from core.config import load_config
from core.logger import setup_logging, get_logger
from core.state import StateTracker
from office.outlook import OutlookClient, EmailFilter

# Load configuration
config = load_config()
config.ensure_directories()

# Set up logging
setup_logging(log_dir=config.log_dir, level=config.log_level)
logger = get_logger(__name__)

# Initialize state tracking
tracker = StateTracker(config.state_dir, "my_tool")

# Use Outlook client
with OutlookClient(account=config.outlook_account) as client:
    emails = client.get_inbox_emails(
        filter=EmailFilter(unread_only=True, limit=100)
    )
    
    for email in emails:
        if tracker.is_processed(email.unique_id):
            continue
            
        # Your business logic here
        process_email(email)
        
        tracker.mark_processed(email.unique_id)
    
    tracker.save()
```

### Sending Emails Safely

```python
from office.outlook import OutlookSender
from office.outlook.models import NewEmail
from core.state import ContentHashTracker

tracker = ContentHashTracker(config.state_dir, "email_send")

with OutlookSender(
    account=config.outlook_account,
    dry_run=config.dry_run,
    state_tracker=tracker
) as sender:
    email = NewEmail(
        to=["recipient@example.com"],
        subject="Automated Email",
        body="Hello from the automation system.",
    )
    
    sender.send(email)  # Won't send duplicates
```

### Parsing Email Content

```python
from parsing import TextParser, HtmlParser

# Plain text parsing
text_parser = TextParser()
result = text_parser.parse(email.body_text)
print(result.data["key_values"])  # Extracted key-value pairs

# HTML parsing
html_parser = HtmlParser()
result = html_parser.parse(email.body_html)
print(result.data["text"])   # Plain text version
print(result.data["tables"]) # Extracted tables
```

### Custom Parser

```python
from parsing.base import BaseParser, ParseResult
import re

class InvoiceParser(BaseParser):
    def parse(self, content, **kwargs):
        data = {}
        
        # Extract invoice number
        match = re.search(r"Invoice #?(\d+)", content)
        if match:
            data["invoice_number"] = match.group(1)
        
        # Extract amount
        match = re.search(r"\$[\d,]+\.?\d*", content)
        if match:
            data["amount"] = match.group(0)
        
        return ParseResult(
            success=bool(data),
            data=data,
            raw_content=content
        )
```

## What's NOT Included

This foundation intentionally excludes:

- Business-specific logic (HR workflows, payslip parsing)
- GUI/interface
- Meetings/calendar automation
- Microsoft Graph API integration
- Shared mailbox access
- Executable packaging

These should be built as separate tools on top of this foundation.

## Troubleshooting

### "Outlook not found" or connection errors

- Ensure Outlook Desktop is running (not just Outlook web)
- Check that your account is configured in Outlook
- Try restarting Outlook

### "Account not found"

- Verify `OUTLOOK_ACCOUNT` matches exactly with your Outlook profile
- Run the example script to see available accounts

### Slow performance

- Reduce `BATCH_SIZE` in `.env`
- Use more specific filters
- Large mailboxes are inherently slow via COM

### Permission errors

- Don't run as administrator
- Ensure the script has read/write access to project directories

## Dependencies

- `pywin32` - Windows COM automation
- `python-dotenv` - Environment configuration
- `openpyxl` - Excel file handling
- `pandas` - Data manipulation
- `beautifulsoup4` - HTML parsing
- `PyYAML` - YAML support

## Contributing

When adding new features:

1. Follow the existing architecture patterns
2. Add proper logging
3. Support dry-run mode for mutations
4. Write docstrings and comments
5. Update this README if needed

## License

Internal use only. Not for distribution.
