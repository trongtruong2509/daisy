## Daisy Automation Platform — Project Overview

**Daisy** is a Windows-only office automation framework for business applications that require COM (Component Object Model) integration with Microsoft Excel and Outlook. It provides a reusable foundation for building reliable, auditable automation tools that generate documents, read spreadsheets, and send emails at scale.

---

## Architecture

Daisy is structured as a **layered stack** with shared infrastructure at the bottom and isolated, independently-runnable tools at the top:

```
┌─ tools/ ────────────────────────────────────────────────┐
│  • payslip-phuclong-ecom/  — Generate & send payslips  │
│  • get-attachment/         — Download email attachments│
│  • [future tools...]                                     │
└──────────────────────────────────────────────────────────┘
        ↓ imports from ↓
┌─ office/ ───────────────────────────────────────────────┐
│  • excel/      — ExcelComReader, PdfConverter          │
│  • outlook/    — OutlookSender, OutlookReader          │
│  • utils/      — COM lifecycle, helpers                │
└──────────────────────────────────────────────────────────┘
        ↓ imports from ↓
┌─ core/ ─────────────────────────────────────────────────┐
│  • config_manager   — .env loading, prompting          │
│  • console          — cprint() for unified output      │
│  • logger           — dual-output file + console      │
│  • state            — persistent state tracking       │
│  • retry            — exponential backoff decorator   │
└──────────────────────────────────────────────────────────┘
```

### Core Layer (`core/`)

Shared infrastructure used by **all** tools:

| Module              | Purpose                                                                                              |
| ------------------- | ---------------------------------------------------------------------------------------------------- |
| `config_manager.py` | Two-layer `.env` loading (global + tool-local), typed getters, interactive prompting                 |
| `console.py`        | Single `cprint()` function for all user-facing output with color support and automatic log mirroring |
| `logger.py`         | File + console dual-output logging with a custom `CONSOLE` level for auditability                    |
| `state.py`          | Persistent state tracking in JSON files for crash-safe resume and duplicate prevention               |
| `retry.py`          | Decorator-based retry logic with exponential backoff for transient COM errors                        |

### Office Layer (`office/`)

Windows COM automation abstractions:

| Module                      | Purpose                                                                  |
| --------------------------- | ------------------------------------------------------------------------ |
| `office/excel/reader.py`    | Read data from Excel via COM with automatic formula recalculation        |
| `office/excel/converter.py` | Convert Excel files to password-protected PDFs using Excel COM export    |
| `office/outlook/sender.py`  | Send emails via Outlook with dry-run, retry, and duplicate prevention    |
| `office/outlook/reader.py`  | Read Outlook mailbox folders and download attachments with deduplication |
| `office/outlook/models.py`  | Typed data classes for Outlook accounts, emails, and attachments         |
| `office/utils/com.py`       | COM initialisation/cleanup context manager (`com_initialized()`)         |
| `office/utils/helpers.py`   | Lifecycle helpers (detect or create Excel/Outlook instances)             |

### Tools Layer (`tools/`)

Self-contained, independently-runnable applications:

| Tool                     | Purpose                                                                       | Config               |
| ------------------------ | ----------------------------------------------------------------------------- | -------------------- |
| `payslip-phuclong-ecom/` | Generate individual payslip PDFs from Excel template + data, send via Outlook | `.env` + `config.py` |
| `get-attachment/`        | Download email attachments from Outlook Inbox with date/keyword filtering     | `.env` + `config.py` |

Each tool is a complete Python package: entry point (`main.py`), config loading (`config.py`), utilities, tests, and local `.env` file.

---

## Key Features

### Unified Configuration

- **Two-layer `.env` system** — global defaults can be overridden per-tool.
- **Interactive prompting** — missing required values are collected from the user at startup with validation.
- **Type-safe access** — `ConfigManager` provides typed getters (`get()`, `get_bool()`, `get_path()`, etc.) and built-in validators.

### Reliable Output & Logging

- **Single `cprint()` function** — all user-facing text goes through one path for consistency, colour, and log mirroring.
- **Dual-output logging** — every action is logged to a timestamped file AND printed to the console with appropriate verbosity.
- **Custom `CONSOLE` level** — user-facing progress messages have their own log level (25, between INFO and WARNING) to remain visible and auditable without cluttering debug logs.

### Crash-Safe Batch Operations

- **Persistent state tracking** — tools resume automatically from the last checkpoint. On restart, already-processed items are skipped.
- **Atomic writes** — state is written to a temp file first, then renamed, preventing corruption on unexpected crashes.
- **Duplicate prevention** — combined with state tracking, prevents re-sending emails or re-generating files.

### COM Automation Reliability

- **Centralised COM lifecycle** — all Windows COM objects are acquired and released through a single context manager (`com_initialized()`).
- **Automatic retry with backoff** — transient COM errors (busy, RPC timeout) are retried up to 3 times with exponential delay.
- **App instance detection** — tools detect if Excel/Outlook is already running and preserve the user's existing windows/data.
- **Safe cleanup** — Excel/Outlook processes are only terminated if the tool started them; existing user sessions are never interrupted.

### Safety by Default

- **Dry-run mode is the default** — all mutation operations (`DRY_RUN=true` by default) log what would happen without actually sending emails or modifying files.
- **User confirmation** — live-mode operations prompt for explicit confirmation before proceeding.
- **Fail-fast validation** — data validation happens before any file generation or email sending.

---

## Use Cases

### Monthly Payslip Distribution

Use the **Payslip Generator** tool to:

1. Read employee ID, name, email, and password from a Data sheet in Excel.
2. Use a template sheet with VLOOKUP/XLOOKUP formulas to generate individual payslip data.
3. Export each employee's payslip as a PDF with their ID as the password.
4. Send personalised emails with the PDF attached, tracking delivery via Outlook.
5. Resume mid-batch if interrupted (via state tracking).

### Automated Attachment Harvesting

Use the **Get Attachment** tool to:

1. Connect to a shared mailbox and search for emails on a specific date.
2. Filter by subject keywords (e.g. `"invoice"`, `"report"`).
3. Download all non-inline attachments to a local folder.
4. Handle filename collisions automatically by appending sender or timestamp.
5. Audit every email and file saved in the log file.

### Future Tools

The platform is designed to support arbitrary office automation tasks:

- Bulk email campaigns with template personalisation.
- Report generation and distribution.
- Data import/validation from email attachments.
- Scheduled document merging (mail merge).

---

## Getting Started

### Prerequisites

- **Operating System:** Windows only (Windows 10 or later recommended).
- **Python:** 3.9+
- **Dependencies:** `pywin32`, `pikepdf`, `python-dotenv`, `colorama` (see `requirements.txt`).

## Requirements & Constraints

### Windows Dependency

Daisy is **Windows-only**. It uses COM (Component Object Model) which is unique to Windows. The `pywin32` package and Windows COM runtime are non-negotiable dependencies.

### Single-Threaded COM

COM objects are **not thread-safe**. All mutations to Excel/Outlook must occur in a single thread. Tools that need parallelism must:

1. Run each batch item serially in one thread.
2. Create separate Excel/Outlook instances for truly independent parallel work (not recommended for this tool).

For the Payslip tool, employees are processed in batches of at most 50, with garbage collection between batches, to prevent COM memory exhaustion.

### Outlook & Exchange Requirements

- **Exchange account required** — Outlook must be configured with a company Exchange account (not a personal Outlook.com account). Tools detect the account list from Outlook's profile.
- **Email sending via COM** — this tool respects company-wide email policies, signatures, and send-as/send-on-behalf settings enforced by IT.
- **COM vs. IMAP/SMTP** — this tool does **not** use IMAP or SMTP directly. All interaction is through COM, which is more reliable for Exchange accounts but requires Outlook to be running.

### File Encoding

- `.env` files: UTF-8 (standard).
- Log files: UTF-8 with error replacement (invalid sequences replaced with `?`).
- Data files: UTF-8 by default; Excel files may be XLSX (OpenOffice format) or XLS (legacy).

---

## Documentation Structure

High-level design and requirements are documented in `docs/`:

- **`project-overview.md`** (this file) — architecture, use cases, getting started.
- **`project-requirements/`** — detailed functional and non-functional requirements for each core and office module:
  - [COM Handling](project-requirements/com-handle.md) — Windows COM lifecycle, `com_initialized()`, thread safety.
  - [Logging](project-requirements/logging.md) — unified console output, file logging, CONSOLE level.
  - [Console Output](project-requirements/console.md) — `cprint()` levels, colours, logging mirror.
  - [Configuration](project-requirements/config.md) — two-layer `.env`, prompting, typed getters.
  - [State Tracking](project-requirements/state.md) — `StateTracker`, atomic writes, resume logic.
  - [Retry Logic](project-requirements/retry.md) — transient error detection, exponential backoff.
  - [Excel Automation](project-requirements/excel.md) — `ExcelComReader`, formula recalculation, PDF conversion.
  - [PDF Conversion](project-requirements/pdf.md) — Excel COM export, `pikepdf` password protection.
  - [Outlook Automation](project-requirements/outlook.md) — email sending, reading, duplicate prevention.
- **`tools/`** — tool-specific requirements:
  - [Payslip Generator](tools/payslip-phuclong-ecom/tool-requirements.md) — end-to-end workflow, batching, resume.
  - [Get Attachment](tools/get-attachment/tool-requirements.md) — folder search, filtering, deduplication.

Each requirement file is formatted with clear sections: **Rationale**, numbered `REQ-XXX-NN` items with code examples, and **Non-functional constraints**.
