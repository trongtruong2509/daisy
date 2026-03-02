## Get Attachment

Download all email attachments from your Outlook Inbox for a specific date,
with optional subject keyword filtering.

---

## Quick Start

### First-time Setup

1. Copy `.env.example` to `.env` in this directory (optional — the tool prompts for any missing value).

2. Run the tool from the project root:

   ```cmd
   run.bat get-attachment
   ```

   Or use the interactive menu:

   ```cmd
   run.bat
   ```

3. Follow the prompts:
   - Choose your Outlook account (or set `OUTLOOK_ACCOUNT` in `.env`)
   - Enter the start date in `DD/MM/YYYY` format (or set `START_DATE`)
   - Enter subject keywords, comma-separated, or leave blank for all emails
   - Enter the directory where attachments should be saved (or set `ATTACHMENT_SAVE_PATH`)

---

## Configuration

All settings can be placed in a `.env` file inside this directory.
The tool will prompt interactively for anything that is not already set.
Prompted values are used for the current run only and are **not** written back to `.env`.

| Setting                | Description                                                      | Example                    |
| ---------------------- | ---------------------------------------------------------------- | -------------------------- |
| `OUTLOOK_ACCOUNT`      | SMTP address of your Outlook account                             | `user@company.com`         |
| `START_DATE`           | Start date for email search (DD/MM/YYYY)                         | `02/03/2026`               |
| `SUBJECT_KEYWORDS`     | Comma-separated subject keywords (OR logic); leave blank for all | `invoice,report`           |
| `ATTACHMENT_SAVE_PATH` | Directory where files are saved                                  | `D:\Downloads\attachments` |
| `LOG_DIR`              | Log output directory                                             | `./logs`                   |
| `LOG_LEVEL`            | Logging verbosity                                                | `INFO`                     |

### Example `.env`

```dotenv
OUTLOOK_ACCOUNT=user@company.com
START_DATE=02/03/2026
SUBJECT_KEYWORDS=invoice,monthly report
ATTACHMENT_SAVE_PATH=D:\Downloads\attachments
LOG_DIR=./logs
LOG_LEVEL=INFO
```

---

## Behaviour

### Email filtering

- Emails are searched in the **Inbox** of the configured account.
- Only emails **received on the target date** (00:00:00 – 23:59:59) are considered.
- If `SUBJECT_KEYWORDS` is set, an email is included when its subject contains
  **any** of the keywords (case-insensitive, OR logic).
- Leaving keywords blank processes attachments from **all** emails on that date.

### Filename deduplication

All attachments land in a single flat directory. When two attachments share
the same filename the following strategy is applied in order:

1. `{original_name}` — saved as-is if the name is free.
2. `{stem}_{sender_address}{suffix}` — sender address appended (characters
   unsafe for Windows filenames are replaced with underscores).
3. `{stem}_{sender_address}_{YYYYMMDD_HHMMSS}{suffix}` — timestamp appended.
4. `{stem}_{sender_address}_{YYYYMMDD_HHMMSS}_{n}{suffix}` — integer counter
   as a final fallback.

### Error handling

- COM errors when saving individual attachments are retried up to three times
  with exponential back-off.
- Failures for one attachment do not stop processing of the remaining attachments
  or emails.
- All errors are logged to the run log file and printed to the console.
- The tool exits with code `2` when at least one attachment could not be saved.

---

## Logs

Each run creates a timestamped log file in the configured `LOG_DIR`
(default: `./logs`). Log files capture every email processed and every
file saved, making them useful for audit purposes.

---

## Running Tests

Unit tests do not require Outlook or Windows COM:

```cmd
cd tools\get-attachment
..\..\..\venv\Scripts\pytest
```

To include integration tests (requires a live Outlook session):

```cmd
..\..\..\venv\Scripts\pytest -m integration
```

---

## Requirements

- Windows with Microsoft Outlook installed and running.
- An active Outlook profile with the target email account configured.
- Python virtual environment with `pywin32` installed (run `setup.bat` from the
  project root to create it).
