## Tool Requirements — Get Attachment (`get-attachment`)

### Overview

This tool connects to a configured Outlook account, searches a mailbox folder for emails received within a specified date range, optionally filters by subject keywords, and saves all non-inline attachments to a local directory. It is designed for recurring daily/weekly batch downloads (e.g. collecting invoice files or report attachments from automated senders).

---

### Functional Requirements

#### REQ-GA-01: Outlook account selection at startup

The tool must prompt the user to choose an Outlook account if `OUTLOOK_ACCOUNT` is not set in `.env`. The prompt must:

1. Call `OutlookClient.get_available_accounts()` to list configured accounts.
2. Display a numbered list of accounts; the user selects by number.
3. Fall back to manual email entry (with format validation) if no accounts are detected.
4. Loop until a valid selection is made.

#### REQ-GA-02: Date range filtering — `START_DATE` to `END_DATE`

The tool must search for emails whose received time falls within the range `[START_DATE 00:00:00, END_DATE 23:59:59]` (inclusive). Both dates must be in `DD/MM/YYYY` format and validated against the `_validate_date_ddmmyyyy` validator.

- `END_DATE` defaults to today if not set.
- If `START_DATE == END_DATE`, only emails received on that single day are included.

Date inputs must support both interactive prompting and `.env` configuration.

#### REQ-GA-03: Folder selection — configurable inbox path

The target folder must default to the Outlook Inbox. An alternative path may be specified via `OUTLOOK_FOLDER` (e.g. `"Inbox/Reports/2026"`) using the slash-separated format accepted by `OutlookReader.get_folder_by_path()`. If the specified path does not exist, the tool must raise a descriptive error and exit with code 1.

#### REQ-GA-04: Subject keyword filtering — OR logic, case-insensitive

When `SUBJECT_KEYWORDS` is set, only emails whose subjects contain **at least one** of the keywords (case-insensitive substring match) must be selected. Leaving `SUBJECT_KEYWORDS` empty processes all emails on the target date range.

Keywords are comma-separated in `.env` or at the interactive prompt (e.g. `invoice,monthly report`).

#### REQ-GA-05: Skip inline (embedded) attachments by default

Attachments whose `content_id` is non-empty (CID-linked inline images used in HTML signatures or decorative email bodies) must be skipped silently. Only `ATTACH_BY_VALUE` (type 1) attachments must be downloaded. This behaviour must not be configurable — inline attachments are never user-facing files.

#### REQ-GA-06: Filename deduplication strategy

All attachments are saved in a single flat directory. When a filename collision occurs, the following strategy must be applied in order:

1. `{original_name}` — saved as-is if the filename is free.
2. `{stem}_{safe_sender}{suffix}` — sanitised sender address appended.
3. `{stem}_{safe_sender}_{YYYYMMDD_HHMMSS}{suffix}` — received timestamp appended.
4. `{stem}_{safe_sender}_{YYYYMMDD_HHMMSS}_{n}{suffix}` — integer counter as final fallback.

Characters unsafe for Windows filenames (`\`, `/`, `:`, `*`, `?`, `"`, `<`, `>`, `|`, `@`, `.`, ` `) in the sender address must be replaced with `_`. Consecutive underscores must be collapsed to a single `_`.

#### REQ-GA-07: Retry COM save operations

Each `attachment.SaveAsFile()` COM call must be wrapped with `@retry_operation(RetryConfig(max_attempts=3, base_delay=1.0, max_delay=10.0))`. A failure to save one attachment must not stop the processing of remaining attachments or emails.

#### REQ-GA-08: Individual attachment failures are non-fatal

When saving an attachment fails after all retries, the tool must:

1. Log the error at `ERROR` level.
2. Print the error via `cprint(..., level="ERROR")`.
3. Append a human-readable error message to `DownloadResult.errors`.
4. Increment `DownloadResult.attachments_failed`.
5. Continue to the next attachment.

The tool must exit with code `2` if at least one attachment could not be saved, even if others succeeded.

#### REQ-GA-09: Pre-run summary and confirmation

Before starting the download, the tool must display a compact summary box via `cprint_summary_box_lite()` showing:

- Outlook account, target folder, start date, end date, subject keywords, save path.

The user must confirm with `yes` / `no` before download begins.

#### REQ-GA-10: Existing-files prompt before download

If the `ATTACHMENT_SAVE_PATH` directory already contains files, the tool must ask the user whether to delete all existing files before downloading. The choices must be:

- **yes** — delete all files, then download fresh.
- **no** — keep existing files; new downloads will overwrite on name clash.

#### REQ-GA-11: `DownloadResult` aggregates outcome statistics

The `AttachmentDownloader.run()` method must return a `DownloadResult` dataclass containing:

- `emails_found` — total emails on the target date range (before keyword filter).
- `emails_matched` — emails that passed the keyword filter.
- `emails_with_attachments` — matched emails that had at least one attachment.
- `attachments_saved` — attachments successfully written to disk.
- `attachments_failed` — attachments that couldn't be saved.
- `saved_files` — list of absolute `Path` objects for every saved file.
- `errors` — list of human-readable error strings.

#### REQ-GA-12: End-of-run summary

After all processing the tool must display a summary box (via `cprint_summary_box()`) showing the `DownloadResult` statistics and the save path.

---

### Configuration Reference

All settings can be placed in `tools/get-attachment/.env`. Absent required values are prompted interactively.

| Key                    | Default           | Description                                             |
| ---------------------- | ----------------- | ------------------------------------------------------- |
| `OUTLOOK_ACCOUNT`      | (prompted)        | SMTP address of the Outlook account to read from        |
| `OUTLOOK_FOLDER`       | `""` (Inbox)      | Slash-separated subfolder path relative to account root |
| `START_DATE`           | (prompted)        | Start date for email search in `DD/MM/YYYY` format      |
| `END_DATE`             | today             | End date for email search in `DD/MM/YYYY` format        |
| `SUBJECT_KEYWORDS`     | `""` (all emails) | Comma-separated keywords; OR logic; blank = no filter   |
| `ATTACHMENT_SAVE_PATH` | `./attachments`   | Directory where attachments are saved                   |
| `LOG_DIR`              | `./logs`          | Log output directory                                    |
| `LOG_LEVEL`            | `INFO`            | Logging verbosity                                       |

---

### Non-functional Constraints

- The tool is Windows-only; `pywin32` is a required dependency.
- There is no dry-run mode — the tool only reads from Outlook and writes files locally, so there is no risk of mutation.
- The tool must not modify, move, or delete any Outlook emails or folders.
- `ATTACHMENT_SAVE_PATH` is created automatically if it does not exist.
- Log files capture every email processed and every file saved, suitable for audit purposes.
- Integration tests must use `@pytest.mark.integration` and are excluded from the default run.
