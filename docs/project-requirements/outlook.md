## Outlook Automation Requirements for Daisy Automation Platform

### Rationale

Outlook COM automation is the only reliable way to send emails from a company Exchange account while respecting signature, send-on-behalf, and other account-level policies enforced by the IT department. COM is inherently fragile — it can fail transiently when Outlook is busy, when RPC calls time out, or when the user interacts with Outlook simultaneously. Centralising all Outlook interaction in `office/outlook/` isolates this complexity and provides a consistent safety layer (dry-run, duplicate prevention, retry) for every tool.

### Requirements

#### REQ-OL-01: `OutlookClient` is the base class for all Outlook interactions

All Outlook COM interaction must go through a subclass of `office.outlook.client.OutlookClient`. Tools must never call `win32com.client.Dispatch("Outlook.Application")` directly. Two concrete subclasses are defined:

- `OutlookSender` — for composing and sending emails.
- `OutlookReader` — for reading folders, searching emails, and downloading attachments.

#### REQ-OL-02: Context manager protocol is mandatory

All `OutlookClient` subclasses must be used exclusively via `with` statements. Calling `connect()` / `disconnect()` directly in production code is forbidden (see REQ-COM-07).

```python
with OutlookSender(account="user@company.com", dry_run=True) as sender:
    sender.send(email)

with OutlookReader(account="user@company.com") as reader:
    emails = reader.get_inbox_emails()
```

#### REQ-OL-03: `com_initialized()` wraps the connection lifecycle

`OutlookClient.connect()` must enter `com_initialized()` at the start, and `OutlookClient.disconnect()` must exit it (see REQ-COM-02). The COM context must be stored as `self._com_ctx` and exited in `__exit__`/`disconnect()` after all COM objects have been released.

#### REQ-OL-04: `get_or_create_outlook()` helper manages application lifecycle

`office.outlook.helpers.get_or_create_outlook()` must implement the detect-or-create pattern from REQ-COM-08. `OutlookClient.connect()` must use this helper. The returned `was_already_running` flag must be stored and used in `disconnect()` to decide whether to call `outlook.Quit()`.

#### REQ-OL-05: `OutlookSender` enforces dry-run by default

`OutlookSender.__init__()` must accept `dry_run: bool = True`. When `dry_run=True`:

- No emails are sent.
- All `send()` calls log at `WARNING` level: `[DRY RUN] Would send: <subject> → <recipients>`.
- Send statistics (`sent_count`, `skipped_count`) are still maintained.
- A `WARNING` banner must be printed at construction time to alert the user.

#### REQ-OL-06: Duplicate send prevention via `StateTracker`

`OutlookSender` must accept an optional `state_tracker: StateTracker` argument. When provided:

1. Before sending, compute a content hash from (recipient, subject, date).
2. Call `tracker.is_processed(hash)`. If `True`, skip and increment `skipped_count`.
3. After a successful send, call `tracker.mark_processed(hash)`.

This prevents re-sending if the tool is restarted mid-batch.

#### REQ-OL-07: Retry on transient COM errors

`OutlookSender._send_outlook_item()` (the internal method that calls `mail_item.Send()`) must be decorated with `@retry_operation` using a `RetryConfig` supplied via the constructor. The default retry config must be `RetryConfig(max_attempts=3, base_delay=2.0)`.

#### REQ-OL-08: `NewEmail` dataclass for outgoing mail

All outgoing emails must be constructed using `office.outlook.models.NewEmail`. Direct manipulation of COM `MailItem` properties outside `OutlookSender` is forbidden.

```python
from office.outlook.models import NewEmail

email = NewEmail(
    to=["recipient@example.com"],
    cc=[],
    subject="Payslip January 2026",
    body="<html>...</html>",
    attachments=[Path("payslip_EMP001.pdf")],
    is_html=True,
)
sender.send(email)
```

`NewEmail` must support: `to` (required), `cc`, `bcc`, `subject`, `body`, `attachments` (list of `Path`), `is_html`, and `importance`.

#### REQ-OL-09: `OutlookReader` resolves the account delivery store root

`OutlookReader.connect()` must call the base `connect()` and then resolve `self._account_folder` from `account_obj.DeliveryStore.GetRootFolder()`. All folder navigation builds on this root.

#### REQ-OL-10: Folder navigation via `get_folder_by_path()`

`OutlookReader` must expose `get_folder_by_path(path)` accepting slash- or backslash-separated relative folder paths (e.g. `"Inbox/Reports/2026"`). Traversal must start from the account root and raise `OutlookFolderNotFoundError` if any segment is not found.

#### REQ-OL-11: Email filter via `EmailFilter` dataclass

Email search/retrieval methods must accept an `office.outlook.models.EmailFilter` dataclass rather than positional date/keyword arguments. The filter must support:

- `start_date` / `end_date` — received-time bounds.
- `subject_keywords` — list of strings; OR logic (any keyword matches).
- `sender_email` — exact sender address filter.
- `has_attachments` — boolean flag.

#### REQ-OL-12: Attachment download with deduplication strategy

`OutlookReader.save_attachment()` (or equivalent) must apply the following filename deduplication strategy when two attachments share the same name in the output directory:

1. `{original_name}` — used as-is if available.
2. `{stem}_{sender_address}{suffix}` — sender address appended (unsafe chars replaced with `_`).
3. `{stem}_{sender_address}_{YYYYMMDD_HHMMSS}{suffix}` — timestamp appended.
4. `{stem}_{sender_address}_{YYYYMMDD_HHMMSS}_{n}{suffix}` — integer counter as final fallback.

#### REQ-OL-13: Inline attachments must be skipped by default

Attachments with a non-empty `Content-ID` header (inline images referenced in HTML body via `cid:`) must be skipped during download unless the caller explicitly opts in. These are typically signature logos or decoration graphics, not user-facing files.

#### REQ-OL-14: `get_available_accounts()` class method

`OutlookClient` must expose a `@classmethod get_available_accounts() -> list[str]` that returns the SMTP addresses of all accounts configured in the running Outlook profile. This is used by tool config modules to present an account selection menu to the user.

#### REQ-OL-15: `OutlookFolderNotFoundError` and `OutlookSendError` exceptions

`office.outlook.exceptions` must define at minimum:

- `OutlookFolderNotFoundError` — raised when a folder path cannot be resolved.
- `OutlookSendError` — raised when a send attempt fails after all retries.
- `OutlookConnectionError` — raised when `connect()` cannot acquire an Outlook COM object.

### Non-functional constraints

- `office/outlook/` modules are Windows-only. Non-Windows environments must receive a clear `ImportError`.
- `OutlookSender` must track and expose `sent_count`, `skipped_count`, and `error_count` for use in end-of-run summaries.
- Unit tests must mock `OutlookSender` and `OutlookReader` at their class boundaries, not at the `pythoncom` level.
- Integration tests must be tagged `@pytest.mark.integration` and excluded from the default `pytest` run.
