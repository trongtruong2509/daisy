## State Tracking Requirements for Daisy Automation Platform

### Rationale

Long-running batch operations (sending hundreds of emails, generating many files) must be resumable after a crash or interruption. Without persistent state tracking, restarting after a partial run risks duplicate sends, duplicate file writes, or reprocessed records. State files also serve as an audit trail showing exactly which items were processed in a given run.

### Requirements

#### REQ-STATE-01: All duplicate-prevention logic uses `StateTracker`

Any operation that must not be repeated (email send, file write, record processing) must be gated by a `StateTracker` instance from `core.state`. Inline duplicate-checking logic (e.g. checking an in-memory list) is insufficient on its own; it must always be backed by `StateTracker` for crash-safety.

```python
from core.state import StateTracker

tracker = StateTracker(state_dir=config.state_dir, state_name="email_send")

if not tracker.is_processed(employee_id):
    send_email(employee)
    tracker.mark_processed(employee_id, metadata={"email": employee.email})

tracker.save()
```

#### REQ-STATE-02: State files are JSON in `tools/<tool>/state/`

State files must be stored in the tool's configured `state_dir` (default: `tools/<tool>/state/`) and named `{state_name}_state.json`. The JSON format must include:

```json
{
  "state_name": "email_send",
  "created_at": "2026-01-15T09:30:00",
  "last_modified": "2026-01-15T09:45:12",
  "total_processed": 47,
  "processed_ids": ["EMP001", "EMP002", "..."],
  "metadata": {
    "EMP001": { "email": "emp1@company.com" },
    "EMP002": { "email": "emp2@company.com" }
  }
}
```

Human readability is required — the file must be pretty-printed with `indent=2`.

#### REQ-STATE-03: Atomic writes via temp file

`StateTracker.save()` must write to a temporary `.json.tmp` file first and then rename it over the target state file. This prevents partial writes leaving the state file corrupt on a sudden crash or power loss.

#### REQ-STATE-04: Corrupt state file recovery

If the state file fails JSON parsing on load, `StateTracker` must:

1. Rename the corrupt file to `{name}.json.corrupt`.
2. Log a warning with the backup path.
3. Start a fresh, empty state.

The tool must not crash because of a corrupt state file.

#### REQ-STATE-05: Auto-save on interval

`StateTracker` must support an `auto_save_interval` (default: 10). After every `auto_save_interval` calls to `mark_processed()`, the tracker must automatically call `save()`. Auto-save must be enabled by default (`auto_save=True`) and must be disableable via the constructor for test isolation.

#### REQ-STATE-06: `ContentHashTracker` for content-based deduplication

`core.state` must expose a `ContentHashTracker` subclass that generates a stable identifier from the content of an operation (e.g. email subject + recipient + date) using MD5 or SHA-256, rather than relying on an externally supplied ID. This is used by `OutlookSender` to prevent re-sending emails whose Message-ID is unavailable.

```python
from core.state import ContentHashTracker

tracker = ContentHashTracker(state_dir, "email_send")
content_id = tracker.compute_hash(subject=email.subject, to=email.to[0], date=date)
if not tracker.is_processed(content_id):
    sender.send(email)
    tracker.mark_processed(content_id)
```

#### REQ-STATE-07: `clear()` method for fresh starts

`StateTracker` must expose a `clear()` method that deletes the state file and resets in-memory state. This is used when a user explicitly chooses to restart a batch from scratch (e.g. via interactive prompt at tool startup).

#### REQ-STATE-08: State directory is created automatically

`StateTracker.__init__()` must call `state_dir.mkdir(parents=True, exist_ok=True)`. The tool must not need to pre-create the directory.

### Non-functional constraints

- State files must not include sensitive data (passwords, email body contents). Only IDs, timestamps, and non-sensitive metadata.
- `StateTracker` must not import any `office/` module.
- Unit tests must be able to use a `tmp_path` fixture-based directory and disable auto-save for deterministic behaviour.
