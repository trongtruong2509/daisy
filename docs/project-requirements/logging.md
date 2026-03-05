## Logging Requirements for Daisy Automation Platform

### Rationale

Every automation run must produce a traceable, auditable record of what happened. Daisy uses Python's standard `logging` module extended with a custom `CONSOLE` level so that user-facing output and internal diagnostic output share a single, unified pipeline. Log files are the primary post-mortem tool; console output is the real-time progress indicator. Both must appear in the same log file so that support and debugging does not require re-running the tool.

### Requirements

#### REQ-LOG-01: Centralised logging initialisation via `setup_logging()`

All logging initialisation must go through `core.logger.setup_logging()`. No tool or office module may call `logging.basicConfig()`, add handlers to the root logger directly, or configure any handler before `setup_logging()` has been called.

```python
from core.logger import setup_logging

log_file = setup_logging(
    log_dir=config.log_dir,
    level=config.log_level,
    run_name="payslip",
)
```

`setup_logging()` is idempotent — calling it more than once on the same process is safe (subsequent calls are no-ops that return the original log file path).

#### REQ-LOG-02: Module loggers via `get_logger()`

Every module that needs a logger must obtain one via `core.logger.get_logger(__name__)`. No module may call `logging.getLogger()` directly.

```python
# Correct
from core.logger import get_logger
logger = get_logger(__name__)

# Forbidden
import logging
logger = logging.getLogger(__name__)
```

#### REQ-LOG-03: Custom `CONSOLE` level at 25

A custom level named `CONSOLE` (numeric value 25, between `INFO=20` and `WARNING=30`) is registered at import time in `core.logger`. It is used exclusively by `cprint()` (see REQ-CON-01) to record user-facing messages in the log file. No other code may call `logger.log(CONSOLE, ...)` directly.

#### REQ-LOG-04: Dual-output handlers

`setup_logging()` must configure exactly two handlers on the root logger:

| Handler                  | Formatter                                             | Minimum level      | Notes                                                                |
| ------------------------ | ----------------------------------------------------- | ------------------ | -------------------------------------------------------------------- |
| `FileHandler`            | `FileFormatter` (timestamp, level, module, message)   | Configured `level` | UTF-8 encoded                                                        |
| `StreamHandler` → stdout | `ConsoleFormatter` (concise prefix + optional colour) | `CONSOLE`          | Suppresses `CONSOLE`-level records — `cprint()` already prints those |

The `CONSOLE`-level filter on the `StreamHandler` prevents doubled console output: `cprint()` prints directly to stdout, and separately logs at `CONSOLE` level for the file record.

#### REQ-LOG-05: Timestamped log filenames

Each run produces a new log file in the configured `log_dir` using the naming pattern:

```
{run_name}_{YYYYMMDD_HHMMSS}.log    # when run_name is provided
run_{YYYYMMDD_HHMMSS}.log           # when run_name is omitted
```

The directory is created automatically if it does not exist. One log file per process run is the expected contract.

#### REQ-LOG-06: File format — `FileFormatter`

Log records written to the file must use this format:

```
YYYY-MM-DD HH:MM:SS | LEVEL    | module.name | message text
```

Example:

```
2026-01-15 09:30:42 | INFO     | core.state  | Loaded state from ./state/email_send_state.json: 47 processed items
2026-01-15 09:30:43 | CONSOLE  | core.console | ✓ Configuration loaded
2026-01-15 09:30:44 | ERROR    | office.outlook.sender | Failed to send email: RPC server unavailable
```

#### REQ-LOG-07: Console format — `ConsoleFormatter`

Records printed to the console stream must use short human-readable prefixes and ANSI colours (via `colorama`) when the output stream is a TTY:

| Level      | Prefix                 | Colour  |
| ---------- | ---------------------- | ------- |
| `DEBUG`    | `[DEBUG]`              | Cyan    |
| `INFO`     | `[*]`                  | Green   |
| `CONSOLE`  | (suppressed by filter) | —       |
| `WARNING`  | `[!]`                  | Yellow  |
| `ERROR`    | `[ERROR]`              | Red     |
| `CRITICAL` | `[CRITICAL]`           | Magenta |

Colour codes must be stripped when `sys.stdout.isatty()` returns `False` (e.g. when output is redirected to a file).

#### REQ-LOG-08: UTF-8 console stream

The `StreamHandler` must wrap `sys.stdout.buffer` with an explicit `io.TextIOWrapper(encoding="utf-8", errors="replace")` to avoid `UnicodeEncodeError` on Windows consoles that default to a non-UTF-8 code page. `line_buffering=True` must be set to ensure output is flushed promptly.

#### REQ-LOG-09: Log level configuration

The file log level is controlled by the `LOG_LEVEL` environment variable (default: `INFO`). Valid values are `DEBUG`, `INFO`, `WARNING`, and `ERROR`. Invalid values must be silently normalised to `INFO`.

### Non-functional constraints

- Log files are plain UTF-8 text — no binary formats, no rotation at this time.
- Log directories are always relative to the tool's working directory unless an absolute path is configured.
- In unit tests, the logging system must not require real file I/O. Tests must mock or suppress handler initialisation rather than creating actual log files in the workspace.
