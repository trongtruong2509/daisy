## Retry Requirements for Daisy Automation Platform

### Rationale

COM automation calls to Excel and Outlook fail transiently due to timing, application busyness, or RPC stack congestion. Rather than requiring every call site to implement its own retry loop, a centralised retry decorator with configurable exponential backoff handles these failures consistently. Permanent errors (wrong arguments, access denied) must not be retried indefinitely; a maximum attempts cap is mandatory.

### Requirements

#### REQ-RETRY-01: All COM operation retries use `@retry_operation`

Any `office/` method that performs a COM call susceptible to transient failure must be decorated with `@retry_operation`. Inline `try/except` loops with `time.sleep()` outside the `core.retry` module are forbidden.

```python
from core.retry import retry_operation, RetryConfig

@retry_operation(RetryConfig(max_attempts=3, base_delay=2.0))
def _send_outlook_item(self, mail_item):
    mail_item.Send()
```

#### REQ-RETRY-02: `RetryConfig` is the sole configuration object

Retry behaviour must be configured via the `core.retry.RetryConfig` dataclass. Magic numbers for delays or attempt counts embedded directly in `except` blocks are forbidden.

| Field                 | Type               | Default        | Constraint     |
| --------------------- | ------------------ | -------------- | -------------- |
| `max_attempts`        | `int`              | 3              | ≥ 1            |
| `base_delay`          | `float`            | 2.0 s          | ≥ 0            |
| `max_delay`           | `float`            | 30.0 s         | ≥ `base_delay` |
| `exponential_base`    | `float`            | 2.0            | —              |
| `retry_on_exceptions` | `tuple[type, ...]` | `(Exception,)` | —              |

`RetryConfig.__post_init__()` must normalise invalid values silently (e.g. `max_attempts < 1` → `1`).

#### REQ-RETRY-03: Exponential backoff — `calculate_delay()`

The delay before retry attempt `n` (1-indexed) is:

$$
\text{delay} = \min\bigl(\text{base\_delay} \times \text{exponential\_base}^{n-1},\ \text{max\_delay}\bigr)
$$

`core.retry.calculate_delay(attempt, config)` must implement this formula and be usable independently for testing.

#### REQ-RETRY-04: Transient COM error detection — `is_transient_error()`

`core.retry.is_transient_error(exc)` must return `True` for:

- `pywintypes.com_error` exceptions whose `args[0]` is in `TRANSIENT_ERROR_CODES`:
  - `-2147352567` — "call rejected by callee" (Outlook/Excel busy)
  - `-2147023174` — "RPC server unavailable"
  - `-2147024891` — "access denied" (sometimes transient)
- Any exception whose string representation contains any of: `"timeout"`, `"timed out"`, `"temporarily unavailable"`, `"connection refused"`, `"rpc"`, `"busy"`.

The check must not import `pywintypes` directly; it must use string-based type name matching to remain importable in non-Windows test environments.

#### REQ-RETRY-05: `RetryExhaustedError` on final failure

When all attempts are exhausted, `retry_operation` must raise `core.retry.RetryExhaustedError` with:

- A human-readable message stating the function name and number of attempts.
- `last_exception` attribute holding the final underlying exception.
- `attempts` attribute holding the total attempt count.

```python
except RetryExhaustedError as e:
    logger.error(f"Sending failed after {e.attempts} attempts: {e.last_exception}")
```

#### REQ-RETRY-06: `on_retry` callback for observability

`retry_operation` must accept an optional `on_retry: Callable[[int, Exception], None]` argument. When provided, it is called before each retry (not before the first attempt) with the current attempt number and the exception that triggered the retry. This enables tools to log or display retry progress without modifying the decorated function.

#### REQ-RETRY-07: Retry must log at `WARNING` level

Each retry attempt must log a `WARNING` message including the attempt number, max attempts, delay, and exception summary. The final failure before raising `RetryExhaustedError` must be logged at `ERROR` level.

### Non-functional constraints

- `core.retry` must not import `pythoncom`, `win32com`, or `pywintypes`. COM-specific error codes are plain integer constants.
- `core.retry` must not import any `office/` module.
- Unit tests must be able to verify retry behaviour using mock functions that raise controlled exceptions without needing a Windows COM environment.
