## Console Output Requirements for Daisy Automation Platform

### Rationale

Consistent, coloured console output gives users clear progress feedback during long-running automation tasks. A single `cprint()` function serves as the only approved output path, ensuring every user-visible message is also captured in the log file for post-run auditability. Bare `print()` calls scattered across the codebase produce output that is invisible to the log system and inconsistent in style.

### Requirements

#### REQ-CON-01: `cprint()` is the sole console output mechanism

All user-facing output must go through `core.console.cprint()`. Bare `print()` calls are forbidden in production code except for interactive prompt responses (e.g. `input()` companion lines such as `print("Example: ...")` in `prompt_for_value()`).

```python
from core.console import cprint

cprint("Starting process", level="PHASE")
cprint("File loaded successfully", level="SUCCESS")
cprint("Something went wrong", level="ERROR")
```

#### REQ-CON-02: Supported output levels

`cprint()` must accept the following `level` values (case-insensitive):

| Level         | Visual style                | Use case                          |
| ------------- | --------------------------- | --------------------------------- |
| `INFO`        | Plain white text            | General informational messages    |
| `BANNER`      | Double-line box (╔═╗ / ╚═╝) | Tool start / section headers      |
| `PHASE`       | `>> ` prefix, cyan          | Top-level workflow phases         |
| `SUCCESS`     | `✓ ` prefix, green          | Completed operations              |
| `ERROR`       | `✗ ` prefix, red            | Failed operations                 |
| `WARNING`     | `⚠ ` prefix, yellow         | Caution / user attention required |
| `PROGRESS`    | Plain white, no prefix      | Incremental item-level updates    |
| `PRE_SUMMARY` | Plain (no colour)           | Pre-summary separator text        |
| `SUMMARY`     | Green text                  | Summary table rows                |

Unknown level values must fall back to `INFO`.

#### REQ-CON-03: Every `cprint()` call also logs at `CONSOLE` level

Internally, every formatting function called by `cprint()` must call `logger.log(CONSOLE, message)` after printing to stdout, so the message appears in the run log file (see REQ-LOG-03). The `cprint` module must use `logging.getLogger("core.console")` for this purpose.

#### REQ-CON-04: Cross-platform colour via `colorama`

Colour formatting must be implemented using `colorama` (`Fore.*`, `Style.*`) with `init(autoreset=True)` called at module load time. This ensures ANSI codes work on Windows PowerShell and Windows Terminal without manual reset calls.

#### REQ-CON-05: `indent` parameter for nested output

`cprint()` must accept an optional `indent: int = 0` parameter that prepends the specified number of spaces before the formatted message. This is used for visually nesting sub-step output beneath a phase header.

```python
cprint("Phase: Reading data", level="PHASE")
cprint("Row 1 processed", level="SUCCESS", indent=2)
cprint("Row 2 processed", level="SUCCESS", indent=2)
```

#### REQ-CON-06: High-level helpers — `cprint_banner()` and `cprint_summary_box()`

The console module must expose:

- `cprint_banner(title, subtitle="")` — prints a `BANNER`-level box with title and optional subtitle.
- `cprint_summary_box(title, stats)` — prints a structured summary box with labelled rows, used at end-of-run.
- `cprint_summary_box_lite(title, stats)` — a compact variant used for pre-run configuration summaries.

These helpers must internally call `cprint()`, not `print()` directly, so all output is logged.

#### REQ-CON-07: Box width and drawing characters are fixed constants

The banner box must use double-line box-drawing characters (`╔`, `═`, `╗`, `╚`, `╝`, `║`) with a fixed minimum inner width of 64 characters. Lines shorter than the minimum width are left-padded with spaces.

### Non-functional constraints

- `cprint()` must be safe to call before `setup_logging()` is invoked (messages will still print to stdout; the log call will silently fail or queue).
- The module must not import any `office/` or `tools/` module.
- In unit tests, `cprint()` must be patchable at the `core.console.cprint` symbol without side-effects on the logging system.
