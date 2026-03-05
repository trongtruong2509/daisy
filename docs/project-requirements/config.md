## Configuration Requirements for Daisy Automation Platform

### Rationale

All tools share the same runtime environment but have independent settings. A two-layer `.env` system (global → tool-local) allows shared defaults to coexist with per-tool overrides without duplication. Typed getters eliminate boiler-plate `os.getenv()` + casting scattered across the codebase. Interactive prompting with validation closes the gap between minimal `.env` setup and a fully guided first-run experience.

### Requirements

#### REQ-CFG-01: `ConfigManager` is the sole `.env` loading mechanism

All configuration loading must go through `core.config_manager.ConfigManager`. No tool or module may call `os.getenv()` directly for settings that should be user-configurable. No tool may call `dotenv.load_dotenv()` directly.

```python
from core.config_manager import ConfigManager

mgr = ConfigManager()
mgr.load_env([global_env_path, local_env_path])
```

#### REQ-CFG-02: Two-layer `.env` loading — global then local

When a tool loads configuration it must supply both the project-level `.env` (global defaults) and the tool-level `.env` (tool-specific overrides) in that order. Later files override earlier ones.

```python
global_env = PROJECT_ROOT / ".env"
local_env  = TOOL_DIR / ".env"
mgr.load_env([global_env, local_env])
```

If a file does not exist it must be silently skipped — `load_env()` must not raise for missing files.

#### REQ-CFG-03: Typed getters — `get()`, `get_bool()`, `get_int()`, `get_path()`, `get_list()`

All environment variable access must use the typed static methods on `ConfigManager`:

| Method                                     | Return type | Notes                                                       |
| ------------------------------------------ | ----------- | ----------------------------------------------------------- |
| `get(key, default="")`                     | `str`       | Raw string                                                  |
| `get_bool(key, default=False)`             | `bool`      | Truthy: `"true"`, `"1"`, `"yes"`, `"on"` (case-insensitive) |
| `get_int(key, default=0)`                  | `int`       | Returns `default` on parse failure                          |
| `get_path(key, default="", base_dir=None)` | `Path`      | Relative paths resolved against `base_dir`                  |
| `get_list(key, default=[], separator=",")` | `list[str]` | Strips whitespace from each item                            |

#### REQ-CFG-04: Path normalisation — `_normalize_path_input()`

All path values (from `.env` or user input) must be normalised before use via `ConfigManager._normalize_path_input()`. The normaliser must:

- Strip leading and trailing whitespace.
- Remove surrounding single or double quotes (common from batch-file variable expansion).
- Strip carriage-return (`\r`) and null (`\x00`) characters embedded by Windows batch pipelines.

This normalisation is applied automatically inside `get_path()` and `prompt_for_value()`.

#### REQ-CFG-05: Interactive prompting via `prompt_for_value()`

When a required value is absent from both `.env` files the tool must call `ConfigManager.prompt_for_value()` to collect it interactively. The method must:

- Print the description at `WARNING` level via `cprint()`.
- Optionally display an example value.
- Loop until the user provides a non-empty, validated input.
- Accept an optional `validator` callable with signature `(str) -> (bool, str)`.

```python
date = mgr.prompt_for_value(
    "DATE",
    "Enter payroll month",
    example="01/2026",
    validator=ConfigManager.validate_date,
)
```

#### REQ-CFG-06: Built-in validators

`ConfigManager` must provide these static validators for reuse across tools:

| Validator                                  | Rule                             |
| ------------------------------------------ | -------------------------------- |
| `validate_date(value)`                     | `MM/YYYY` format, month 01–12    |
| `validate_email(value)`                    | Standard email regex             |
| `validate_file_path(value, base_dir=None)` | File must exist at resolved path |

All validators return `(bool, str)` — `(True, "")` on success, `(False, error_message)` on failure.

#### REQ-CFG-07: Persist values with `save_to_env()`

When the tool persists a prompted value back to a `.env` file for subsequent runs, it must use `ConfigManager.save_to_env(env_file, key, value)`. Direct `open()` + write to `.env` files is forbidden.

#### REQ-CFG-08: Tool config dataclasses extend or wrap `ConfigManager`

Each tool must define a typed config dataclass (e.g. `PayslipConfig`, `GetAttachmentConfig`) populated by `ConfigManager` getters. The dataclass must expose:

- A `validate() -> list[str]` method returning human-readable error messages (empty list = valid).
- An `ensure_directories()` method that calls `mkdir(parents=True, exist_ok=True)` on all path-type fields.
- A `is_valid() -> bool` convenience property.

#### REQ-CFG-09: `DRY_RUN` defaults to `True` everywhere

All tools must default `dry_run=True`. The setting must be explicitly set to `false` in `.env` to enable live mode. This ensures a misconfigured or first-time run never accidentally mutates Outlook or the file system.

#### REQ-CFG-10: Config validation at startup — fail fast

Each tool must call `config.validate()` immediately after loading config and before any COM or file I/O. If errors are present, print them via `cprint()` at `ERROR` level and exit with code 1.

### Non-functional constraints

- `ConfigManager` must not import any `office/` module to avoid circular dependencies.
- Config dataclasses must be serialisable to plain dictionaries for logging/debugging purposes.
- Testing: unit tests must be able to instantiate config objects with constructor arguments without loading any `.env` files.
