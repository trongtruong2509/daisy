# Daisy Automation Platform — AI Coding Instructions

## Architecture Overview

**Daisy** is a Windows-only office automation platform structured as a shared foundation (`core/`, `office/`) with isolated tools under `tools/<tool-name>/`.

```
docs/          # Documentation, including project requirements and tool-specific requirements
core/          # Shared infrastructure: config, logging, state tracking, retry, console output
office/        # Wrappers around Windows COM: Excel (openpyxl + COM) and Outlook
parsing/       # HTML/text parsers (shared utilities)
tools/         # Self-contained tools; each has its own config.py, main.py, .env, tests/
```

Each tool in `tools/` is independently runnable and imports from the project root via:

```python
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
```

## Developer Workflows

**Setup (once):** `setup.bat` — creates `venv/` and installs `requirements.txt`.

**Run interactively:** `run.bat` — shows tool menu; activates `venv` automatically.

**Run a specific tool:** `run.bat payslip-phuclong`

**Run tests (from tool directory):**

```cmd
cd tools/payslip-phuclong-ecom
..\..\..\venv\Scripts\pytest
```

Default config (`pytest.ini`) excludes `integration` tests. To include them: `pytest -m integration` (requires live Excel/Outlook COM on Windows).

**Direct execution during development:**

```cmd
cd tools/payslip-phuclong-ecom
..\..\..\venv\Scripts\python main.py
```

## Key Conventions

### Configuration: two-layer `.env` + `ConfigManager`

- Global `.env` lives at project root; tool-specific `.env` lives in `tools/<tool>/`.
- `ConfigManager.load_env([global_env, local_env])` — later files override earlier ones.
- Typed getters: `get()`, `get_bool()`, `get_path()`. Prompt with `prompt_for_value(key, label, default, validator=...)`.
- `dry_run=True` is the default safety default everywhere. **Always test with `DRY_RUN=true` first.**

### Console output: always use `cprint()`

Use `core.console.cprint(message, level=...)` instead of bare `print()`. Levels: `PHASE`, `SUCCESS`, `ERROR`, `WARNING`, `INFO`, `PROGRESS`, `SUMMARY`. Output is automatically mirrored to log files.

### State tracking: idempotent operations via `StateTracker`

`core.state.StateTracker` and `ContentHashTracker` prevent duplicate operations (e.g., re-sending emails after a crash). Call `tracker.is_processed(id)` before acting, `tracker.mark_processed(id)` after, and `tracker.save()` periodically. State files live in `tools/<tool>/state/` as JSON.

### Retry: COM operations use `@retry_operation`

Outlook/Excel COM calls are wrapped with `core.retry.retry_operation(RetryConfig(...))`. Common transient COM error codes are listed in `core/retry.py`.

### Outlook sending: always use `OutlookSender` as context manager

```python
with OutlookSender(account=config.outlook_account, dry_run=config.dry_run, state_tracker=tracker) as sender:
    sender.send(NewEmail(to=[...], subject=..., body=...))
```

### Excel COM: set B3=MNV, let Excel recalculate

`PayslipGenerator` sets `TBKQ!B3 = MNV` to trigger VLOOKUP/XLOOKUP formulas, copies the sheet to a new workbook, and pastes as values. This replicates the original VBA macro approach and handles any formula complexity. Never attempt to replicate formula logic in Python.

## Adding a New Tool

1. Create `tools/<tool-name>/` with `__init__.py`, `config.py`, `main.py`.
2. Extend `core.config.Config` or instantiate `ConfigManager` directly for tool-specific settings.
3. Add tool entry to `run.bat`'s menu section.
4. Create `tools/<tool-name>/tests/` with `unit/`, `component/`, `integration/` subdirectories and a `pytest.ini` mirroring the existing one.

## External Dependencies

- **pywin32** — required for all Excel/Outlook COM automation; Windows-only.
- **pikepdf** — PDF password protection for generated payslips.
- **python-dotenv** — `.env` loading.
- **colorama** — cross-platform colored terminal output.
- No web services, databases, or cloud dependencies.
