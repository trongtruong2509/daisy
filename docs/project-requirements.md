## COM Initialisation and Lifecycle

### Rationale

Windows COM (Component Object Model) underpins all Excel and Outlook automation via `win32com`. COM uses the Single-Threaded Apartment (STA) model by default: `pythoncom.CoInitialize()` must be called once per thread before any COM object is created on that thread, and `pythoncom.CoUninitialize()` must be called on the same thread after all COM objects have been released.

Failing to follow these rules produces `RPC_E_WRONG_THREAD` errors, memory leaks, or silent data corruption in multi-threaded scenarios.

### Requirements

#### REQ-COM-01: Centralised COM bootstrapping

All COM initialisation and uninitialisation logic must live in `office/com.py`. No other file in the project may import `pythoncom` directly or call `CoInitialize` / `CoUninitialize`.

The module must expose:

- `is_available() -> bool` — returns `True` when `pywin32` / `pythoncom` is installed.
- `com_initialized()` — a context manager that calls `CoInitialize()` on entry and `CoUninitialize()` on exit, in the calling thread.

#### REQ-COM-02: `office/` classes use `com_initialized()`

Every `office/` class that creates COM objects (`ExcelComReader`, `PdfConverter`, `OutlookClient`, and all subclasses) must wrap its COM lifecycle inside `com_initialized()`. The context must be entered in `__enter__` (or the equivalent `open()` / `connect()` method) and exited in `__exit__` (or `close()` / `disconnect()`).

```python
# Correct pattern for any office/ COM class
from office.com import com_initialized

class ExcelComReader:
    def __enter__(self):
        self._com_ctx = com_initialized()
        self._com_ctx.__enter__()
        self._open_workbook()
        return self

    def __exit__(self, *exc):
        self._close_workbook()
        self._com_ctx.__exit__(*exc)
        return False
```

#### REQ-COM-03: Tools never import `pythoncom` or `win32com` directly

Code under `tools/` must not import `pythoncom`, `win32com`, or `pywintypes`. All COM interaction must go through `office/` library classes. If a tool-layer class (e.g., `PayslipGenerator`) currently contains raw `win32com.client.Dispatch()` calls, that logic must be refactored into an `office/` class.

#### REQ-COM-04: One COM instance per thread — never share across threads

COM objects are not thread-safe. Each thread that performs COM work must:

1. Enter `com_initialized()` at the start of the thread function (outermost scope).
2. Create its own `office/` class instances (e.g., its own `ExcelComReader`, `OutlookSender`).
3. Release all COM objects (exit context managers) before the thread function returns.
4. Never pass COM object references to another thread.

```python
# Correct multi-threaded pattern
from concurrent.futures import ThreadPoolExecutor
from office.com import com_initialized
from office.excel.reader import ExcelComReader

def worker(employee, source_path):
    with com_initialized():            # COM init for this thread
        with ExcelComReader(source_path) as r:  # own COM instance
            return r.read_cell("Sheet1", "A1")

with ThreadPoolExecutor(max_workers=4) as pool:
    results = list(pool.map(lambda e: worker(e, path), employees))
```

#### REQ-COM-05: Recommended thread pool sizes

| Workload                              | Recommended `max_workers` | Reason                                                                      |
| ------------------------------------- | ------------------------- | --------------------------------------------------------------------------- |
| Excel file generation (CPU/COM bound) | 4–8                       | Excel spawns a separate process per `Dispatch()`; more than 8 wastes memory |
| Outlook email sending (COM/IO bound)  | 2–4                       | Outlook serialises internally; additional threads queue behind the COM lock |

Never use `ProcessPoolExecutor` for COM work — COM objects are not picklable and cannot be passed between processes.

#### REQ-COM-06: `HAS_PYTHONCOM` / `HAS_WIN32COM` guards forbidden outside `office/`

The pattern below is forbidden in any file other than `office/com.py`:

```python
# FORBIDDEN everywhere except office/com.py
try:
    import pythoncom
    HAS_PYTHONCOM = True
except ImportError:
    HAS_PYTHONCOM = False
    pythoncom = None
```

All availability checking is done once in `office/com.py` via `is_available()`.

#### REQ-COM-07: Context managers are mandatory — no bare `open()` / `close()` in production code

All production code that uses `office/` classes must do so via the context manager protocol (`with` statement). Calling `open()` / `connect()` without a corresponding `close()` / `disconnect()` is not permitted, as it leaks COM references and Excel/Outlook processes.

#### REQ-COM-08: Never forcefully quit running applications — preserve user state

Calling `.Quit()` on an Excel or Outlook COM object terminates the **entire application** (not just the tool's session), which closes all user workbooks and emails. To protect user data and allow the tool to coexist with the user's own Excel/Outlook sessions:

1. **For Excel:** Detect whether the application was running before the automation started using `GetObject("Excel.Application")`. Only call `excel.Quit()` if the tool created a new instance; if the user already had Excel open, close only the workbook(s) and release the COM reference without calling `Quit()`.

2. **For Outlook:** Same pattern — use `GetObject("Outlook.Application")` to detect if Outlook was already running. Only call `outlook.Quit()` if the tool started it.

3. **Implementation location:** Helper functions `get_or_create_excel()` and `get_or_create_outlook()` must live in `office/excel/helpers.py` and `office/outlook/helpers.py` respectively. These helpers return a tuple `(app_object, was_already_running)`. All `office/` classes (`ExcelComReader`, `PdfConverter`, `OutlookClient`, etc.) use these helpers in their `__enter__` and `__exit__` methods.

**Pseudocode pattern (actual implementation in `office/excel/helpers.py`):**

```python
# office/excel/helpers.py
def get_or_create_excel():
    """
    Get a reference to Excel.Application if running, or create a new instance.

    Returns:
        (excel_app, was_already_running)' tuple.
        was_already_running is True if Excel was already running before this call.
    """
    try:
        # Try to get the running instance without creating a new one
        excel = win32com.client.GetObject(class_="Excel.Application")
        return excel, True
    except (pythoncom.com_error, AttributeError):
        # No running instance; create a new one
        excel = win32com.client.Dispatch("Excel.Application")
        return excel, False
```

Then in `ExcelComReader.__exit__`, use the flag:

```python
def __exit__(self, *exc):
    if self._workbook:
        self._workbook.Close(SaveChanges=False)
    if self._excel and not self._was_already_running:
        self._excel.Quit()  # Only quit if we created it
```

4. **Test requirement:** Integration tests must verify that running the tool with an open Excel/Outlook window does not close that window.

### Non-functional constraints

- This project is Windows-only. COM is not available on Linux or macOS. All `office/` classes must raise a clear `ImportError` (or `NotImplementedError`) with an actionable message when `pywin32` is not installed, rather than failing silently.
- Tests that do not require live COM (unit and component tests) must mock the `office/` classes at the boundary, not mock `pythoncom` internals.
