"""
Backward-compatibility shim — do not import directly.

All COM bootstrapping has moved to ``office.utils.com``.
This module re-exports the full public API so that any remaining
caller works without changes during a transition period.
"""
from office.utils.com import (  # noqa: F401
    HAS_COM,
    is_available,
    ensure_com_available,
    com_initialized,
    get_pythoncom,
    get_win32com_client,
    get_pywintypes,
)


from __future__ import annotations

import contextlib
import logging
from typing import Generator

logger = logging.getLogger(__name__)

# ── Availability detection (single location) ────────────────────

try:
    import pythoncom as _pythoncom
    import win32com.client as _win32com_client
    import pywintypes as _pywintypes

    HAS_COM = True
except ImportError:
    _pythoncom = None  # type: ignore[assignment]
    _win32com_client = None  # type: ignore[assignment]
    _pywintypes = None  # type: ignore[assignment]
    HAS_COM = False


# ── Public helpers ───────────────────────────────────────────────

def is_available() -> bool:
    """Return ``True`` when pywin32 / pythoncom is installed."""
    return HAS_COM


def ensure_com_available() -> None:
    """Raise ``ImportError`` with an actionable message if pywin32 is absent.

    Call this at the top of any ``office/`` class ``__init__`` that requires
    COM so the user gets a clear message instead of a silent failure.
    """
    if not HAS_COM:
        raise ImportError(
            "pywin32 is required for COM automation on Windows. "
            "Install it with:  pip install pywin32"
        )


@contextlib.contextmanager
def com_initialized() -> Generator[None, None, None]:
    """Context manager that brackets COM apartment initialisation.

    Calls ``pythoncom.CoInitialize()`` on entry and
    ``pythoncom.CoUninitialize()`` on exit **in the calling thread**.

    Safe to nest — ``CoInitialize`` is reference-counted on Windows;
    each successful ``CoInitialize`` must be paired with a
    ``CoUninitialize``.

    Raises:
        ImportError: If pywin32 is not installed.

    Usage::

        from office.com import com_initialized

        with com_initialized():
            excel = win32com.client.Dispatch("Excel.Application")
            ...
    """
    ensure_com_available()

    _pythoncom.CoInitialize()
    logger.debug("COM apartment initialised on current thread")
    try:
        yield
    finally:
        try:
            _pythoncom.CoUninitialize()
            logger.debug("COM apartment uninitialised on current thread")
        except Exception:
            # Safe to swallow — may already be uninitialised if an earlier
            # CoUninitialize happened (e.g. nested context managers).
            pass


# ── Internal accessors (for office/ classes only) ────────────────
# These give office/ modules access to the underlying libraries without
# importing pythoncom/win32com/pywintypes directly.  They are NOT part
# of the public API and must never be imported by tools/.

def get_pythoncom():
    """Return the ``pythoncom`` module.  For ``office/`` internal use only."""
    ensure_com_available()
    return _pythoncom


def get_win32com_client():
    """Return ``win32com.client``.  For ``office/`` internal use only."""
    ensure_com_available()
    return _win32com_client


def get_pywintypes():
    """Return the ``pywintypes`` module.  For ``office/`` internal use only."""
    ensure_com_available()
    return _pywintypes
