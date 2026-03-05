"""
Generic COM application lifecycle helpers.

Two usage patterns are supported:

**Background / isolated** (``create_app_background`` / ``create_*_background``):
    Always calls ``Dispatch`` to spin up a **new, separate process**.
    The new instance starts with ``Visible = False`` and is completely
    independent of any window the user has open.  The tool always owns it
    and always calls ``Quit()`` when finished.
    Use this for **all Excel batch processing**: file reading, PDF conversion,
    payslip generation, etc.  The user''s existing Excel is untouched because
    we are running in a *different process*.

**Attach-or-create** (``get_or_create_app`` / ``get_or_create_*``):
    Tries ``GetObject`` first so it can reuse an already-running instance.
    If the application was already running the tool will *not* call ``Quit()``
    when finished, preserving the user''s session.
    Intended for applications that are single-instance by nature, such as
    **Outlook** (where you always want to send through the running session).

Usage -- background (Excel processing)::

    from office.utils.helpers import create_excel_background, safe_quit_excel

    excel, was_running = create_excel_background()   # was_running is always False
    try:
        excel.Visible = False
        ...
    finally:
        safe_quit_excel(excel, was_running)          # always calls Quit()

Usage -- attach-or-create (Outlook sending)::

    from office.utils.helpers import get_or_create_outlook, safe_quit_outlook

    outlook, was_running = get_or_create_outlook()
    try:
        ...
    finally:
        safe_quit_outlook(outlook, was_running)      # skips Quit() if was_running

Requirement references: REQ-COM-08.
"""

from __future__ import annotations

import logging
from typing import Tuple

from office.utils.com import ensure_com_available, get_win32com_client, get_pythoncom

logger = logging.getLogger(__name__)


# -- Background / isolated (always new process) -----------------------------------------------

def create_app_background(com_class: str) -> Tuple[object, bool]:
    """Spin up a **new, isolated** COM application instance via ``Dispatch``.

    Never attaches to a running instance -- always creates a separate process.
    This is the correct choice for any background batch processing task
    (reading files, converting PDFs, generating payslips, etc.) because
    the new process starts invisible and does not share windows or state
    with whatever the user may have open.

    Args:
        com_class: The COM ProgID, e.g. ``"Excel.Application"``.

    Returns:
        ``(app, False)`` -- the second element is always ``False`` to signal
        that ``safe_quit_app()`` must call ``Quit()`` when finished.

    Raises:
        ImportError: If pywin32 is not installed.
    """
    ensure_com_available()
    win32 = get_win32com_client()
    # DispatchEx bypasses the COM ROT (Running Object Table) and always
    # creates a fresh out-of-process server via CoCreateInstance, regardless
    # of whether the application is already running on the desktop.
    # Dispatch() checks the ROT first and can silently return the user's
    # existing visible instance — DispatchEx never does this.
    app = win32.DispatchEx(com_class)
    logger.debug("Created isolated background %s process via DispatchEx", com_class)
    return app, False


def create_excel_background() -> Tuple[object, bool]:
    """Convenience alias: ``create_app_background("Excel.Application")``."""
    return create_app_background("Excel.Application")


# -- Attach-or-create (reuse running session) -------------------------------------------------

def get_or_create_app(com_class: str) -> Tuple[object, bool]:
    """Get a reference to a running COM application or create a new instance.

    Tries ``GetObject(Class=com_class)`` first (attaches to the running ROT entry
    without spawning a new process).  Falls back to ``Dispatch(com_class)`` when
    no running instance is found.

    Prefer ``create_app_background()`` for any **background processing** task.
    This function is suitable when you intentionally want to reuse the user''s
    running session, e.g. sending emails through the running Outlook.

    Args:
        com_class: The COM ProgID, e.g. ``"Outlook.Application"``.

    Returns:
        ``(app, was_already_running)`` tuple.
        *was_already_running* is ``True`` if the application was running before
        this call -- in that case ``safe_quit_app()`` will **not** call ``Quit()``.

    Raises:
        ImportError: If pywin32 is not installed.
    """
    ensure_com_available()
    win32 = get_win32com_client()
    pythoncom = get_pythoncom()

    try:
        app = win32.GetObject(Class=com_class)
        logger.debug("Attached to existing %s instance", com_class)
        return app, True
    except (pythoncom.com_error, AttributeError, TypeError):
        app = win32.Dispatch(com_class)
        logger.debug("Created new %s instance", com_class)
        return app, False


# -- Shared quit helper -----------------------------------------------------------------------

def safe_quit_app(app: object, was_already_running: bool) -> None:
    """Quit a COM application only if the tool created it.

    Args:
        app: The COM application object (any Office app).
        was_already_running: Flag returned by ``get_or_create_app()`` or
                             ``create_app_background()`` (always ``False``).
    """
    if app is None:
        return

    if was_already_running:
        logger.debug(
            "%r was already running -- releasing reference without Quit()", app
        )
    else:
        try:
            app.Quit()
            logger.debug("Quit() called on tool-created COM instance")
        except Exception:
            pass  # Application may have been closed by the user in the meantime


# -- Convenience aliases ----------------------------------------------------------------------

def get_or_create_excel() -> Tuple[object, bool]:
    """Attach-or-create alias for ``Excel.Application``.

    .. warning::
        Do **not** use this for background file processing -- it may attach to
        the user''s visible Excel window.  Use ``create_excel_background()``
        instead.
    """
    return get_or_create_app("Excel.Application")


def safe_quit_excel(excel_app: object, was_already_running: bool) -> None:
    """Convenience alias: ``safe_quit_app()`` for Excel."""
    safe_quit_app(excel_app, was_already_running)


def get_or_create_outlook() -> Tuple[object, bool]:
    """Convenience alias: ``get_or_create_app("Outlook.Application")``."""
    return get_or_create_app("Outlook.Application")


def safe_quit_outlook(outlook_app: object, was_already_running: bool) -> None:
    """Convenience alias: ``safe_quit_app()`` for Outlook."""
    safe_quit_app(outlook_app, was_already_running)
