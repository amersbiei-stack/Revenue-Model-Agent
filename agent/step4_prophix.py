"""Step 4 — Prophix Analyzer refresh (fully automated).

Drives the Prophix Analyzer CV2 pane end-to-end, no human interaction:

  1. Open the workbook and activate `Units Bookings DV`.
  2. Force Excel to the foreground (SendKeys requires it).
  3. Send Alt+N+Y+1 via Excel's Application.SendKeys to open the pane.
  4. Use UI Automation (pywinauto) to find the pane's refresh split-button,
     expand its dropdown, and click "All Sheets".
  5. Sleep PROPHIX_WAIT_SECONDS (default 600) for the refresh to complete.

Step 5 is the safety net: if the refresh didn't actually run, its sum
checks on all 5 DV tabs will fail and a Step 5 failure email will draft.

Assumptions (documented in SETUP.md):
  - Prophix Analyzer CV2 is installed and signed in on this machine.
  - Excel Trust Center includes the project folder (no macro prompt).
"""
import time
from pathlib import Path

import win32con
import win32gui

from agent import config, excel_com
from agent.logging_setup import get_logger


def _bring_excel_to_front(app, log) -> None:
    """Restore + foreground the Excel window. SendKeys won't land otherwise."""
    try:
        hwnd = app.Hwnd
    except Exception as e:
        log.warning("Could not read Application.Hwnd: %s", e)
        return
    if not hwnd:
        return
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.3)
    except Exception as e:
        log.warning("Could not bring Excel to foreground: %s", e)


def _analyzer_pane(main_window, timeout: float):
    """Return the pywinauto wrapper for the Prophix Analyzer pane."""
    return main_window.child_window(
        title_re="Prophix Analyzer.*",
        control_type="Pane",
        found_index=0,
    ).wait("exists visible", timeout=timeout)


def _open_analyzer_pane(app, main_window, log) -> None:
    """If the pane isn't already open, send Alt+N+Y+1 and wait for it."""
    from pywinauto.keyboard import send_keys
    from pywinauto.timings import TimeoutError as PWATimeoutError

    pane = main_window.child_window(
        title_re="Prophix Analyzer.*",
        control_type="Pane",
    )
    if pane.exists(timeout=1):
        log.info("Prophix Analyzer pane already open")
        return

    # Three-attempt retry. pywinauto's send_keys drives OS-level keyboard
    # input (unlike Excel.Application.SendKeys, which quietly fails for
    # ribbon KeyTip navigation). Focus must belong to Excel for ribbon
    # shortcuts to register.
    for attempt in range(1, 4):
        try:
            main_window.set_focus()
        except Exception as e:
            log.warning("Excel set_focus failed (attempt %d): %s", attempt, e)
        time.sleep(0.8)

        log.info("Sending Alt+N, then Y, then 1 (attempt %d/3)", attempt)
        # Explicit VK_MENU down/up makes the Alt boundary unambiguous —
        # the %n shorthand sometimes releases Alt before 'n' registers.
        send_keys("{VK_MENU down}n{VK_MENU up}", pause=0.15)
        time.sleep(1.5)  # let the Insert tab render its command KeyTips
        # Two-char KeyTips are matched by Excel's state machine after each
        # key; split Y and 1 with a pause or the second char can be lost.
        send_keys("y", pause=0.2)
        time.sleep(0.35)
        send_keys("1", pause=0.2)

        try:
            pane.wait("exists visible", timeout=12)
            log.info("Prophix Analyzer pane opened")
            return
        except PWATimeoutError:
            log.warning(
                "Prophix Analyzer pane did not appear within 12s "
                "(attempt %d/3)", attempt,
            )
            # Drop out of any stuck KeyTip mode before retrying.
            try:
                send_keys("{ESC}{ESC}", pause=0.1)
            except Exception:
                pass
            time.sleep(1.0)

    # Dump what UIA can see so we can debug the next run.
    log.error("Could not open Prophix Analyzer pane via Alt+N+Y+1. "
              "Dumping top-level children of the Excel window:")
    try:
        for child in main_window.children()[:40]:
            try:
                log.error("  title=%r class=%s",
                          child.window_text(),
                          child.friendly_class_name())
            except Exception:
                pass
    except Exception as e:
        log.error("  (could not enumerate children: %s)", e)
    raise RuntimeError(
        "Step 4: Prophix Analyzer pane never opened after Alt+N+Y+1. "
        "Confirm the shortcut still applies and that Excel keeps focus "
        "during Step 4 (do not click away while the agent runs)."
    )


def _click_refresh_all_sheets(main_window, log) -> None:
    """Expand the pane's Refresh split-button and click 'All Sheets'."""
    pane = _analyzer_pane(main_window, timeout=10)

    # The refresh control at the bottom of the pane is a split-button named
    # "Refresh". Try a few selector shapes in case the UIA tree differs by
    # Prophix build.
    candidates = [
        {"title": "Refresh", "control_type": "SplitButton"},
        {"title_re": "Refresh.*", "control_type": "SplitButton"},
        {"title": "Refresh", "control_type": "Button"},
        {"title_re": "Refresh.*", "control_type": "Button"},
    ]
    refresh_btn = None
    for kwargs in candidates:
        try:
            candidate = pane.child_window(**kwargs)
            if candidate.exists(timeout=2):
                refresh_btn = candidate
                break
        except Exception:
            continue
    if refresh_btn is None:
        raise RuntimeError(
            "Step 4: could not locate the 'Refresh' control in the Prophix "
            "Analyzer pane. UI tree may have changed in this Prophix build."
        )

    log.info("Expanding Refresh split-button dropdown")
    try:
        refresh_btn.expand()  # standard UIA ExpandCollapse pattern
    except Exception:
        # Split buttons sometimes only respond to a mouse click on the arrow.
        try:
            refresh_btn.click_input()
        except Exception as e:
            raise RuntimeError(
                f"Step 4: could not open the Refresh dropdown: {e}"
            ) from e

    time.sleep(0.5)

    # "All Sheets" is a MenuItem that appears on the desktop, not inside the
    # pane, once the split-button is expanded.
    all_sheets_selectors = [
        {"title": "All Sheets", "control_type": "MenuItem"},
        {"title": "All Sheets", "control_type": "ListItem"},
        {"title": "All Sheets"},
    ]
    target = None
    for kwargs in all_sheets_selectors:
        try:
            candidate = main_window.child_window(**kwargs)
            if candidate.exists(timeout=3):
                target = candidate
                break
        except Exception:
            continue
    if target is None:
        raise RuntimeError(
            "Step 4: Refresh dropdown opened but 'All Sheets' item not found."
        )

    try:
        target.invoke()
    except Exception:
        target.click_input()
    log.info("Clicked 'All Sheets' — refresh started")


def run(workbook_path: Path, **_ignored) -> None:
    """Fully-automated Prophix refresh. Raises on any automation failure."""
    log = get_logger()
    log.info("=== STEP 4: Prophix Analyzer refresh (automated) ===")

    # Import lazily so the rest of the agent can import on machines that
    # don't have pywinauto installed yet.
    from pywinauto import Application as UIApp

    with excel_com.excel_session(visible=True) as app:
        with excel_com.open_workbook(app, workbook_path,
                                     read_only=False, save_on_close=True):
            wb = app.ActiveWorkbook
            wb.Worksheets("Units Bookings DV").Activate()
            log.info("Activated tab: Units Bookings DV")

            _bring_excel_to_front(app, log)

            ui = UIApp(backend="uia").connect(class_name="XLMAIN", timeout=15)
            main = ui.top_window()

            _open_analyzer_pane(app, main, log)
            _bring_excel_to_front(app, log)
            _click_refresh_all_sheets(main, log)

            log.info("Sleeping %d seconds for Prophix refresh to complete",
                     config.PROPHIX_WAIT_SECONDS)
            time.sleep(config.PROPHIX_WAIT_SECONDS)
            log.info("Sleep complete; saving workbook")
    log.info("Step 4 complete")
