"""Step 4 — Prophix Analyzer refresh (fully automated).

Drives the Prophix Analyzer CV2 pane end-to-end, no human interaction:

  1. Open the workbook and activate `Units Bookings DV`.
  2. Click the Insert > Prophix > Analyzer ribbon button via UI Automation
     (falls back to Alt+N+Y+1 keystrokes if the UIA click can't find the
     button — some Prophix builds expose the control under different names).
  3. Use UI Automation to find the pane's refresh split-button, expand its
     dropdown, and click "All Sheets".
  4. Sleep PROPHIX_WAIT_SECONDS (default 600) for the refresh to complete.

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


def _activate_insert_tab(main_window, log) -> bool:
    """Select the Insert ribbon tab via UIA so its buttons are reachable."""
    for kwargs in (
        {"title": "Insert", "control_type": "TabItem"},
        {"title_re": "Insert.*", "control_type": "TabItem"},
        {"title": "Insert", "control_type": "Button"},
    ):
        try:
            tab = main_window.child_window(**kwargs)
            if not tab.exists(timeout=2):
                continue
            try:
                tab.select()
            except Exception:
                tab.click_input()
            time.sleep(0.6)
            log.info("Insert tab activated via UIA (%s)", kwargs)
            return True
        except Exception:
            continue
    log.warning("Could not activate Insert tab via UIA")
    return False


def _find_prophix_analyzer_button(main_window, log):
    """Find the leftmost 'Analyzer' ribbon button (Prophix Analyzer CV2).

    The Prophix group on the Insert tab has Analyzer / Contributor /
    Analyzer buttons; the CV2 one is the leftmost Analyzer. pywinauto's
    descendants traversal is left-to-right, so the first match is the
    right target on every Prophix build we've seen.
    """
    for kwargs in (
        {"title": "Analyzer", "control_type": "Button"},
        {"title_re": r"^Analyzer$", "control_type": "Button"},
        {"title_re": r".*Analyzer.*", "control_type": "Button"},
    ):
        try:
            matches = main_window.descendants(**kwargs)
        except Exception as e:
            log.warning("Descendant search failed for %s: %s", kwargs, e)
            continue
        if matches:
            log.info("Found %d 'Analyzer' ribbon button(s) via %s",
                     len(matches), kwargs)
            return matches[0]
    return None


def _click_via_uia(main_window, log) -> bool:
    """Primary path: click the Analyzer ribbon button via UIA Invoke.

    Returns True if the Prophix Analyzer pane appears within 12s.
    """
    from pywinauto.timings import TimeoutError as PWATimeoutError

    pane = main_window.child_window(
        title_re="Prophix Analyzer.*", control_type="Pane",
    )

    for attempt in range(1, 3):
        try:
            main_window.set_focus()
        except Exception as e:
            log.warning("set_focus failed (UIA attempt %d): %s", attempt, e)
        time.sleep(0.4)

        _activate_insert_tab(main_window, log)

        btn = _find_prophix_analyzer_button(main_window, log)
        if btn is None:
            log.warning("Analyzer ribbon button not found (attempt %d/2)", attempt)
            continue

        log.info("Clicking Analyzer ribbon button via UIA (attempt %d/2)", attempt)
        clicked = False
        try:
            btn.invoke()  # UIA Invoke pattern — a real click without mouse
            clicked = True
        except Exception as e:
            log.warning("btn.invoke() failed: %s; trying click_input", e)
            try:
                btn.click_input()
                clicked = True
            except Exception as e2:
                log.warning("btn.click_input() failed: %s", e2)

        if not clicked:
            continue

        try:
            pane.wait("exists visible", timeout=12)
            log.info("Prophix Analyzer pane opened (UIA click)")
            return True
        except PWATimeoutError:
            log.warning("Pane didn't appear after UIA click (attempt %d/2)",
                        attempt)

    return False


def _click_via_keyboard(main_window, log) -> bool:
    """Fallback path: Alt+N+Y+1 via OS-level keystrokes."""
    from pywinauto.keyboard import send_keys
    from pywinauto.timings import TimeoutError as PWATimeoutError

    pane = main_window.child_window(
        title_re="Prophix Analyzer.*", control_type="Pane",
    )

    for attempt in range(1, 3):
        try:
            main_window.set_focus()
        except Exception as e:
            log.warning("set_focus failed (kbd attempt %d): %s", attempt, e)
        time.sleep(0.8)

        log.info("Sending Alt+N, then Y, then 1 (kbd attempt %d/2)", attempt)
        send_keys("{VK_MENU down}n{VK_MENU up}", pause=0.15)
        time.sleep(1.5)
        send_keys("y", pause=0.2)
        time.sleep(0.35)
        send_keys("1", pause=0.2)

        try:
            pane.wait("exists visible", timeout=12)
            log.info("Prophix Analyzer pane opened (keyboard)")
            return True
        except PWATimeoutError:
            log.warning("Pane didn't appear after keystrokes (attempt %d/2)",
                        attempt)
            try:
                send_keys("{ESC}{ESC}", pause=0.1)
            except Exception:
                pass
            time.sleep(1.0)
    return False


def _open_analyzer_pane(app, main_window, log) -> None:
    """Open the Prophix Analyzer pane. UIA click first, keyboard fallback."""
    pane = main_window.child_window(
        title_re="Prophix Analyzer.*", control_type="Pane",
    )
    if pane.exists(timeout=1):
        log.info("Prophix Analyzer pane already open")
        return

    if _click_via_uia(main_window, log):
        return

    log.info("UIA click could not open the pane; falling back to Alt+N+Y+1")
    if _click_via_keyboard(main_window, log):
        return

    # Both paths failed — dump diagnostics so the next run can be tuned.
    log.error("Could not open Prophix Analyzer pane by UIA click or keyboard.")
    log.error("Top-level children of the Excel window:")
    try:
        for child in main_window.children()[:40]:
            try:
                log.error("  title=%r class=%s",
                          child.window_text(), child.friendly_class_name())
            except Exception:
                pass
    except Exception as e:
        log.error("  (could not enumerate children: %s)", e)
    log.error("Ribbon buttons visible (first 40):")
    try:
        buttons = main_window.descendants(control_type="Button")
        for b in buttons[:40]:
            try:
                log.error("  button title=%r", b.window_text())
            except Exception:
                pass
    except Exception as e:
        log.error("  (could not enumerate buttons: %s)", e)
    raise RuntimeError(
        "Step 4: Prophix Analyzer pane never opened (tried UIA click on the "
        "Insert > Analyzer ribbon button, then Alt+N+Y+1 keystrokes). Check "
        "the log for visible UI elements."
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
