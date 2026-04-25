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


def _ensure_excel_maximized(app, log) -> None:
    """Maximize Excel + expand the ribbon so every ribbon button is in the
    UIA tree. Narrow windows collapse the Prophix group into an overflow
    menu and the Analyzer button becomes invisible to descendants search.
    """
    XL_MAXIMIZED = -4137
    try:
        app.WindowState = XL_MAXIMIZED
        log.info("Excel maximized via COM (WindowState=xlMaximized)")
    except Exception as e:
        log.warning("COM maximize failed: %s", e)
    try:
        hwnd = app.Hwnd
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(hwnd)
            log.info("Excel maximized via Win32")
    except Exception as e:
        log.warning("Win32 maximize failed: %s", e)
    # If the ribbon was minimized to a single row, un-minimize it so ribbon
    # buttons render at full size and appear in the UIA tree.
    try:
        ribbon = app.CommandBars("Ribbon")
        h = getattr(ribbon, "Height", 0)
        log.info("Ribbon height: %s (minimized if <80)", h)
        if h and h < 80:
            log.info("Ribbon appears minimized; toggling via ExecuteMso")
            app.CommandBars.ExecuteMso("MinimizeRibbon")
    except Exception as e:
        log.warning("Could not check/toggle ribbon minimize state: %s", e)
    time.sleep(0.8)


# (legacy _analyzer_pane removed — pywinauto 0.6.9 on Python 3.14 rejects
# title_re. See _wait_for_pane / _find_pane_descendant below for the
# title_re-free replacements.)


def _activate_insert_tab(main_window, log) -> bool:
    """Click the Insert ribbon tab via UIA so its buttons populate.

    We physically click_input() the tab rather than calling .select() —
    select() was returning success without actually switching tabs on
    some installs (UIA tree never showed any Insert-tab buttons).
    """
    for kwargs in (
        {"title": "Insert", "control_type": "TabItem"},
        {"title": "Insert"},
    ):
        try:
            tab = main_window.child_window(**kwargs)
            if not tab.exists(timeout=2):
                continue
            try:
                tab.click_input()
            except Exception:
                try:
                    tab.select()
                except Exception as e:
                    log.warning("Both click_input and select failed: %s", e)
                    continue
            time.sleep(1.5)  # let the Insert tab render its buttons
            log.info("Insert tab clicked via UIA (%s)", kwargs)
            return True
        except Exception:
            continue
    log.warning("Could not click Insert tab via UIA")
    return False


def _find_prophix_analyzer_button(main_window, log):
    """Enumerate all descendants and filter for a visible 'Analyzer' button.

    Avoids pywinauto's title_re kwarg (unsupported on 0.6.9). We enumerate
    all descendants, keep anything whose display text contains 'Analyzer'
    and has a non-zero rectangle, then pick the leftmost — the Prophix
    CV2 button (which sits to the left of Contributor and the duplicate
    second Analyzer in the ribbon Prophix group).
    """
    # Pull every descendant once; cheaper than multiple searches.
    try:
        all_descendants = main_window.descendants()
    except Exception as e:
        log.warning("Could not enumerate descendants: %s", e)
        return None

    candidates = []
    for d in all_descendants:
        try:
            title = d.window_text() or ""
        except Exception:
            continue
        if "Analyzer" not in title:
            continue
        try:
            rect = d.rectangle()
            if rect.width() <= 0 or rect.height() <= 0:
                continue
        except Exception:
            continue
        try:
            ctype = d.element_info.control_type
        except Exception:
            ctype = "?"
        candidates.append((rect.left, d, title, ctype))

    log.info("Analyzer candidates (visible, by X):")
    candidates.sort(key=lambda t: t[0])
    for left, d, title, ctype in candidates[:10]:
        log.info("  left=%d title=%r ctype=%s", left, title, ctype)

    if not candidates:
        return None
    return candidates[0][1]


def _try_open_via_com_addin(app, log) -> bool:
    """Try to open the Prophix Analyzer pane by invoking the COM add-in.

    Prophix Analyzer CV2 registers itself in Application.COMAddIns. If
    its automation Object exposes a show-pane method, this bypasses the
    ribbon entirely (no UIA, no keyboard). Pure COM.
    """
    try:
        addins = app.COMAddIns
    except Exception as e:
        log.warning("Could not read COMAddIns: %s", e)
        return False

    try:
        count = addins.Count
    except Exception:
        count = 0
    log.info("COMAddIns available: %d", count)

    for i in range(1, count + 1):
        try:
            addin = addins.Item(i)
            progid = str(getattr(addin, "ProgID", "") or "")
            desc = str(getattr(addin, "Description", "") or "")
            connected = bool(getattr(addin, "Connect", False))
        except Exception as e:
            log.warning("  [%d] inspection failed: %s", i, e)
            continue
        log.info("  [%d] ProgID=%s | Desc=%s | Connected=%s",
                 i, progid, desc, connected)
        if "prophix" not in progid.lower() and "prophix" not in desc.lower():
            continue
        if "analyzer" not in progid.lower() and "analyzer" not in desc.lower():
            # Prophix Contributor is also a COM addin; skip non-analyzer ones.
            continue

        log.info("Found Prophix Analyzer addin: %s", progid)
        try:
            if not connected:
                addin.Connect = True
                time.sleep(1.0)
                log.info("Connected %s", progid)
            obj = addin.Object
        except Exception as e:
            log.warning("Could not connect/read Object on %s: %s", progid, e)
            continue
        if obj is None:
            log.info("%s exposes no automation Object", progid)
            continue

        for method_name in (
            "ShowPane", "ShowTaskPane", "OpenTaskPane",
            "OpenAnalyzer", "OpenPane", "Show", "Activate",
            "ToggleTaskPane", "ToggleAnalyzer",
        ):
            try:
                if hasattr(obj, method_name):
                    log.info("Invoking %s.%s()", progid, method_name)
                    getattr(obj, method_name)()
                    return True
            except Exception as e:
                log.warning("%s.%s() failed: %s", progid, method_name, e)
    return False


def _pane_is_open(main_window) -> bool:
    """Poll the UIA tree for any descendant whose title starts with
    'Prophix Analyzer'. Avoids title_re (broken on pywinauto 0.6.9 +
    Python 3.14)."""
    try:
        for d in main_window.descendants():
            try:
                title = d.window_text() or ""
            except Exception:
                continue
            if title.startswith("Prophix Analyzer"):
                return True
    except Exception:
        pass
    return False


def _wait_for_pane(main_window, log, timeout: float = 12.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _pane_is_open(main_window):
            return True
        time.sleep(0.5)
    return False


def _click_via_uia(main_window, log) -> bool:
    """Click the Analyzer ribbon button via UIA Invoke."""
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
            btn.invoke()
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

        if _wait_for_pane(main_window, log, timeout=12):
            log.info("Prophix Analyzer pane opened (UIA click)")
            return True
        log.warning("Pane didn't appear after UIA click (attempt %d/2)", attempt)

    return False


def _click_via_keyboard(main_window, log) -> bool:
    """Fallback path: Alt+N+Y+1 via OS-level keystrokes."""
    from pywinauto.keyboard import send_keys

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

        if _wait_for_pane(main_window, log, timeout=12):
            log.info("Prophix Analyzer pane opened (keyboard)")
            return True
        log.warning("Pane didn't appear after keystrokes (attempt %d/2)", attempt)
        try:
            send_keys("{ESC}{ESC}", pause=0.1)
        except Exception:
            pass
        time.sleep(1.0)
    return False


def _open_analyzer_pane(app, main_window, log) -> None:
    """Open the Prophix Analyzer pane.

    Order of attempts (most reliable first):
      1. Direct COM-add-in invocation (no UI at all).
      2. UIA click on the Insert > Analyzer ribbon button.
      3. Alt+N+Y+1 keystrokes.
    """
    if _pane_is_open(main_window):
        log.info("Prophix Analyzer pane already open")
        return

    log.info("Trying COM add-in invocation first (no UI)")
    if _try_open_via_com_addin(app, log):
        if _wait_for_pane(main_window, log, timeout=15):
            log.info("Prophix Analyzer pane opened (COM add-in)")
            return
        log.warning("COM add-in invoked but pane never appeared")

    log.info("Trying UIA click on Analyzer ribbon button")
    if _click_via_uia(main_window, log):
        return

    log.info("UIA click failed; falling back to Alt+N+Y+1 keystrokes")
    if _click_via_keyboard(main_window, log):
        return

    # All paths failed — dump diagnostics so the next run can be tuned.
    log.error("Could not open Prophix Analyzer pane by COM, UIA, or keyboard.")
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


def _find_pane_descendant(main_window, log):
    """Return a wrapper for whichever descendant has the Prophix pane title."""
    for d in main_window.descendants():
        try:
            title = d.window_text() or ""
        except Exception:
            continue
        if title.startswith("Prophix Analyzer"):
            return d
    return None


def _click_refresh_all_sheets(main_window, log) -> None:
    """Find the pane's Refresh split-button by walking descendants, then
    invoke it and click 'All Sheets'. Avoids title_re entirely."""
    pane = _find_pane_descendant(main_window, log)
    if pane is None:
        raise RuntimeError(
            "Step 4: Prophix Analyzer pane is not in the UIA tree."
        )

    refresh_btn = None
    for d in pane.descendants():
        try:
            title = (d.window_text() or "").strip()
        except Exception:
            continue
        if title == "Refresh" or title.startswith("Refresh"):
            try:
                if d.rectangle().width() > 0:
                    refresh_btn = d
                    log.info("Found Refresh control: title=%r ctype=%s",
                             title, d.element_info.control_type)
                    break
            except Exception:
                refresh_btn = d
                break
    if refresh_btn is None:
        raise RuntimeError(
            "Step 4: could not locate the 'Refresh' control in the Prophix "
            "Analyzer pane."
        )

    log.info("Expanding Refresh split-button dropdown")
    try:
        refresh_btn.expand()
    except Exception:
        try:
            refresh_btn.click_input()
        except Exception as e:
            raise RuntimeError(
                f"Step 4: could not open the Refresh dropdown: {e}"
            ) from e

    time.sleep(0.6)

    # "All Sheets" appears as a desktop-level menu item once the split-button
    # is expanded. Walk the desktop's descendants, not just main_window's.
    target = None
    try:
        from pywinauto import Desktop
        desktop = Desktop(backend="uia")
        for d in desktop.descendants():
            try:
                if (d.window_text() or "").strip() == "All Sheets":
                    if d.rectangle().width() > 0:
                        target = d
                        break
            except Exception:
                continue
    except Exception as e:
        log.warning("Could not enumerate Desktop for All Sheets: %s", e)

    if target is None:
        # Fallback: search main_window descendants directly.
        for d in main_window.descendants():
            try:
                if (d.window_text() or "").strip() == "All Sheets":
                    target = d
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
            _ensure_excel_maximized(app, log)

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
