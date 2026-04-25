"""Step 4 — Prophix Analyzer refresh (fully automated).

Drives the Prophix Analyzer CV2 pane end-to-end, no human interaction:

  1. Launch EXCEL.EXE as a subprocess (not via COM Dispatch) so Prophix
     Analyzer CV2 loads — Dispatch starts Excel in automation/embedding
     mode, which skips many COM add-ins including Prophix.
  2. Attach to the workbook via COM and activate `Units Bookings DV`.
  3. Click the Insert > Prophix > Analyzer ribbon button via UI Automation
     (tries every visible 'Analyzer' button — the CV2 one's position
     in the group varies by install). Falls back to Alt+N+Y+1 if needed.
  4. Use UI Automation to find the pane's refresh split-button, expand
     its dropdown, and click "All Sheets".
  5. Sleep PROPHIX_WAIT_SECONDS (default 600) for the refresh to complete.

Step 5 is the safety net: if the refresh didn't actually run, its sum
checks on all 5 DV tabs will fail and a Step 5 failure email will draft.

Assumptions (documented in SETUP.md):
  - Prophix Analyzer CV2 is installed and signed in on this machine.
  - Excel Trust Center includes the project folder (no macro prompt).
"""
import os
import subprocess
import time
from pathlib import Path

import win32com.client
import win32con
import win32gui

from agent import config, excel_com
from agent.logging_setup import get_logger


def _find_excel_exe(log) -> str | None:
    """Locate EXCEL.EXE via App Paths registry key, then common install dirs."""
    try:
        import winreg
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
        ) as k:
            val, _ = winreg.QueryValueEx(k, None)
            if val and os.path.exists(val):
                log.info("Found EXCEL.EXE via registry: %s", val)
                return val
    except Exception as e:
        log.debug("registry lookup for excel.exe failed: %s", e)

    bases = [
        r"C:\Program Files\Microsoft Office",
        r"C:\Program Files (x86)\Microsoft Office",
    ]
    for base in bases:
        if not os.path.isdir(base):
            continue
        for root, _dirs, files in os.walk(base):
            if "EXCEL.EXE" in files:
                p = os.path.join(root, "EXCEL.EXE")
                log.info("Found EXCEL.EXE by filesystem walk: %s", p)
                return p
    return None


def _launch_excel_natively(workbook_path: Path, log):
    """Start EXCEL.EXE in a subprocess, wait for the workbook to become
    COM-accessible, and return (app, wb).

    This replaces win32.Dispatch("Excel.Application") for Step 4 because
    Dispatch puts Excel in embedding/automation mode, which skips loading
    Prophix Analyzer CV2 (and other COM add-ins). A subprocess launch
    behaves exactly like double-clicking the file.
    """
    excel_exe = _find_excel_exe(log)
    if not excel_exe:
        raise RuntimeError(
            "Step 4: could not locate EXCEL.EXE. Tried the App Paths "
            "registry key and common install dirs."
        )

    log.info("Launching Excel natively: %s  %s", excel_exe, workbook_path)
    # DETACHED_PROCESS so Excel keeps running even if the agent exits
    # unexpectedly mid-step (belt and suspenders — we still Save/Quit
    # in the finally block).
    creationflags = 0x00000008  # DETACHED_PROCESS
    subprocess.Popen(
        [excel_exe, str(workbook_path)],
        creationflags=creationflags,
        close_fds=True,
    )

    target_path_lower = str(workbook_path).lower()
    deadline = time.time() + 90
    wb = None
    app = None
    log.info("Waiting for workbook to register in the running Excel instance")
    while time.time() < deadline:
        try:
            # GetObject on a file path returns the Workbook COM object from
            # whichever Excel has it open.
            wb_candidate = win32com.client.GetObject(str(workbook_path))
            app = wb_candidate.Application
            # Scan Workbooks collection on that app to get the workbook
            # wrapper we'll use for the rest of the step.
            for i in range(1, app.Workbooks.Count + 1):
                try:
                    w = app.Workbooks.Item(i)
                    if str(w.FullName).lower() == target_path_lower:
                        wb = w
                        break
                except Exception:
                    continue
            if wb is not None:
                log.info("Excel attached and workbook open")
                break
        except Exception:
            pass
        time.sleep(1.5)

    if wb is None or app is None:
        raise RuntimeError(
            "Step 4: Excel subprocess started but the workbook did not "
            "become COM-accessible within 90s."
        )
    # Extra settle time so COM add-ins (Prophix) finish registering.
    time.sleep(5.0)
    return app, wb


def _shutdown_excel(app, wb, log) -> None:
    """Save the workbook, close it, and quit Excel. Best-effort."""
    try:
        if wb is not None:
            wb.Save()
            log.info("Workbook saved")
    except Exception as e:
        log.warning("Save failed: %s", e)
    try:
        if wb is not None:
            wb.Close(SaveChanges=False)
    except Exception as e:
        log.warning("Workbook close failed: %s", e)
    try:
        if app is not None:
            app.Quit()
            log.info("Excel quit")
    except Exception as e:
        log.warning("Excel quit failed: %s", e)





def _bring_excel_to_front(app, log) -> None:
    """Restore + foreground the Excel window. SendKeys won't land otherwise.

    app.Hwnd can raise 'application is busy' when Excel is mid-operation
    (common right after opening a task pane). Retry with backoff instead
    of bailing on the first transient failure.
    """
    hwnd = None
    for i in range(5):
        try:
            hwnd = app.Hwnd
            break
        except Exception as e:
            msg = str(e).lower()
            if "busy" in msg or "message filter" in msg:
                time.sleep(0.8 * (i + 1))
                continue
            log.warning("Could not read Application.Hwnd: %s", e)
            return
    if not hwnd:
        log.warning("Application.Hwnd unavailable after retries; skipping foreground")
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


def _com_retry(fn, log, retries: int = 5, initial_delay: float = 1.0):
    """Call a COM function, retrying if Excel returns 'application is busy'.

    Excel raises pywintypes.com_error -2147417846 (message filter) when the
    COM server is mid-operation — common right after triggering a task-pane
    open or a refresh.
    """
    delay = initial_delay
    last = None
    for i in range(retries):
        try:
            return fn()
        except Exception as e:
            last = e
            msg = str(e).lower()
            if ("message filter" in msg or
                    "busy" in msg or
                    "0x8001010a" in msg or
                    "rpc_e_serverfaulty" in msg):
                log.debug("COM retry %d/%d: busy; sleeping %.1fs", i + 1, retries, delay)
                time.sleep(delay)
                delay *= 1.5
                continue
            raise
    raise RuntimeError(f"COM call remained busy after {retries} retries: {last}")


def _wait_for_addins_to_load(app, log, wait_seconds: int = 15) -> None:
    """Pause for Excel-via-COM add-ins to finish loading, then log what's
    actually available. Force-connect any Prophix entry we spot."""
    log.info("Waiting %ds for Excel add-ins to initialize after workbook open",
             wait_seconds)
    time.sleep(wait_seconds)

    # COM add-ins
    try:
        com_addins = app.COMAddIns
        count = com_addins.Count
    except Exception as e:
        log.warning("Could not read COMAddIns: %s", e)
        count = 0
    log.info("COM add-ins currently registered: %d", count)
    for i in range(1, count + 1):
        try:
            a = com_addins.Item(i)
            progid = str(getattr(a, "ProgID", "") or "")
            desc = str(getattr(a, "Description", "") or "")
            connected = bool(getattr(a, "Connect", False))
            log.info("  COM[%d] %s | %s | Connect=%s",
                     i, progid, desc, connected)
            if ("prophix" in progid.lower() or "prophix" in desc.lower()) and not connected:
                log.info("  -> connecting %s", progid)
                try:
                    a.Connect = True
                    time.sleep(1.5)
                except Exception as e:
                    log.warning("  -> could not connect %s: %s", progid, e)
        except Exception as e:
            log.warning("  COM[%d] inspection failed: %s", i, e)

    # Native Excel add-ins (XLAM/XLL)
    try:
        xl_addins = app.AddIns
        xl_count = xl_addins.Count
    except Exception as e:
        log.warning("Could not read Application.AddIns: %s", e)
        xl_count = 0
    log.info("Excel (XLL/XLAM) add-ins registered: %d", xl_count)
    for i in range(1, xl_count + 1):
        try:
            a = xl_addins.Item(i)
            name = str(getattr(a, "Name", "") or "")
            installed = bool(getattr(a, "Installed", False))
            log.info("  XL[%d] %s | Installed=%s", i, name, installed)
            if "prophix" in name.lower() and not installed:
                log.info("  -> installing %s", name)
                try:
                    a.Installed = True
                    time.sleep(1.5)
                except Exception as e:
                    log.warning("  -> could not install %s: %s", name, e)
        except Exception as e:
            log.warning("  XL[%d] inspection failed: %s", i, e)


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


def _find_prophix_analyzer_buttons(main_window, log):
    """Return every visible 'Analyzer' descendant, sorted left→right.

    The Prophix group has Contributor + 2x Analyzer. Which Analyzer is
    the CV2 one varies by install (leftmost on one, rightmost on another).
    Caller should try each until the Prophix Analyzer pane actually opens.
    """
    try:
        all_descendants = main_window.descendants()
    except Exception as e:
        log.warning("Could not enumerate descendants: %s", e)
        return []

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

    candidates.sort(key=lambda t: t[0])
    log.info("Analyzer candidates (%d, visible, by X):", len(candidates))
    for left, _d, title, ctype in candidates[:10]:
        log.info("  left=%d title=%r ctype=%s", left, title, ctype)
    return [c[1] for c in candidates]


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
    """Click each visible 'Analyzer' ribbon button via UIA Invoke until the
    Prophix Analyzer pane opens.

    The Prophix group has Contributor + 2x Analyzer buttons. We don't know
    a priori which Analyzer is CV2 vs the other build, so we try each and
    the first one that opens the pane wins.
    """
    try:
        main_window.set_focus()
    except Exception as e:
        log.warning("set_focus failed: %s", e)
    time.sleep(0.4)

    _activate_insert_tab(main_window, log)

    buttons = _find_prophix_analyzer_buttons(main_window, log)
    if not buttons:
        log.warning("No visible 'Analyzer' ribbon buttons found")
        return False

    for idx, btn in enumerate(buttons, start=1):
        log.info("Clicking Analyzer candidate %d/%d via UIA", idx, len(buttons))
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

        if _wait_for_pane(main_window, log, timeout=10):
            log.info("Prophix Analyzer pane opened (UIA click, candidate %d)",
                     idx)
            return True
        log.info("Candidate %d didn't open the pane; trying next", idx)

    return False


def _click_via_keyboard(main_window, log) -> bool:
    """Fallback path: Alt+N+Y+1 via OS-level keystrokes.

    Each attempt first clicks the Insert ribbon tab — this nudges Prophix
    to lazy-load its ribbon controls when Excel was launched via COM
    (add-ins don't always auto-initialize the same way they do on a
    double-click launch).
    """
    from pywinauto.keyboard import send_keys

    for attempt in range(1, 5):
        # Clicking Insert is cheap and often the nudge Prophix needs.
        _activate_insert_tab(main_window, log)

        try:
            main_window.set_focus()
        except Exception as e:
            log.warning("set_focus failed (kbd attempt %d): %s", attempt, e)
        time.sleep(0.8)

        log.info("Sending Alt+N, then Y, then 1 (kbd attempt %d/4)", attempt)
        send_keys("{VK_MENU down}n{VK_MENU up}", pause=0.15)
        time.sleep(1.5)
        send_keys("y", pause=0.2)
        time.sleep(0.35)
        send_keys("1", pause=0.2)

        if _wait_for_pane(main_window, log, timeout=15):
            log.info("Prophix Analyzer pane opened (keyboard, attempt %d/4)",
                     attempt)
            return True
        log.warning("Pane didn't appear after keystrokes (attempt %d/4)", attempt)
        try:
            send_keys("{ESC}{ESC}", pause=0.1)
        except Exception:
            pass
        time.sleep(2.0)
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

    # Native subprocess launch (NOT win32.Dispatch) so Prophix loads.
    app, wb = _launch_excel_natively(workbook_path, log)
    try:
        try:
            wb.Worksheets("Units Bookings DV").Activate()
            log.info("Activated tab: Units Bookings DV")
        except Exception as e:
            log.warning("Could not activate Units Bookings DV: %s", e)

        _bring_excel_to_front(app, log)
        _ensure_excel_maximized(app, log)

        # Even with a native launch, give add-ins a moment to register
        # and log what's present for diagnostics.
        _wait_for_addins_to_load(app, log, wait_seconds=10)

        ui = UIApp(backend="uia").connect(class_name="XLMAIN", timeout=20)
        main = ui.top_window()

        _open_analyzer_pane(app, main, log)

        # Give the Prophix pane time to populate its UIA subtree before
        # we walk it looking for the Refresh split-button.
        time.sleep(4.0)

        _click_refresh_all_sheets(main, log)

        log.info("Sleeping %d seconds for Prophix refresh to complete",
                 config.PROPHIX_WAIT_SECONDS)
        time.sleep(config.PROPHIX_WAIT_SECONDS)
        log.info("Sleep complete; saving workbook")
    finally:
        _shutdown_excel(app, wb, log)
    log.info("Step 4 complete")
