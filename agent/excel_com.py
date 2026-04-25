"""Excel COM helpers. Open workbook, run macro, read ranges with label-check guards."""
import subprocess
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

import win32com.client as win32
from win32com.client import gencache

from agent.logging_setup import get_logger

# msoAutomationSecurity values
MSO_AUTO_SECURITY_LOW = 1     # auto-enable macros (required for our .xlsm)
MSO_AUTO_SECURITY_DISABLE = 3


def _make_excel_app(log):
    """Create Excel.Application via early binding (gencache.EnsureDispatch)
    so AutomationSecurity and other typed properties are reachable. Fall
    back to late-bound Dispatch if EnsureDispatch fails (Python 3.14 +
    bleeding-edge pywin32 sometimes can't generate the cache)."""
    try:
        app = gencache.EnsureDispatch("Excel.Application")
        log.debug("Excel.Application via gencache.EnsureDispatch (early binding)")
        return app
    except Exception as e:
        log.warning("EnsureDispatch failed (%s); falling back to Dispatch", e)
        return win32.Dispatch("Excel.Application")


def unblock_file(path: Path, log) -> None:
    """Remove the Zone.Identifier alternate data stream from a file.

    Office's 'Block macros from the Internet' policy disables VBA in any
    .xlsm carrying Mark of the Web — even when AutomationSecurity=Low.
    Stripping the ADS lets the macro run again. Idempotent: silently
    no-op if there's no MOTW to remove.
    """
    ads_path = f"{path}:Zone.Identifier"
    try:
        # `del` on the ADS path; using cmd.exe because del's ADS-aware.
        result = subprocess.run(
            ["cmd", "/c", "del", "/f", ads_path],
            capture_output=True, text=True, timeout=10,
        )
        if result.returncode == 0:
            log.info("Stripped Mark of the Web from %s", path.name)
        else:
            # rc=2 typically means the ADS didn't exist; not an error.
            log.debug("unblock-file: rc=%d stderr=%r", result.returncode,
                      result.stderr.strip())
    except Exception as e:
        log.warning("unblock-file failed (continuing anyway): %s", e)


def ensure_trusted_location(folder_path: Path, log) -> bool:
    """Register a folder under HKCU\\Software\\Microsoft\\Office\\<ver>\\
    Excel\\Security\\Trusted Locations so macros run without the per-file
    security gate. Idempotent: silently skips if the folder is already
    listed. Returns True if a Trusted Location entry exists for the
    folder when this function returns.
    """
    import os
    import winreg

    folder = os.path.normpath(str(folder_path))
    if not folder.endswith("\\"):
        folder_with_sep = folder + "\\"
    else:
        folder_with_sep = folder
    folder_norm = os.path.normcase(folder.rstrip("\\"))

    success_any_ver = False
    for ver in ("16.0", "15.0", "14.0"):
        base = rf"Software\Microsoft\Office\{ver}\Excel\Security\Trusted Locations"
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, base, 0,
                winreg.KEY_READ | winreg.KEY_WRITE,
            ) as parent:
                # Already listed?
                already = False
                i = 0
                while True:
                    try:
                        sub_name = winreg.EnumKey(parent, i)
                    except OSError:
                        break
                    try:
                        with winreg.OpenKey(parent, sub_name) as sk:
                            try:
                                p, _ = winreg.QueryValueEx(sk, "Path")
                            except FileNotFoundError:
                                p = ""
                        if (p and
                                os.path.normcase(str(p).rstrip("\\"))
                                == folder_norm):
                            already = True
                            break
                    except Exception:
                        pass
                    i += 1

                if already:
                    log.info("Trusted Location already present for Office %s: %s",
                             ver, folder)
                    success_any_ver = True
                    continue

                # Create our own entry. Use a stable name so we don't
                # multiply-register on subsequent runs.
                with winreg.CreateKey(parent, "RevenueModelAgent") as sk:
                    winreg.SetValueEx(sk, "Path", 0, winreg.REG_SZ,
                                      folder_with_sep)
                    winreg.SetValueEx(sk, "AllowSubFolders", 0,
                                      winreg.REG_DWORD, 1)
                    winreg.SetValueEx(sk, "Description", 0, winreg.REG_SZ,
                                      "Revenue Model Agent (auto-added)")
                    winreg.SetValueEx(sk, "Date", 0, winreg.REG_SZ,
                                      "Auto-added by Revenue Model Agent")
                log.info("Registered Trusted Location for Office %s: %s",
                         ver, folder)
                success_any_ver = True
        except FileNotFoundError:
            # That Office version isn't installed; try the next.
            continue
        except PermissionError as e:
            log.warning(
                "Cannot write Trusted Locations for Office %s "
                "(permission denied): %s", ver, e,
            )
        except Exception as e:
            log.warning("Trusted Locations registration failed for Office %s: %s",
                        ver, e)

    if not success_any_ver:
        log.warning(
            "Could not register a Trusted Location in any Office version. "
            "If macros are blocked, add the folder manually: Excel → File → "
            "Options → Trust Center → Trust Center Settings → Trusted Locations."
        )
    return success_any_ver


def clear_disabled_item(workbook_path: Path, log) -> None:
    """Remove the workbook from Excel's per-file 'Disabled Items' cache.

    After a crash or failed automation Excel sometimes adds a file to
    HKCU\\Software\\Microsoft\\Office\\<ver>\\Excel\\Resiliency\\
    DisabledItems, which disables macros for THAT FILE regardless of
    Trust Center / AutomationSecurity. Best-effort cleanup.
    """
    import winreg

    target = str(workbook_path).lower()
    cleared = 0
    for ver in ("16.0", "15.0", "14.0"):
        base = (rf"Software\Microsoft\Office\{ver}\Excel\Resiliency"
                r"\DisabledItems")
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, base, 0,
                winreg.KEY_READ | winreg.KEY_WRITE,
            ) as k:
                values_to_delete = []
                i = 0
                while True:
                    try:
                        name, data, _kind = winreg.EnumValue(k, i)
                    except OSError:
                        break
                    try:
                        as_text = bytes(data).decode(
                            "utf-16-le", errors="ignore",
                        ).lower() if isinstance(data, (bytes, bytearray)) else str(data).lower()
                    except Exception:
                        as_text = ""
                    if target in as_text or workbook_path.name.lower() in as_text:
                        values_to_delete.append(name)
                    i += 1
                for name in values_to_delete:
                    try:
                        winreg.DeleteValue(k, name)
                        cleared += 1
                    except Exception as e:
                        log.warning("Could not delete DisabledItems entry %s: %s",
                                    name, e)
        except FileNotFoundError:
            continue
        except Exception as e:
            log.warning("DisabledItems lookup failed for Office %s: %s", ver, e)
    if cleared:
        log.info("Cleared %d DisabledItems entry(ies) for this workbook",
                 cleared)
    else:
        log.debug("No DisabledItems entries found for this workbook")


@contextmanager
def excel_session(visible: bool = True) -> Iterator:
    """Context-managed Excel.Application. Auto-enables macros, suppresses alerts."""
    log = get_logger()
    app = _make_excel_app(log)
    # AutomationSecurity is not always exposed by late-bound dynamic dispatch
    # (seen on Python 3.14 + pywin32). If we can't touch it, Trust Center must
    # be configured instead (see SETUP.md). Don't let this block the run.
    prev_security = None
    try:
        prev_security = app.AutomationSecurity
        app.AutomationSecurity = MSO_AUTO_SECURITY_LOW
        sec_note = "AutomationSecurity=Low"
    except AttributeError:
        sec_note = "AutomationSecurity unavailable (relying on Trust Center)"
    prev_alerts = app.DisplayAlerts
    prev_screen = app.ScreenUpdating
    app.Visible = visible
    app.DisplayAlerts = False
    log.debug("Excel session opened (visible=%s, %s)", visible, sec_note)
    try:
        yield app
    finally:
        try:
            app.ScreenUpdating = prev_screen
            app.DisplayAlerts = prev_alerts
            if prev_security is not None:
                try:
                    app.AutomationSecurity = prev_security
                except AttributeError:
                    pass
            app.Quit()
            log.debug("Excel session closed")
        except Exception as e:
            log.warning("Error during Excel cleanup: %s", e)


@contextmanager
def open_workbook(app, path: Path, read_only: bool = False, save_on_close: bool = False) -> Iterator:
    """Open a workbook, yield it, close with the chosen save behavior."""
    log = get_logger()
    log.debug("Opening workbook: %s (read_only=%s)", path, read_only)
    wb = app.Workbooks.Open(str(path), ReadOnly=read_only, UpdateLinks=0)
    try:
        yield wb
    finally:
        wb.Close(SaveChanges=save_on_close)
        log.debug("Workbook closed (save_on_close=%s)", save_on_close)


def run_macro(app, macro_name: str, *args) -> None:
    """Invoke a workbook-bound VBA macro by name."""
    log = get_logger()
    log.info("Running macro: %s", macro_name)
    app.Run(macro_name, *args)
    log.info("Macro returned: %s", macro_name)


def find_column_by_label(ws, header_row: int, label: str, max_search_cols: int = 500) -> int:
    """Return 1-based column where ws.Cells(header_row, c).Value == label. 0 if not found."""
    for c in range(1, max_search_cols + 1):
        v = ws.Cells(header_row, c).Value
        if v is not None and str(v).strip() == label:
            return c
    return 0


def assert_label(ws, row: int, col: int, expected: str) -> None:
    """Raise AssertionError if cell text doesn't match expected. Used for label-check guards."""
    actual = ws.Cells(row, col).Value
    actual_s = "" if actual is None else str(actual).strip()
    if actual_s != expected:
        raise AssertionError(
            f"Label-check failed at {ws.Name}!R{row}C{col}: "
            f"expected {expected!r}, got {actual_s!r}"
        )


def read_cell(ws, row: int, col: int):
    """Read a single cell value (raw, no coercion)."""
    return ws.Cells(row, col).Value


def sum_column_range(ws, start_row: int, end_row: int, col: int) -> float:
    """SUM of numeric cells in [start_row..end_row, col]. One COM call."""
    rng = ws.Range(ws.Cells(start_row, col), ws.Cells(end_row, col))
    values = rng.Value
    total = 0.0
    if values is None:
        return 0.0
    if isinstance(values, tuple):
        # Multi-row, single-column: tuple of 1-tuples
        for row in values:
            v = row[0] if isinstance(row, tuple) else row
            if isinstance(v, (int, float)):
                total += float(v)
    elif isinstance(values, (int, float)):
        # Single cell
        total = float(values)
    return total
