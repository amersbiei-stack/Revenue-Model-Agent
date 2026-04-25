"""Step 4 — Prophix Analyzer refresh (manual confirmation).

Prophix Analyzer CV2 is an Office Web Add-in (Office.js task pane) hosted at
https://cmcf.ca.prophix.cloud. It runs in a sandboxed WebView and exposes
no programmatic API to VBA, COM, or external automation — confirmed by the
manifest's `xsi:type="TaskPaneApp"` and `<Action xsi:type="ShowTaskpane">`.
The only way to trigger its 'Refresh > All Sheets' action is a real user
click in the task pane.

This step:
  1. Opens the workbook and activates `Units Bookings DV`.
  2. Prints clear instructions and waits for the user to:
       - press Alt+N+Y+1 to open the Analyzer pane,
       - click Refresh ▾ -> All Sheets in the pane,
       - press Enter on the console.
  3. Sleeps PROPHIX_WAIT_SECONDS (default 600) for the refresh to finish.
  4. Saves and closes.

Step 5 is the safety net: if the user pressed Enter prematurely or Prophix
didn't actually populate the new month, Step 5's per-tab sum checks fail and
draft a Step 5 failure email instead of letting bad data flow downstream.
"""
import time
from pathlib import Path

from agent import config, excel_com
from agent.logging_setup import get_logger


def run(workbook_path: Path) -> None:
    log = get_logger()
    log.info("=== STEP 4: Prophix Analyzer refresh ===")

    with excel_com.excel_session(visible=True) as app:
        with excel_com.open_workbook(app, workbook_path,
                                     read_only=False, save_on_close=True):
            wb = app.ActiveWorkbook
            try:
                wb.Worksheets("Units Bookings DV").Activate()
                log.info("Activated tab: Units Bookings DV")
            except Exception as e:
                log.warning("Could not activate Units Bookings DV: %s", e)

            msg = (
                "\n" + "=" * 70 + "\n"
                "STEP 4 — MANUAL ACTION REQUIRED\n"
                "  1. Excel is open with 'Units Bookings DV' active.\n"
                "  2. Press Alt + N + Y + 1 to open the Prophix Analyzer pane.\n"
                "  3. In the pane's bottom toolbar, click the arrow next to\n"
                "     the Refresh button and choose 'All Sheets'.\n"
                "  4. Once the refresh has STARTED, press Enter here.\n"
                "  5. The agent will then sleep 10 minutes for the refresh to\n"
                "     complete, then continue automatically through Steps 5-7.\n"
                + "=" * 70 + "\n"
            )
            print(msg)
            input("Press Enter once Prophix refresh has started... ")

            log.info("Sleeping %d seconds for Prophix refresh",
                     config.PROPHIX_WAIT_SECONDS)
            time.sleep(config.PROPHIX_WAIT_SECONDS)
            log.info("Sleep complete; saving workbook")
    log.info("Step 4 complete")
