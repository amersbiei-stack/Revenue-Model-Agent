"""Step 4 — Prophix Analyzer refresh.

The Prophix Analyzer CV2 add-in does not expose a programmatic completion signal.
We open the workbook, navigate to `Units Bookings DV`, prompt the human to perform
the manual Alt+N+Y1 → All Sheets refresh, then sleep 10 minutes per spec.

If the user wants fully unattended operation later, the manual prompt block can
be removed and replaced with a SendKeys/UI-Automation sequence — but those are
fragile under Task Scheduler (no keyboard focus). Manual confirmation is safer
for now.
"""
import time
from pathlib import Path

from agent import config, excel_com
from agent.logging_setup import get_logger


def run(workbook_path: Path, *, manual_confirm: bool = True) -> None:
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

            if manual_confirm:
                msg = (
                    "\n" + "=" * 70 + "\n"
                    "STEP 4 — MANUAL ACTION REQUIRED\n"
                    "  1. Excel should be open with 'Units Bookings DV' active.\n"
                    "  2. Press Alt + N + Y1 to open Prophix Analyzer.\n"
                    "  3. In the bottom toolbar, click the arrow next to the\n"
                    "     refresh button and choose 'All Sheets'.\n"
                    "  4. Once the refresh has STARTED, press Enter here.\n"
                    "  5. The agent will then sleep 10 minutes.\n"
                    + "=" * 70 + "\n"
                )
                print(msg)
                input("Press Enter once Prophix refresh has started... ")

            log.info("Sleeping %d seconds for Prophix refresh", config.PROPHIX_WAIT_SECONDS)
            time.sleep(config.PROPHIX_WAIT_SECONDS)
            log.info("Sleep complete; saving workbook")
    log.info("Step 4 complete")
