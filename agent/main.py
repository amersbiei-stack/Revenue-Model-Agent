"""Top-level orchestrator for the Revenue Model monthly rollover agent.

Usage:
  py -m agent.main                            # auto-detect latest workbook, auto-derive close
  py -m agent.main --file "Revenue Model - 04212026 (Internal).xlsm"
  py -m agent.main --close-year 2026 --close-month 3   # override close
  py -m agent.main --skip-step1                # skip archive/rename (e.g. after a retry)
  py -m agent.main --only-step 7               # run a single step for debugging
"""
import argparse
import sys
from datetime import date
from pathlib import Path

from agent import config, date_utils
from agent import (
    step1_archive_rename,
    step3_roll_formulas,
    step4_prophix,
    step5_verify,
    step6_validate,
    step7_email,
)
from agent.excel_com import excel_session, open_workbook
from agent.logging_setup import setup_run_logger


def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--file", type=Path, default=None,
                   help="Path to live workbook (default: latest in ROOT)")
    p.add_argument("--close-year", type=int, default=None,
                   help="Expected close year (default: derived from today)")
    p.add_argument("--close-month", type=int, default=None,
                   help="Expected close month (default: derived from today)")
    p.add_argument("--skip-step1", action="store_true",
                   help="Skip archive+rename (file already in expected state)")
    p.add_argument("--skip-step3", action="store_true")
    p.add_argument("--skip-step4", action="store_true")
    p.add_argument("--only-step", type=int, default=None, choices=[1, 3, 4, 5, 6, 7],
                   help="Run a single step (for debugging)")
    return p.parse_args()


def _resolve_expected_close(close_year: int | None, close_month: int | None) -> tuple[int, int]:
    """If not provided, expected close = today - 1 month."""
    today = date.today()
    derived = date_utils.derive_run_dates(today)
    y = close_year if close_year else derived.new_year
    m = close_month if close_month else derived.new_month
    return y, m


def main() -> int:
    args = _parse_args()
    log = setup_run_logger()
    log.info("Revenue Model Rollover Agent — starting")

    # Resolve input workbook
    wb_path = args.file or date_utils.find_latest_workbook(config.ROOT)
    log.info("Input workbook: %s", wb_path)
    file_date = date_utils.parse_filename(wb_path)
    run_dates = date_utils.derive_run_dates(file_date)
    log.info("Derived: file_date=%s new_month_label=%s src_month_label=%s",
             file_date, run_dates.new_month_label, run_dates.src_month_label)

    close_y, close_m = _resolve_expected_close(args.close_year, args.close_month)
    log.info("Expected close: %d-%02d (%s %d)",
             close_y, close_m, date_utils.month_name(close_m), close_y)

    only = args.only_step

    # === STEP 1 ===
    if only in (None, 1) and not args.skip_step1:
        wb_path = step1_archive_rename.run(wb_path)
        log.info("Post-step1 workbook path: %s", wb_path)
        # Re-derive run_dates from the renamed file
        file_date = date_utils.parse_filename(wb_path)
        run_dates = date_utils.derive_run_dates(file_date)
        log.info("Re-derived after rename: new_month_label=%s", run_dates.new_month_label)
    if only == 1:
        return 0

    # === STEP 3 ===
    if only in (None, 3) and not args.skip_step3:
        step3_roll_formulas.run(wb_path, close_y, close_m)
    if only == 3:
        return 0

    # === STEP 4 ===
    if only in (None, 4) and not args.skip_step4:
        step4_prophix.run(wb_path, manual_confirm=True)
    if only == 4:
        return 0

    # === STEP 5 ===
    if only in (None, 5):
        step5_verify.run(wb_path, run_dates)
    if only == 5:
        return 0

    # === STEP 6 ===
    step6_checks: list = []
    if only in (None, 6):
        step6_checks = step6_validate.run(wb_path, run_dates)
    if only == 6:
        return 0

    # === STEP 7 ===
    if only in (None, 7):
        step7_email.run(wb_path, run_dates, step6_checks)
    if only == 7:
        return 0

    # === STEP 8 — save and close ===
    log.info("=== STEP 8: Save and close ===")
    with excel_session(visible=True) as app:
        with open_workbook(app, wb_path, read_only=False, save_on_close=True):
            pass
    log.info("Step 8 complete")

    log.info("Revenue Model Rollover Agent — done")
    return 0


if __name__ == "__main__":
    sys.exit(main())
