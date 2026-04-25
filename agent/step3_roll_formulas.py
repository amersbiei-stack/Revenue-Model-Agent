"""Step 3 — Roll formulas month-over-month via the existing VBA macro.

Includes a date-math pre-flight guard so that if the filename's date doesn't
imply the expected close month, we stop before populating the wrong column.
"""
from datetime import date
from pathlib import Path

from agent import config, excel_com, date_utils
from agent.logging_setup import get_logger


def date_math_guard(file_date: date,
                    expected_close_year: int,
                    expected_close_month: int) -> None:
    log = get_logger()
    derived = date_utils.derive_run_dates(file_date)
    log.info(
        "Date-math guard: filename=%s -> derived close=%s (%d-%02d). expected=%d-%02d",
        file_date, derived.new_month_label,
        derived.new_year, derived.new_month,
        expected_close_year, expected_close_month,
    )
    if derived.new_month != expected_close_month or derived.new_year != expected_close_year:
        raise ValueError(
            f"Date-math guard FAILED: filename {file_date.isoformat()} implies "
            f"close month {derived.new_year}-{derived.new_month:02d}, but you "
            f"specified close {expected_close_year}-{expected_close_month:02d}. "
            "Stopping — running would roll the wrong months."
        )


def run(workbook_path: Path,
        expected_close_year: int,
        expected_close_month: int) -> None:
    log = get_logger()
    log.info("=== STEP 3: Roll formulas month-over-month ===")
    file_date = date_utils.parse_filename(workbook_path)
    date_math_guard(file_date, expected_close_year, expected_close_month)

    # Strip Mark of the Web before opening. Office's 'Block macros from
    # the Internet' policy disables VBA in any .xlsm carrying MOTW — even
    # when AutomationSecurity=Low — and OneDrive-synced files sometimes
    # acquire MOTW after sync round-trips.
    excel_com.unblock_file(workbook_path, log)

    with excel_com.excel_session(visible=True) as app:
        with excel_com.open_workbook(app, workbook_path,
                                     read_only=False, save_on_close=True):
            excel_com.run_macro(app, config.MACRO_STEP2_ROLL_FORMULAS)
    log.info("Step 3 complete (macro reports success via TimedPopup)")
