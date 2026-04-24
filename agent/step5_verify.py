"""Step 5 — Verify Prophix refresh actually populated the new month.

For each of the 5 DV tabs, finds the new-month column by label lookup in the
header row and sums the data range. If any sum <= 0, halts and drafts a
Step 5 failure email; otherwise returns the per-tab sums for logging.
"""
from dataclasses import dataclass
from pathlib import Path

from agent import config, excel_com, outlook_draft, date_utils
from agent.logging_setup import get_logger


@dataclass
class TabCheck:
    tab: str
    header_row: int
    data_start: int
    data_end: int
    new_month_col: int  # 0 if not found
    sum_value: float
    passed: bool


def _check_all_tabs(wb, new_month_label: str) -> list[TabCheck]:
    log = get_logger()
    results: list[TabCheck] = []
    for tab, (header_row, start, end) in config.STEP5_TABS.items():
        ws = wb.Worksheets(tab)
        col = excel_com.find_column_by_label(ws, header_row, new_month_label)
        if col == 0:
            log.error("Tab %r: month label %r not found in row %d",
                      tab, new_month_label, header_row)
            results.append(TabCheck(tab, header_row, start, end, 0, 0.0, False))
            continue
        s = excel_com.sum_column_range(ws, start, end, col)
        ok = s > 0
        log.info("Tab %r: col=%d sum=%s passed=%s",
                 tab, col, f"{s:,.2f}", ok)
        results.append(TabCheck(tab, header_row, start, end, col, s, ok))
    return results


def _failure_email_html(new_month_label: str, failed: list[TabCheck]) -> str:
    lines = [
        f"<p>Step 5 failure — Prophix data did not populate for <b>{new_month_label}</b>.</p>",
        "<p><b>Failed tabs:</b></p>",
        "<ul>",
    ]
    for c in failed:
        if c.new_month_col == 0:
            lines.append(
                f"<li><b>{c.tab}</b>: month column {new_month_label!r} not found in "
                f"row {c.header_row}</li>"
            )
        else:
            lines.append(
                f"<li><b>{c.tab}</b>: SUM(rows {c.data_start}-{c.data_end}, "
                f"col {c.new_month_col}) = ${c.sum_value:,.2f}</li>"
            )
    lines += [
        "</ul>",
        "<p>Do not resume until the Prophix refresh has been re-run and all 5 tabs pass.</p>",
    ]
    return "\n".join(lines)


def run(workbook_path: Path, run_dates: date_utils.RunDates) -> list[TabCheck]:
    log = get_logger()
    log.info("=== STEP 5: Verify Prophix data populated (label=%s) ===",
             run_dates.new_month_label)

    with excel_com.excel_session(visible=True) as app:
        with excel_com.open_workbook(app, workbook_path,
                                     read_only=True, save_on_close=False) as wb:
            results = _check_all_tabs(wb, run_dates.new_month_label)

    failed = [r for r in results if not r.passed]
    if failed:
        month_name = date_utils.month_name(run_dates.new_month)
        subject = config.EMAIL_SUBJECT_STEP5_FAIL.format(
            month=month_name, year=run_dates.new_year)
        html = _failure_email_html(run_dates.new_month_label, failed)
        entry_id = outlook_draft.create_draft(
            to=config.EMAIL_RECIPIENT, subject=subject, html_body=html)
        log.error("Step 5 failed — draft email saved (EntryID=%s)", entry_id)
        raise RuntimeError(
            f"Step 5 failed: {len(failed)}/{len(results)} tabs did not populate. "
            "Alert email drafted in Outlook."
        )

    log.info("Step 5 complete — all %d tabs passed", len(results))
    return results
