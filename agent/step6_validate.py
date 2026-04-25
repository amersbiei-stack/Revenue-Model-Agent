"""Step 6 — Q&A validation tie-outs: Summary By Region vs source data tabs.

Classifies each check into PASS / WARN / BLOCK per the uniform $1,000 threshold.
On any BLOCK-tier failure: drafts variance-alert email and halts.
On WARN-tier only: returns results so the exec email can show the banner.
"""
from dataclasses import dataclass
from enum import Enum
from pathlib import Path

from agent import config, excel_com, outlook_draft, date_utils
from agent.logging_setup import get_logger


class Tier(Enum):
    PASS = "PASS"
    WARN = "WARN"
    BLOCK = "BLOCK"


@dataclass
class Check:
    label: str
    summary_row: int
    source_start: int
    source_end: int
    is_total: bool
    summary_value: float
    source_sum: float
    variance: float      # summary - source, signed
    tier: Tier


def _classify(variance: float, is_total: bool) -> Tier:
    tol = config.TOLERANCE_TOTAL if is_total else config.TOLERANCE_SUBREGIONAL
    a = abs(variance)
    if a <= tol:
        return Tier.PASS
    if a <= config.BLOCK_LIMIT:
        return Tier.WARN
    return Tier.BLOCK


def _run_dollar_checks(sr_ws, bk_ws, summary_col: int, bookings_col: int) -> list[Check]:
    log = get_logger()
    out: list[Check] = []
    for label, s_row, b_start, b_end, is_total in config.STEP6_DOLLAR_CHECKS:
        s_val = excel_com.read_cell(sr_ws, s_row, summary_col)
        s_val = float(s_val) if isinstance(s_val, (int, float)) else 0.0
        b_sum = excel_com.sum_column_range(bk_ws, b_start, b_end, bookings_col)
        variance = s_val - b_sum
        tier = _classify(variance, is_total)
        log.info("[%s] %s: summary=%s, source=%s, variance=%s",
                 tier.value, label,
                 f"${s_val:,.2f}", f"${b_sum:,.2f}", f"${variance:+,.2f}")
        out.append(Check(label, s_row, b_start, b_end, is_total,
                         s_val, b_sum, variance, tier))
    return out


def _run_units_checks(sr_ws, source_tabs: dict, summary_col: int,
                      source_cols: dict) -> list[Check]:
    """One Check per entry in STEP6_UNITS_CHECKS. Each entry can hit a
    different source tab (e.g. Units Bookings DV vs Migration Units).
    """
    log = get_logger()
    out: list[Check] = []
    for label, s_row, tab_name, u_start, u_end in config.STEP6_UNITS_CHECKS:
        ws = source_tabs[tab_name]
        col = source_cols[tab_name]
        s_val = excel_com.read_cell(sr_ws, s_row, summary_col)
        s_val = float(s_val) if isinstance(s_val, (int, float)) else 0.0
        u_sum = excel_com.sum_column_range(ws, u_start, u_end, col)
        variance = s_val - u_sum
        tier = _classify(variance, is_total=True)
        log.info("[%s] %s: summary=%s, source=%s ('%s' rows %d-%d), variance=%s",
                 tier.value, label,
                 f"{s_val:,.2f}", f"{u_sum:,.2f}",
                 tab_name, u_start, u_end,
                 f"{variance:+,.2f}")
        out.append(Check(label, s_row, u_start, u_end, True,
                         s_val, u_sum, variance, tier))
    return out


def _apply_latam_offset_rule(checks: list[Check], log) -> None:
    """If LATAM PX and LATAM AS variances offset each other (sum to within
    the total tolerance), the LATAM dollar total is correct — money is just
    misclassified between the two subregions. The downstream LATAM total
    check already passes, so the subregion-level BLOCKs are a false alarm.
    Downgrade both subregion Checks to PASS in that case.
    """
    px = next((c for c in checks if c.label == "LATAM PX"), None)
    ax = next((c for c in checks if c.label == "LATAM AS"), None)
    if px is None or ax is None:
        return
    combined = px.variance + ax.variance
    if abs(combined) <= config.TOLERANCE_TOTAL:
        log.info(
            "[OFFSET] LATAM PX (%s) + LATAM AS (%s) = %s — variances offset "
            "within $%.2f, downgrading both subregion checks to PASS.",
            f"${px.variance:+,.2f}", f"${ax.variance:+,.2f}",
            f"${combined:+,.2f}", config.TOLERANCE_TOTAL,
        )
        px.tier = Tier.PASS
        ax.tier = Tier.PASS


def _variance_email_html(new_month_label: str,
                         month_name: str, year: int,
                         blockers: list[Check], warnings: list[Check],
                         file_path: Path) -> str:
    def line(c: Check) -> str:
        return (f"<li><b>{c.label}</b>: Variance of ${c.variance:+,.2f} "
                f"between Summary By Region Row {c.summary_row} and source rows "
                f"{c.source_start}:{c.source_end}</li>")

    lines = [
        f"<p>Executive summary for {month_name} {year} close was "
        "<b>NOT sent</b> due to material variance(s) in validation.</p>",
        f"<p><b>BLOCKING variances (exceeded ${config.BLOCK_LIMIT:,.0f} — "
        "manual review required):</b></p>",
        "<ul>", *[line(c) for c in blockers], "</ul>",
    ]
    if warnings:
        lines += [
            "<p><b>Warning-tier variances (informational):</b></p>",
            "<ul>", *[line(c) for c in warnings], "</ul>",
        ]
    lines += [
        f"<p>File: <code>{file_path}</code></p>",
        "<p><b>Next steps:</b></p>",
        "<ol>",
        "<li>Review the Summary By Region tab and the source tabs for the listed rows in the new month column.</li>",
        "<li>If the issue is a failed/partial Prophix refresh, re-run Step 4 (Prophix refresh) and Step 5 (data verification).</li>",
        "<li>If the issue is a formula/FX drift, correct the source and re-run Step 6 only.</li>",
        "<li>Once checks pass (or are below the blocking threshold), the executive summary email can be sent.</li>",
        "</ol>",
        "<p>The workflow is halted. Do not re-run from Step 1 unless the underlying file is corrupted — the archive backup exists if needed.</p>",
    ]
    return "\n".join(lines)


def run(workbook_path: Path, run_dates: date_utils.RunDates) -> list[Check]:
    log = get_logger()
    log.info("=== STEP 6: Q&A validation tie-outs (label=%s) ===",
             run_dates.new_month_label)

    # Build the set of source tabs we'll need to open. $$ Bookings DV is
    # always required for dollar checks; the units checks may hit Units
    # Bookings DV and/or Migration Units depending on STEP6_UNITS_CHECKS.
    units_source_tab_names = {entry[2] for entry in config.STEP6_UNITS_CHECKS}

    with excel_com.excel_session(visible=True) as app:
        with excel_com.open_workbook(app, workbook_path,
                                     read_only=True, save_on_close=False) as wb:
            sr = wb.Worksheets("Summary By Region")
            bk = wb.Worksheets("$$ Bookings DV")

            # Open every source tab the units checks reference, plus find
            # the new-month column on each.
            source_tabs: dict = {}
            source_cols: dict[str, int] = {}
            for tab_name in units_source_tab_names:
                ws = wb.Worksheets(tab_name)
                col = excel_com.find_column_by_label(
                    ws, 2, run_dates.new_month_label)
                if col == 0:
                    raise RuntimeError(
                        f"Could not find new-month column "
                        f"{run_dates.new_month_label!r} on tab {tab_name!r}"
                    )
                source_tabs[tab_name] = ws
                source_cols[tab_name] = col

            summary_col = excel_com.find_column_by_label(
                sr, config.SUMMARY_HEADER_ROW, run_dates.new_month_label)
            bookings_col = excel_com.find_column_by_label(
                bk, 2, run_dates.new_month_label)
            log.info("Columns: summary=%d, bookings=%d, units_tabs=%s",
                     summary_col, bookings_col,
                     {n: source_cols[n] for n in units_source_tab_names})
            if 0 in (summary_col, bookings_col):
                raise RuntimeError(
                    f"Could not find new-month column {run_dates.new_month_label!r} "
                    f"(summary={summary_col}, bookings={bookings_col})"
                )

            dollar_checks = _run_dollar_checks(sr, bk, summary_col, bookings_col)
            units_checks = _run_units_checks(sr, source_tabs, summary_col,
                                             source_cols)

    all_checks = dollar_checks + units_checks

    # Apply LATAM PX/AS offset rule before tier filtering. If the two
    # subregion variances cancel, the dollar total is fine and we skip
    # the BLOCK email.
    _apply_latam_offset_rule(all_checks, log)

    blockers = [c for c in all_checks if c.tier is Tier.BLOCK]
    warnings = [c for c in all_checks if c.tier is Tier.WARN]

    if blockers:
        month_name = date_utils.month_name(run_dates.new_month)
        subject = config.EMAIL_SUBJECT_VARIANCE.format(
            month=month_name, year=run_dates.new_year)
        html = _variance_email_html(
            run_dates.new_month_label, month_name, run_dates.new_year,
            blockers, warnings, workbook_path)
        entry_id = outlook_draft.create_draft(
            to=config.EMAIL_RECIPIENT, subject=subject, html_body=html)
        log.error("Step 6 BLOCK tier — variance alert drafted (EntryID=%s)", entry_id)
        raise RuntimeError(
            f"Step 6 halted: {len(blockers)} BLOCK-tier variances. Alert drafted."
        )

    log.info("Step 6 complete — %d PASS, %d WARN, %d BLOCK",
             sum(1 for c in all_checks if c.tier is Tier.PASS),
             len(warnings), len(blockers))
    return all_checks
