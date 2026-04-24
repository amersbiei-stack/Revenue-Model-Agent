"""Step 7 — Executive summary email: region metrics, bullets, 3 charts, Outlook draft."""
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Iterable

from dateutil.relativedelta import relativedelta

from agent import config, excel_com, outlook_draft, charts, date_utils
from agent import step6_validate
from agent.logging_setup import get_logger


@dataclass
class PeriodMetrics:
    actual: float = 0.0
    plan: float = 0.0

    @property
    def variance(self) -> float:
        return self.actual - self.plan

    @property
    def variance_pct(self) -> float:
        if self.plan == 0:
            return 0.0
        return (self.variance / abs(self.plan)) * 100.0


@dataclass
class RegionMetrics:
    region: str
    mtd: PeriodMetrics = field(default_factory=PeriodMetrics)
    qtd: PeriodMetrics = field(default_factory=PeriodMetrics)
    ytd: PeriodMetrics = field(default_factory=PeriodMetrics)
    mtd_prior_actual: float = 0.0  # for Bullet 9


@dataclass
class OverallMetrics:
    mtd: PeriodMetrics = field(default_factory=PeriodMetrics)
    qtd: PeriodMetrics = field(default_factory=PeriodMetrics)
    ytd: PeriodMetrics = field(default_factory=PeriodMetrics)
    mtd_prior_actual: float = 0.0
    full_quarter_plan: float = 0.0
    months_remaining_in_qtr: int = 0


# ----------------------- helpers -----------------------

def _fmt_money(v: float, sign: bool = False) -> str:
    if sign:
        return f"${v:+,.0f}"
    return f"${v:,.0f}"


def _fmt_pct(v: float, sign: bool = False) -> str:
    if sign:
        return f"{v:+.1f}%"
    return f"{abs(v):.1f}%"


def _above_below(variance: float) -> str:
    return "above" if variance >= 0 else "below"


def _increased_decreased(variance: float) -> str:
    return "increased" if variance >= 0 else "decreased"


def _ahead_behind(variance: float) -> str:
    return "ahead of" if variance >= 0 else "behind"


def _sum_cells(ws, row: int, cols: list[int]) -> float:
    total = 0.0
    for c in cols:
        if c <= 0:
            continue
        v = excel_com.read_cell(ws, row, c)
        if isinstance(v, (int, float)):
            total += float(v)
    return total


def _find_month_columns(ws, header_row: int, labels: Iterable[str]) -> dict[str, int]:
    return {lbl: excel_com.find_column_by_label(ws, header_row, lbl) for lbl in labels}


def _verify_region_labels(sr, fp) -> None:
    """Defensive label checks so a row shift stops us before bad data ships."""
    for region, sr_row, fp_row, header_row, expected_hdr in config.CHART_REGIONS:
        excel_com.assert_label(sr, header_row, 2, expected_hdr)
        excel_com.assert_label(sr, sr_row, 2, "Total")
        excel_com.assert_label(fp, header_row, 2, expected_hdr)
        excel_com.assert_label(fp, fp_row, 2, "Total")


# ----------------------- metric computation -----------------------

def _compute_region_metrics(sr, fp,
                            sr_cols: dict[str, int], fp_cols: dict[str, int],
                            run_dates: date_utils.RunDates) -> list[RegionMetrics]:
    close_m = run_dates.new_month
    close_y = run_dates.new_year
    mtd_label = run_dates.new_month_label
    qtd_labels = date_utils.quarter_month_labels(close_m, close_y)
    ytd_labels = date_utils.ytd_month_labels(close_m, close_y)
    prior_label = date_utils.prior_month_label(run_dates.new_month_date)

    def sr_cs(labels: Iterable[str]) -> list[int]:
        return [sr_cols[l] for l in labels]

    def fp_cs(labels: Iterable[str]) -> list[int]:
        return [fp_cols[l] for l in labels]

    regions: list[RegionMetrics] = []
    for region, sr_row, fp_row, _hdr_row, _exp in config.CHART_REGIONS:
        rm = RegionMetrics(region=region)
        rm.mtd.actual = _sum_cells(sr, sr_row, sr_cs([mtd_label]))
        rm.mtd.plan   = _sum_cells(fp, fp_row, fp_cs([mtd_label]))
        rm.qtd.actual = _sum_cells(sr, sr_row, sr_cs(qtd_labels))
        rm.qtd.plan   = _sum_cells(fp, fp_row, fp_cs(qtd_labels))
        rm.ytd.actual = _sum_cells(sr, sr_row, sr_cs(ytd_labels))
        rm.ytd.plan   = _sum_cells(fp, fp_row, fp_cs(ytd_labels))
        if prior_label in sr_cols and sr_cols[prior_label] > 0:
            v = excel_com.read_cell(sr, sr_row, sr_cols[prior_label])
            rm.mtd_prior_actual = float(v) if isinstance(v, (int, float)) else 0.0
        regions.append(rm)
    return regions


def _compute_overall_metrics(sr, fp,
                             sr_cols: dict[str, int], fp_cols: dict[str, int],
                             run_dates: date_utils.RunDates) -> OverallMetrics:
    close_m = run_dates.new_month
    close_y = run_dates.new_year
    mtd_label = run_dates.new_month_label
    qtd_labels = date_utils.quarter_month_labels(close_m, close_y)
    qtd_full_labels = date_utils.quarter_full_month_labels(close_m, close_y)
    ytd_labels = date_utils.ytd_month_labels(close_m, close_y)
    prior_label = date_utils.prior_month_label(run_dates.new_month_date)

    sr_row = config.TOTAL_BOOKINGS_SUMMARY_ROW
    fp_row = config.TOTAL_BOOKINGS_PLAN_ROW

    def sr_cs(labels): return [sr_cols[l] for l in labels]
    def fp_cs(labels): return [fp_cols[l] for l in labels]

    om = OverallMetrics()
    om.mtd.actual = _sum_cells(sr, sr_row, sr_cs([mtd_label]))
    om.mtd.plan   = _sum_cells(fp, fp_row, fp_cs([mtd_label]))
    om.qtd.actual = _sum_cells(sr, sr_row, sr_cs(qtd_labels))
    om.qtd.plan   = _sum_cells(fp, fp_row, fp_cs(qtd_labels))
    om.ytd.actual = _sum_cells(sr, sr_row, sr_cs(ytd_labels))
    om.ytd.plan   = _sum_cells(fp, fp_row, fp_cs(ytd_labels))
    om.full_quarter_plan = _sum_cells(fp, fp_row, fp_cs(qtd_full_labels))

    if prior_label in sr_cols and sr_cols[prior_label] > 0:
        v = excel_com.read_cell(sr, sr_row, sr_cols[prior_label])
        om.mtd_prior_actual = float(v) if isinstance(v, (int, float)) else 0.0

    quarter_first_m = config.quarter_first_month(config.quarter_of_month(close_m))
    months_elapsed = close_m - quarter_first_m + 1
    om.months_remaining_in_qtr = 3 - months_elapsed
    return om


# ----------------------- bullet composition -----------------------

def _bold_money(v: float) -> str:
    return f"<b>{_fmt_money(v)}</b>"


def _variance_phrase(variance: float, pct: float) -> str:
    return f"{_above_below(variance)} Plan by {_fmt_money(abs(variance))} ({_fmt_pct(pct)})"


def _compose_bullets(overall: OverallMetrics,
                     regions: list[RegionMetrics],
                     run_dates: date_utils.RunDates,
                     warnings: list[step6_validate.Check]) -> list[str]:
    month = date_utils.month_name(run_dates.new_month)
    prior = date_utils.month_name(run_dates.new_month - 1 if run_dates.new_month > 1 else 12)
    year = run_dates.new_year
    qnum = run_dates.quarter

    by_region = {r.region: r for r in regions}
    NA = by_region["NA"]; EMEA = by_region["EMEA"]
    APAC = by_region["APAC"]; LATAM = by_region["LATAM"]

    # MoM
    mom_variance = overall.mtd.actual - overall.mtd_prior_actual
    mom_pct = (mom_variance / abs(overall.mtd_prior_actual) * 100.0
               if overall.mtd_prior_actual else 0.0)

    # Strongest / weakest by MTD variance $
    ordered = sorted(regions, key=lambda r: r.mtd.variance, reverse=True)
    strongest = ordered[0]
    weakest = ordered[-1]

    # MoM-driver region (largest absolute change vs prior month)
    mom_by_region = sorted(
        regions,
        key=lambda r: abs(r.mtd.actual - r.mtd_prior_actual),
        reverse=True,
    )
    mom_driver = mom_by_region[0].region if mom_by_region else "all regions"

    # Forward-looking: on track if QTD actual >= QTD plan
    on_track = "on track" if overall.qtd.variance >= 0 else "at risk"
    full_q_plan = overall.full_quarter_plan
    remaining_needed = max(full_q_plan - overall.qtd.actual, 0.0)
    months_remaining = overall.months_remaining_in_qtr
    mo_word = "month" if months_remaining == 1 else "months"

    bullets: list[str] = []

    # Bullet 1 — MTD Headline
    bullets.append(
        f"Total ACV Bookings for {month} came in at {_bold_money(overall.mtd.actual)} CAD, "
        f"{_variance_phrase(overall.mtd.variance, overall.mtd.variance_pct)} "
        f"and {_above_below(mom_variance)} Prior Month by "
        f"{_fmt_money(abs(mom_variance))} ({_fmt_pct(mom_pct)})."
    )

    # Bullet 2 — YTD Headline
    bullets.append(
        f"Year-to-date ACV Bookings through {month} stand at {_bold_money(overall.ytd.actual)} CAD, "
        f"tracking {_variance_phrase(overall.ytd.variance, overall.ytd.variance_pct)}, "
        f"reflecting {_ahead_behind(overall.ytd.variance)} pace for full-year target."
    )

    # Bullet 3 — QTD Headline
    bullets.append(
        f"Q{qnum} quarter-to-date Bookings of {_bold_money(overall.qtd.actual)} CAD are "
        f"{_variance_phrase(overall.qtd.variance, overall.qtd.variance_pct)}, "
        f"with {months_remaining} {mo_word} remaining in the quarter."
    )

    # Bullet 4 — Strongest Regional (MTD)
    bullets.append(
        f"{strongest.region} was the strongest contributor in {month}, delivering "
        f"{_fmt_money(strongest.mtd.actual)} CAD against a Plan of {_fmt_money(strongest.mtd.plan)}, "
        f"{_above_below(strongest.mtd.variance)} by {_fmt_money(abs(strongest.mtd.variance))} "
        f"({_fmt_pct(strongest.mtd.variance_pct)})."
    )

    # Bullet 5 — Weakest Regional (MTD)
    bullets.append(
        f"{weakest.region} was the largest miss in {month} at {_fmt_money(weakest.mtd.actual)} CAD "
        f"vs a Plan of {_fmt_money(weakest.mtd.plan)}, "
        f"{_above_below(weakest.mtd.variance)} by {_fmt_money(abs(weakest.mtd.variance))} "
        f"({_fmt_pct(weakest.mtd.variance_pct)})."
    )

    # Bullet 6 — NA (MTD and YTD)
    bullets.append(
        f"North America posted MTD Bookings of {_fmt_money(NA.mtd.actual)} CAD "
        f"({_above_below(NA.mtd.variance)} Plan by {_fmt_pct(NA.mtd.variance_pct)}) "
        f"and YTD of {_fmt_money(NA.ytd.actual)} CAD "
        f"({_above_below(NA.ytd.variance)} Plan by {_fmt_pct(NA.ytd.variance_pct)})."
    )

    # Bullet 7 — EMEA (MTD and YTD)
    bullets.append(
        f"EMEA delivered MTD Bookings of {_fmt_money(EMEA.mtd.actual)} CAD "
        f"({_above_below(EMEA.mtd.variance)} Plan by {_fmt_pct(EMEA.mtd.variance_pct)}) "
        f"and YTD of {_fmt_money(EMEA.ytd.actual)} CAD "
        f"({_above_below(EMEA.ytd.variance)} Plan by {_fmt_pct(EMEA.ytd.variance_pct)})."
    )

    # Bullet 8 — APAC and LATAM (MTD)
    bullets.append(
        f"APAC contributed {_fmt_money(APAC.mtd.actual)} CAD MTD "
        f"({_above_below(APAC.mtd.variance)} Plan by {_fmt_pct(APAC.mtd.variance_pct)}) "
        f"while LATAM delivered {_fmt_money(LATAM.mtd.actual)} CAD "
        f"({_above_below(LATAM.mtd.variance)} Plan by {_fmt_pct(LATAM.mtd.variance_pct)})."
    )

    # Bullet 9 — MoM
    bullets.append(
        f"Compared to {prior}, total Bookings {_increased_decreased(mom_variance)} by "
        f"{_fmt_money(abs(mom_variance))} ({_fmt_pct(mom_pct)}), "
        f"driven primarily by {mom_driver}."
    )

    # Bullet 10 — Forward-Looking
    bullets.append(
        f"Based on current QTD trajectory, the business is {on_track} to achieve the "
        f"Q{qnum} Plan of {_fmt_money(full_q_plan)} CAD, requiring "
        f"{_fmt_money(remaining_needed)} in {months_remaining} remaining {mo_word}."
    )

    # WARN-tier override: insert WARNING bullet as Bullet 1, pushing others to 2-11
    if warnings:
        regions_flagged = ", ".join(w.label for w in warnings)
        warn_bullet = (
            f"[WARNING] Validation checks detected variances in {regions_flagged}. "
            "Numbers should be reviewed before distribution."
        )
        bullets.insert(0, warn_bullet)

    return bullets


# ----------------------- HTML assembly -----------------------

def _warning_banner_html(warnings: list[step6_validate.Check]) -> str:
    if not warnings:
        return ""
    lines = [
        "<p><b>Check the below warnings (variance within tolerated band, please review):</b></p>",
        "<ul>",
    ]
    for w in warnings:
        lines.append(
            f"<li><b>{w.label}</b>: Variance of ${w.variance:+,.2f} between Summary By "
            f"Region Row {w.summary_row} and source rows {w.source_start}:{w.source_end}</li>"
        )
    lines.append("</ul>")
    return "\n".join(lines)


def _build_html_body(bullets: list[str],
                     warn_banner_html: str,
                     chart_cids: list[str],
                     file_path: Path) -> str:
    bullets_html = "\n".join(f"<li>{b}</li>" for b in bullets)
    charts_html = "\n".join(
        f'<img src="cid:{cid}" style="display:block; margin:12px 0; max-width:720px;">'
        for cid in chart_cids
    )
    return f"""
<html>
<body style="font-family: Calibri, Arial, sans-serif; font-size: 11pt;">
{warn_banner_html}
<p><b>Executive Summary</b></p>
<ul>
{bullets_html}
</ul>
{charts_html}
<p style="font-size: 10pt; color: #555;">File: {file_path}</p>
</body>
</html>
""".strip()


# ----------------------- main entry -----------------------

def run(workbook_path: Path,
        run_dates: date_utils.RunDates,
        step6_checks: list[step6_validate.Check]) -> str:
    log = get_logger()
    log.info("=== STEP 7: Executive summary email ===")

    warnings = [c for c in step6_checks if c.tier is step6_validate.Tier.WARN]

    with excel_com.excel_session(visible=True) as app:
        with excel_com.open_workbook(app, workbook_path,
                                     read_only=True, save_on_close=False) as wb:
            sr = wb.Worksheets("Summary By Region")
            fp = wb.Worksheets("FY26 Plan")

            _verify_region_labels(sr, fp)

            close_m = run_dates.new_month
            close_y = run_dates.new_year
            mtd_label = run_dates.new_month_label
            prior_label = date_utils.prior_month_label(run_dates.new_month_date)
            all_labels = set(date_utils.ytd_month_labels(close_m, close_y))
            all_labels.update(date_utils.quarter_full_month_labels(close_m, close_y))
            all_labels.add(mtd_label); all_labels.add(prior_label)

            sr_cols = _find_month_columns(sr, config.SUMMARY_HEADER_ROW, all_labels)
            fp_cols = _find_month_columns(fp, config.PLAN_HEADER_ROW, all_labels)
            missing_sr = [l for l, c in sr_cols.items() if c == 0 and l != prior_label]
            missing_fp = [l for l, c in fp_cols.items() if c == 0 and l != prior_label]
            if missing_sr:
                raise RuntimeError(f"Missing Summary By Region columns: {missing_sr}")
            if missing_fp:
                raise RuntimeError(f"Missing FY26 Plan columns: {missing_fp}")

            regions = _compute_region_metrics(sr, fp, sr_cols, fp_cols, run_dates)
            overall = _compute_overall_metrics(sr, fp, sr_cols, fp_cols, run_dates)

    log.info("Overall MTD: actual=%s plan=%s var=%s",
             _fmt_money(overall.mtd.actual), _fmt_money(overall.mtd.plan),
             _fmt_money(overall.mtd.variance, sign=True))
    log.info("Overall YTD: actual=%s plan=%s var=%s",
             _fmt_money(overall.ytd.actual), _fmt_money(overall.ytd.plan),
             _fmt_money(overall.ytd.variance, sign=True))
    for r in regions:
        log.info("%-5s MTD: actual=%s plan=%s var=%s",
                 r.region, _fmt_money(r.mtd.actual),
                 _fmt_money(r.mtd.plan), _fmt_money(r.mtd.variance, sign=True))

    # Charts
    region_labels = [r.region for r in regions]
    mtd_png = charts.clustered_bar_png(
        f"MTD: {date_utils.month_name(run_dates.new_month)} {run_dates.new_year} vs Plan",
        region_labels, [r.mtd.plan for r in regions], [r.mtd.actual for r in regions],
    )
    qtd_png = charts.clustered_bar_png(
        f"QTD: Q{run_dates.quarter} {run_dates.new_year} vs Plan",
        region_labels, [r.qtd.plan for r in regions], [r.qtd.actual for r in regions],
    )
    ytd_png = charts.clustered_bar_png(
        f"YTD: FY{str(run_dates.new_year)[-2:]} vs Plan",
        region_labels, [r.ytd.plan for r in regions], [r.ytd.actual for r in regions],
    )

    bullets = _compose_bullets(overall, regions, run_dates, warnings)
    banner = _warning_banner_html(warnings)
    html = _build_html_body(
        bullets, banner,
        chart_cids=["mtd_chart", "qtd_chart", "ytd_chart"],
        file_path=workbook_path,
    )

    month_name = date_utils.month_name(run_dates.new_month)
    subject = config.EMAIL_SUBJECT_EXEC.format(month=month_name, year=run_dates.new_year)
    entry_id = outlook_draft.create_draft(
        to=config.EMAIL_RECIPIENT,
        subject=subject,
        html_body=html,
        cid_pngs={"mtd_chart": mtd_png, "qtd_chart": qtd_png, "ytd_chart": ytd_png},
    )
    log.info("Step 7 complete — exec summary draft saved (EntryID=%s)", entry_id)
    return entry_id
