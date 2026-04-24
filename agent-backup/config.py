"""Central configuration. Single source of truth for paths, thresholds, tab/row maps."""
from pathlib import Path

# === Paths ===
ROOT = Path(r"C:\Users\amers\Downloads\Revenue Model")
ARCHIVE_DIR = ROOT / "Archive"
LOG_DIR = ROOT / "agent" / "logs"

# === Filename pattern ===
# e.g. "Revenue Model - 04062026 (Internal).xlsm"
FILE_PREFIX = "Revenue Model - "
FILE_SUFFIX = " (Internal)"
FILE_EXT = ".xlsm"
FILE_DATE_FORMAT = "%m%d%Y"

# === Email ===
EMAIL_RECIPIENT = "asbiei@prophix.com"
EMAIL_SUBJECT_EXEC = "Revenue Model - {month} {year} Close | Executive Summary"
EMAIL_SUBJECT_VARIANCE = "Revenue Model - {month} {year} Close | VARIANCE ALERT - Exec Summary Blocked"
EMAIL_SUBJECT_STEP5_FAIL = "Revenue Model - {month} {year} Close | Step 5 FAILURE - Prophix data did not populate"

# === Step 4 ===
PROPHIX_WAIT_SECONDS = 600  # 10 minutes

# === Step 6 tolerance tiers (uniform $1,000 BLOCK threshold) ===
TOLERANCE_SUBREGIONAL = 0.01
TOLERANCE_TOTAL = 1.00
BLOCK_LIMIT = 1000.00

# === VBA macro names ===
# Step 1 is handled in pure Python (see step1_archive_rename.py) — no macro needed.
MACRO_STEP2_ROLL_FORMULAS = "Step2_RollFormulas_TwoMonthsBack_SourceToTarget_Reliable"

# === Step 3 — per-country tabs whose formulas the Step2 macro rolls ===
STEP3_TABS = [
    "NA", "NA_USD", "NA_CAD", "UK_GBP", "Europe & MEA", "Beneleux",
    "DACH", "Nordics", "Western Europe", "Eastern Europe", "ME&A",
    "LATAM PX", "LATAM AS", "Asia", "Pacific (FG)",
]

# === Step 5 — Prophix-refresh verification tabs ===
# tab_name -> (header_row, data_start_row, data_end_row)
STEP5_TABS = {
    "Units Bookings DV": (2, 3, 152),
    "$$ Bookings DV":    (2, 3, 602),
    "Sub Europe Units":  (2, 3, 182),
    "Migration Units":   (2, 3, 17),
    "AS":                (2, 3, 122),
}

# === Step 6 dollar tie-out checks ===
# (label, summary_row, bookings_start, bookings_end, is_total)
STEP6_DOLLAR_CHECKS = [
    ("NA",             22, 3,   182, False),
    ("UK",             29, 183, 242, False),
    ("Europe & MEA",   36, 243, 362, False),
    ("APAC",           64, 363, 482, False),
    ("LATAM",          85, 483, 602, False),
    ("Total Bookings", 92, 3,   602, True),
    # Subregional drill-downs
    ("Asia",           50, 363, 422, False),
    ("Pacific",        57, 423, 482, False),
    ("LATAM PX",       71, 483, 542, False),
    ("LATAM AS",       78, 543, 602, False),
]

# === Step 6 units tie-out check ===
# (label, summary_row, units_start, units_end)
STEP6_UNITS_CHECK = ("Total Units", 169, 3, 152)

# === Step 7 chart source rows (per-region CAD totals) ===
# (region, summary_row, plan_row, header_row, expected_header_label)
# Rows are identical on Summary By Region and FY26 Plan Internal block.
CHART_REGIONS = [
    ("NA",    22, 22, 17, "NA $$"),
    ("EMEA",  43, 43, 38, "EMEA $$"),
    ("APAC",  64, 64, 59, "APAC $$"),
    ("LATAM", 85, 85, 80, "LATAM $$"),
]

# Top-line totals used for exec summary headlines (MTD/QTD/YTD overall)
TOTAL_BOOKINGS_SUMMARY_ROW = 92  # Summary By Region: "Total Regions $$" Total
TOTAL_BOOKINGS_PLAN_ROW = 92     # FY26 Plan Internal: "Total Regions $$" Total

# === Header rows for month-label lookup ===
SUMMARY_HEADER_ROW = 17
PLAN_HEADER_ROW = 2
PER_COUNTRY_HEADER_ROW = 4

# === Chart styling ===
CHART_BUDGET_COLOR = "#E9EAEF"
CHART_ACTUAL_COLOR = "#095A87"
CHART_DPI = 150
CHART_FIGSIZE = (8, 4)

# === Calendar fiscal year (verified from FY26 Plan!C14='Q1-2026') ===
def quarter_of_month(m: int) -> int:
    """1->Q1, 2->Q1, 3->Q1, 4->Q2, ..., 12->Q4."""
    return (m - 1) // 3 + 1


def quarter_first_month(q: int) -> int:
    """Q1->1, Q2->4, Q3->7, Q4->10."""
    return (q - 1) * 3 + 1
