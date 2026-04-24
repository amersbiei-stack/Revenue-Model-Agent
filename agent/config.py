"""Central configuration for the Revenue Model rollover agent.

Machine-specific values (paths, email, Prophix wait) are loaded from
`config.json` at the project root. Data-specific values (tab names,
row numbers, tolerance thresholds, chart styling) live in this file.

See SETUP.md for what to edit when moving to a new machine.
"""
from __future__ import annotations

import json
from pathlib import Path

# === JSON config loader ================================================
# Project root is inferred from this file's location: agent/config.py -> parent is project root.
# config.json must live at the project root. Users can override the project_root inside the JSON
# if they move the folder without moving the agent source (unusual but supported).

_CODE_PROJECT_ROOT = Path(__file__).resolve().parents[1]
_CONFIG_JSON_PATH = _CODE_PROJECT_ROOT / "config.json"

if not _CONFIG_JSON_PATH.exists():
    raise FileNotFoundError(
        f"Missing {_CONFIG_JSON_PATH}.\n"
        f"  Copy {_CODE_PROJECT_ROOT / 'config.example.json'} to config.json "
        f"and fill in the machine-specific values. See SETUP.md."
    )

with open(_CONFIG_JSON_PATH, "r", encoding="utf-8") as _f:
    _cfg: dict = json.load(_f)


def _required(key: str):
    if key not in _cfg or _cfg[key] is None or str(_cfg[key]).startswith("<REPLACE_ME"):
        raise ValueError(
            f"config.json is missing required key '{key}' or still has a <REPLACE_ME> placeholder. "
            f"See SETUP.md."
        )
    return _cfg[key]


# === Paths (machine-specific) ==========================================
ROOT = Path(_required("project_root"))
ARCHIVE_DIR = Path(_cfg["archive_dir"]) if _cfg.get("archive_dir") else ROOT / "Archive"
LOG_DIR = Path(_cfg["log_dir"]) if _cfg.get("log_dir") else ROOT / "agent" / "logs"

# === Filename pattern (data-specific, rarely changed) ==================
FILE_PREFIX = _cfg.get("file_prefix", "Revenue Model - ")
FILE_SUFFIX = _cfg.get("file_suffix", " (Internal)")
FILE_EXT = _cfg.get("file_ext", ".xlsm")
FILE_DATE_FORMAT = _cfg.get("file_date_format", "%m%d%Y")

# === Email ==============================================================
EMAIL_RECIPIENT = _required("email_recipient")
EMAIL_SUBJECT_EXEC = "Revenue Model - {month} {year} Close | Executive Summary"
EMAIL_SUBJECT_VARIANCE = "Revenue Model - {month} {year} Close | VARIANCE ALERT - Exec Summary Blocked"
EMAIL_SUBJECT_STEP5_FAIL = "Revenue Model - {month} {year} Close | Step 5 FAILURE - Prophix data did not populate"

# === Step 4 =============================================================
PROPHIX_WAIT_SECONDS = int(_cfg.get("prophix_wait_seconds", 600))  # 10 minutes default

# === Step 6 tolerance tiers (uniform $1,000 BLOCK threshold) ============
TOLERANCE_SUBREGIONAL = 0.01
TOLERANCE_TOTAL = 1.00
BLOCK_LIMIT = 1000.00

# === VBA macro names ====================================================
# Step 1 is handled in pure Python (see step1_archive_rename.py) — no macro needed.
MACRO_STEP2_ROLL_FORMULAS = "Step2_RollFormulas_TwoMonthsBack_SourceToTarget_Reliable"

# === Step 3 — per-country tabs whose formulas the Step2 macro rolls =====
STEP3_TABS = [
    "NA", "NA_USD", "NA_CAD", "UK_GBP", "Europe & MEA", "Beneleux",
    "DACH", "Nordics", "Western Europe", "Eastern Europe", "ME&A",
    "LATAM PX", "LATAM AS", "Asia", "Pacific (FG)",
]

# === Step 5 — Prophix-refresh verification tabs =========================
# tab_name -> (header_row, data_start_row, data_end_row)
STEP5_TABS = {
    "Units Bookings DV": (2, 3, 152),
    "$$ Bookings DV":    (2, 3, 602),
    "Sub Europe Units":  (2, 3, 182),
    "Migration Units":   (2, 3, 17),
    "AS":                (2, 3, 122),
}

# === Step 6 dollar tie-out checks =======================================
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

# === Step 6 units tie-out check =========================================
# (label, summary_row, units_start, units_end)
STEP6_UNITS_CHECK = ("Total Units", 169, 3, 152)

# === Step 7 chart source rows (per-region CAD totals) ===================
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

# === Header rows for month-label lookup =================================
SUMMARY_HEADER_ROW = 17
PLAN_HEADER_ROW = 2
PER_COUNTRY_HEADER_ROW = 4

# === Chart styling ======================================================
CHART_BUDGET_COLOR = "#E9EAEF"
CHART_ACTUAL_COLOR = "#095A87"
CHART_DPI = 150
CHART_FIGSIZE = (8, 4)

# === Calendar fiscal year (verified from FY26 Plan!C14='Q1-2026') =======
def quarter_of_month(m: int) -> int:
    """1->Q1, 2->Q1, 3->Q1, 4->Q2, ..., 12->Q4."""
    return (m - 1) // 3 + 1


def quarter_first_month(q: int) -> int:
    """Q1->1, Q2->4, Q3->7, Q4->10."""
    return (q - 1) * 3 + 1
