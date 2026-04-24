"""Filename parsing, month-label derivation, period column ranges."""
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from dateutil.relativedelta import relativedelta

from agent import config

_FILENAME_RE = re.compile(
    r"Revenue Model - (\d{8}) \(Internal\)\.xlsm$",
    re.IGNORECASE,
)

MONTH_NAMES = ["", "January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November", "December"]


@dataclass(frozen=True)
class RunDates:
    file_date: date          # parsed from filename
    new_month_date: date     # close month, first day (file month - 1)
    new_month_label: str     # "YYYYMmm"
    src_month_label: str     # "YYYYMmm" — one month before new

    @property
    def new_month(self) -> int:
        return self.new_month_date.month

    @property
    def new_year(self) -> int:
        return self.new_month_date.year

    @property
    def quarter(self) -> int:
        return config.quarter_of_month(self.new_month)


def parse_filename(path: Path) -> date:
    """Parse 'Revenue Model - MMDDYYYY (Internal).xlsm' → date."""
    m = _FILENAME_RE.search(path.name)
    if not m:
        raise ValueError(f"Filename does not match expected pattern: {path.name}")
    s = m.group(1)
    return date(int(s[4:8]), int(s[0:2]), int(s[2:4]))


def month_label(d: date) -> str:
    """Format a date as 'YYYYMmm' (e.g., 2026M04)."""
    return f"{d.year}M{d.month:02d}"


def month_name(m: int) -> str:
    return MONTH_NAMES[m]


def derive_run_dates(file_date: date) -> RunDates:
    """new = file - 1mo (close month); src = file - 2mo (source month for Step 2 macro)."""
    base = date(file_date.year, file_date.month, 1)
    new_month_date = base - relativedelta(months=1)
    src_month_date = base - relativedelta(months=2)
    return RunDates(
        file_date=file_date,
        new_month_date=new_month_date,
        new_month_label=month_label(new_month_date),
        src_month_label=month_label(src_month_date),
    )


def prior_month_label(close_date: date) -> str:
    """The month before the close month (used for MoM comparison)."""
    return month_label(close_date - relativedelta(months=1))


def quarter_month_labels(close_month: int, year: int) -> list[str]:
    """Months in the containing quarter, up to and including close_month."""
    q = config.quarter_of_month(close_month)
    first = config.quarter_first_month(q)
    return [month_label(date(year, m, 1)) for m in range(first, close_month + 1)]


def quarter_full_month_labels(close_month: int, year: int) -> list[str]:
    """All 3 months of the containing quarter (used for Plan QTD when picking the precomputed col)."""
    q = config.quarter_of_month(close_month)
    first = config.quarter_first_month(q)
    return [month_label(date(year, m, 1)) for m in range(first, first + 3)]


def ytd_month_labels(close_month: int, year: int) -> list[str]:
    """Jan through close_month of year."""
    return [month_label(date(year, m, 1)) for m in range(1, close_month + 1)]


def find_latest_workbook(folder: Path) -> Path:
    """Most recent Revenue Model file by filename date."""
    candidates = []
    for p in folder.glob("Revenue Model - * (Internal).xlsm"):
        try:
            d = parse_filename(p)
            candidates.append((d, p))
        except ValueError:
            continue
    if not candidates:
        raise FileNotFoundError(f"No Revenue Model files found in {folder}")
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def expected_renamed_path(today: date | None = None) -> Path:
    """Where the file will live after Step 1 renames it to today's MMDDYYYY."""
    if today is None:
        today = date.today()
    name = f"{config.FILE_PREFIX}{today.strftime(config.FILE_DATE_FORMAT)}{config.FILE_SUFFIX}{config.FILE_EXT}"
    return config.ROOT / name
