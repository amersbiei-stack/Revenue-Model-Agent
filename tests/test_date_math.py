"""Date-math + filename parsing tests."""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from agent import date_utils, config  # noqa: E402


def test_filename_parse_apr2026():
    d = date_utils.parse_filename(Path("Revenue Model - 04062026 (Internal).xlsm"))
    assert d == date(2026, 4, 6)


def test_filename_parse_may2026():
    d = date_utils.parse_filename(Path("Revenue Model - 05072026 (Internal).xlsm"))
    assert d == date(2026, 5, 7)


def test_filename_parse_jan_wraps():
    d = date_utils.parse_filename(Path("Revenue Model - 01052027 (Internal).xlsm"))
    assert d == date(2027, 1, 5)


def test_filename_parse_rejects():
    try:
        date_utils.parse_filename(Path("Revenue Model External.xlsm"))
    except ValueError:
        return
    raise AssertionError("Expected ValueError")


def test_month_label():
    assert date_utils.month_label(date(2026, 4, 1)) == "2026M04"
    assert date_utils.month_label(date(2026, 12, 1)) == "2026M12"
    assert date_utils.month_label(date(2027, 1, 1)) == "2027M01"


def test_derive_run_dates_apr6_2026():
    """File dated 4/6/2026 → close = March 2026, src = February 2026."""
    rd = date_utils.derive_run_dates(date(2026, 4, 6))
    assert rd.new_month_label == "2026M03"
    assert rd.src_month_label == "2026M02"
    assert rd.new_month == 3 and rd.new_year == 2026
    assert rd.quarter == 1


def test_derive_run_dates_may_for_april_close():
    """Running in May → April is close month."""
    rd = date_utils.derive_run_dates(date(2026, 5, 4))
    assert rd.new_month_label == "2026M04"
    assert rd.src_month_label == "2026M03"
    assert rd.quarter == 2


def test_derive_run_dates_jan_wraps_year():
    """January filename → close is prior December."""
    rd = date_utils.derive_run_dates(date(2027, 1, 5))
    assert rd.new_month_label == "2026M12"
    assert rd.src_month_label == "2026M11"
    assert rd.quarter == 4


def test_quarter_month_labels():
    # April close in Q2 → only April
    assert date_utils.quarter_month_labels(4, 2026) == ["2026M04"]
    # June close → full Q2
    assert date_utils.quarter_month_labels(6, 2026) == ["2026M04", "2026M05", "2026M06"]
    # January → only January
    assert date_utils.quarter_month_labels(1, 2026) == ["2026M01"]


def test_quarter_full_month_labels():
    assert date_utils.quarter_full_month_labels(4, 2026) == ["2026M04", "2026M05", "2026M06"]
    assert date_utils.quarter_full_month_labels(11, 2026) == ["2026M10", "2026M11", "2026M12"]


def test_ytd_month_labels():
    assert date_utils.ytd_month_labels(4, 2026) == [
        "2026M01", "2026M02", "2026M03", "2026M04"
    ]
    assert date_utils.ytd_month_labels(1, 2026) == ["2026M01"]


def test_prior_month_label():
    assert date_utils.prior_month_label(date(2026, 4, 1)) == "2026M03"
    assert date_utils.prior_month_label(date(2026, 1, 1)) == "2025M12"


def test_quarter_of_month():
    assert config.quarter_of_month(1) == 1
    assert config.quarter_of_month(3) == 1
    assert config.quarter_of_month(4) == 2
    assert config.quarter_of_month(7) == 3
    assert config.quarter_of_month(12) == 4


def test_quarter_first_month():
    assert config.quarter_first_month(1) == 1
    assert config.quarter_first_month(2) == 4
    assert config.quarter_first_month(3) == 7
    assert config.quarter_first_month(4) == 10


def test_month_name():
    assert date_utils.month_name(1) == "January"
    assert date_utils.month_name(4) == "April"
    assert date_utils.month_name(12) == "December"


if __name__ == "__main__":
    fns = [v for k, v in list(globals().items()) if k.startswith("test_")]
    passed = failed = 0
    for fn in fns:
        try:
            fn(); passed += 1
            print(f"  OK    {fn.__name__}")
        except Exception as e:
            failed += 1
            print(f"  FAIL  {fn.__name__}: {e!r}")
    print(f"\n{passed} passed, {failed} failed")
    sys.exit(0 if failed == 0 else 1)
