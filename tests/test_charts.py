"""Smoke tests for chart PNG generation."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from agent import charts  # noqa: E402


def test_returns_png_bytes():
    png = charts.clustered_bar_png(
        title="MTD: April 2026 vs Plan",
        regions=["NA", "EMEA", "APAC", "LATAM"],
        budget_values=[898_432, 426_638, 254_625, 342_759],
        actual_values=[910_500, 412_000, 245_000, 360_000],
    )
    assert isinstance(png, bytes)
    assert png[:8] == b"\x89PNG\r\n\x1a\n", "Not a valid PNG header"
    assert len(png) > 5_000


def test_negative_values_render():
    png = charts.clustered_bar_png(
        "Test — with negatives",
        ["A", "B"], [100_000, -50_000], [75_000, -30_000],
    )
    assert isinstance(png, bytes) and len(png) > 1_000


def test_zero_plan_renders():
    png = charts.clustered_bar_png(
        "Zero plan case",
        ["Only"], [0.0], [100_000.0],
    )
    assert isinstance(png, bytes) and len(png) > 1_000


def test_mismatched_lengths_raises():
    try:
        charts.clustered_bar_png("bad", ["A", "B"], [1.0], [2.0, 3.0])
    except ValueError:
        return
    raise AssertionError("Expected ValueError for mismatched lengths")


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
