"""Step 6 tolerance classification tests."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from agent.step6_validate import _classify, Tier  # noqa: E402


def test_subregional_pass_at_threshold():
    assert _classify(0.01, is_total=False) == Tier.PASS
    assert _classify(-0.01, is_total=False) == Tier.PASS


def test_subregional_warn_just_above_tolerance():
    assert _classify(0.02, is_total=False) == Tier.WARN
    assert _classify(-0.02, is_total=False) == Tier.WARN


def test_subregional_warn_at_block_boundary():
    assert _classify(1000.00, is_total=False) == Tier.WARN
    assert _classify(-1000.00, is_total=False) == Tier.WARN


def test_subregional_block_just_over():
    assert _classify(1000.01, is_total=False) == Tier.BLOCK
    assert _classify(-1000.01, is_total=False) == Tier.BLOCK


def test_total_pass_at_threshold():
    assert _classify(1.00, is_total=True) == Tier.PASS
    assert _classify(-1.00, is_total=True) == Tier.PASS


def test_total_warn_just_above():
    assert _classify(1.01, is_total=True) == Tier.WARN


def test_total_block_above_uniform_limit():
    assert _classify(1000.01, is_total=True) == Tier.BLOCK
    assert _classify(-5000.00, is_total=True) == Tier.BLOCK


def test_zero_variance_is_pass():
    assert _classify(0.0, is_total=False) == Tier.PASS
    assert _classify(0.0, is_total=True) == Tier.PASS


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
