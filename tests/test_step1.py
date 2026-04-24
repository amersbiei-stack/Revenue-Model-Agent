"""Step 1 archive+rename tests — pure Python file ops, no Excel."""
import sys
import tempfile
from datetime import date
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from agent import step1_archive_rename  # noqa: E402


def _make_workbook(root: Path, name: str, content: bytes = b"xlsm-stub") -> Path:
    p = root / name
    p.write_bytes(content)
    return p


def _with_patched_root(tmp_root: Path):
    """Return a context patching both config.ROOT and config.ARCHIVE_DIR."""
    archive = tmp_root / "Archive"
    return patch.multiple("agent.config", ROOT=tmp_root, ARCHIVE_DIR=archive)


def test_archive_and_rename_happy_path():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        src = _make_workbook(root, "Revenue Model - 04062026 (Internal).xlsm")
        with _with_patched_root(root):
            out = step1_archive_rename.run(src, today=date(2026, 4, 22))
        assert out.name == "Revenue Model - 04222026 (Internal).xlsm"
        assert out.exists()
        assert not src.exists(), "Source should have been renamed away"
        backup = root / "Archive" / "Revenue Model - 04062026 (Internal) Backup.xlsm"
        assert backup.exists(), "Archive copy missing"
        assert backup.read_bytes() == b"xlsm-stub"


def test_archive_collision_gets_numeric_suffix():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        (root / "Archive").mkdir()
        # Pre-seed colliding backups
        (root / "Archive" / "Revenue Model - 04062026 (Internal) Backup.xlsm").write_bytes(b"old1")
        (root / "Archive" / "Revenue Model - 04062026 (Internal) Backup (2).xlsm").write_bytes(b"old2")
        src = _make_workbook(root, "Revenue Model - 04062026 (Internal).xlsm", b"fresh")
        with _with_patched_root(root):
            step1_archive_rename.run(src, today=date(2026, 4, 22))
        expected = root / "Archive" / "Revenue Model - 04062026 (Internal) Backup (3).xlsm"
        assert expected.exists()
        assert expected.read_bytes() == b"fresh"


def test_same_day_rerun_overwrites_target():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        src = _make_workbook(root, "Revenue Model - 04062026 (Internal).xlsm", b"new")
        # Stale file already at today's dated name from an earlier run
        stale = _make_workbook(root, "Revenue Model - 04222026 (Internal).xlsm", b"stale")
        with _with_patched_root(root):
            out = step1_archive_rename.run(src, today=date(2026, 4, 22))
        assert out == stale.parent / "Revenue Model - 04222026 (Internal).xlsm"
        assert out.read_bytes() == b"new", "Stale file should have been overwritten"
        assert not src.exists()


def test_source_is_already_todays_name_is_noop_rename():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        src = _make_workbook(root, "Revenue Model - 04222026 (Internal).xlsm", b"payload")
        with _with_patched_root(root):
            out = step1_archive_rename.run(src, today=date(2026, 4, 22))
        assert out == src
        assert src.exists()
        backup = root / "Archive" / "Revenue Model - 04222026 (Internal) Backup.xlsm"
        assert backup.exists()


def test_missing_source_raises():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        ghost = root / "Revenue Model - 04062026 (Internal).xlsm"
        with _with_patched_root(root):
            try:
                step1_archive_rename.run(ghost, today=date(2026, 4, 22))
            except FileNotFoundError:
                return
        raise AssertionError("Expected FileNotFoundError")


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
