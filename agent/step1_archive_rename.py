"""Step 1 — Archive + rename the live workbook via pure Python file ops.

1a. Copy the live workbook to ROOT/Archive with a " Backup" suffix inserted
    before the extension. Collision-safe: " Backup (2)", " Backup (3)", ...
1b. Rename the live workbook to today's MMDDYYYY date.
    - If source is already today's dated name, skip the rename (idempotent).
    - If a different file at today's dated path already exists (same-day
      re-run), overwrite it.

Post-conditions:
    - Backup copy exists in ROOT/Archive.
    - Live file has today's MMDDYYYY in the name.
    - Source path no longer exists (unless source was already today's name).
"""
import shutil
from datetime import date
from pathlib import Path

from agent import config, date_utils
from agent.logging_setup import get_logger


def _archive_copy_path(source: Path, archive_dir: Path) -> Path:
    """Return a non-colliding '<stem> Backup<suffix>.ext' path inside archive_dir."""
    stem, ext = source.stem, source.suffix  # "Revenue Model - 04062026 (Internal)", ".xlsm"
    base = archive_dir / f"{stem} Backup{ext}"
    if not base.exists():
        return base
    i = 2
    while True:
        candidate = archive_dir / f"{stem} Backup ({i}){ext}"
        if not candidate.exists():
            return candidate
        i += 1


def _archive(source: Path, log) -> Path:
    archive_dir = config.ARCHIVE_DIR
    archive_dir.mkdir(parents=True, exist_ok=True)
    dst = _archive_copy_path(source, archive_dir)
    log.info("1a. Archiving: %s -> %s", source.name, dst)
    shutil.copy2(source, dst)
    if not dst.exists():
        raise RuntimeError(f"Archive copy not found after copy: {dst}")
    log.info("1a. Archive copy confirmed (%d bytes)", dst.stat().st_size)
    return dst


def _rename_to_today(source: Path, today: date, log) -> Path:
    target = date_utils.expected_renamed_path(today)
    if source.resolve() == target.resolve():
        log.info("1b. Source already matches today's name; no rename needed: %s",
                 source.name)
        return source
    if target.exists():
        log.warning("1b. Target already exists — overwriting: %s", target.name)
        target.unlink()
    log.info("1b. Renaming: %s -> %s", source.name, target.name)
    source.rename(target)
    if not target.exists():
        raise RuntimeError(f"Rename failed: {target} missing after rename")
    if source.exists() and source.resolve() != target.resolve():
        raise RuntimeError(f"Source still exists after rename: {source}")
    return target


def run(workbook_path: Path, today: date | None = None) -> Path:
    """Archive + rename. Returns the post-rename live path."""
    log = get_logger()
    log.info("=== STEP 1: Archive + rename (pure Python) ===")
    log.info("Source workbook: %s", workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    _archive(workbook_path, log)
    renamed = _rename_to_today(workbook_path, today or date.today(), log)
    log.info("Step 1 complete. Live file: %s", renamed.name)
    return renamed
