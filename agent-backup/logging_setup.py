"""Per-run rotating log file. One file per agent invocation."""
import logging
from datetime import datetime
from pathlib import Path

from agent import config


def setup_run_logger() -> logging.Logger:
    config.LOG_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = config.LOG_DIR / f"run_{ts}.log"

    logger = logging.getLogger("revmodel.agent")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()  # avoid duplicates if re-entered

    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    logger.info(f"Log file: {log_path}")
    return logger


def get_logger() -> logging.Logger:
    return logging.getLogger("revmodel.agent")
