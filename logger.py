"""Structured logging to a dated file and the console.

One log per run, named for the day, alongside the PDFs that run produced. The
duplicate-handler guard is carried over from venire_3.0's logger.py: setup is
called once from main, but a second call must not double every line.

Names appear in these logs. A rap sheet is a person's criminal history, so the
log carries the defendant name from the spreadsheet row and the case number,
because without them nobody can tell which row failed. It carries nothing out
of the PDF itself.
"""

import logging
import sys
from pathlib import Path

LOGGER_NAME = "RapSheet"
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)-7s - %(message)s"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"


def setup_logging(log_path: Path) -> logging.Logger:
    """Attach a file handler and a console handler to the root RAP SHEET logger.

    Args:
        log_path: Full path to the run's .log file. Parent folders are created.

    Returns:
        The configured logger.

    Example:
        log = setup_logging(Path("logs/2026-10-02.log"))
    """
    log_path = Path(log_path)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.DEBUG)

    # Called twice, every line would appear twice. Venire learned this one.
    if logger.handlers:
        return logger

    formatter = logging.Formatter(LOG_FORMAT, datefmt=DATE_FORMAT)

    to_file = logging.FileHandler(log_path, encoding="utf-8")
    to_file.setLevel(logging.DEBUG)
    to_file.setFormatter(formatter)
    logger.addHandler(to_file)

    to_console = logging.StreamHandler(sys.stdout)
    to_console.setLevel(logging.INFO)
    to_console.setFormatter(formatter)
    logger.addHandler(to_console)

    return logger


def get_logger(module: str = "") -> logging.Logger:
    """Return the run logger, or a child of it named for one module.

    Args:
        module: Module name, e.g. "stac". Empty returns the parent.

    Example:
        logger = get_logger(__name__)
    """
    if not module:
        return logging.getLogger(LOGGER_NAME)
    return logging.getLogger(LOGGER_NAME).getChild(module)
