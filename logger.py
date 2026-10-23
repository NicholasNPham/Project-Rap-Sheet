"""Structured logging to the console and two files in the log folder.

main.log holds every line from every run. <date>.log holds only that day's
lines, so a single day can be read or handed over without scrolling through
the whole history. The duplicate-handler guard is carried over from
venire_3.0's logger.py: setup is called once from main, but a second call must
not double every line.

Names appear in these logs. A rap sheet is a person's criminal history, so the
log carries the defendant name from the spreadsheet row and the case number,
because without them nobody can tell which row failed. It carries nothing out
of the PDF itself.
"""

import logging
import os
import sys
from datetime import date
from pathlib import Path

LOGGER_NAME = "RapSheet"
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)-7s - %(message)s"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
MAIN_LOG_NAME = "main.log"


class DayFileHandler(logging.FileHandler):
    """Writes to <folder>/<date>.log and moves to a new file when the day changes.

    A --watch is started once and runs for days, so the file is chosen by each
    line's own date rather than once at setup.
    """

    def __init__(self, folder: Path) -> None:
        """Open today's file in `folder`.

        Args:
            folder: The log folder. It must already exist.
        """
        self.folder = Path(folder)
        self.day = date.today()
        super().__init__(self._path_for(self.day), encoding="utf-8")

    def _path_for(self, day: date) -> Path:
        """Return the day file's path for `day`.

        Args:
            day: The date the file is for.

        Returns:
            <folder>/<YYYY-MM-DD>.log
        """
        return self.folder / f"{day.isoformat()}.log"

    def emit(self, record: logging.LogRecord) -> None:
        """Write one line, switching to a new day's file first if needed.

        Args:
            record: The log record to write.
        """
        day = date.fromtimestamp(record.created)
        if day != self.day:
            # Close only the stream. Handler.close() would also unregister the
            # handler, and FileHandler.emit reopens the stream from
            # baseFilename on the next write.
            if self.stream:
                self.stream.close()
                self.stream = None
            self.day = day
            self.baseFilename = os.path.abspath(self._path_for(day))
        super().emit(record)


def setup_logging(log_dir: Path) -> logging.Logger:
    """Attach the main log, the day log and the console to the RAP SHEET logger.

    Args:
        log_dir: The log folder. Created if missing.

    Returns:
        The configured logger.

    Example:
        log = setup_logging(Path("logs"))
    """
    log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.DEBUG)

    # Called twice, every line would appear twice. Venire learned this one.
    if logger.handlers:
        return logger

    formatter = logging.Formatter(LOG_FORMAT, datefmt=DATE_FORMAT)

    to_main = logging.FileHandler(log_dir / MAIN_LOG_NAME, encoding="utf-8")
    to_day = DayFileHandler(log_dir)
    for handler in (to_main, to_day):
        handler.setLevel(logging.DEBUG)
        handler.setFormatter(formatter)
        logger.addHandler(handler)

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
