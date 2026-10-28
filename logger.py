"""Structured logging to the console and two files in the log folder.

main.log holds every line from every run. <date>.log holds only that day's
lines, so a single day can be read or handed over without scrolling through
the whole history. The duplicate-handler guard is carried over from
venire_3.0's logger.py: setup is called once from main, but a second call must
not double every line.

No names or dates of birth go into these logs. A row is told apart by its
row number and UCN, which is enough to find it in the working copy. Callers
are written that way, and as a backstop every handler runs a Redactor: each
name and date of birth the run reads is registered with hide_names and
hide_values, and any of them that still turns up in a line, an error message
or a traceback is replaced before it is written. Nothing out of the PDF itself
is ever logged.
"""

import logging
import os
import re
import sys
import threading
from collections.abc import Iterable
from datetime import date
from pathlib import Path

LOGGER_NAME = "RapSheet"
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)-7s - %(message)s"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
MAIN_LOG_NAME = "main.log"

# What a hidden name or value is replaced with.
NAME_MASK = "[name]"
VALUE_MASK = "[dob]"

# Name parts shorter than this are not hidden. A middle initial on its own
# would otherwise mask every lone "A" or "J" in every line, and says nothing
# about who the person is.
MIN_NAME_PART = 2

# How a name is cut into the parts that are hidden one by one, so that
# "DOE, John" is caught as DOE_John in a PDF filename or John Doe in a
# message. Apostrophes stay inside a part (O'BRIEN), hyphens split it.
NAME_PART_SPLIT = re.compile(r"[^A-Za-z']+")


class Redactor(logging.Filter):
    """Replaces registered names and values in every record it sees.

    Attached to the handlers, not the logger: a filter on the logger only
    sees records logged on it directly, not those its children pass up.
    """

    def __init__(self) -> None:
        """Start with nothing registered."""
        super().__init__()
        self._lock = threading.Lock()
        self._names: set[str] = set()
        self._values: set[str] = set()
        self._name_pattern: re.Pattern | None = None
        self._value_pattern: re.Pattern | None = None

    def hide_names(self, names: Iterable[str]) -> None:
        """Register names to be replaced with NAME_MASK.

        Each name is cut into its parts, and each part is matched on its own,
        whole word only and ignoring case, so it is caught however a line
        happens to arrange the name.

        Args:
            names: Names as the report or STAC spells them. Blanks are ignored.
        """
        parts = set()
        for name in names:
            for part in NAME_PART_SPLIT.split(str(name or "")):
                part = part.strip("'")
                if len(part) >= MIN_NAME_PART:
                    parts.add(part.upper())
                    # O'BRIEN is also caught as OBRIEN and BRIEN, in case a
                    # filename or a search dropped the apostrophe.
                    parts.add(part.replace("'", "").upper())
                    parts.update(
                        piece.upper() for piece in part.split("'")
                        if len(piece) >= MIN_NAME_PART
                    )
        with self._lock:
            if parts - self._names:
                self._names |= parts
                self._name_pattern = _whole_word_pattern(self._names, r"[A-Za-z]")

    def hide_values(self, values: Iterable[str]) -> None:
        """Register exact values, such as dates of birth, to be replaced with VALUE_MASK.

        Args:
            values: Values as they may appear in a line. Blanks are ignored.
        """
        new = {str(value).strip() for value in values if str(value or "").strip()}
        with self._lock:
            if new - self._values:
                self._values |= new
                self._value_pattern = _whole_word_pattern(self._values, r"[0-9]")

    def redact(self, text: str) -> str:
        """Return `text` with every registered name and value replaced.

        Args:
            text: Any text about to be logged.

        Returns:
            The text with NAME_MASK and VALUE_MASK in place of what was hidden.
        """
        with self._lock:
            name_pattern, value_pattern = self._name_pattern, self._value_pattern
        if value_pattern:
            text = value_pattern.sub(VALUE_MASK, text)
        if name_pattern:
            text = name_pattern.sub(NAME_MASK, text)
        return text

    def filter(self, record: logging.LogRecord) -> bool:
        """Redact the record's message and any traceback, then let it through.

        The message is formatted here and its arguments dropped, since a name
        is just as often in an argument as in the format string. The
        traceback is formatted here too, because an exception's text can
        carry a name the message never mentioned.

        Args:
            record: The record about to be written.

        Returns:
            Always True. Nothing is dropped, only masked.
        """
        message = record.getMessage()
        redacted = self.redact(message)
        if redacted != message:
            record.msg = redacted
            record.args = None
        if record.exc_info and not record.exc_text:
            record.exc_text = logging.Formatter().formatException(record.exc_info)
        if record.exc_text:
            record.exc_text = self.redact(record.exc_text)
        return True


def _whole_word_pattern(words: set[str], word_char: str) -> re.Pattern:
    """A case-insensitive pattern matching any of `words` as a whole word.

    Longest first, so DOE-SMITH style overlaps take the longer match.

    Args:
        words: What to match.
        word_char: A character class that must not touch either side of a
            match, so DOE is caught in DOE_John but not inside DOES.

    Returns:
        The compiled pattern.
    """
    alternatives = "|".join(re.escape(word) for word in sorted(words, key=len, reverse=True))
    return re.compile(
        rf"(?<!{word_char})(?:{alternatives})(?!{word_char})", re.IGNORECASE
    )


# One for the whole run, shared by every handler, so a name registered by any
# module is hidden in every log and on the console.
_REDACTOR = Redactor()


def hide_names(names: Iterable[str]) -> None:
    """Register names that must never appear in the logs. See Redactor.hide_names.

    Args:
        names: Names as the report or STAC spells them.
    """
    _REDACTOR.hide_names(names)


def hide_values(values: Iterable[str]) -> None:
    """Register exact values, such as dates of birth, that must never appear in the logs.

    Args:
        values: Values as they may appear in a line.
    """
    _REDACTOR.hide_values(values)


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
        handler.addFilter(_REDACTOR)
        logger.addHandler(handler)

    to_console = logging.StreamHandler(sys.stdout)
    to_console.setLevel(logging.INFO)
    to_console.setFormatter(formatter)
    to_console.addFilter(_REDACTOR)
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
