"""Reads rows from the spreadsheet and writes each row's outcome back.

Ported from venire_3.0's excel_handler.py, with three changes.

The row carries a UCN. That is the whole reason this project exists: the case
number sits in the same row as the name and date of birth, so the PDF that
CCIS produces for that person has somewhere to be filed.

The workbook is opened once and held, not reloaded for every write. venire
reloads and re-saves the entire file per row, which is what its own
"DO NOT END PROCESS. IT WILL CORRUPT THE VENIRE EXCEL SHEET" warning is about.

Saves are atomic: written to a temp file beside the real one and moved into
place. os.replace is atomic on the same filesystem, so an interrupt leaves
either the old workbook or the new one and never a truncated one. This matters
more here than in venire, because the outcome column is the only record of
which rap sheets are already filed in a production legal system.

The outcome column doubles as the resume record. venire needs a separate
progress.txt because it pre-fills every row with an error outcome before the
loop starts, so a filled cell cannot mean "done". Nothing is pre-filled here,
so a blank cell means not yet attempted and a run picks up where it stopped.
"""

import os
from pathlib import Path

import openpyxl

from exceptions import ConfigProblem, SystemProblem
from logger import get_logger

logger = get_logger(__name__)

# PARSE CONSTANTS, from venire's excel_handler
MAX_SPLIT_PARAMETER = 1
FIRST_NAME_SPLIT_PARAMETER = 0
COMMA_DELIMITER = ","

# Generational suffixes, which the sheet puts in the surname field:
#
# They are stripped for the CCIS search and kept everywhere else. CCIS indexes
# the surname on its own, so searching it for a last name of "DOE III"
# finds nothing, and the row would then be recorded as a compound-surname
# warning when the real surname is one plain word. The full cell is still what
# STAC's defendant name is checked against, where the suffix is harmless: the
# name check accepts one name's words being a subset of the other's.
#
# Only a trailing suffix is removed, and only when something is left over, so
# a genuine compound surname like "LOPEZ CRUZ" is untouched and still gets the
# warning it deserves.
NAME_SUFFIXES = frozenset({
    "JR", "JR.", "SR", "SR.", "I", "II", "III", "IV", "V", "VI", "VII",
})


class Outcome:
    """What happened to one row, written into the outcome column.

    A blank cell means the row has not been attempted, which is what makes
    resuming work. Every value here is a terminal state: a run that sees one
    skips that row.

    CHECK_BY_HAND is the dangerous one. It means Save was pressed and STAC
    never confirmed, so the rap sheet may or may not be on the case. A rerun
    must not touch that row, because filing it twice is the other way to get
    this wrong.
    """

    FILED = "Filed in STAC"
    REACHED_SAVE = "Reached save - not pressed"
    REHEARSED = "Rehearsed - nothing uploaded"
    NO_CCIS_RECORD = "No matches found in CCIS"
    COMPOUND_NAME = "Warning - compound last name, check manually"
    NO_UCN = "No case number in this row"
    BAD_FORMAT = "Warning - could not read this row, check manually"
    STAC_PROBLEM = "Not filed - check manually"
    CHECK_BY_HAND = "CHECK BY HAND - save unconfirmed, do not re-run"
    ERROR = "Error - check manually"

    #: Outcomes that mean the row is finished and must not be retried.
    TERMINAL = frozenset()


# Filled after the class body so the values can reference each other.
Outcome.TERMINAL = frozenset({
    Outcome.FILED,
    Outcome.REACHED_SAVE,
    Outcome.REHEARSED,
    Outcome.NO_CCIS_RECORD,
    Outcome.COMPOUND_NAME,
    Outcome.NO_UCN,
    Outcome.BAD_FORMAT,
    Outcome.STAC_PROBLEM,
    Outcome.CHECK_BY_HAND,
    Outcome.ERROR,
})


def parse_name(name: str) -> tuple:
    """Split 'Last, First Middle' into first and last name.

    venire's logic, with one addition: a name holding more than one comma is
    refused rather than guessed at.

    venire splits on the first comma, so 'Vale, Jr., Sam' gives a last name of
    'Vale' and a first name of 'Jr.'. In venire that costs one juror lookup. It
    is worse here, because CCIS would then be searched for a person whose first
    name is 'Jr.', find nothing, and the row would be recorded as a clean
    "no matches found" for somebody who may well have a record. A false
    negative that looks like a real answer is the one outcome worth refusing
    outright, so an ambiguous name goes to a person instead.

    Args:
        name: Full name string ex: 'Smith, John Michael'

    Returns:
        (first_name, last_name) ex: ('John', 'Smith')
        (None, None) if the name is empty, has no comma, or has more than one.

    Example:
        parse_name("Smith, John Michael")  ->  ("John", "Smith")
        parse_name("Vale, Jr., Sam")       ->  (None, None)
    """
    if not name or COMMA_DELIMITER not in name:
        return None, None

    if name.count(COMMA_DELIMITER) > MAX_SPLIT_PARAMETER:
        return None, None

    left, right = name.split(COMMA_DELIMITER, MAX_SPLIT_PARAMETER)

    last_name = left.strip()
    right = right.strip()
    if not last_name or not right:
        return None, None

    first_name = right.split()[FIRST_NAME_SPLIT_PARAMETER]

    return first_name, last_name


def strip_suffix(last_name: str) -> str:
    """Drop a trailing generational suffix, for the CCIS surname search.

    Only ever removes from the end, and never returns empty.
    """
    if not last_name:
        return last_name

    words = last_name.split()
    while len(words) > 1 and words[-1].upper().rstrip(".") in {
        suffix.rstrip(".") for suffix in NAME_SUFFIXES
    }:
        words.pop()

    return " ".join(words)


def format_dob(dob) -> str | None:
    """Format a date of birth as mm/dd/yyyy for CCIS.

    Handles both the datetime object openpyxl returns for a real date cell and
    a plain string for a text cell.

    Args:
        dob: datetime or string.

    Returns:
        '04/23/1985', or None if the value is neither.
    """
    if hasattr(dob, "strftime"):
        return dob.strftime("%m/%d/%Y")

    if isinstance(dob, str):
        return dob.strip()

    return None


def normalise_ucn(value) -> str:
    """Return the UCN as a clean string, or "" if the cell holds nothing usable.

    Excel will happily store a case number as a number, as a date, or with a
    stray space from a paste. Only the obviously-unusable cases are rejected
    here; whether the number names a real case is STAC's answer to give, and
    its "No case in STAC with case number X" is a far more useful error than
    anything this function could invent.
    """
    if value is None:
        return ""

    # A date cell means Excel reinterpreted something like 25-2026-MM as a
    # date. The original digits are gone, so there is nothing to recover.
    if hasattr(value, "strftime"):
        return ""

    # A number means Excel treated the cell as numeric. A real UCN carries
    # letters and a division code, so it is always text; a numeric one is a
    # partial or mistyped number. It is passed through rather than rejected,
    # because STAC's own "No case in STAC with case number X" names the value
    # and is more useful than a guess here, but it is worth a warning.
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        logger.warning(
            "Case number %r is stored as a number, not text. A full UCN contains "
            "letters, so this one is probably incomplete. Format that column as "
            "Text in Excel.",
            value,
        )

    return str(value).strip()


class Spreadsheet:
    """The workbook, held open for the run, with atomic per-row saves.

    Used as `with Spreadsheet(config) as sheet:` so the workbook is closed
    even when the run blows up.
    """

    def __init__(self, config: dict) -> None:
        excel = config["excel"]
        self.path = Path(excel["file"])
        self.sheet_name = excel["sheet"]
        self.dob_column = excel["dob_column"]
        self.name_column = excel["name_column"]
        self.ucn_column = excel["ucn_column"]
        self.outcome_column = excel["outcome_column"]
        self.first_data_row = excel["first_data_row"]
        # Both optional, 0 meaning the sheet has no such column. The real
        # sheet has no id column at all, and the case number is only used as
        # a fallback when the UCN finds nothing.
        self.id_column = excel.get("id_column") or 0
        self.case_number_column = excel.get("case_number_column") or 0

        self._workbook = None
        self._sheet = None

    def __enter__(self) -> "Spreadsheet":
        self.open()
        return self

    def __exit__(self, *_) -> None:
        self.close()

    def open(self) -> None:
        """Load the workbook and find the sheet.

        Raises:
            ConfigProblem: The file is gone, unreadable, or has no sheet by
                the configured name. The message lists the sheets it does
                have, since a renamed tab is the usual cause.
        """
        if not self.path.is_file():
            raise ConfigProblem(f"No spreadsheet at {self.path}.")

        try:
            self._workbook = openpyxl.load_workbook(self.path)
        except Exception as error:  # noqa: BLE001 - openpyxl raises several types
            raise ConfigProblem(
                f"Could not open {self.path}: {type(error).__name__}: {error}"
            ) from error

        if self.sheet_name not in self._workbook.sheetnames:
            available = ", ".join(self._workbook.sheetnames)
            self.close()
            raise ConfigProblem(
                f"{self.path.name} has no sheet called {self.sheet_name!r}. "
                f"It has: {available}. Set excel.sheet in config.yaml."
            )

        self._sheet = self._workbook[self.sheet_name]
        logger.info(
            "Opened %s, sheet %r, %s row(s)",
            self.path.name, self.sheet_name, self._sheet.max_row,
        )

    def close(self) -> None:
        """Close the workbook. Safe to call twice, and never raises."""
        if self._workbook is None:
            return
        try:
            self._workbook.close()
        except Exception as error:  # noqa: BLE001
            logger.warning("Workbook did not close cleanly: %s", error)
        finally:
            self._workbook = None
            self._sheet = None

    def _cell(self, row: int, column: int):
        return self._sheet.cell(row=row, column=column).value

    def load_rows(self) -> list:
        """Read every row of the sheet into a list of dicts, in sheet order.

        Nothing is written here, and nothing is skipped for being unparseable:
        a row that cannot be read still comes back, carrying `problem`, so the
        caller can write an outcome against it rather than silently passing
        over it. venire skips such rows inside the parser, which means a
        malformed row leaves no trace unless you go looking in the log.

        Returns:
            One dict per non-empty row:
                row          sheet row number
                person_id    the id column if the sheet has one, else the row
                             number, so every row has something to name a file
                             with
                name         the name cell, verbatim, for the STAC check
                first_name   parsed, for the CCIS search
                last_name    parsed surname, suffix included
                search_last  the surname with a generational suffix removed,
                             which is what CCIS is actually searched for
                dob          formatted mm/dd/yyyy
                ucn          the UCN, cleaned
                case_number  the case number, cleaned. The fallback identifier
                             when the UCN finds no case in STAC
                outcome      what is already in the outcome column, or ""
                problem      why this row cannot be processed, or ""
        """
        rows = []

        for number in range(self.first_data_row, self._sheet.max_row + 1):
            person_id = self._cell(number, self.id_column) if self.id_column else None
            dob = self._cell(number, self.dob_column)
            name = self._cell(number, self.name_column)
            ucn = self._cell(number, self.ucn_column)
            outcome = self._cell(number, self.outcome_column)
            case_number = (
                self._cell(number, self.case_number_column)
                if self.case_number_column
                else None
            )

            # Entirely empty row. Trailing blanks are normal in a sheet that
            # has had rows deleted, so these are dropped rather than reported.
            if not any((person_id, dob, name, ucn, case_number)):
                continue

            first_name, last_name = parse_name(str(name) if name else "")
            formatted_dob = format_dob(dob)
            clean_ucn = normalise_ucn(ucn)
            clean_case_number = normalise_ucn(case_number)

            problem = ""
            if not name:
                problem = "no name in this row"
            elif str(name).count(COMMA_DELIMITER) > MAX_SPLIT_PARAMETER:
                # Refused rather than guessed. See parse_name for why.
                problem = (
                    f"{str(name)!r} has more than one comma, so which part is the "
                    "surname is ambiguous. Searching CCIS for the wrong first name "
                    "would come back as a clean 'no record'. Rewrite it as "
                    "'Last, First' with a single comma."
                )
            elif not first_name or not last_name:
                problem = (
                    f"could not split {str(name)!r} into a first and last name; "
                    "it needs to read 'Last, First'"
                )
            elif not formatted_dob:
                problem = f"could not read the date of birth {dob!r}"
            elif not clean_ucn and not clean_case_number:
                # Only a problem when BOTH identifiers are unusable. Either one
                # on its own can find the case, so a blank UCN with a good case
                # number is a row that still works.
                if ucn is None and case_number is None:
                    problem = "no case number in this row"
                elif hasattr(ucn, "strftime") or hasattr(case_number, "strftime"):
                    problem = (
                        f"Excel stored the case number as a date (ucn {ucn}, case "
                        f"number {case_number}). Format those columns as Text and "
                        "paste the numbers again."
                    )
                else:
                    problem = (
                        f"could not read either identifier (ucn {ucn!r}, case "
                        f"number {case_number!r})"
                    )

            rows.append({
                "row": number,
                # The real sheet has no id column, so the row number stands in.
                # Every row needs something to name its PDF with.
                "person_id": (
                    str(person_id).strip() if person_id is not None else str(number)
                ),
                "name": str(name).strip() if name else "",
                "first_name": first_name,
                "last_name": last_name,
                "search_last": strip_suffix(last_name) if last_name else last_name,
                "dob": formatted_dob,
                "ucn": clean_ucn,
                "case_number": clean_case_number,
                "outcome": str(outcome).strip() if outcome else "",
                "problem": problem,
            })

        logger.info("Read %s row(s) with data", len(rows))
        return rows

    def write_outcome(self, row: int, outcome: str) -> None:
        """Write one row's outcome and save the workbook atomically.

        Saved per row on purpose. The alternative, saving every N rows, would
        mean a crash loses the record of rap sheets that are already in STAC,
        and a rerun would file them again. At roughly twenty seconds a row the
        save is not the slow part.

        Raises:
            SystemProblem: If the workbook cannot be saved. Carrying on would
                mean filing documents whose outcome is not being recorded, so
                this stops the run.
        """
        self._sheet.cell(row=row, column=self.outcome_column).value = outcome
        self._save()

    def _save(self) -> None:
        """Save to a temp file beside the workbook, then move it into place.

        os.replace is atomic within a filesystem, so the real file is never
        partially written. The temp file sits in the same directory
        deliberately: a different drive, or the system temp folder, would make
        the move a copy and lose the atomicity.
        """
        temp_path = self.path.with_name(f".{self.path.name}.saving")
        try:
            self._workbook.save(temp_path)
            os.replace(temp_path, self.path)
        except Exception as error:  # noqa: BLE001 - openpyxl and OS errors both
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass
            raise SystemProblem(
                f"Could not save {self.path.name}: {type(error).__name__}: {error}. "
                "Usually the file is open in Excel. Nothing further will be filed "
                "until the outcome can be recorded."
            ) from error


def is_done(row: dict) -> bool:
    """True when this row already carries a terminal outcome.

    Anything in the outcome column counts, not only the strings this project
    writes: a cell someone filled in by hand is a deliberate instruction to
    leave that row alone.
    """
    return bool(row.get("outcome"))


def clear_outcomes(sheet: Spreadsheet, rows: list) -> int:
    """Blank every outcome cell so the next run starts over. Returns the count.

    This is what --fresh does. Rows marked CHECK_BY_HAND are left alone: the
    rap sheet may already be on the case, and clearing that cell would send a
    rerun straight into filing it twice.
    """
    cleared = 0
    for row in rows:
        if not row["outcome"]:
            continue
        if row["outcome"] == Outcome.CHECK_BY_HAND:
            logger.warning(
                "Row %s is marked %r and is NOT being cleared. Check that case in "
                "STAC by hand before running it again.",
                row["row"], Outcome.CHECK_BY_HAND,
            )
            continue
        sheet._sheet.cell(row=row["row"], column=sheet.outcome_column).value = None
        row["outcome"] = ""
        cleared += 1

    if cleared:
        sheet._save()
    return cleared