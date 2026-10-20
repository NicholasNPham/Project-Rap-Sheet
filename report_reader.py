"""Reads the STAC report export and keeps a working copy that outcomes go into.

The report viewer's Export button downloads a binary Excel 97-2003 file
(.xls). openpyxl cannot open that format, so it is read with xlrd, and the
rows are copied into a .xlsx working copy beside it. The working copy is what
the run writes outcomes into, through excel_handler.Spreadsheet, so the
per-row atomic save and the outcome column work exactly as they did for the
hand-built sheet. The downloaded .xls is never modified.

The export has a title row, then a header, then one row per case:

    (blank) | UCN | OPEN Date | Name | (blank) | DOB

Columns are found by their header label, not by position, because Crystal
puts blank spacer columns between them and a report design change would move
everything along. Every UCN cell carries a hyperlink to the case in STAC. That
link is the way into the case: it points at STAC's own incident id, which
cannot be built from the UCN.

Dates of birth arrive as text in yyyy-mm-dd and are passed to CCIS as
mm/dd/yyyy. A DOB in any other shape is refused rather than passed on, because CCIS treats a malformed date as "no
record", which would be filed on the case as if it were true.
"""

from datetime import date, datetime
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit

import openpyxl

from exceptions import ConfigProblem, SystemProblem
from excel_handler import (
    COMMA_DELIMITER,
    MAX_SPLIT_PARAMETER,
    normalise_ucn,
    parse_name,
    strip_suffix,
)
from logger import get_logger

logger = get_logger(__name__)

# The export's header labels, matched without regard to case or spacing.
UCN_HEADER = "UCN"
NAME_HEADER = "NAME"
DOB_HEADER = "DOB"
OPEN_DATE_HEADER = "OPEN DATE"
REQUIRED_HEADERS = (UCN_HEADER, NAME_HEADER, DOB_HEADER)

# How far down the export to look for the header row. Crystal puts a title
# and sometimes a run date above it.
HEADER_SEARCH_ROWS = 50

# The layout CCIS is searched with.
DOB_FORMAT = "%m/%d/%Y"

# Layouts a text DOB in the export is accepted in. The export writes
# yyyy-mm-dd (ISO, so year, month, day with no ambiguity). The two
# month-first layouts are accepted too, so a report change between them does
# not flag every row. Day-first is never accepted: 01-02 would be a guess.
DOB_INPUT_FORMATS = ("%Y-%m-%d", "%m-%d-%Y", "%m/%d/%Y")

# The first eight bytes of an OLE compound file, which is what a binary .xls
# is. Checked so a changed export format fails with a clear message rather
# than an xlrd traceback.
OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

# Working copy layout. Fixed, because this module writes it.
WORKING_SHEET = "Report"
WORKING_HEADERS = ("UCN", "OPEN Date", "Name", "DOB", "Outcome")
WC_UCN_COLUMN = 1
WC_OPEN_DATE_COLUMN = 2
WC_NAME_COLUMN = 3
WC_DOB_COLUMN = 4
WC_OUTCOME_COLUMN = 5
WC_FIRST_DATA_ROW = 2

# One export cell: its value, and the link target on it ("" when none).
Cell = tuple[object, str]


# ------------------------------------------------------------ reading


def read_xls(path: Path) -> list[list[Cell]]:
    """Read the first sheet of a binary .xls into rows of (value, link).

    Args:
        path: The downloaded export.

    Returns:
        Every row of the sheet, top to bottom. A date cell comes back as a
        datetime; everything else as xlrd gives it.

    Raises:
        ConfigProblem: The file is missing, is not a binary .xls, or xlrd is
            not installed.
    """
    if not path.is_file():
        raise ConfigProblem(f"No report export at {path}.")

    with path.open("rb") as handle:
        if handle.read(len(OLE_SIGNATURE)) != OLE_SIGNATURE:
            raise ConfigProblem(
                f"{path.name} is not an Excel 97-2003 file. The report export format "
                "may have changed. Run inspect_export.py on it to see what it is."
            )

    try:
        import xlrd
    except ImportError as error:
        raise ConfigProblem(
            "Reading the report export needs xlrd. Run: pip install -r requirements.txt"
        ) from error

    try:
        book = xlrd.open_workbook(str(path), on_demand=True)
    except Exception as error:  # noqa: BLE001 - xlrd raises several types
        raise ConfigProblem(
            f"Could not open {path.name}: {type(error).__name__}"
        ) from error

    try:
        sheet = book.sheet_by_index(0)
        links = sheet.hyperlink_map
        rows = []
        for r in range(sheet.nrows):
            row = []
            for c in range(sheet.ncols):
                cell = sheet.cell(r, c)
                if cell.ctype == xlrd.XL_CELL_DATE:
                    value: object = xlrd.xldate_as_datetime(cell.value, book.datemode)
                else:
                    value = cell.value
                link = links.get((r, c))
                row.append((value, (link.url_or_path or "").strip() if link else ""))
            rows.append(row)
        return rows
    finally:
        book.release_resources()


def _label(value: object) -> str:
    """A cell as an upper-case label with its spacing collapsed, for header matching."""
    return " ".join(str(value or "").split()).upper()


def find_header(rows: list[list[Cell]]) -> tuple[int, dict[str, int]]:
    """Find the header row and the column of each label the tool reads.

    Args:
        rows: The export, as read_xls returns it.

    Returns:
        (index of the header row, {label: column index}). OPEN DATE is
        included only when the export has it.

    Raises:
        ConfigProblem: No row in the first HEADER_SEARCH_ROWS has all of
            UCN, Name and DOB. The message names the labels that row had, if
            a row with UCN was found, so a renamed column is a quick fix.
    """
    for index, row in enumerate(rows[:HEADER_SEARCH_ROWS]):
        labels = [_label(value) for value, _ in row]
        if UCN_HEADER not in labels:
            continue

        missing = [wanted for wanted in REQUIRED_HEADERS if wanted not in labels]
        if missing:
            raise ConfigProblem(
                f"The report's header row has no {', '.join(missing)} column. "
                f"It has: {', '.join(label for label in labels if label)}."
            )

        columns = {wanted: labels.index(wanted) for wanted in REQUIRED_HEADERS}
        if OPEN_DATE_HEADER in labels:
            columns[OPEN_DATE_HEADER] = labels.index(OPEN_DATE_HEADER)
        return index, columns

    raise ConfigProblem(
        f"No header row in the report export: no cell reads {UCN_HEADER!r} in the "
        f"first {HEADER_SEARCH_ROWS} rows."
    )


def normalise_dob(value: object) -> str | None:
    """Return a DOB as mm/dd/yyyy for CCIS, or None when it is not a real date.

    Reads yyyy-mm-dd, which is what the export holds, or mm-dd-yyyy or
    mm/dd/yyyy. Strict otherwise on purpose: 1/1/80 or a day-first date
    would be quietly turned into something by a lenient parser, and a wrong
    date reaches CCIS as a clean "no record".

    Args:
        value: The DOB cell: text, or a datetime if Excel stored a real date.

    Returns:
        '01/01/1980', or None.
    """
    if isinstance(value, (datetime, date)):
        return value.strftime(DOB_FORMAT)
    if not isinstance(value, str) or not value.strip():
        return None
    for layout in DOB_INPUT_FORMATS:
        try:
            return datetime.strptime(value.strip(), layout).strftime(DOB_FORMAT)
        except ValueError:
            continue
    return None


def _text(value: object) -> str:
    """A cell value as stripped text."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _cell(row: list[Cell], column: int | None) -> Cell:
    """The cell at `column`, or an empty one when the row is short."""
    if column is None or column >= len(row):
        return None, ""
    return row[column]


def build_rows(rows: list[list[Cell]], header_at: int, columns: dict[str, int]) -> list[dict]:
    """Turn the export's data rows into the row dicts the run works from.

    Same keys as excel_handler.Spreadsheet.load_rows, plus case_url and
    export_row, so process_row needs nothing new to read a row. `row` is the
    row number in the working copy, which is where outcomes are written.

    A row that cannot be searched still comes back, carrying `problem`, so
    the run writes an outcome against it instead of passing over it.

    Args:
        rows: The export, as read_xls returns it.
        header_at: Index of the header row.
        columns: Label to column index, from find_header.

    Returns:
        One dict per non-empty row below the header, in export order.
    """
    built = []
    for index in range(header_at + 1, len(rows)):
        row = rows[index]
        if not any(_text(value) for value, _ in row):
            continue

        ucn_value, case_url = _cell(row, columns[UCN_HEADER])
        name = _text(_cell(row, columns[NAME_HEADER])[0])
        dob_value = _cell(row, columns[DOB_HEADER])[0]
        open_date = _text(_cell(row, columns.get(OPEN_DATE_HEADER))[0])

        first_name, last_name = parse_name(name)
        dob = normalise_dob(dob_value)
        ucn = normalise_ucn(ucn_value)

        problem = ""
        if not name:
            problem = "no name in this row"
        elif name.count(COMMA_DELIMITER) > MAX_SPLIT_PARAMETER:
            problem = (
                f"{name!r} has more than one comma, so which part is the surname is "
                "ambiguous. Searching CCIS for the wrong first name would come back "
                "as a clean 'no record'."
            )
        elif not first_name or not last_name:
            problem = f"could not split {name!r} into a first and last name"
        elif not dob:
            problem = (
                f"could not read the date of birth {_text(dob_value)!r}; it needs "
                "to be yyyy-mm-dd or mm-dd-yyyy"
            )
        elif not ucn:
            problem = "no UCN in this row"

        built.append({
            "row": WC_FIRST_DATA_ROW + len(built),
            "export_row": index + 1,
            "person_id": str(WC_FIRST_DATA_ROW + len(built)),
            "name": name,
            "first_name": first_name,
            "last_name": last_name,
            "search_last": strip_suffix(last_name) if last_name else last_name,
            "dob": dob,
            "ucn": ucn,
            # The export has no separate case number; the UCN is the only
            # identifier, and the link is how the case is reached.
            "case_number": "",
            "case_url": case_url,
            "open_date": open_date,
            "outcome": "",
            "problem": problem,
        })

    logger.info("Read %s row(s) from the report export", len(built))
    return built


# ------------------------------------------------------ link retargeting


def retarget_link(url: str, report_base: str, filing_base: str) -> str:
    """Point a report link at the STAC instance the run files into.

    The report is run on prod STAC, so its links name prod. Filing goes to
    the instance in stac.url, which can be test STAC, and test holds the same
    incident ids as prod, so swapping the scheme and host is enough to reach
    the same case there. The defendant check on the case page still runs, so
    an id that does not line up is caught rather than filed on.

    Only a link on the report's own host is changed. Anything else comes back
    as it was, and the host check in StacSession.link_is_ours then refuses to
    open it.

    Args:
        url: The link from the report row. May be "".
        report_base: The STAC url the report was run on.
        filing_base: The STAC url the run files into.

    Returns:
        The link on the filing host, or `url` unchanged.
    """
    if not url:
        return url
    link = urlsplit(url)
    source = urlsplit(report_base)
    target = urlsplit(filing_base)
    if link.netloc.lower() != source.netloc.lower():
        return url
    return urlunsplit((target.scheme, target.netloc, link.path, link.query, link.fragment))


# ------------------------------------------------------- working copy


def working_copy_path(export_path: Path) -> Path:
    """Where the working copy of an export goes: beside it, as .xlsx."""
    return export_path.with_suffix(".xlsx")


def read_previous_outcomes(path: Path) -> dict[str, str]:
    """Outcomes already in an existing working copy, keyed by UCN.

    This is what stops a run that is restarted on the same export from
    filing a case twice. A case filed before the restart still has its
    outcome here even though the .xls it came from has not changed.

    Args:
        path: The working copy. Need not exist.

    Returns:
        {UCN: outcome} for every row that has both. Empty if there is no
        working copy yet.

    Raises:
        SystemProblem: The working copy exists but cannot be read. Starting
            over without its outcomes could file cases twice, so the run
            stops rather than guess.
    """
    if not path.is_file():
        return {}

    try:
        book = openpyxl.load_workbook(path, read_only=True)
    except Exception as error:  # noqa: BLE001 - openpyxl raises several types
        raise SystemProblem(
            f"Could not read the existing working copy {path.name} "
            f"({type(error).__name__}). It holds which cases are already filed, so "
            "the run will not start over without it. Close it in Excel, or move it "
            "aside if you are sure nothing in it was filed."
        ) from error

    try:
        sheet = book[WORKING_SHEET] if WORKING_SHEET in book.sheetnames else book.worksheets[0]
        outcomes = {}
        for values in sheet.iter_rows(min_row=WC_FIRST_DATA_ROW, values_only=True):
            ucn = _text(values[WC_UCN_COLUMN - 1]) if len(values) >= WC_UCN_COLUMN else ""
            outcome = (
                _text(values[WC_OUTCOME_COLUMN - 1])
                if len(values) >= WC_OUTCOME_COLUMN else ""
            )
            if ucn and outcome:
                outcomes[ucn] = outcome
        return outcomes
    finally:
        book.close()


def write_working_copy(path: Path, rows: list[dict]) -> None:
    """Write the working copy: one row per case, the link kept on the UCN cell.

    Written to a temp file and moved into place, the same as every save in
    excel_handler, so an interrupt never leaves half a file.

    Args:
        path: Where the working copy goes.
        rows: From build_rows, with any carried-over outcomes already set.

    Raises:
        SystemProblem: The file cannot be written, usually because it is
            open in Excel.
    """
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = WORKING_SHEET
    sheet.append(list(WORKING_HEADERS))

    for row in rows:
        number = row["row"]
        sheet.cell(row=number, column=WC_UCN_COLUMN, value=row["ucn"] or None)
        if row["case_url"]:
            sheet.cell(row=number, column=WC_UCN_COLUMN).hyperlink = row["case_url"]
        sheet.cell(row=number, column=WC_OPEN_DATE_COLUMN, value=row["open_date"] or None)
        sheet.cell(row=number, column=WC_NAME_COLUMN, value=row["name"] or None)
        # Stored as text so Excel does not turn it back into a date serial.
        sheet.cell(row=number, column=WC_DOB_COLUMN, value=row["dob"] or None)
        sheet.cell(row=number, column=WC_OUTCOME_COLUMN, value=row["outcome"] or None)

    temp_path = path.with_name(f".{path.name}.saving")
    try:
        book.save(temp_path)
        temp_path.replace(path)
    except OSError as error:
        temp_path.unlink(missing_ok=True)
        raise SystemProblem(
            f"Could not write the working copy {path.name}: {error}. Usually it is "
            "open in Excel."
        ) from error


def working_copy_config(path: Path) -> dict:
    """A config fragment that points excel_handler.Spreadsheet at the working copy.

    Lets the run reuse Spreadsheet.write_outcome, with its per-row atomic
    save, instead of a second implementation of the same thing.
    """
    return {
        "excel": {
            "file": str(path),
            "sheet": WORKING_SHEET,
            "ucn_column": WC_UCN_COLUMN,
            "name_column": WC_NAME_COLUMN,
            "dob_column": WC_DOB_COLUMN,
            "outcome_column": WC_OUTCOME_COLUMN,
            "first_data_row": WC_FIRST_DATA_ROW,
            "case_number_column": 0,
            "id_column": 0,
        }
    }


def prepare_report(export_path: Path) -> tuple[Path, list[dict]]:
    """Read the export, carry over earlier outcomes, and write the working copy.

    Args:
        export_path: The downloaded .xls.

    Returns:
        (working copy path, rows). Rows that already had an outcome in an
        earlier working copy carry it, so the run skips them.

    Raises:
        ConfigProblem: The export cannot be read or has no usable header.
        SystemProblem: The working copy cannot be read or written.
    """
    cells = read_xls(export_path)
    header_at, columns = find_header(cells)
    rows = build_rows(cells, header_at, columns)

    target = working_copy_path(export_path)
    previous = read_previous_outcomes(target)
    carried = 0
    for row in rows:
        outcome = previous.get(row["ucn"])
        if outcome:
            row["outcome"] = outcome
            carried += 1
    if carried:
        logger.info(
            "%s row(s) already have an outcome from an earlier run on this export",
            carried,
        )

    write_working_copy(target, rows)
    logger.info("Working copy: %s", target.name)
    return target, rows
