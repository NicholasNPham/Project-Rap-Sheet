"""RAP SHEET entry point.

One pass down the spreadsheet. For each row: search CCIS by name and date of
birth, capture the case summary as a PDF, then file that PDF on the row's case
number in STAC under one Type/Subtype.

Two sessions, both opened once and held for the whole run: CCIS in one Chrome,
STAC in another. They are driven one after the other, not at the same time.
Overlapping them would save the four seconds CCIS takes while STAC is being
prepared, and would cost a speculative upload box opened on every row that
turns out to have no CCIS record at all. PROJECT-DALYN already has a function
for cleaning up exactly that mess, and its comment says why: an unsaved upload
bleeds into the next document.

The outcome column is the record of what happened and the resume marker. A row
with anything in that cell is left alone, so a run that stops halfway picks up
where it left off and never files the same rap sheet twice.

Nothing reaches STAC until config says so. stac.upload_enabled and
stac.save_enabled are both false by default: a first run finds every case,
checks every name and picks the Type/Subtype without sending a single file.
"""

import argparse
import re
import sys
from collections import Counter
from datetime import date
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent

# The modules live in src/ in the documented layout, but work equally well
# sitting flat beside main.py. Both are put on the path so either arrangement
# imports, and neither needs the files moved to run.
for _candidate in (PROJECT_ROOT / "src", PROJECT_ROOT):
    if _candidate.is_dir() and str(_candidate) not in sys.path:
        sys.path.insert(0, str(_candidate))

# src/ modules import each other flat (from exceptions import ...), so the
# path insert above has to run before any of these.
import excel_handler  # noqa: E402
from ccis import CcisSession  # noqa: E402
from config_loader import load_config  # noqa: E402
from exceptions import ConfigProblem, RowProblem, SystemProblem  # noqa: E402
from excel_handler import Outcome, Spreadsheet  # noqa: E402
from logger import get_logger, setup_logging  # noqa: E402
from stac import SaveMayHaveHappened, StacRunner  # noqa: E402

logger = get_logger("main")

EXIT_OK = 0
EXIT_SYSTEM_PROBLEM = 1
EXIT_CONFIG_PROBLEM = 2

# How many rows in a row may fail with a SystemProblem before the run gives up.
# One case with a page that will not settle should not end a run; STAC or CCIS
# being down should. Reset by any row that gets through.
MAX_CONSECUTIVE_FAILURES = 5

# Characters that cannot go in a filename on Windows, plus the ones that make
# a path ambiguous anywhere. Names in this spreadsheet carry apostrophes,
# full stops and the occasional slash (O'BRIEN, JR., SMITH/JONES), and venire
# writes them into a path unexamined.
UNSAFE_IN_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _safe_piece(value: str) -> str:
    """Make one part of a filename safe, without losing which person it is.

    Trailing dots and commas are stripped as well as the illegal characters:
    Windows refuses a filename ending in a full stop, which 'Vale, Jr.' would
    otherwise produce. Apostrophes and hyphens are kept, since O'Brien and
    Mary-Jane are legal in a filename and dropping them makes the file harder
    to match back to its row.
    """
    cleaned = UNSAFE_IN_FILENAME.sub("", str(value or ""))
    cleaned = re.sub(r"\s+", "_", cleaned.strip())
    cleaned = cleaned.strip("._,")
    return cleaned or "unknown"


NO_RECORD_FILENAME_TAG = "NO_RECORD"


def pdf_path_for(folder: Path, row: dict, has_record: bool = True) -> Path:
    """Where this row's PDF goes.

    Named the way venire names them, person id first, so the folder sorts in
    spreadsheet order and a file can be matched back to a row by eye. A
    no-results PDF carries a NO_RECORD tag so the two kinds are told apart in
    the folder without opening them.

    Args:
        folder: This run's PDF folder.
        row: One dict from Spreadsheet.load_rows.
        has_record: False when the PDF is the CCIS no-results page.

    Returns:
        The path to write the PDF to.
    """
    pieces = [
        _safe_piece(row["person_id"]),
        _safe_piece(row["last_name"]),
        _safe_piece(row["first_name"]),
    ]
    if not has_record:
        pieces.append(NO_RECORD_FILENAME_TAG)
    return folder / f"{'_'.join(pieces)}.pdf"


def _describe(row: dict) -> str:
    """How a row is named in the log: enough to find it, nothing extra."""
    identifier = row["ucn"] or row["case_number"] or "no case number"
    return f"row {row['row']} ({row['last_name']}, {row['first_name']} / {identifier})"


def process_row(row: dict, ccis: CcisSession, runner: StacRunner, pdf_folder: Path) -> str:
    """Take one row from CCIS to STAC and return the outcome to write.

    This is the whole per-row sequence, and it is one function on purpose: it
    is the only place that knows CCIS runs before STAC, so it is the only
    place that would change if the two were ever overlapped.

    Args:
        row: One dict from Spreadsheet.load_rows.
        ccis: The open CCIS session.
        runner: The open STAC runner.
        pdf_folder: Where this run's PDFs are written.

    Returns:
        One of the Outcome strings.

    Raises:
        SystemProblem: CCIS or STAC is broken, rather than this row being odd.
            The caller counts these and stops the run when they pile up.
    """
    # CCIS first. Every search yields a PDF, either the case summary or the
    # printed no-results page, and either one goes on to STAC.
    if row["search_last"] != row["last_name"]:
        logger.info(
            "%s: searching CCIS for surname %r, with the suffix from %r dropped",
            _describe(row), row["search_last"], row["last_name"],
        )

    try:
        result = ccis.fetch_pdf(row["first_name"], row["search_last"], row["dob"])
    except RowProblem as error:
        logger.warning("%s: CCIS problem, %s", _describe(row), error)
        return Outcome.ERROR

    has_record = result.has_record
    if not has_record and len(row["search_last"].split()) >= 2:
        # venire's insight, kept as a warning: a compound last name that
        # returns nothing may be indexed under only part of the surname. The
        # no-results PDF is filed regardless; this is for whoever reviews the log.
        logger.warning(
            "%s: no CCIS record, but %r is a compound last name and may be "
            "indexed under part of it. Filing the no-results PDF anyway.",
            _describe(row), row["search_last"],
        )

    # Selenium uploads from a path, so the bytes have to land on disk.
    pdf_folder.mkdir(parents=True, exist_ok=True)
    pdf_file = pdf_path_for(pdf_folder, row, has_record)
    pdf_file.write_bytes(result.pdf_bytes)
    logger.info(
        "%s: %s captured, %s KB -> %s",
        _describe(row),
        "CCIS record" if has_record else "no-results page",
        len(result.pdf_bytes) // 1024, pdf_file.name,
    )

    # Then STAC. The PDF stays on disk either way: it is the audit trail, and
    # the thing somebody would file by hand if this row needs it.
    try:
        runner.enter_row(row["ucn"], row["case_number"], row["name"], [pdf_file])
    except SaveMayHaveHappened as error:
        # The one outcome that must never be retried.
        logger.error("%s: %s", _describe(row), error)
        return Outcome.CHECK_BY_HAND
    except RowProblem as error:
        logger.warning("%s: not filed, %s", _describe(row), error)
        return Outcome.STAC_PROBLEM

    session = runner.session
    if session.save_enabled:
        return Outcome.FILED if has_record else Outcome.FILED_NO_RECORD
    if session.upload_enabled:
        return Outcome.REACHED_SAVE if has_record else Outcome.REACHED_SAVE_NO_RECORD
    return Outcome.REHEARSED if has_record else Outcome.REHEARSED_NO_RECORD


def check(config: dict) -> int:
    """Report what the config resolved to and what the spreadsheet holds, then stop.

    Opens no browser and signs in to nothing, so it is safe to run against a
    live config at any time. This is the first thing to run after filling in
    config.yaml: it answers where the spreadsheet is being looked for, whether
    the columns line up, and how many rows would actually be attempted, none of
    which needs CCIS or STAC to be reachable.
    """
    stac = config["stac"]
    excel = config["excel"]
    paths = config["paths"]

    logger.info("=== config ===")
    logger.info("  CCIS url            %s", config["ccis"]["url"])
    logger.info("  CCIS login          loaded from Credential Manager")
    logger.info("  STAC url            %s", stac["url"])
    logger.info("  STAC login          loaded from Credential Manager")
    logger.info("  test instance       %s", stac["is_test_instance"])
    logger.info("  Type/Subtype        %s", stac["document_type"] + "/" + stac["subtype"]
                if stac.get("pair_is_set") else "NOT SET")
    logger.info("  uploads enabled     %s", stac["upload_enabled"])
    logger.info("  saves enabled       %s", stac["save_enabled"])
    logger.info("  name check          %s", stac["check_name"])
    logger.info("  fresh browser       %s", stac["fresh_browser"])

    logger.info("=== paths ===")
    logger.info("  spreadsheet         %s", excel["file"])
    logger.info("  PDFs                %s", Path(paths["pdfs"]) / str(date.today()))
    logger.info("  logs                %s", paths["logs"])
    logger.info("  chromedriver        %s", paths["chromedriver"] or "(let Selenium find one)")

    logger.info("=== spreadsheet ===")
    logger.info(
        "  sheet %r, data from row %s", excel["sheet"], excel["first_data_row"]
    )
    for label, key in (
        ("case number", "case_number_column"),
        ("ucn", "ucn_column"),
        ("name", "name_column"),
        ("dob", "dob_column"),
        ("outcome", "outcome_column"),
        ("id", "id_column"),
    ):
        number = excel.get(key) or 0
        where = f"column {chr(ord('A') + number - 1)}" if number else "not in this sheet"
        logger.info("    %-12s %s", label, where)

    with Spreadsheet(config) as sheet:
        rows = sheet.load_rows()

        done = [row for row in rows if excel_handler.is_done(row)]
        problems = [row for row in rows if not excel_handler.is_done(row) and row["problem"]]
        todo = [row for row in rows if not excel_handler.is_done(row) and not row["problem"]]

        logger.info("  %s row(s) with data", len(rows))
        logger.info("  %s would be attempted", len(todo))
        logger.info("  %s already carry an outcome and would be skipped", len(done))
        logger.info("  %s cannot be read and would be flagged", len(problems))

        if problems:
            logger.info("=== rows that would be flagged ===")
            for row in problems:
                logger.warning("  row %s: %s", row["row"], row["problem"])

        if todo:
            # The first few, so the column mapping can be eyeballed. If the
            # name and the case number have landed in each other's columns,
            # this is where it shows.
            logger.info("=== first rows that would be attempted ===")
            for row in todo[:5]:
                ccis_note = (
                    f" (CCIS surname {row['search_last']!r})"
                    if row["search_last"] != row["last_name"]
                    else ""
                )
                logger.info(
                    "  row %s | %s, %s%s | dob %s | ucn %s | case %s",
                    row["row"], row["last_name"], row["first_name"], ccis_note,
                    row["dob"], row["ucn"] or "-", row["case_number"] or "-",
                )
            if len(todo) > 5:
                logger.info("  ... and %s more", len(todo) - 5)

    logger.info(
        "Config and spreadsheet are readable. Nothing was signed in to and nothing "
        "was changed."
    )
    return EXIT_OK


def run(config: dict, limit: int | None, fresh: bool, only_row: int | None) -> int:
    """One pass down the spreadsheet.

    Raises:
        SystemProblem: CCIS or STAC could not be opened, the spreadsheet could
            not be saved, or MAX_CONSECUTIVE_FAILURES rows failed in a row.
            Every outcome written before that point is already in the file.
    """
    stac = config["stac"]
    pdf_folder = Path(config["paths"]["pdfs"]) / str(date.today())

    if not stac["is_test_instance"]:
        logger.warning(
            "LIVE STAC: %s. is_test_instance is false, so nothing is holding this "
            "back from real cases.", stac["url"],
        )
    if not stac["upload_enabled"]:
        logger.info(
            "Rehearsal: stac.upload_enabled is false, so no file will reach STAC. "
            "Every case is still found and every name still checked, which is what "
            "this pass is for."
        )
        if not stac.get("pair_is_set"):
            logger.info(
                "No Type/Subtype set, which is fine while uploads are off: it is "
                "only read when a file is actually being filed."
            )
    else:
        logger.info(
            "Filing everything under %s/%s", stac["document_type"], stac["subtype"]
        )
        if not stac["save_enabled"]:
            logger.info(
                "Uploads ON, saves OFF: files will reach STAC's server but Save "
                "will not be pressed."
            )

    with Spreadsheet(config) as sheet:
        rows = sheet.load_rows()

        if fresh:
            cleared = excel_handler.clear_outcomes(sheet, rows)
            logger.info("--fresh cleared %s outcome cell(s)", cleared)

        if only_row is not None:
            rows = [row for row in rows if row["row"] == only_row]
            if not rows:
                raise SystemProblem(
                    f"--only-row {only_row} does not name a row with data in it."
                )
            logger.info("--only-row %s: one row this pass", only_row)

        # Rows that cannot be processed get their outcome written now, before
        # either browser opens. There is no point signing in to two systems to
        # discover the sheet has a column problem.
        todo = []
        for row in rows:
            if excel_handler.is_done(row):
                continue
            if row["problem"]:
                outcome = (
                    Outcome.NO_UCN
                    if "case number in this row" in row["problem"]
                    else Outcome.BAD_FORMAT
                )
                logger.warning("Row %s: %s", row["row"], row["problem"])
                sheet.write_outcome(row["row"], outcome)
                row["outcome"] = outcome
                continue
            todo.append(row)

        already_done = sum(1 for row in rows if excel_handler.is_done(row))
        if limit is not None:
            todo = todo[:limit]
            logger.info("--limit %s: %s row(s) this pass", limit, len(todo))

        logger.info(
            "%s row(s) to do, %s already carry an outcome", len(todo), already_done
        )

        if not todo:
            logger.info("Nothing to do. Use --fresh to clear the outcome column.")
            return EXIT_OK

        failures = 0
        with CcisSession(config) as ccis, StacRunner(config) as runner:
            for position, row in enumerate(todo, start=1):
                logger.info("--- %s of %s | %s", position, len(todo), _describe(row))

                try:
                    outcome = process_row(row, ccis, runner, pdf_folder)
                    failures = 0
                except SystemProblem as error:
                    failures += 1
                    logger.error(
                        "%s: stopped on a system problem (%s in a row): %s",
                        _describe(row), failures, error,
                    )
                    outcome = Outcome.ERROR

                # Written before the next row starts, so an interrupt leaves a
                # correct record of exactly how far the run got.
                sheet.write_outcome(row["row"], outcome)
                row["outcome"] = outcome
                logger.info("%s: %s", _describe(row), outcome)

                if failures >= MAX_CONSECUTIVE_FAILURES:
                    raise SystemProblem(
                        f"{failures} rows in a row failed. Stopping rather than "
                        "working down the rest of the sheet against something that "
                        "is not answering."
                    )

        _summarise(todo)

    return EXIT_OK


def _summarise(rows: list) -> None:
    """Count what happened, and say plainly what still needs a person."""
    counts = Counter(row["outcome"] for row in rows)
    logger.info("Rows this pass: %s", len(rows))
    for outcome, count in counts.most_common():
        logger.info("  %-45s %s", outcome, count)

    unconfirmed = counts.get(Outcome.CHECK_BY_HAND, 0)
    if unconfirmed:
        logger.error(
            "%s row(s) are marked %r. Save was pressed and STAC never confirmed, so "
            "those rap sheets may or may not be on their cases. Check each one by "
            "hand. A rerun will NOT retry them and --fresh will NOT clear them.",
            unconfirmed, Outcome.CHECK_BY_HAND,
        )

    needs_a_person = sum(
        counts.get(outcome, 0)
        for outcome in (
            Outcome.STAC_PROBLEM,
            Outcome.COMPOUND_NAME,
            Outcome.BAD_FORMAT,
            Outcome.NO_UCN,
            Outcome.ERROR,
        )
    )
    if needs_a_person:
        logger.info(
            "%s row(s) need a look. Filter the outcome column for anything that is "
            "not %r.", needs_a_person, Outcome.FILED,
        )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="File CCIS case summaries onto their STAC cases, row by row."
    )
    parser.add_argument(
        "--limit",
        type=int,
        help="Do only the first N rows that still need doing. For testing.",
    )
    parser.add_argument(
        "--fresh",
        action="store_true",
        help="Clear the outcome column first and start over. Rows marked "
             "'CHECK BY HAND' are left alone.",
    )
    parser.add_argument(
        "--only-row",
        type=int,
        metavar="N",
        help="Do one spreadsheet row by its row number, whatever its outcome says.",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Report what the config resolved to and what the spreadsheet holds, "
             "then stop. Opens no browser and signs in to nothing. Run this first.",
    )
    parser.add_argument(
        "--config",
        type=Path,
        help="Path to a config file other than config/config.yaml.",
    )
    args = parser.parse_args(argv)

    if args.limit is not None and args.limit < 1:
        parser.error("--limit must be at least 1")
    if args.only_row is not None and args.only_row < 1:
        parser.error("--only-row must be at least 1")
    if args.only_row is not None and args.limit is not None:
        parser.error("--only-row and --limit do not make sense together")

    # Logging is not set up yet, because the log folder comes from config.
    try:
        config = load_config(args.config)
    except ConfigProblem as error:
        print(f"Config problem: {error}", file=sys.stderr)
        return EXIT_CONFIG_PROBLEM

    log_path = Path(config["paths"]["logs"]) / f"{date.today()}.log"
    setup_logging(log_path)

    try:
        if args.check:
            return check(config)
        return run(config, args.limit, args.fresh, args.only_row)
    except ConfigProblem as error:
        logger.error("Config problem: %s", error)
        return EXIT_CONFIG_PROBLEM
    except SystemProblem as error:
        logger.error("Stopped: %s", error)
        return EXIT_SYSTEM_PROBLEM
    except KeyboardInterrupt:
        logger.warning(
            "Interrupted. Every row already attempted has its outcome in the "
            "spreadsheet, so the next run resumes from here."
        )
        return EXIT_SYSTEM_PROBLEM
    except Exception:
        logger.exception("Unexpected error. This is a bug, not a row problem.")
        return EXIT_SYSTEM_PROBLEM


if __name__ == "__main__":
    sys.exit(main())