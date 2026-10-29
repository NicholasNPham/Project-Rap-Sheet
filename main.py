"""RAP SHEET entry point.

One pass over the STAC AutoCCIS report. Sign in to the report's STAC
(report.url, prod), run the report, download it as Excel, and close that
browser. Then, signed in to the filing STAC (stac.url, which can be test),
for each case on the report: search CCIS by name and date of birth, capture
the case summary (or the printed no-results page) as a PDF, open the case
from the report's link in a new tab, and file the PDF under one Type/Subtype.
The links name the report's STAC; test holds the same incident ids, so they
are pointed at the filing STAC before use.

Two sessions for the filing, both opened once and held for the whole run:
STAC in one Chrome, CCIS in another. They are driven one after the other, not at the same
time. Overlapping them would save the few seconds CCIS takes while STAC is
being prepared, and would cost a speculative upload box opened on every row
that turns out to need a person.

A filed case drops off the report, so a rerun after a crash fetches the report
again and finds only what is left. Each export gets a .xlsx working copy that
holds the outcome of every row, so the end of a run says exactly which cases
need a person. --report reruns on an export already downloaded; outcomes in
its working copy are kept, so nothing filed is filed twice. --watch keeps
running, with a fresh pass every 30 minutes through the working day.

Nothing reaches STAC until config says so. stac.upload_enabled and
stac.save_enabled are both false by default: a first run finds every case,
checks every name and picks the Type/Subtype without sending a single file.
"""

import argparse
import re
import sys
import time
from collections import Counter
from datetime import date, datetime, timedelta
from datetime import time as clock
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
import report_reader  # noqa: E402
from ccis import CcisSession  # noqa: E402
from config_loader import load_config  # noqa: E402
from exceptions import ConfigProblem, RowProblem, SystemProblem  # noqa: E402
from excel_handler import Outcome, Spreadsheet, search_variants  # noqa: E402
from logger import get_logger, hide_names, hide_values, setup_logging  # noqa: E402
from report_fetch import fetch_report  # noqa: E402
from stac import SaveMayHaveHappened, StacRunner, StacSession  # noqa: E402

logger = get_logger("main")

EXIT_OK = 0
EXIT_SYSTEM_PROBLEM = 1
EXIT_CONFIG_PROBLEM = 2
# A pass pressed Save on some row and STAC never confirmed it. A person has to
# look before anything runs again, so --watch stops on this.
EXIT_CHECK_BY_HAND = 3

# --watch runs a pass every WATCH_EVERY_MINUTES from WATCH_START to WATCH_END
# inclusive, on these weekdays (Monday is 0). A pass that runs past the next
# slot makes it skip that slot rather than start a second pass late.
WATCH_WEEKDAYS = frozenset(range(5))
WATCH_START = clock(7, 30)
WATCH_END = clock(18, 0)
WATCH_EVERY_MINUTES = 30

# How often --watch looks at the clock while waiting. Short sleeps rather than
# one long one, so a PC that wakes from sleep catches up on its own.
WATCH_POLL_SECONDS = 30

# How many rows in a row may fail with a SystemProblem before the run gives up.
# One case with a page that will not settle should not end a run; STAC or CCIS
# being down should. Reset by any row that gets through.
MAX_CONSECUTIVE_FAILURES = 5

# How many rows --check lists so the column mapping can be eyeballed.
CHECK_SAMPLE_ROWS = 5

# Characters that cannot go in a filename on Windows, plus the ones that make
# a path ambiguous anywhere. Names carry apostrophes, full stops and the
# occasional slash (O'BRIEN, JR., SMITH/JONES), and venire writes them into a
# path unexamined.
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
    report order and a file can be matched back to a row by eye. A no-results
    PDF carries a NO_RECORD tag so the two kinds are told apart in the folder
    without opening them.

    Args:
        folder: This run's PDF folder.
        row: One row dict from report_reader.
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
    """How a row is named in the log: its row number and UCN, never the person's name."""
    identifier = row["ucn"] or row["case_number"] or "no case number"
    return f"row {row['row']} ({identifier})"


def _dob_forms(dob: str | None) -> list[str]:
    """Every way a row's date of birth may be written in a log line.

    Args:
        dob: As report_reader leaves it, mm/dd/yyyy, or None.

    Returns:
        mm/dd/yyyy, yyyy-mm-dd and mm-dd-yyyy, or just `dob` if it is not in
        the expected shape, or nothing if it is empty.
    """
    if not dob:
        return []
    try:
        month, day, year = dob.split("/")
    except ValueError:
        return [dob]
    return [dob, f"{year}-{month}-{day}", f"{month}-{day}-{year}"]


def _hide_from_logs(rows: list[dict]) -> None:
    """Register every row's names and date of birth so the logs never show them.

    Done as soon as the export is read, before any line about a row is
    written. The log lines themselves name rows by row number and UCN; this
    is the backstop for error messages, filenames and anything else that
    carries a name without being asked to.

    Args:
        rows: Row dicts from report_reader.
    """
    hide_names(
        part for row in rows
        for part in (row["name"], row["first_name"], row["last_name"], row["search_last"])
    )
    hide_values(form for row in rows for form in _dob_forms(row["dob"]))


def report_session_config(config: dict) -> dict:
    """A copy of config that points a StacSession at the report's STAC, read only.

    The report is run on prod STAC. That session only ever runs the report
    and downloads it, and the switches are forced off here as well, so even
    a future change that handed it to the filing code could not upload or
    save on prod.

    Args:
        config: From load_config.

    Returns:
        A new config dict. The original is not modified.
    """
    stac = dict(config["stac"])
    stac.update(
        url=config["report"]["url"],
        upload_enabled=False,
        save_enabled=False,
        is_test_instance=False,
    )
    return {**config, "stac": stac}


def _fetch_from_report_stac(config: dict) -> Path:
    """Sign in to the report's STAC, download the report, and close that browser.

    Closed before the filing STAC is opened, so there is never a prod window
    sitting open while files are being uploaded.
    """
    logger.info("Fetching the report from %s (read only)", config["report"]["url"])
    with StacSession(report_session_config(config)) as session:
        return fetch_report(
            session,
            config["paths"]["reports"],
            config["report"]["report_timeout"],
            config["report"]["download_timeout"],
        )


def _problem_outcome(problem: str) -> str:
    """The outcome to record for a row the reader could not use."""
    return Outcome.NO_UCN if problem.startswith("no UCN") else Outcome.BAD_FORMAT


def process_row(row: dict, ccis: CcisSession, runner: StacRunner, pdf_folder: Path) -> str:
    """Take one row from CCIS to STAC and return the outcome to write.

    This is the whole per-row sequence, and it is one function on purpose: it
    is the only place that knows CCIS runs before STAC, so it is the only
    place that would change if the two were ever overlapped.

    Args:
        row: One row dict from report_reader.
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
        logger.info("%s: searching CCIS with the surname's suffix dropped", _describe(row))

    # venire's insight: a compound name that returns nothing may be indexed
    # under only part of it. So a miss is retried with the other variants, and
    # the first hit is the one filed. The fallback only matters for a row
    # built without a parseable name, which the readers refuse before here.
    variants = search_variants(row["name"]) or [(row["first_name"], row["search_last"])]
    first_miss = None
    for matched_on, (first, last) in enumerate(variants):
        if matched_on:
            logger.info(
                "%s: no CCIS record yet, trying name variant %d of %d",
                _describe(row), matched_on + 1, len(variants),
            )
        # Counted before the search, so one that fails still counts. Read by
        # _summarise to show whether the extra searches are finding anything.
        row["ccis_searches"] = matched_on + 1
        try:
            result = ccis.fetch_pdf(first, last, row["dob"])
        except RowProblem as error:
            # Even after earlier misses: "no record" is only filed once every
            # variant has actually been searched.
            logger.warning("%s: CCIS problem, %s", _describe(row), error)
            return Outcome.ERROR
        if result.has_record:
            break
        if first_miss is None:
            first_miss = result
    else:
        # Every variant missed. The full-name page is the one filed, because it
        # shows the name exactly as the sheet has it.
        result = first_miss
        if len(variants) > 1:
            logger.info(
                "%s: no CCIS record under any of %d name variants",
                _describe(row), len(variants),
            )

    has_record = result.has_record
    if has_record and matched_on:
        logger.info(
            "%s: CCIS record found on name variant %d of %d, after the full name missed",
            _describe(row), matched_on + 1, len(variants),
        )

    # Selenium uploads from a path, so the bytes have to land on disk.
    pdf_folder.mkdir(parents=True, exist_ok=True)
    pdf_file = pdf_path_for(pdf_folder, row, has_record)
    pdf_file.write_bytes(result.pdf_bytes)
    # The folder, not the filename: the filename carries the person's name.
    # It starts with the row number, which is how to find it.
    logger.info(
        "%s: %s captured, %s KB, saved in %s",
        _describe(row),
        "CCIS record" if has_record else "no-results page",
        len(result.pdf_bytes) // 1024, pdf_folder.name,
    )

    # Then STAC. The PDF stays on disk either way: it is the audit trail, and
    # the thing somebody would file by hand if this row needs it.
    try:
        runner.enter_row(
            row["ucn"], row["case_number"], row["name"], [pdf_file],
            row.get("case_url", ""),
        )
    except SaveMayHaveHappened as error:
        # Never retried from this working copy. The next report decides.
        logger.error("%s: %s", _describe(row), error)
        return Outcome.CHECK_BY_HAND
    except RowProblem as error:
        logger.warning("%s: not filed, %s", _describe(row), error)
        return Outcome.STAC_PROBLEM

    if not has_record:
        filed, reached_save, rehearsed = (
            Outcome.FILED_NO_RECORD, Outcome.REACHED_SAVE_NO_RECORD, Outcome.REHEARSED_NO_RECORD,
        )
    elif matched_on:
        filed, reached_save, rehearsed = (
            Outcome.FILED_PARTIAL_NAME, Outcome.REACHED_SAVE_PARTIAL_NAME,
            Outcome.REHEARSED_PARTIAL_NAME,
        )
    else:
        filed, reached_save, rehearsed = (
            Outcome.FILED, Outcome.REACHED_SAVE, Outcome.REHEARSED,
        )

    session = runner.session
    if session.save_enabled:
        return filed
    if session.upload_enabled:
        return reached_save
    return rehearsed


def check(config: dict, report: Path | None) -> int:
    """Report what the config resolved to and, given an export, what it holds. Then stop.

    Opens no browser, signs in to nothing and writes nothing, so it is safe
    to run at any time. Without --report it checks only the config, since
    fetching the report means signing in to STAC.
    """
    stac = config["stac"]
    paths = config["paths"]

    logger.info("=== config ===")
    logger.info("  CCIS url            %s", config["ccis"]["url"])
    logger.info("  CCIS login          loaded from Credential Manager")
    logger.info("  report from STAC    %s (read only)", config["report"]["url"])
    logger.info("  filing to STAC      %s", stac["url"])
    logger.info("  STAC login          loaded from Credential Manager")
    logger.info("  test instance       %s", stac["is_test_instance"])
    logger.info("  Type/Subtype        %s", stac["document_type"] + "/" + stac["subtype"]
                if stac.get("pair_is_set") else "NOT SET")
    logger.info("  uploads enabled     %s", stac["upload_enabled"])
    logger.info("  saves enabled       %s", stac["save_enabled"])
    logger.info("  name check          %s", stac["check_name"])
    logger.info("  fresh browser       %s", stac["fresh_browser"])

    logger.info(
        "  report timeouts     %ss to load, %ss to download",
        config["report"]["report_timeout"], config["report"]["download_timeout"],
    )

    logger.info("=== paths ===")
    logger.info("  reports             %s", paths["reports"])
    logger.info("  PDFs                %s", Path(paths["pdfs"]) / str(date.today()))
    logger.info("  logs                %s", paths["logs"])
    logger.info("  chromedriver        %s", paths["chromedriver"] or "(let Selenium find one)")

    if report is None:
        logger.info(
            "No --report given, so no export was read. Pass --report <file> to "
            "check one already downloaded."
        )
        return EXIT_OK

    logger.info("=== report export ===")
    cells = report_reader.read_xls(report)
    header_at, columns = report_reader.find_header(cells)
    rows = report_reader.build_rows(cells, header_at, columns)
    _hide_from_logs(rows)

    problems = [row for row in rows if row["problem"]]
    todo = [row for row in rows if not row["problem"]]
    linked = [row for row in todo if row["case_url"]]

    logger.info("  %s", report.name)
    logger.info("  header on row %s", header_at + 1)
    logger.info("  %s row(s) with data", len(rows))
    logger.info("  %s would be attempted", len(todo))
    logger.info("  %s of those have a case link", len(linked))
    logger.info("  %s cannot be read and would be flagged", len(problems))

    if problems:
        logger.info("=== rows that would be flagged ===")
        for row in problems:
            logger.warning("  row %s: %s", row["export_row"], row["problem"])

    if todo:
        # Names and dates of birth are left out, as everywhere in the logs.
        # To eyeball how the columns were read, inspect_export.py --show-rows
        # prints them to the terminal without logging anything.
        logger.info("=== first rows that would be attempted ===")
        for row in todo[:CHECK_SAMPLE_ROWS]:
            logger.info(
                "  row %s | ucn %s | suffix dropped for CCIS: %s | link %s",
                row["export_row"], row["ucn"],
                "yes" if row["search_last"] != row["last_name"] else "no",
                "yes" if row["case_url"] else "no",
            )
        if len(todo) > CHECK_SAMPLE_ROWS:
            logger.info("  ... and %s more", len(todo) - CHECK_SAMPLE_ROWS)

    logger.info("Config and export are readable. Nothing was signed in to and nothing was changed.")
    return EXIT_OK


def run(config: dict, limit: int | None, only_row: int | None, report: Path | None) -> int:
    """One pass over the report.

    Args:
        config: From load_config.
        limit: Do only the first N rows that still need doing.
        only_row: Do one row, by its row number in the working copy.
        report: An export already downloaded. None fetches a fresh one.

    Raises:
        SystemProblem: STAC or CCIS could not be opened, the report could not
            be fetched or read, the working copy could not be saved, or
            MAX_CONSECUTIVE_FAILURES rows failed in a row. Every outcome
            written before that point is already in the working copy.
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
            "Every case is still opened and every name still checked, which is "
            "what this pass is for."
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

    # The report first, from its own STAC, with that browser closed again
    # before anything is filed. Nothing else is signed in to until the report
    # says there is work to do.
    if report is None:
        report = _fetch_from_report_stac(config)
    else:
        logger.info("Using the export already downloaded: %s", report.name)

    working_copy, rows = report_reader.prepare_report(report)
    _hide_from_logs(rows)

    # The report's links name the STAC it was run on. Filing may be on a
    # different instance with the same incident ids, so the links are pointed
    # there. Only in memory: the working copy keeps the links as exported.
    retargeted = 0
    for row in rows:
        moved = report_reader.retarget_link(
            row["case_url"], config["report"]["url"], stac["url"]
        )
        if moved != row["case_url"]:
            row["case_url"] = moved
            retargeted += 1
    if retargeted:
        logger.info(
            "%s case link(s) pointed from the report's STAC at %s", retargeted, stac["url"]
        )

    with Spreadsheet(report_reader.working_copy_config(working_copy)) as sheet:
        if only_row is not None:
            rows = [row for row in rows if row["row"] == only_row]
            if not rows:
                raise SystemProblem(
                    f"--only-row {only_row} does not name a row in {working_copy.name}."
                )
            # Whatever its outcome says, as the option promises.
            rows[0]["outcome"] = ""
            logger.info("--only-row %s: one row this pass", only_row)

        already_done = sum(1 for row in rows if excel_handler.is_done(row))

        # Rows that cannot be searched get their outcome written now, before
        # either browser opens.
        todo = []
        for row in rows:
            if excel_handler.is_done(row):
                continue
            if row["problem"]:
                outcome = _problem_outcome(row["problem"])
                logger.warning("Row %s: %s", row["row"], row["problem"])
                sheet.write_outcome(row["row"], outcome)
                row["outcome"] = outcome
                continue
            todo.append(row)

        if limit is not None:
            todo = todo[:limit]
            logger.info("--limit %s: %s row(s) this pass", limit, len(todo))

        logger.info(
            "%s row(s) to do, %s already carry an outcome", len(todo), already_done
        )

        if not todo:
            logger.info("Nothing to do.")
            return EXIT_OK

        failures = 0
        # The filing STAC (stac.url) and CCIS, signed in to only now.
        with StacRunner(config) as runner, CcisSession(config) as ccis:
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
                        "working down the rest of the report against something "
                        "that is not answering."
                    )

        _summarise(todo, working_copy)

    if any(row["outcome"] == Outcome.CHECK_BY_HAND for row in todo):
        return EXIT_CHECK_BY_HAND
    return EXIT_OK


def next_slot(now: datetime) -> datetime:
    """The first --watch slot at or after `now`.

    Args:
        now: The current time.

    Returns:
        A weekday time between WATCH_START and WATCH_END on the
        WATCH_EVERY_MINUTES grid. A slot whose minute has already begun is
        past, so a pass that finishes quickly does not run twice.
    """
    day = now.date()
    while True:
        if day.weekday() in WATCH_WEEKDAYS:
            slot = datetime.combine(day, WATCH_START)
            last = datetime.combine(day, WATCH_END)
            while slot <= last:
                if slot >= now:
                    return slot
                slot += timedelta(minutes=WATCH_EVERY_MINUTES)
        day += timedelta(days=1)


def _sleep_until(moment: datetime) -> None:
    """Poll the clock until `moment` has passed."""
    while (remaining := (moment - datetime.now()).total_seconds()) > 0:
        time.sleep(min(remaining, WATCH_POLL_SECONDS))


def watch(config: dict, limit: int | None) -> int:
    """Run a fresh pass at every slot until stopped.

    Each pass fetches the report again. Filed cases drop off it, so a pass
    only sees what earlier passes have not done.

    A SystemProblem ends that pass only: STAC being down at 9:00 should not
    cost the 9:30 pass. Anything else ends the watch, through main's handlers:
    a ConfigProblem will not fix itself, and a bug would repeat every slot.

    An unconfirmed save does not stop the watch either. The report lists only
    cases with no CRIM HIST image, so a save that landed drops the case off
    the next report, and one that did not brings it back to be filed again.

    Args:
        config: From load_config. Read once, so a config change needs a restart.
        limit: Passed to every pass, as with --limit.

    Returns:
        Never returns normally. Ctrl+C or a non-SystemProblem error ends it.
    """
    while True:
        slot = next_slot(datetime.now())
        logger.info("Watching. Next pass at %s. Ctrl+C to stop.", f"{slot:%a %Y-%m-%d %H:%M}")
        _sleep_until(slot)

        logger.info("=== Pass for the %s slot ===", f"{slot:%H:%M}")
        try:
            code = run(config, limit, None, None)
        except SystemProblem as error:
            logger.error("Pass stopped: %s. Trying again at the next slot.", error)
            continue

        if code == EXIT_CHECK_BY_HAND:
            logger.warning(
                "A save was never confirmed. Carrying on: the next report will list "
                "that case again if the rap sheet did not land."
            )


def _summarise(rows: list, working_copy: Path) -> None:
    """Count what happened, and say plainly what still needs a person."""
    counts = Counter(row["outcome"] for row in rows)
    logger.info("Rows this pass: %s", len(rows))
    for outcome, count in counts.most_common():
        logger.info("  %-45s %s", outcome, count)

    # Searches past each row's first. Rows that never reached CCIS have no
    # count, and the first search is the one every row gets anyway.
    retried = [row for row in rows if row.get("ccis_searches", 0) > 1]
    if retried:
        found = sum(
            1 for row in retried
            if row["outcome"] in (
                Outcome.FILED_PARTIAL_NAME, Outcome.REACHED_SAVE_PARTIAL_NAME,
                Outcome.REHEARSED_PARTIAL_NAME,
            )
        )
        logger.info(
            "Extra CCIS searches: %s on %s row(s), %s found a record",
            sum(row["ccis_searches"] - 1 for row in retried), len(retried), found,
        )

    unconfirmed = counts.get(Outcome.CHECK_BY_HAND, 0)
    if unconfirmed:
        logger.warning(
            "%s row(s) are marked %r. Save was pressed and STAC never confirmed. "
            "The next report lists only cases still missing their rap sheet, so "
            "any that did not land will be filed then.",
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
            "%s row(s) need a look. Open %s and filter the Outcome column for "
            "anything that is not %r.", needs_a_person, working_copy.name, Outcome.FILED,
        )


def main(argv: list[str] | None = None) -> int:
    """Parse the command line, load config, and run or check.

    Args:
        argv: Command-line arguments, without the program name. None reads
            sys.argv.

    Returns:
        EXIT_OK, EXIT_SYSTEM_PROBLEM, EXIT_CONFIG_PROBLEM or EXIT_CHECK_BY_HAND.
    """
    parser = argparse.ArgumentParser(
        description="File CCIS case summaries onto the cases in STAC's AutoCCIS report."
    )
    parser.add_argument(
        "--limit",
        type=int,
        help="Do only the first N rows that still need doing. For testing.",
    )
    parser.add_argument(
        "--only-row",
        type=int,
        metavar="N",
        help="Do one row by its row number in the working copy, whatever its outcome says.",
    )
    parser.add_argument(
        "--report",
        type=Path,
        metavar="FILE",
        help="Use an export already downloaded instead of fetching the report. "
             "Outcomes already in its working copy are kept.",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Report what the config resolved to, and with --report what the "
             "export holds, then stop. Opens no browser and signs in to nothing.",
    )
    parser.add_argument(
        "--config",
        type=Path,
        help="Path to a config file other than config/config.yaml.",
    )
    parser.add_argument(
        "--watch",
        action="store_true",
        help="Keep running: a fresh pass every 30 minutes, 7:30 to 18:00, "
             "Monday to Friday. Ctrl+C to stop.",
    )
    args = parser.parse_args(argv)

    if args.watch and (args.report or args.only_row is not None or args.check):
        parser.error("--watch fetches a fresh report each pass, so it cannot take "
                     "--report, --only-row or --check")

    if args.limit is not None and args.limit < 1:
        parser.error("--limit must be at least 1")
    if args.only_row is not None and args.only_row < 1:
        parser.error("--only-row must be at least 1")
    if args.only_row is not None and args.limit is not None:
        parser.error("--only-row and --limit do not make sense together")
    if args.only_row is not None and args.report is None:
        parser.error(
            "--only-row needs --report: row numbers belong to one export's working copy"
        )

    report = args.report.resolve() if args.report else None

    # Logging is not set up yet, because the log folder comes from config.
    try:
        config = load_config(args.config)
    except ConfigProblem as error:
        print(f"Config problem: {error}", file=sys.stderr)
        return EXIT_CONFIG_PROBLEM

    setup_logging(Path(config["paths"]["logs"]))

    try:
        if args.check:
            return check(config, report)
        if args.watch:
            return watch(config, args.limit)
        return run(config, args.limit, args.only_row, report)
    except ConfigProblem as error:
        logger.error("Config problem: %s", error)
        return EXIT_CONFIG_PROBLEM
    except SystemProblem as error:
        logger.error("Stopped: %s", error)
        return EXIT_SYSTEM_PROBLEM
    except KeyboardInterrupt:
        logger.warning(
            "Interrupted. Every row already attempted has its outcome in the "
            "working copy. Filed cases drop off the report, so the next run picks "
            "up the rest."
        )
        return EXIT_SYSTEM_PROBLEM
    except Exception:
        logger.exception("Unexpected error. This is a bug, not a row problem.")
        return EXIT_SYSTEM_PROBLEM


if __name__ == "__main__":
    sys.exit(main())
