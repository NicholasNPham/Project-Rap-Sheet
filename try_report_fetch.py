"""Fetch the AutoCCIS report once and check it is a usable export. Nothing else.

Signs in to STAC, runs the report, downloads the Excel export, closes the
browser, then reads the file the way a real pass would and says whether it
worked. No CCIS, no case pages, no uploads, no saves. For working on the
export step without starting a whole pass.

    python try_report_fetch.py
        Fetch from report.url, the same STAC a real pass uses.

    python try_report_fetch.py --filing-stac
        Fetch from stac.url instead, to see whether that instance's export
        behaves differently.

Either way the session is read only: upload and save are forced off, the
same as for a real pass's report fetch.

The export holds defendant names and dates of birth, so only counts are
printed: how many rows, how many have a case link, how many the run could
not search. The file is left in paths.reports, named as a real pass names it.
"""

import argparse
import sys
from pathlib import Path

from config_loader import load_config
from exceptions import ConfigProblem, SystemProblem
from logger import get_logger, setup_logging
from main import report_session_config
from report_fetch import fetch_report
from report_reader import build_rows, find_header, read_xls
from stac import StacSession

logger = get_logger("try_report_fetch")

EXIT_OK = 0
EXIT_FETCH_FAILED = 1
EXIT_CONFIG_PROBLEM = 2
EXIT_UNREADABLE = 3


def session_config(config: dict, filing_stac: bool) -> dict:
    """The read-only session config to fetch with.

    Args:
        config: From load_config.
        filing_stac: True to fetch from stac.url instead of report.url.

    Returns:
        A new config dict with upload and save forced off.
    """
    chosen = report_session_config(config)
    if filing_stac:
        chosen["stac"]["url"] = config["stac"]["url"]
    return chosen


def fetch(config: dict) -> Path:
    """Sign in, download the report, and close the browser.

    Args:
        config: From session_config.

    Returns:
        The saved export.

    Raises:
        SystemProblem: From fetch_report or the sign-in.
    """
    with StacSession(config) as session:
        return fetch_report(
            session,
            config["paths"]["reports"],
            config["report"]["report_timeout"],
            config["report"]["download_timeout"],
        )


def check_export(path: Path) -> None:
    """Read the export as a real pass would and log counts only.

    Args:
        path: The saved export.

    Raises:
        ConfigProblem: The file is not an .xls, or has no header row the
            reader recognises.
    """
    rows = read_xls(path)
    header_at, columns = find_header(rows)
    built = build_rows(rows, header_at, columns)

    with_link = sum(1 for row in built if row["case_url"])
    with_problem = sum(1 for row in built if row["problem"])
    logger.info(
        "Export %s: %s bytes, %s row(s), %s with a case link, %s the run could not search",
        path.name, path.stat().st_size, len(built), with_link, with_problem,
    )
    if built and not with_link:
        # Every row reached through its link, so an export without them
        # reads fine here but would fail every row of a real pass.
        logger.warning("No row has a case link. A real pass could not open any case.")


def main(argv: list[str] | None = None) -> int:
    """Parse the command line, fetch the report, and check it.

    Args:
        argv: Command-line arguments, without the program name. None reads
            sys.argv.

    Returns:
        EXIT_OK, EXIT_FETCH_FAILED, EXIT_CONFIG_PROBLEM or EXIT_UNREADABLE.
    """
    parser = argparse.ArgumentParser(
        description="Fetch the AutoCCIS report once and check it. Files nothing."
    )
    parser.add_argument(
        "--filing-stac",
        action="store_true",
        help="Fetch from stac.url instead of report.url.",
    )
    parser.add_argument(
        "--config",
        type=Path,
        help="Path to a config file other than config/config.yaml.",
    )
    args = parser.parse_args(argv)

    try:
        config = load_config(args.config)
    except ConfigProblem as error:
        print(f"Config problem: {error}", file=sys.stderr)
        return EXIT_CONFIG_PROBLEM

    setup_logging(Path(config["paths"]["logs"]))
    chosen = session_config(config, args.filing_stac)
    logger.info("Fetching the report from %s (read only)", chosen["stac"]["url"])

    try:
        export = fetch(chosen)
    except SystemProblem as error:
        logger.error("Fetch failed: %s", error)
        return EXIT_FETCH_FAILED
    except KeyboardInterrupt:
        logger.warning("Interrupted before the report arrived.")
        return EXIT_FETCH_FAILED

    try:
        check_export(export)
    except ConfigProblem as error:
        logger.error("Downloaded %s but it is not a usable export: %s", export.name, error)
        return EXIT_UNREADABLE

    logger.info("Report fetch works.")
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
