"""Runs the AutoCCIS report in STAC and downloads it as Excel.

The clicks, from the signed-in page:

    Cases (sidebar)  ->  Report (case menu)  ->  Search  ->  Excel export

The report form arrives already filled in, so Search needs nothing typed.
After Search the report goes through several loading phases before the export
button does anything, so the wait is on the button being usable with no
loading mask on screen, not on a fixed sleep.

Chrome is told, through the DevTools protocol, to download into a folder this
run chooses, so the file never has to be fished out of the user's Downloads.
Chrome names it AutoCCIS.xls, or AutoCCIS (1).xls and so on when that name is
taken, so the file is found by being new, not by its name, then renamed with
the date and time so each run's export is kept and none overwrites another.

The browser stays on the report page afterwards. Each case opens in a tab of
its own, so nothing needs the main tab to be anywhere in particular.
"""

import time
from datetime import datetime
from pathlib import Path

from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from exceptions import SystemProblem
from logger import get_logger
from stac import CASES_SIDEBAR_CSS, StacSession

logger = get_logger(__name__)

# The clicks. Selectors from STAC's own markup.
REPORT_MENU_CSS = "li[data-button-id='report']"
REPORT_SEARCH_BUTTON_ID = "pagesIncidentReportsIndexFrm_report_button"
EXCEL_EXPORT_CSS = "button.report-export-btn[data-format='Excel']"

# The dialog STAC shows when Search is pressed before the form has filled in
# its default report. Matched on its text, and Ok as the next button after
# it, so nothing depends on the dialog's own markup.
NOT_READY_TEXT_XPATH = "//*[normalize-space(text())='Please select a report.']"
NOT_READY_OK_XPATH = f"{NOT_READY_TEXT_XPATH}/following::button[normalize-space(.)='Ok'][1]"

# Seconds to watch for that dialog after pressing Search, and to wait before
# pressing it again. Search is tried this many times in all.
NOT_READY_WATCH_SECONDS = 3
SEARCH_RETRY_PAUSE_SECONDS = 2
SEARCH_ATTEMPTS = 3

# Anything Kendo or the report viewer shows while it is still working. The
# export is only clicked once none of these is visible.
LOADING_CSS = ".k-loading-mask, .k-loading-image, .k-loader, .k-i-loading"

# Seconds. The report takes a while to build; the download is small.
DEFAULT_REPORT_TIMEOUT = 180
DEFAULT_DOWNLOAD_TIMEOUT = 60

# Seconds to wait for a download to begin after clicking export. If nothing
# appears in that time the click most likely landed before the report was
# ready, and it is tried once more.
DOWNLOAD_START_SECONDS = 20
EXPORT_CLICK_ATTEMPTS = 2

# How often to look at the folder while waiting.
POLL_SECONDS = 0.5

# Files Chrome writes while a download is still in progress.
PARTIAL_SUFFIXES = (".crdownload", ".tmp", ".part")

# The report's name, which Chrome uses for the file. Only used to name the
# renamed copy and to warn if some other report came down instead.
REPORT_NAME = "AutoCCIS"


def fetch_report(
    session: StacSession,
    download_dir: Path,
    report_timeout: int = DEFAULT_REPORT_TIMEOUT,
    download_timeout: int = DEFAULT_DOWNLOAD_TIMEOUT,
) -> Path:
    """Run the report and download it, returning the path of the saved file.

    Args:
        session: A signed-in StacSession, on its main tab.
        download_dir: Where the export is saved. Created if missing.
        report_timeout: Seconds to wait for the report to finish loading.
        download_timeout: Seconds to wait for the download to finish once it
            has started.

    Returns:
        The export, renamed to AutoCCIS_<date>_<time> with its original
        extension.

    Raises:
        SystemProblem: A click target never appeared, the report never
            finished loading, or no file arrived. Without the report there is
            nothing to work on, so the run stops.
    """
    download_dir = Path(download_dir).resolve()
    download_dir.mkdir(parents=True, exist_ok=True)
    _send_downloads_to(session, download_dir)

    driver = session.driver
    wait = WebDriverWait(driver, session.wait_timeout)

    _click(wait, driver, (By.CSS_SELECTOR, CASES_SIDEBAR_CSS), "the Cases sidebar")
    _click(wait, driver, (By.CSS_SELECTOR, REPORT_MENU_CSS), "Report in the case menu")
    _start_search(wait, driver)
    logger.info("Report search started, waiting for it to load")

    before = _files_in(download_dir)
    downloaded = None
    for attempt in range(1, EXPORT_CLICK_ATTEMPTS + 1):
        button = _wait_for_export(driver, report_timeout)
        driver.execute_script("arguments[0].click();", button)
        # The button may have been inside the viewer's iframe. Everything
        # after this works from the top-level page.
        driver.switch_to.default_content()
        logger.info("Clicked Excel export (try %s of %s)", attempt, EXPORT_CLICK_ATTEMPTS)

        if _wait_for_start(download_dir, before, DOWNLOAD_START_SECONDS):
            downloaded = _wait_for_finish(download_dir, before, download_timeout)
            break
        logger.warning(
            "No download started within %ss of clicking export", DOWNLOAD_START_SECONDS
        )

    if downloaded is None:
        raise SystemProblem(
            f"Clicked the Excel export {EXPORT_CLICK_ATTEMPTS} times and no file "
            f"arrived in {download_dir}."
        )

    if REPORT_NAME.lower() not in downloaded.name.lower():
        # Still used: the reader checks the header, which is the real test.
        # Said here because a different report would explain a header error.
        logger.warning(
            "The export is called %r, not %s. Check the report form's default.",
            downloaded.name, REPORT_NAME,
        )

    final = download_dir / (
        f"{REPORT_NAME}_{datetime.now():%Y-%m-%d_%H%M%S}{downloaded.suffix.lower()}"
    )
    try:
        downloaded.replace(final)
    except OSError as error:
        raise SystemProblem(f"Could not rename the export {downloaded.name}: {error}") from error

    logger.info("Report saved as %s", final.name)
    return final


def _send_downloads_to(session: StacSession, folder: Path) -> None:
    """Tell this Chrome to save downloads into `folder` without asking.

    Raises:
        SystemProblem: Chrome refused. Without this the file would land in
            the user's Downloads, where picking out the right one is guesswork.
    """
    try:
        session.driver.execute_cdp_cmd(
            "Browser.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": str(folder), "eventsEnabled": False},
        )
    except WebDriverException as error:
        raise SystemProblem(
            f"Could not set Chrome's download folder to {folder}: {error}"
        ) from error


def _click(wait: WebDriverWait, driver, locator: tuple, what: str) -> None:
    """Wait for an element to be clickable and click it through JavaScript.

    JavaScript, as everywhere else in STAC: Kendo overlays intercept native
    clicks often enough to matter.

    Raises:
        SystemProblem: It never became clickable.
    """
    try:
        element = wait.until(EC.element_to_be_clickable(locator))
        driver.execute_script("arguments[0].click();", element)
    except TimeoutException as error:
        raise SystemProblem(f"Could not click {what}; it never became clickable.") from error


def _start_search(wait: WebDriverWait, driver) -> None:
    """Press the report's Search, again if STAC says no report is selected yet.

    The Search button is clickable before the form has loaded its default
    report. Pressed that early, STAC shows "Please select a report." and runs
    nothing, which used to look like a report that never finished loading.

    Raises:
        SystemProblem: Search could not be clicked, the dialog could not be
            dismissed, or the form still had no report after every attempt.
    """
    for attempt in range(1, SEARCH_ATTEMPTS + 1):
        _click(wait, driver, (By.ID, REPORT_SEARCH_BUTTON_ID), "the report's Search button")
        if not _dismiss_not_ready(driver):
            return
        logger.warning(
            "Search was pressed before the report form had loaded (try %s of %s)",
            attempt, SEARCH_ATTEMPTS,
        )
        time.sleep(SEARCH_RETRY_PAUSE_SECONDS)

    raise SystemProblem(
        f"Pressed the report's Search {SEARCH_ATTEMPTS} times and STAC said "
        "'Please select a report.' every time. The form is not filling in its "
        "default report."
    )


def _dismiss_not_ready(driver) -> bool:
    """Close the "Please select a report." dialog if it appears.

    Returns:
        True if the dialog appeared and was closed, False if it never showed.

    Raises:
        SystemProblem: The dialog appeared but its Ok button could not be
            clicked, which would leave it blocking every later click.
    """
    try:
        WebDriverWait(driver, NOT_READY_WATCH_SECONDS).until(
            EC.visibility_of_element_located((By.XPATH, NOT_READY_TEXT_XPATH))
        )
    except TimeoutException:
        return False

    try:
        ok = WebDriverWait(driver, NOT_READY_WATCH_SECONDS).until(
            EC.element_to_be_clickable((By.XPATH, NOT_READY_OK_XPATH))
        )
        driver.execute_script("arguments[0].click();", ok)
        WebDriverWait(driver, NOT_READY_WATCH_SECONDS).until(
            EC.invisibility_of_element_located((By.XPATH, NOT_READY_TEXT_XPATH))
        )
    except TimeoutException as error:
        raise SystemProblem(
            "STAC said 'Please select a report.' and its Ok button could not be "
            "clicked to close it."
        ) from error
    return True


def _find_export(driver) -> WebElement | None:
    """The Excel export button if it is usable and nothing is loading, else None.

    Looks in the page and then in each iframe on it, since the report viewer
    lives in one and the button may sit inside it. Always leaves the driver
    on the top-level page, except when it returns a button inside a frame,
    in which case it leaves the driver in that frame so the button can be
    clicked.
    """
    driver.switch_to.default_content()
    contexts = [None] + driver.find_elements(By.TAG_NAME, "iframe")

    for frame in contexts:
        try:
            driver.switch_to.default_content()
            if frame is not None:
                driver.switch_to.frame(frame)
            if any(el.is_displayed() for el in driver.find_elements(By.CSS_SELECTOR, LOADING_CSS)):
                driver.switch_to.default_content()
                return None
            for button in driver.find_elements(By.CSS_SELECTOR, EXCEL_EXPORT_CSS):
                if button.is_displayed() and button.is_enabled():
                    return button
        except WebDriverException:
            continue

    driver.switch_to.default_content()
    return None


def _wait_for_export(driver, timeout: int) -> WebElement:
    """Wait until the report has loaded and the Excel export can be clicked.

    Raises:
        SystemProblem: The report did not finish within `timeout` seconds.
    """
    try:
        return WebDriverWait(driver, timeout, poll_frequency=1).until(_find_export)
    except TimeoutException as error:
        driver.switch_to.default_content()
        raise SystemProblem(
            f"The report did not finish loading within {timeout}s, or its Excel "
            "export button never appeared."
        ) from error


def _files_in(folder: Path) -> set[str]:
    """Names of the files in `folder` right now."""
    return {path.name for path in folder.iterdir() if path.is_file()}


def _new_files(folder: Path, before: set[str]) -> list[Path]:
    """Files in `folder` that were not there before, partial downloads included."""
    return [
        path for path in folder.iterdir()
        if path.is_file() and path.name not in before
    ]


def _wait_for_start(folder: Path, before: set[str], seconds: float) -> bool:
    """True once any new file, finished or not, shows up in `folder`."""
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if _new_files(folder, before):
            return True
        time.sleep(POLL_SECONDS)
    return False


def _wait_for_finish(folder: Path, before: set[str], seconds: float) -> Path:
    """Wait for the new download to finish, and return it.

    Finished means no partial file is left and the new file's size has held
    steady across two looks.

    Raises:
        SystemProblem: Still downloading after `seconds`.
    """
    deadline = time.monotonic() + seconds
    last_size = -1
    while time.monotonic() < deadline:
        new = _new_files(folder, before)
        partial = [p for p in new if p.suffix.lower() in PARTIAL_SUFFIXES]
        done = [p for p in new if p.suffix.lower() not in PARTIAL_SUFFIXES]
        if done and not partial:
            newest = max(done, key=lambda p: p.stat().st_mtime)
            size = newest.stat().st_size
            if size > 0 and size == last_size:
                return newest
            last_size = size
        time.sleep(POLL_SECONDS)

    raise SystemProblem(f"The export was still downloading after {seconds}s.")
