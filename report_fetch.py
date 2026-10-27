"""Runs the AutoCCIS report in STAC and downloads it as Excel.

The clicks, from the signed-in page:

    Cases (sidebar)  ->  Report (case menu)  ->  Search  ->  Excel export

The report form arrives already filled in, so Search needs nothing typed.
Search opens a Report Preview window at once, with its Save As toolbar and
Excel button ready, but the viewer frame under the toolbar stays blank for a
while before the report appears in it. Clicking export in that gap makes STAC
build the whole report before it sends anything, so the wait is on the viewer
having filled in, not on the button or a fixed sleep.

Chrome is told, through the DevTools protocol, to download into an incoming
folder inside the reports folder, so the file never has to be fished out of
the user's Downloads. Chrome names it AutoCCIS.xls and writes over a file of
that name, so a raw export left behind by a failed run once hid every later
one. The incoming folder is emptied before the export is clicked, which
leaves the download as the only thing that can arrive there. It is then
renamed with the date and time and moved up into the reports folder, so each
run's export is kept and none overwrites another.

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
EXCEL_EXPORT_CSS = "#reportPreviewWindow button.report-export-btn[data-format='Excel']"

# Says whether the report list has loaded with the wanted report selected.
# The Search button is clickable about three seconds before the list exists,
# and Search pressed in that gap gets "Please select a report." Watched by
# hand, the row arrived already selected in the grid, so this waits for that
# rather than selecting anything itself. Matched on the row's text: its id
# and data-uid are fresh GUIDs on every load.
REPORT_SELECTED_JS = r"""
var wanted = arguments[0];
var cells = document.querySelectorAll('tr.k-master-row td');
var row = null;
for (var i = 0; i < cells.length; i++) {
  if ((cells[i].textContent || '').trim() === wanted) { row = cells[i].closest('tr'); break; }
}
if (!row) { return 'no-row'; }
var gridEl = row.closest('.k-grid');
var grid = (gridEl && window.jQuery) ? window.jQuery(gridEl).data('kendoGrid') : null;
if (!grid) { return 'no-grid'; }
var picked = grid.select();
if (!picked || picked.length === 0) { return 'nothing-selected'; }
return picked[0] === row ? 'ok' : 'other-selected';
"""

# Seconds to wait for the report list. It took about three by hand; this is
# well past that so a slow STAC is not mistaken for a missing report.
REPORT_LIST_SECONDS = 60

# What each answer from REPORT_SELECTED_JS means, for the error if the wait
# runs out.
REPORT_LIST_STATES = {
    "no-row": "the report list never showed a row named {name}",
    "no-grid": "the {name} row is not in a grid STAC's code can read",
    "nothing-selected": "the {name} row is there but no report is selected",
    "other-selected": "a report other than {name} is selected",
}

# The viewer frame inside the preview window. Matched on the page name only:
# the rest of its src is session tokens that change on every load.
VIEWER_FRAME_CSS = "#reportPreviewWindow iframe[src*='viewer.aspx']"

# How much text the viewer shows. Only the length comes back, never the text,
# which is the report's rows.
VIEWER_TEXT_LENGTH_JS = "return document.body ? document.body.innerText.trim().length : 0;"

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
# export is only clicked once none of these is visible inside the viewer.
# The viewer has not been seen to show one, so filling in is the real test;
# this is a guard in case it ever does.
LOADING_CSS = ".k-loading-mask, .k-loading-image, .k-loader, .k-i-loading"

# Seconds. The report takes a while to build. The download is small, but STAC
# builds the export after the click, and its partial file can sit unchanged
# for over a minute before the whole thing arrives at once. A slow export and
# a stuck one look the same from the folder, so the download wait is a hard
# cap set well past the slowest seen rather than a check for progress.
DEFAULT_REPORT_TIMEOUT = 180
DEFAULT_DOWNLOAD_TIMEOUT = 300

# Seconds to wait for a download to begin after clicking export. If nothing
# appears in that time the click most likely landed before the report was
# ready, and it is tried once more.
DOWNLOAD_START_SECONDS = 20
EXPORT_CLICK_ATTEMPTS = 3

# Seconds to keep retrying the rename while Chrome still holds the file.
RENAME_SECONDS = 10

# How often to look at the folder while waiting.
POLL_SECONDS = 0.5

# Files Chrome writes while a download is still in progress.
PARTIAL_SUFFIXES = (".crdownload", ".tmp", ".part")

# The report's name, which Chrome uses for the file. Only used to name the
# renamed copy and to warn if some other report came down instead.
REPORT_NAME = "AutoCCIS"

# The subfolder of the reports folder Chrome downloads into. Only raw,
# unrenamed downloads ever sit here, so emptying it loses nothing a pass used.
INCOMING_FOLDER_NAME = "incoming"

# Chrome's own downloads page. When an export attempt fails it is read in a
# tab of its own, because a download Chrome cancels or blocks has its file
# deleted, leaving nothing in the folder to say why.
CHROME_DOWNLOADS_URL = "chrome://downloads"

# Fields of each download worth logging. A list of names to take, not a list
# to leave out, so the download's URL, which carries session tokens, can never
# slip into the log.
CHROME_DOWNLOAD_FIELDS = (
    "fileName", "filePath", "state", "dangerType", "lastReasonText",
    "percent", "progressStatusText", "isDangerous", "isInsecure",
)

# Reads those fields, plus what each row shows a person: its status tag
# ("Canceled", "Failed - Network error"), its description line, and the
# buttons it offers. The buttons say what Chrome is waiting on better than
# its state number, which changes meaning between Chrome versions: Pause and
# Cancel for a download under way, Keep and Discard for one held for a
# decision. The row's URL element is never read. null until the page has
# drawn its list.
CHROME_DOWNLOADS_JS = r"""
var fields = arguments[0];
var manager = document.querySelector('downloads-manager');
if (!manager || !manager.shadowRoot) { return null; }
var rows = manager.shadowRoot.querySelectorAll('downloads-item');
var out = [];
for (var i = 0; i < rows.length; i++) {
  var data = rows[i].data || {};
  var entry = {};
  for (var f = 0; f < fields.length; f++) {
    var value = data[fields[f]];
    if (value !== undefined && value !== null && typeof value !== 'object') {
      entry[fields[f]] = String(value);
    }
  }
  var root = rows[i].shadowRoot;
  var tag = root ? root.querySelector('#tag') : null;
  entry.status = tag ? tag.textContent.trim() : '';
  var description = root ? root.querySelector('#description') : null;
  entry.description = description ? description.textContent.trim() : '';
  var labels = [];
  var buttons = root ? root.querySelectorAll('cr-button, button, cr-icon-button') : [];
  for (var b = 0; b < buttons.length; b++) {
    if (buttons[b].offsetParent === null) { continue; }
    var label = (buttons[b].getAttribute('aria-label') || buttons[b].textContent || '').trim();
    if (label) { labels.push(label); }
  }
  entry.buttons = labels.join(' | ');
  out.push(entry);
}
return out;
"""

# Seconds to give the downloads page to list its rows. The list fills in a
# moment after the page loads, and an empty list is a real answer, so this
# is how long an empty list is waited on before being believed.
CHROME_DOWNLOADS_SECONDS = 5


def fetch_report(
    session: StacSession,
    download_dir: Path,
    report_timeout: int = DEFAULT_REPORT_TIMEOUT,
    download_timeout: int = DEFAULT_DOWNLOAD_TIMEOUT,
) -> Path:
    """Run the report and download it, returning the path of the saved file.

    Args:
        session: A signed-in StacSession, on its main tab.
        download_dir: Where the export is saved. Created if missing, along
            with the incoming folder inside it that Chrome downloads into.
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
    incoming = download_dir / INCOMING_FOLDER_NAME
    incoming.mkdir(parents=True, exist_ok=True)
    _send_downloads_to(session, incoming)

    driver = session.driver
    wait = WebDriverWait(driver, session.wait_timeout)

    _click(wait, driver, (By.CSS_SELECTOR, CASES_SIDEBAR_CSS), "the Cases sidebar")
    _click(wait, driver, (By.CSS_SELECTOR, REPORT_MENU_CSS), "Report in the case menu")
    _wait_for_report_selected(driver, REPORT_LIST_SECONDS)
    _start_search(wait, driver)
    logger.info("Report search started, waiting for it to load")

    # Emptied once, not before every try: a download from an earlier try that
    # arrives late is still this pass's report.
    _empty_incoming(incoming)
    before = _files_in(incoming)
    downloaded = None
    for attempt in range(1, EXPORT_CLICK_ATTEMPTS + 1):
        button = _wait_for_export(driver, report_timeout)
        driver.execute_script("arguments[0].click();", button)
        logger.info("Clicked Excel export (try %s of %s)", attempt, EXPORT_CLICK_ATTEMPTS)

        if not _wait_for_start(incoming, before, DOWNLOAD_START_SECONDS):
            logger.warning(
                "No download started within %ss of clicking export", DOWNLOAD_START_SECONDS
            )
            _log_chrome_downloads(driver)
            continue

        try:
            arrived = _wait_for_finish(incoming, before, download_timeout)
        except SystemProblem:
            _log_chrome_downloads(driver)
            raise
        if REPORT_NAME.lower() in arrived.name.lower():
            downloaded = arrived
            break

        # Chrome names an HTML response with no filename downloads.htm. That
        # is STAC answering the export with a page, not the report. It is left
        # in the incoming folder, ignored from here on, and the export clicked
        # again. The next pass empties it.
        logger.warning(
            "The export came back as %r, not %s. Clicking export again.",
            arrived.name, REPORT_NAME,
        )
        _log_chrome_downloads(driver)
        before[arrived.name] = arrived.stat().st_mtime_ns

    if downloaded is None:
        raise SystemProblem(
            f"Clicked the Excel export {EXPORT_CLICK_ATTEMPTS} times and no "
            f"{REPORT_NAME} file arrived in {incoming}."
        )

    final = download_dir / (
        f"{REPORT_NAME}_{datetime.now():%Y-%m-%d_%H%M%S}{downloaded.suffix.lower()}"
    )
    _rename(downloaded, final)

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


def _wait_for_report_selected(driver, timeout: int) -> None:
    """Wait until the report list has loaded with the AutoCCIS report selected.

    Args:
        driver: The report STAC's WebDriver, on the report page.
        timeout: Seconds to wait.

    Raises:
        SystemProblem: The list did not reach that state in time. Searching
            anyway would run nothing, or run the wrong report.
    """
    last_state = "no-row"

    def selected(driver) -> bool:
        nonlocal last_state
        try:
            last_state = driver.execute_script(REPORT_SELECTED_JS, REPORT_NAME)
        except WebDriverException:
            # The page swapping in the report form looks like this. The next
            # look tries again.
            return False
        return last_state == "ok"

    try:
        WebDriverWait(driver, timeout, poll_frequency=POLL_SECONDS).until(selected)
    except TimeoutException as error:
        why = REPORT_LIST_STATES.get(last_state, f"its state was {last_state!r}")
        raise SystemProblem(
            f"The report page was not ready to search after {timeout}s: "
            f"{why.format(name=REPORT_NAME)}."
        ) from error
    logger.info("Report list loaded with %s selected", REPORT_NAME)


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


def _viewer_text_length(driver) -> int:
    """How much text the report viewer is showing, or 0 if it is not ready.

    0 while the viewer frame is missing, its page is still loading, or a
    loading indicator shows inside it. Always leaves the driver on the
    top-level page.

    Returns:
        The length of the viewer's text, or 0.
    """
    driver.switch_to.default_content()
    try:
        frames = driver.find_elements(By.CSS_SELECTOR, VIEWER_FRAME_CSS)
        if not frames:
            return 0
        driver.switch_to.frame(frames[0])
        if driver.execute_script("return document.readyState") != "complete":
            return 0
        if any(el.is_displayed() for el in driver.find_elements(By.CSS_SELECTOR, LOADING_CSS)):
            return 0
        return int(driver.execute_script(VIEWER_TEXT_LENGTH_JS) or 0)
    except WebDriverException:
        # The frame reloading under us looks like this. Not ready yet; the
        # next look tries again.
        return 0
    finally:
        driver.switch_to.default_content()


def _find_export(driver) -> WebElement | None:
    """The preview window's Excel export button if it is usable, else None."""
    try:
        for button in driver.find_elements(By.CSS_SELECTOR, EXCEL_EXPORT_CSS):
            if button.is_displayed() and button.is_enabled():
                return button
    except WebDriverException:
        pass
    return None


def _wait_for_export(driver, timeout: int) -> WebElement:
    """Wait until the report preview has filled in, then return its Excel button.

    Filled in means the viewer shows some text and the amount has held
    steady across two looks a second apart, so a report still drawing is not
    taken as done.

    Raises:
        SystemProblem: The preview did not fill in within `timeout` seconds,
            or its Excel button never became usable.
    """
    last_length = -1

    def ready(driver) -> WebElement | None:
        nonlocal last_length
        length = _viewer_text_length(driver)
        settled = length > 0 and length == last_length
        last_length = length
        return _find_export(driver) if settled else None

    try:
        return WebDriverWait(driver, timeout, poll_frequency=1).until(ready)
    except TimeoutException as error:
        raise SystemProblem(
            f"The report preview did not fill in within {timeout}s, or its "
            "Excel export button never became usable."
        ) from error


def _log_chrome_downloads(driver) -> None:
    """Log what Chrome's downloads page says about each download. Never raises.

    Opens chrome://downloads in a new tab, logs one line per download, then
    closes the tab and goes back to the tab it started on. Only the fields in
    CHROME_DOWNLOAD_FIELDS and the row's status line are logged, never the
    URL. A diagnostic, so any failure here is logged and swallowed rather
    than ending a pass that might still succeed on the next click.

    Args:
        driver: The report STAC's WebDriver, on the report tab.
    """
    try:
        home = driver.current_window_handle
    except WebDriverException as error:
        logger.warning("Could not read Chrome's downloads page: %s", error)
        return

    try:
        driver.switch_to.new_window("tab")
        driver.get(CHROME_DOWNLOADS_URL)
        rows = None
        deadline = time.monotonic() + CHROME_DOWNLOADS_SECONDS
        while time.monotonic() < deadline:
            rows = driver.execute_script(CHROME_DOWNLOADS_JS, list(CHROME_DOWNLOAD_FIELDS))
            if rows:
                break
            time.sleep(POLL_SECONDS)

        if rows is None:
            logger.warning("Chrome's downloads page never drew its list")
        elif not rows:
            logger.warning("Chrome's downloads page lists no downloads at all")
        else:
            for number, row in enumerate(rows, start=1):
                logger.warning(
                    "Chrome download %s of %s: %s",
                    number, len(rows), ", ".join(f"{k}={v!r}" for k, v in row.items()),
                )
    except WebDriverException as error:
        logger.warning("Could not read Chrome's downloads page: %s", error)
    finally:
        try:
            if driver.current_window_handle != home:
                driver.close()
            driver.switch_to.window(home)
        except WebDriverException as error:
            logger.warning("Could not get back to the report tab: %s", error)


def _empty_incoming(folder: Path) -> None:
    """Delete every file an earlier run left in the incoming folder. Never raises.

    Anything here is a raw download a failed pass never renamed or used: a
    late copy of a report, or STAC's downloads.htm. Each one removed is
    logged by name. One that cannot be removed, because something still has
    it open, is logged and left: the download wait already tells a file
    written after the click from one that was there before, so it cannot be
    mistaken for this pass's export.

    Args:
        folder: The incoming folder.
    """
    for path in folder.iterdir():
        if not path.is_file():
            continue
        try:
            path.unlink()
            logger.warning("Removed %s left in %s by an earlier run", path.name, folder.name)
        except OSError as error:
            logger.warning(
                "Could not remove %s from %s, leaving it: %s", path.name, folder.name, error
            )


def _files_in(folder: Path) -> dict[str, int]:
    """Each file in `folder` right now, with its last-modified time in nanoseconds.

    A file that disappears between being listed and being looked at is left
    out. Chrome renaming its partial file to the finished name looks like
    that, and the next look sees the finished file.
    """
    times = {}
    for path in folder.iterdir():
        try:
            if path.is_file():
                times[path.name] = path.stat().st_mtime_ns
        except FileNotFoundError:
            continue
    return times


def _new_files(folder: Path, before: dict[str, int]) -> list[Path]:
    """Files in `folder` that are new or written since `before`, partials included.

    Written since counts as well as new because Chrome, told where to save,
    writes over a file of the same name. An AutoCCIS.xls left behind by an
    earlier run that never renamed it would otherwise hide every later
    download under a name that was already there.
    """
    return [
        folder / name for name, modified in _files_in(folder).items()
        if before.get(name) != modified
    ]


def _wait_for_start(folder: Path, before: dict[str, int], seconds: float) -> bool:
    """True once any new file, finished or not, shows up in `folder`."""
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if _new_files(folder, before):
            return True
        time.sleep(POLL_SECONDS)
    return False


def _rename(source: Path, target: Path) -> None:
    """Rename the export, retrying while Chrome still has it open.

    A steady size is not proof Chrome has let go of the file, and Windows
    refuses to rename a file another process holds.

    Raises:
        SystemProblem: Still locked after RENAME_SECONDS, or any other OS error.
    """
    deadline = time.monotonic() + RENAME_SECONDS
    while True:
        try:
            source.replace(target)
            return
        except PermissionError as error:
            if time.monotonic() >= deadline:
                raise SystemProblem(
                    f"Could not rename the export {source.name}; it was still in use "
                    f"after {RENAME_SECONDS}s: {error}"
                ) from error
            time.sleep(POLL_SECONDS)
        except OSError as error:
            raise SystemProblem(f"Could not rename the export {source.name}: {error}") from error


def _wait_for_finish(folder: Path, before: dict[str, int], seconds: float) -> Path:
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

    new = _new_files(folder, before)
    if any(p.suffix.lower() in PARTIAL_SUFFIXES for p in new):
        raise SystemProblem(
            f"The export was still downloading after {seconds}s: Chrome had a partial "
            "file that never completed."
        )
    raise SystemProblem(
        f"The export never settled into a finished file within {seconds}s "
        f"({len(new)} new file(s), none complete and non-empty)."
    )
