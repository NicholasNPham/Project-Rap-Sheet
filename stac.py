"""Drives the STAC web interface with Selenium: sign in, find the case, file the PDF.

A port of src/stac.py from PROJECT-DALYN, which is itself a port of
ftp_to_stac.py from PCSO911. Four things change in this port:

    Type and Subtype are ONE pair for the whole run, read from config. DALYN
    took them per document from a rules sheet because it classified unknown
    filings; every row here produces the same kind of document, so there is
    nothing to classify and no rules sheet.

    The defendant check compares STAC's name against the name in the
    spreadsheet row, the way PCSO911 compares it against the FTP directory
    name. DALYN read the name out of the document's own caption, which cannot
    work here: a CCIS case-summary page has no STATE OF FLORIDA vs. caption,
    so the caption parser would return nothing and the check would silently
    skip itself on every single row.

    fresh_browser defaults to false, not true. DALYN closes Chrome after every
    upload box; at one upload box per row that is roughly ten seconds of
    re-signing-in per row, and it is the single largest cost in a run.

    Nothing emails anybody from inside an except block. Failures raise
    RowProblem or SystemProblem for main.py to route.

One browser session covers the whole run. PCSO911 opens and closes Chrome per
case, which is fine at its volume; a rap sheet run walks a whole spreadsheet
and signing in that many times would be both slow and conspicuous in STAC's
audit log.
"""

import re
import time
from difflib import SequenceMatcher
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import (
    NoAlertPresentException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

from exceptions import RowProblem, SystemProblem
from logger import get_logger

logger = get_logger(__name__)

# Seconds. Defaults only; config.yaml overrides both.
DEFAULT_WAIT_TIMEOUT = 10
DEFAULT_UPLOAD_TIMEOUT = 60

# Attempts per row. The second one gets a fresh browser.
DEFAULT_MAX_ATTEMPTS = 2

# Kendo rebuilds its dialogs after each interaction and will drop a click
# that lands too soon after the previous one. Found the hard way in PCSO911.
KENDO_PAUSE_SECONDS = 1

# Sign-in page
LOGIN_PROVIDER_SELECT_ID = "LoginProvider"
LOGIN_PROVIDER_LOCAL = "Local"
USERNAME_FIELD_ID = "Username"
PASSWORD_FIELD_ID = "Password"
SUBMIT_LOGIN_BUTTON_ID = "submitLogin"

# Present only once signed in, so it doubles as the proof that sign-in worked
CASES_SIDEBAR_CSS = "[data-menuid='incident']"

# Cases search
SEARCH_CRITERIA_DROPDOWN_CSS = "button[role='button'][aria-label='select']"

# The open criteria dropdown. Kendo keeps closed listboxes in the DOM with
# display:none, so the style check is what distinguishes the one on screen.
SEARCH_CRITERIA_LISTBOX_OPEN_XPATH = (
    "//ul[@id='incidentsSearchMainCriteria_listbox' and "
    "not(contains(@style,'display: none'))]"
)
SEARCH_CRITERIA_OPTIONS_CSS = "ul#incidentsSearchMainCriteria_listbox span"

# What the Case Number criterion might be called. STAC's own label is not
# known, so each is tried in turn and the first that the dropdown offers
# wins, the same approach TYPE_INPUT_SELECTORS takes to the Type field. The
# options the dropdown really has are logged when none of these match, so the
# right label can be read off a run rather than out of Inspect.
DEFAULT_UCN_CRITERIA = ("UCN",)
DEFAULT_CASE_NUMBER_CRITERIA = (
    "Case Number",
    "Case No",
    "Case #",
    "Agency Case Number",
    "Agency Case No",
    "Local Case Number",
)
SEARCH_FIELD_ID = "incidentsSearchMainSearchValue"
SEARCH_BUTTON_ID = "incidentsSearchMainButton"
NO_RECORDS_CSS = ".k-grid-norecords-template"
CASE_DEFENDANT_NAME_CSS = "td[data-original-column-name='Def_Name'] span.k-button-text"
IMAGES_TAB_ID = "incidentsTab-tab-3"

# Add image: dropzone and tiles
IMAGE_TILE_UNSELECTED_CSS = "#image-manager-listview-name .cipimage:not(.k-selected)"
DROPZONE_PANEL_CSS = ".pagesImagesIndex-upload-drop-zone-element"
PREVIEW_IFRAME_CSS = "#imageTabPageSplitterRightPane iframe.iframe-document"
FILE_INPUT_CSS = (
    "input[id^='cipFileUpload_pagesImagesIndex-upload'][multiple]:not([webkitdirectory])"
)
UPLOAD_SUCCESS_XPATH = (
    "//span[contains(@class,'k-file-validation-message') and "
    "text()='File(s) uploaded successfully.']"
)

# Type/Subtype matrix dialog
MATRIX_FIND_BUTTON_CSS = ".c-button-find-type-subtype"
MATRIX_DIALOG_CSS = "div#codeSearchDialog"
MATRIX_SEARCH_INPUT_CSS = "div#codeSearchDialog input.k-input-inner[placeholder='Search...']"
MATRIX_SELECT_BUTTON_ID = "SelectCodeAndSubCode"
SUBTYPE_INPUT_SELECTOR = "input[name='image_sub_type']"

# Finds the matrix row by both of its cell values and tells Kendo's grid
# widget that this is the selected row, then reads back what the grid now
# says is selected.
#
# Clicking the row and waiting for the k-selected class was not enough. The
# Select button reads the grid's own selection, which a click does not always
# move, so Select kept taking the first row in the list: a search for NOTICES
# lists APPEALS/NOTICES above COURT/NOTICES, and APPEALS is what got filed.
# Setting the selection through the widget and checking it here means the
# wrong row is caught before Select is pressed rather than after.
SELECT_MATRIX_ROW_JS = r"""
var wantedType = arguments[0].toUpperCase();
var wantedSub = arguments[1].toUpperCase();
var dialog = document.querySelector(arguments[2]);
if (!dialog) { return {ok: false, why: 'no-dialog'}; }

// Grouping adds a leading <td class="k-group-cell"> to every data row that
// has no counterpart in the header, so raw cell positions do not line up
// with header positions. Dropping those cells on both sides is what makes
// the indexes mean the same thing.
function isGroupCell(cell) {
  var cls = cell.className || '';
  return cls.indexOf('k-group-cell') !== -1
      || cls.indexOf('k-table-group-td') !== -1
      || cls.indexOf('k-hierarchy-cell') !== -1;
}
function cellsOf(tr) {
  var out = [];
  for (var c = 0; c < tr.cells.length; c++) {
    if (isGroupCell(tr.cells[c])) { continue; }
    out.push((tr.cells[c].innerText || tr.cells[c].textContent || '')
      .trim().toUpperCase().replace(/\s+/g, ' '));
  }
  return out;
}
function isHeaderRow(tr) {
  for (var c = 0; c < tr.cells.length; c++) {
    if (tr.cells[c].tagName === 'TH') { return true; }
  }
  return false;
}

// The dialog is four columns: Image Type, Image Type Desc, Image Sub Type,
// Image Sub Type Desc. Find the two code columns by their headers, because
// matching on "some cell equals PLS RVW" would also hit a description that
// happens to read the same as another pair's code.
var typeAt = -1, subAt = -1, seen = 0;
var headers = dialog.querySelectorAll('th');
for (var h = 0; h < headers.length; h++) {
  if (isGroupCell(headers[h])) { continue; }
  var label = (headers[h].innerText || headers[h].textContent || '')
    .trim().toUpperCase().replace(/\s+/g, ' ');
  if (label === 'IMAGE TYPE') { typeAt = seen; }
  else if (label === 'IMAGE SUB TYPE') { subAt = seen; }
  seen += 1;
}

function isMatch(tr) {
  if (!tr.cells || tr.cells.length === 0) { return false; }
  if (isHeaderRow(tr)) { return false; }
  // Hidden rows count. Kendo keeps a collapsed group's rows in the DOM, and
  // a row being out of sight does not stop the grid selecting it.
  var text = cellsOf(tr);
  if (typeAt >= 0 && subAt >= 0 && text.length > Math.max(typeAt, subAt)) {
    return text[typeAt] === wantedType && text[subAt] === wantedSub;
  }
  return text.indexOf(wantedType) !== -1 && text.indexOf(wantedSub) !== -1;
}

var matches = [];
var rows = dialog.querySelectorAll('tr');
for (var i = 0; i < rows.length; i++) {
  rows[i].removeAttribute('data-rapsheet-pick');
  if (isMatch(rows[i])) { matches.push(rows[i]); }
}
if (matches.length === 0) {
  return {ok: false, why: 'no-row', scanned: rows.length,
          columns: {type: typeAt, subtype: subAt}};
}

var tr = matches[0];
// Marked so Python can click this exact row without having to guess how
// STAC wraps its cell text. The previous version looked for <span> with
// exact text, which only works while every column renders that way.
tr.setAttribute('data-rapsheet-pick', '1');
var gridEl = tr.closest ? tr.closest('.k-grid') : null;
var grid = (gridEl && window.jQuery) ? window.jQuery(gridEl).data('kendoGrid') : null;
if (!grid) { return {ok: false, why: 'no-grid', matches: matches.length}; }

try { grid.clearSelection(); } catch (e) {}
grid.select(tr);
try { grid.trigger('change'); } catch (e) {}

var chosen = grid.select();
var chosenRow = (chosen && chosen.length) ? chosen[0] : null;
return {
  ok: chosenRow ? isMatch(chosenRow) : false,
  why: 'set',
  matches: matches.length,
  cells: chosenRow ? cellsOf(chosenRow) : []
};
"""

# The row the pick script marked, so a native click can land on that exact row.
MATRIX_PICKED_ROW_CSS = "tr[data-rapsheet-pick='1']"

# Puts every pair on one page before anything goes looking for a row.
#
# The dialog arrives grouped by Image Type and paged at about a hundred rows,
# listed alphabetically from APPEALS. COURT lands on page one, which is the
# only reason COURT/ORDERS ever worked. PLS does not, so PLS/RVW was not
# missing from STAC, it was missing from the page being read.
#
# Grouping goes too, because a collapsed group hides its rows and the grid
# then has nothing visible to hand the Select button.
PREPARE_MATRIX_GRID_JS = r"""
var dialog = document.querySelector(arguments[0]);
if (!dialog) { return {ok: false, why: 'no-dialog'}; }
if (!window.jQuery) { return {ok: false, why: 'no-jquery'}; }

var gridEl = dialog.querySelector('.k-grid');
var grid = gridEl ? window.jQuery(gridEl).data('kendoGrid') : null;
if (!grid || !grid.dataSource) { return {ok: false, why: 'no-grid'}; }

var ds = grid.dataSource;
var report = {
  ok: true,
  total: ds.total(),
  pageSizeWas: ds.pageSize() || null,
  grouped: false,
  unpaged: false
};

try {
  var groups = ds.group();
  if (groups && groups.length) {
    report.grouped = true;
    ds.group([]);
  }
} catch (e) { report.groupError = String(e); }

try {
  var total = ds.total();
  var size = ds.pageSize();
  if (total && size && size < total) {
    ds.pageSize(total);
    report.unpaged = true;
  }
} catch (e) { report.pageError = String(e); }

report.pageSizeNow = ds.pageSize() || null;
report.rowsRendered = gridEl.querySelectorAll('tbody tr').length;
return report;
"""

# How many of the dialog's rows to name in an error message. Enough to see
# what the search actually matched, not enough to bury the log.
MATRIX_ROW_SAMPLE = 25

# Reads back what the dialog is listing right now, for the error message when
# the wanted row is not among them.
DESCRIBE_MATRIX_ROWS_JS = r"""
var wantedType = (arguments[1] || '').toUpperCase();
var wantedSub = (arguments[2] || '').toUpperCase();
var dialog = document.querySelector(arguments[0]);
if (!dialog) { return null; }

function isGroupCell(cell) {
  var cls = cell.className || '';
  return cls.indexOf('k-group-cell') !== -1
      || cls.indexOf('k-table-group-td') !== -1
      || cls.indexOf('k-hierarchy-cell') !== -1;
}
function cellsOf(tr) {
  var out = [];
  for (var c = 0; c < tr.cells.length; c++) {
    if (isGroupCell(tr.cells[c])) { continue; }
    out.push((tr.cells[c].innerText || tr.cells[c].textContent || '')
      .trim().toUpperCase().replace(/\s+/g, ' '));
  }
  return out;
}
function isHeaderRow(tr) {
  for (var c = 0; c < tr.cells.length; c++) {
    if (tr.cells[c].tagName === 'TH') { return true; }
  }
  return false;
}

var typeAt = -1, subAt = -1, seen = 0;
var headers = dialog.querySelectorAll('th');
for (var h = 0; h < headers.length; h++) {
  if (isGroupCell(headers[h])) { continue; }
  var label = (headers[h].innerText || headers[h].textContent || '')
    .trim().toUpperCase().replace(/\s+/g, ' ');
  if (label === 'IMAGE TYPE') { typeAt = seen; }
  else if (label === 'IMAGE SUB TYPE') { subAt = seen; }
  seen += 1;
}

var out = {total: 0, types: {}, subtypesOfWantedType: [], typesOfWantedSub: [],
           nearPairs: [], columns: {type: typeAt, subtype: subAt}};
var rows = dialog.querySelectorAll('tr');
for (var i = 0; i < rows.length; i++) {
  var tr = rows[i];
  if (!tr.cells || tr.cells.length === 0) { continue; }
  if (isHeaderRow(tr)) { continue; }

  var text = cellsOf(tr);
  if (typeAt < 0 || subAt < 0 || text.length <= Math.max(typeAt, subAt)) { continue; }

  var t = text[typeAt];
  var s = text[subAt];
  if (!t && !s) { continue; }

  out.total += 1;
  out.types[t] = (out.types[t] || 0) + 1;
  if (t === wantedType) { out.subtypesOfWantedType.push(s); }
  if (s === wantedSub) { out.typesOfWantedSub.push(t); }
  // Near misses, so a pair whose real name only resembles the wanted one
  // says so instead of being reported as absent. This is what would have
  // named PLS RVW/PLS RVW the first time a search for PLS/RVW missed.
  var near = (wantedType && (t.indexOf(wantedType) !== -1 || wantedType.indexOf(t) !== -1))
          || (wantedSub && (s.indexOf(wantedSub) !== -1 || wantedSub.indexOf(s) !== -1));
  if (near && !(t === wantedType && s === wantedSub)) {
    out.nearPairs.push(t + '/' + s);
  }
}
out.typeList = Object.keys(out.types).sort();
return out;
"""

# The Type field. STAC's own naming is not certain here, so several
# spellings are tried and the first that answers wins. If none do, the type
# cannot be confirmed and the document is not saved: a search for ORDERS
# alone matches APPEALS/ORDERS, FEL APP/ORDERS and COURT/ORDERS, so checking
# only the subtype is how the wrong one gets filed.
TYPE_INPUT_SELECTORS = (
    "input[name='image_type']",
    "input[name='image_main_type']",
    "input[name='image_code']",
    "input[name='imageType']",
)

# Seconds to let an existing document's preview render before uploading.
# Short on purpose: it is a courtesy, not a requirement, and a case with no
# documents on it has nothing to show.
PREVIEW_WAIT_SECONDS = 5

# Kendo marks the chosen grid row with this class.
KENDO_SELECTED_CLASS = "k-selected"

# Save
SAVE_BUTTON_ID = "SaveImage"
SAVED_NOTIFICATION_XPATH = "//div[contains(@class,'c-notification-success')]"

# Words STAC hangs off a defendant name that are not part of the name: alert
# flags, programme codes, prosecutor initials. Carried over from PCSO911.
EXCLUDED_NAME_TOKENS = frozenset({
    "AM", "SVP", "AME", "SO", "AMSP", "JLA", "PJLA", "SP", "ALERT", "BKGRDALERT",
    "CP", "DO", "NOT", "USE", "GANG", "NCP", "NO", "CC", "OSCP", "SPCALERT",
    "TTP", "VFOSC", "HA",
})

# A name needs at least this many usable words before it is worth comparing.
# One word is a surname on its own, which matches far too much.
MIN_NAME_TOKENS = 2

# How alike two words must be to count as the same name, once an exact match
# has failed. Tuned in DALYN against both directions: EMERICK against EMRICK
# (0.92) and THOMPSON against THORNPSON (0.82) pass, while SMITH and SMYTHE
# (0.73), CARTER and CARTWRIGHT (0.67) and LEE and LEEDS (0.75) stay apart.
#
# The tolerance matters less here than it did in DALYN, because the name being
# compared comes from a spreadsheet cell rather than OCR, so the two spellings
# should already agree. It is kept because STAC's own name formatting varies:
# SMITH, JOHN A (ALERT) against Smith, John Allen.
#
# Known limit: ROBERT and ROBERTA score 0.92 and would pass.
NAME_TOKEN_SIMILARITY = 0.82


class SaveMayHaveHappened(SystemProblem):
    """Save was clicked but the confirmation never came, so nobody knows.

    The one failure that must never be retried blindly: the rap sheet may
    already be on the case. A person has to look.
    """


class StacSession:
    """One signed-in browser session, reused for every row in a run.

    Opened by the caller with `with StacSession(config) as stac:` so Chrome is
    closed even when a row blows up halfway through.

    The Type/Subtype pair is held on the session rather than passed per call,
    because it is one pair for the whole run and threading it through every
    method would only create a way for one row to use a different one.
    """

    def __init__(self, config: dict) -> None:
        """Read what is needed from config. Opens nothing yet.

        Args:
            config: The dict from load_config.
        """
        stac = config["stac"]
        self.url = str(stac["url"]).rstrip("/")
        self._username = stac["username"]
        self._password = stac["password"]
        self.document_type = stac["document_type"]
        self.subtype = stac["subtype"]
        self.is_test_instance = bool(stac.get("is_test_instance", True))
        self.upload_enabled = bool(stac.get("upload_enabled", False))
        self.save_enabled = bool(stac.get("save_enabled", False))
        self.check_name = bool(stac.get("check_name", True))
        self.wait_timeout = int(stac.get("wait_timeout", DEFAULT_WAIT_TIMEOUT))
        self.upload_timeout = int(stac.get("upload_timeout", DEFAULT_UPLOAD_TIMEOUT))
        self.max_attempts = max(1, int(stac.get("max_attempts", DEFAULT_MAX_ATTEMPTS)))
        self.action_pause = float(stac.get("action_pause", 0) or 0)
        self._chromedriver = _usable_driver_path(config["paths"].get("chromedriver"))

        # Criteria labels to look for in the search dropdown. Tuples because
        # STAC's wording for the case number is unknown and several spellings
        # get tried; a single string in config is accepted and wrapped.
        self.ucn_criteria = _as_tuple(stac.get("ucn_criteria"), DEFAULT_UCN_CRITERIA)
        self.case_number_criteria = _as_tuple(
            stac.get("case_number_criteria"), DEFAULT_CASE_NUMBER_CRITERIA
        )

        self.driver: webdriver.Chrome | None = None
        self.wait: WebDriverWait | None = None
        # STAC keeps the chosen search criterion, so the dropdown only needs
        # touching when it has to change. Reset whenever a new browser opens.
        self._criteria = ""

    @property
    def pair(self) -> str:
        """The Type/Subtype pair as one string, for log lines.

        Says so plainly when it is unset, which is allowed while uploads are
        off. A log line reading "under /" would look like a bug.
        """
        if not self.document_type and not self.subtype:
            return "(no Type/Subtype set)"
        return f"{self.document_type}/{self.subtype}"

    def __enter__(self) -> "StacSession":
        self.open()
        return self

    def __exit__(self, *_) -> None:
        self.close()

    def open(self) -> None:
        """Launch Chrome, load STAC, and sign in.

        Raises:
            SystemProblem: If Chrome will not start, STAC will not load, or
                the credentials are refused. None of these are a problem with
                any one row, so the run should stop.
        """
        logger.info(
            "Opening STAC at %s as %s (test instance: %s, uploading: %s, saving: %s)",
            self.url,
            self.pair,
            self.is_test_instance,
            self.upload_enabled,
            self.save_enabled,
        )

        self.driver = self._start_chrome()
        self._criteria = ""

        self.wait = WebDriverWait(self.driver, self.wait_timeout)

        try:
            self.driver.get(self.url)
            self.driver.maximize_window()
        except WebDriverException as error:
            self.close()
            raise SystemProblem(f"Could not load STAC at {self.url}: {error}") from error

        try:
            self._sign_in()
        except Exception:
            self.close()
            raise

    def _start_chrome(self) -> webdriver.Chrome:
        """Start Chrome, trying each way of finding chromedriver in turn.

        Order matters on a locked-down network: a path in config is used as
        given, then Selenium's own driver resolution, and only last the
        downloader, which needs to reach the internet and often cannot.

        Raises:
            SystemProblem: If none of them produce a working Chrome. The
                message lists what was tried.
        """
        attempts = []

        if self._chromedriver:
            attempts.append(("paths.chromedriver", lambda: Service(self._chromedriver)))

        # Selenium Manager, built into Selenium 4.6 and later. Uses a driver
        # already on the machine when there is one.
        attempts.append(("Selenium's own driver lookup", lambda: None))

        attempts.append(
            ("downloading chromedriver", lambda: Service(ChromeDriverManager().install()))
        )

        failures = []
        for description, build_service in attempts:
            try:
                service = build_service()
                driver = webdriver.Chrome(service=service) if service else webdriver.Chrome()
                logger.debug("Chrome started via %s", description)
                return driver
            except Exception as error:  # noqa: BLE001 - each route fails differently
                failures.append(
                    f"{description}: {type(error).__name__}: {str(error).splitlines()[0]}"
                )

        raise SystemProblem(
            "Could not start Chrome for STAC. Set paths.chromedriver in config.yaml "
            "to a chromedriver.exe matching your installed Chrome version. Tried: "
            + " | ".join(failures)
        )

    def _sign_in(self) -> None:
        """Fill the sign-in form and wait for the Cases sidebar to appear.

        Raises:
            SystemProblem: If any step of sign-in does not complete.
        """
        try:
            provider = self.wait.until(
                EC.presence_of_element_located((By.ID, LOGIN_PROVIDER_SELECT_ID))
            )
            Select(provider).select_by_value(LOGIN_PROVIDER_LOCAL)
        except TimeoutException as error:
            raise SystemProblem(
                f"No 'Authenticate Using' dropdown on the sign-in page at {self.url}. "
                "Either the page has changed or the url is not STAC."
            ) from error

        try:
            self.wait.until(
                EC.element_to_be_clickable((By.ID, USERNAME_FIELD_ID))
            ).send_keys(self._username)
            self.wait.until(
                EC.element_to_be_clickable((By.ID, PASSWORD_FIELD_ID))
            ).send_keys(self._password)
        except TimeoutException as error:
            raise SystemProblem("Could not find the STAC username or password field.") from error

        try:
            submit = self.wait.until(
                EC.element_to_be_clickable((By.ID, SUBMIT_LOGIN_BUTTON_ID))
            )
            # Clicked through JavaScript because the real button sits under an
            # overlay often enough that a normal click intercepts.
            self.driver.execute_script("arguments[0].click();", submit)
        except TimeoutException as error:
            raise SystemProblem("Could not find the STAC sign-in button.") from error

        try:
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, CASES_SIDEBAR_CSS)))
        except TimeoutException as error:
            # The sidebar is the only reliable proof of a successful sign-in:
            # STAC re-renders the same page on a bad password rather than
            # saying so in a way Selenium can read.
            raise SystemProblem(
                "Signed in but the Cases sidebar never appeared. Usually a wrong "
                "username or password in Credential Manager. Rerun store_secret.py. Or an account without access "
                "to this instance."
            ) from error

        logger.info("Signed in to STAC")
        self._pause("signed in")

    def close(self) -> None:
        """Close Chrome. Safe to call twice, and never raises."""
        if self.driver is None:
            return
        try:
            self.driver.quit()
        except WebDriverException as error:
            logger.warning("STAC's Chrome did not close cleanly: %s", error)
        finally:
            self.driver = None
            self.wait = None

    @staticmethod
    def _settle() -> None:
        """Give Kendo a beat to finish rebuilding before the next click."""
        time.sleep(KENDO_PAUSE_SECONDS)

    def _pause(self, what: str) -> None:
        """Hold after a step so a person can see it, when action_pause is set.

        Purely for watching. Nothing in STAC needs it, and production runs
        with action_pause at 0.
        """
        if self.action_pause <= 0:
            return
        logger.info("  ... %s", what)
        time.sleep(self.action_pause)

    # ---------------------------------------------------------------- search

    def find_case(
        self, ucn: str, case_number: str = "", expected_name: str = ""
    ) -> str:
        """Find the case by UCN, falling back to its case number, then open Images.

        The sheet carries two identifiers for the same case and they fail
        independently. A UCN is long and hand-typed, so it is the one more
        likely to be wrong; the case number is shorter and comes off a
        different system. Trying the second when the first finds nothing turns
        a row that would be reported as "not in STAC" into a row that files
        correctly, and it costs an extra search only on rows that were about to
        fail anyway.

        The UCN goes first because it is unique statewide and it is the search
        both PCSO911 and DALYN already prove in production.

        Args:
            ucn: The UCN from the row. May be "" if the sheet has none.
            case_number: The case number from the row, used only if the UCN
                finds nothing. May be "".
            expected_name: The name from the same row, as written in the sheet.
                This is the name CCIS was searched for, so it is also whose rap
                sheet the PDF holds. Pass "" to skip the check.

        Returns:
            The defendant name STAC shows for the case.

        Raises:
            RowProblem: Neither identifier found a case, or STAC's defendant
                does not match the row's name. Either way this row needs a
                person, and the rest of the run carries on.
            SystemProblem: STAC itself did not respond as expected.
        """
        attempts = [
            ("UCN", ucn, self.ucn_criteria),
            ("case number", case_number, self.case_number_criteria),
        ]
        attempts = [attempt for attempt in attempts if attempt[1]]

        if not attempts:
            raise RowProblem(
                "This row has neither a UCN nor a case number, so there is nothing "
                "to find the case with."
            )

        tried = []
        for position, (label, value, criteria) in enumerate(attempts):
            try:
                self._open_case_search(criteria)
            except SystemProblem:
                if position == 0:
                    # The primary identifier's criterion is not selectable, so
                    # nothing can be searched at all. That is STAC being wrong,
                    # not this row.
                    raise
                # A fallback whose criterion STAC does not offer. Logged once
                # by _choose_criteria with the labels the dropdown really has.
                # Giving up on the row is right; stopping the run is not, so
                # this must not stay a SystemProblem.
                logger.info(
                    "Cannot search by %s in this instance, so %s is the only "
                    "identifier available for this row.",
                    label, tried[0][0] if tried else "the UCN",
                )
                break

            self._pause(f"case search open, searching by {label}")

            stac_name = self._run_search(label, value)
            if stac_name is not None:
                if tried:
                    logger.info(
                        "Found the case by %s %s after the %s found nothing",
                        label, value, " and ".join(name for name, _ in tried),
                    )
                self._pause(f"found {value}")
                self._check_defendant(value, stac_name, expected_name)
                self._open_images_tab(value)
                self._pause("images tab open")
                return stac_name

            tried.append((label, value))
            logger.info("No case in STAC for %s %s", label, value)

        described = ", ".join(f"{label} {value}" for label, value in tried)
        raise RowProblem(
            f"No case in STAC matching {described}. Either the identifiers in the "
            "spreadsheet are wrong or the case is not in this instance."
        )

    def _open_case_search(self, criteria: tuple) -> None:
        """Get to a usable Cases search box set to one of `criteria`.

        Clicking the Cases sidebar is enough from most pages, but not from a
        case that has been opened: STAC leaves the search box present and
        disabled, and typing into it fails with "element is not currently
        interactable". PCSO911 never meets this because it opens a new browser
        for every case.

        So the sidebar is tried first, and if the box is not usable the page is
        loaded fresh, which always works.

        Args:
            criteria: Labels to look for in the criteria dropdown, in order.
                The first one STAC offers is selected.

        Raises:
            SystemProblem: If the search box cannot be reached either way, or
                if none of `criteria` is in the dropdown. The message lists
                what the dropdown does offer.
        """
        if self._try_open_case_search(criteria):
            return

        logger.info("Search box was not usable; reloading STAC to clear the page")
        try:
            self.driver.get(self.url)
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, CASES_SIDEBAR_CSS)))
        except (TimeoutException, WebDriverException) as error:
            raise SystemProblem(f"Could not get back to STAC's Cases page: {error}") from error

        # A reload resets the search criteria, so it has to be chosen again.
        self._criteria = ""

        if not self._try_open_case_search(criteria):
            raise SystemProblem(
                "STAC's case search box is still not usable after reloading the "
                f"page, or it does not offer any of {list(criteria)}. The dropdown "
                f"offers: {self._criteria_on_offer()}."
            )

    def _try_open_case_search(self, criteria: tuple) -> bool:
        """One attempt at reaching a usable search box. False if it is not ready."""
        try:
            self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, CASES_SIDEBAR_CSS))
            ).click()
        except (TimeoutException, WebDriverException):
            return False

        # STAC keeps whatever criterion was last chosen, so the dropdown only
        # needs touching when the wanted one is not already set.
        if self._criteria not in criteria and not self._choose_criteria(criteria):
            return False

        # Present is not the same as usable. STAC leaves the box on the page in
        # a disabled state after a case is opened, and only this tells the
        # difference.
        try:
            field = self.wait.until(EC.element_to_be_clickable((By.ID, SEARCH_FIELD_ID)))
            return field.is_enabled() and field.is_displayed()
        except (TimeoutException, WebDriverException):
            return False

    def _choose_criteria(self, criteria: tuple) -> bool:
        """Set the search dropdown to the first of `criteria` that STAC offers.

        False if the dropdown will not open or offers none of them. Several
        labels are tried because STAC's wording for the case number criterion
        is not known from the two existing tools, which only ever use UCN.
        """
        try:
            dropdown = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, SEARCH_CRITERIA_DROPDOWN_CSS))
            )
            self.driver.execute_script("arguments[0].click();", dropdown)
        except (TimeoutException, WebDriverException):
            return False

        for label in criteria:
            option_xpath = (
                f"{SEARCH_CRITERIA_LISTBOX_OPEN_XPATH}"
                f"//span[text()={_xpath_literal(label)}]"
            )
            try:
                option = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, option_xpath))
                )
                self.driver.execute_script("arguments[0].click();", option)
            except (TimeoutException, WebDriverException):
                continue

            self._criteria = label
            logger.debug("Search criteria set to %r", label)
            return True

        logger.warning(
            "STAC's search dropdown offers none of %s. It offers: %s. Set "
            "stac.case_number_criteria in config.yaml to the right label.",
            list(criteria), self._criteria_on_offer(),
        )
        self._close_dropdown()
        return False

    def _criteria_on_offer(self) -> str:
        """The labels STAC's criteria dropdown actually lists, for an error message.

        Reading them is what turns "the Case Number option is not there" into a
        one-line config fix, instead of someone opening Inspect on a production
        system to find out what the option is called.
        """
        try:
            spans = self.driver.find_elements(
                By.CSS_SELECTOR, SEARCH_CRITERIA_OPTIONS_CSS
            )
            labels = [
                text for text in (span.text.strip() for span in spans) if text
            ]
        except WebDriverException:
            return "(could not read the dropdown)"

        if not labels:
            return "(the dropdown listed nothing)"
        return ", ".join(dict.fromkeys(labels))

    def _close_dropdown(self) -> None:
        """Press Escape so a dropdown left open cannot block the next click."""
        try:
            self.driver.switch_to.active_element.send_keys(Keys.ESCAPE)
        except WebDriverException:
            pass

    def _run_search(self, label: str, value: str) -> str | None:
        """Search for one identifier. Returns STAC's defendant name, or None.

        None means the search ran correctly and matched no case, which is an
        ordinary answer: it is what lets find_case try the other identifier
        rather than giving up. Anything that stops the search running at all
        is a SystemProblem.

        Raises:
            SystemProblem: If the search could not be run or the results never
                loaded.
        """
        try:
            field = self.wait.until(EC.element_to_be_clickable((By.ID, SEARCH_FIELD_ID)))
            field.clear()
            field.click()
            field.send_keys(value)
            button = self.driver.find_element(By.ID, SEARCH_BUTTON_ID)
            self.driver.execute_script("arguments[0].click();", button)
        except (TimeoutException, WebDriverException) as error:
            raise SystemProblem(
                f"Could not run the {label} search in STAC: {error}"
            ) from error

        # Wait for either outcome. find_elements returns a list, so this is
        # truthy on a result row and truthy on "no records", and keeps waiting
        # while the grid is still loading.
        try:
            self.wait.until(
                lambda d: d.find_elements(By.CSS_SELECTOR, NO_RECORDS_CSS)
                or d.find_elements(By.CSS_SELECTOR, CASE_DEFENDANT_NAME_CSS)
            )
        except TimeoutException as error:
            raise SystemProblem(
                f"STAC's search results never loaded for {label} {value}."
            ) from error

        if self.driver.find_elements(By.CSS_SELECTOR, NO_RECORDS_CSS):
            return None

        try:
            return self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, CASE_DEFENDANT_NAME_CSS))
            ).text.strip()
        except TimeoutException as error:
            raise SystemProblem(
                f"Search for {label} {value} returned a row with no defendant name."
            ) from error

    def _check_defendant(self, ucn: str, stac_name: str, expected_name: str) -> None:
        """Compare STAC's defendant against the name in the spreadsheet row.

        This is the guard on the row itself. A UCN and a name sit in the same
        row and nothing has checked that they belong together; the names
        disagreeing is how a transcription error or a shifted column shows up,
        and getting it wrong means filing one person's criminal history onto
        another person's case.

        DALYN read this name out of the document's caption. That cannot work
        here, so the name comes from the sheet instead, the way PCSO911 takes
        it from the FTP directory name.

        Raises:
            RowProblem: If both names are readable and they do not match.
        """
        if not self.check_name:
            logger.warning(
                "%s: name check is switched off in config, filing without it", ucn
            )
            return

        if not expected_name:
            # Blank rather than skipped silently. Unlike DALYN's captions,
            # which are legitimately absent from plenty of documents, a row
            # with no name should not have got this far: CCIS was searched by
            # name, so an empty one means something upstream went wrong.
            raise RowProblem(
                f"No name in the spreadsheet row for {ucn}, so there is nothing to "
                "check STAC's defendant against. Filing without that check could "
                "put this rap sheet on the wrong person's case."
            )

        if len(_name_tokens(expected_name)) < MIN_NAME_TOKENS:
            raise RowProblem(
                f"The name {expected_name!r} in the row for {ucn} is only one usable "
                "word, which matches too many defendants to be a real check."
            )

        if names_match(stac_name, expected_name):
            logger.info(
                "%s: defendant matches, STAC %r / sheet %r", ucn, stac_name, expected_name
            )
            return

        # Both names are in the message because without them nobody can tell a
        # real mismatch from a formatting difference. They are already in the
        # log by way of the PDF filenames, so this is not new exposure.
        raise RowProblem(
            f"Defendant mismatch on {ucn}: STAC says {stac_name!r}, the spreadsheet "
            f"row says {expected_name!r}. Comparing words "
            f"{sorted(_name_tokens(stac_name))} against "
            f"{sorted(_name_tokens(expected_name))}."
        )

    def _open_images_tab(self, ucn: str) -> None:
        """Open the case's Images tab, where documents are added."""
        try:
            tab = self.wait.until(EC.element_to_be_clickable((By.ID, IMAGES_TAB_ID)))
            self.driver.execute_script("arguments[0].click();", tab)
        except TimeoutException as error:
            raise SystemProblem(f"Could not open the Images tab for {ucn}.") from error

    # ------------------------------------------------------------- add image

    def add_documents(self, ucn: str, paths: list) -> None:
        """Upload files to the open case under the run's one Type/Subtype.

        A STAC upload box carries a single Type and Subtype for everything in
        it, which is exactly what this run needs: every rap sheet files under
        the same pair, so several rows that share a case number can go up
        together in one box.

        Args:
            ucn: Case number, for log lines only. The case is already open.
            paths: Absolute local paths. Usually one, the row's PDF.

        Raises:
            RowProblem: The Type/Subtype is not in STAC's matrix.
            SaveMayHaveHappened: Save was clicked and never confirmed.
            SystemProblem: Anything else in STAC misbehaved.
        """
        names = ", ".join(Path(path).name for path in paths)

        if not self.upload_enabled:
            logger.info(
                "%s: WOULD upload %s under %s and press Save. Nothing sent, "
                "stac.upload_enabled is false.",
                ucn, names, self.pair,
            )
            return

        self._open_dropzone()
        self._pause("dropzone open")
        self._upload(paths)
        self._pause(f"uploaded {names}")
        self._select_type_subtype()

        if not self.save_enabled:
            self._reach_save_without_pressing(ucn, names)
            return

        self._save(ucn, names)

    def _reach_save_without_pressing(self, ucn: str, names: str) -> None:
        """Go as far as the Save button, prove it is ready, and leave it alone.

        Everything a real save does except the click: the Subtype has landed in
        the form, the button exists and is enabled. If this passes, turning
        save_enabled on should work.

        The upload is then thrown away by reloading the page, so the next row
        does not start on a half-filled form.
        """
        upload_wait = WebDriverWait(self.driver, self.upload_timeout)

        try:
            upload_wait.until(
                lambda d: self._read_kendo_value(SUBTYPE_INPUT_SELECTOR) == self.subtype
            )
        except TimeoutException as error:
            raise SystemProblem(
                f"The Subtype field never showed {self.subtype}, so Save would have "
                "filed this under the wrong type."
            ) from error

        try:
            button = upload_wait.until(EC.element_to_be_clickable((By.ID, SAVE_BUTTON_ID)))
        except TimeoutException as error:
            raise SystemProblem("The Save button never became clickable.") from error

        self._pause("at the Save button, not pressing it")

        logger.info(
            "%s: REACHED SAVE for %s under %s. Button %r is ready and was NOT "
            "pressed, stac.save_enabled is false.",
            ucn, names, self.pair, button.text.strip() or SAVE_BUTTON_ID,
        )

        self._discard_pending_upload()

    def _discard_pending_upload(self) -> None:
        """Reload the page so an unsaved upload cannot bleed into the next one."""
        try:
            self.driver.refresh()
            try:
                alert = self.driver.switch_to.alert
                alert.accept()
            except NoAlertPresentException:
                pass
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, CASES_SIDEBAR_CSS)))
        except (TimeoutException, WebDriverException) as error:
            raise SystemProblem(
                f"Could not clear the unsaved upload before the next row: {error}"
            ) from error

    def _open_dropzone(self) -> None:
        """Get the Images tab into the state where files can be dropped.

        Three steps, and only the middle one is required.

        Clicking an existing image tile is what makes STAC render the dropzone,
        but a case with no documents on it yet has no tile to click, and the
        dropzone is there anyway.

        Waiting for the preview iframe is PCSO911's guard against the file
        input being swapped out mid-upload. It only applies when there is
        something to preview, so a case with no images, or one whose preview
        will not render, waits briefly and moves on rather than failing.

        Raises:
            SystemProblem: If the dropzone itself never appears.
        """
        tiles = self.driver.find_elements(By.CSS_SELECTOR, IMAGE_TILE_UNSELECTED_CSS)
        if tiles:
            try:
                self.driver.execute_script("arguments[0].click();", tiles[0])
            except WebDriverException as error:
                logger.debug("Could not click an image tile: %s", error)
        else:
            logger.info("No existing images on this case, going straight to the dropzone")

        try:
            self.wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, DROPZONE_PANEL_CSS))
            )
        except TimeoutException as error:
            raise SystemProblem("The upload dropzone never appeared.") from error

        self._wait_for_preview()

    def _wait_for_preview(self) -> None:
        """Let an existing document's preview settle, if there is one.

        Not required. When a preview is loading, letting it finish stops STAC
        replacing the file input underneath an upload already in progress. When
        there is nothing to preview, waiting the full timeout would stall every
        row on an empty case.
        """
        brief = WebDriverWait(self.driver, PREVIEW_WAIT_SECONDS)
        try:
            brief.until(
                lambda d: "web/viewer.html?file="
                in (
                    d.find_element(By.CSS_SELECTOR, PREVIEW_IFRAME_CSS).get_attribute("src")
                    or ""
                )
            )
        except (TimeoutException, WebDriverException):
            logger.info("No document preview to wait for, continuing")

    def _upload(self, paths: list) -> None:
        """Send the files to the dropzone and wait for every one to confirm."""
        try:
            file_input = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, FILE_INPUT_CSS))
            )
            # The input is hidden behind a class; Selenium cannot type into it
            # until that is removed.
            self.driver.execute_script("arguments[0].removeAttribute('class')", file_input)
        except TimeoutException as error:
            raise SystemProblem("Could not find STAC's file input.") from error

        upload_wait = WebDriverWait(self.driver, self.upload_timeout)
        try:
            file_input.send_keys("\n".join(str(path) for path in paths))
            upload_wait.until(
                lambda d: len(d.find_elements(By.XPATH, UPLOAD_SUCCESS_XPATH)) >= len(paths)
            )
        except TimeoutException as error:
            raise SystemProblem(
                f"Only some of the {len(paths)} file(s) finished uploading within "
                f"{self.upload_timeout}s."
            ) from error

        self._remove_preview_iframe()

    def _remove_preview_iframe(self) -> None:
        """Drop the preview iframe, which otherwise intercepts the Save click."""
        self.driver.execute_script(
            f"var f=document.querySelector({PREVIEW_IFRAME_CSS!r}); if(f) f.remove();"
        )

    def _select_type_subtype(self) -> None:
        """Open the matrix dialog, select the run's Type/Subtype row, and verify.

        PCSO911 had 911AUDIO written into the row XPath, the search box and the
        save check. DALYN took the pair from a rules sheet per document. Here
        it is one pair for the run, held on the session, but the dialog is just
        as awkward either way and all of DALYN's handling of it is kept.

        Several tries, for two different failures. One is the dialog taking the
        wrong row, which a reopen and a double-click fix. The other is the row
        not turning up at all for a given search term: STAC's search box may be
        matching the pair's description rather than its code, so a search for
        RVW finds nothing while the PLS/RVW row is sitting right there. Each
        term gets a go before the pair is called missing.

        Raises:
            RowProblem: If no search term turns the row up. The message lists
                what the dialog did show, since "not in STAC" and "not found by
                this search" look identical from here and are not the same
                thing at all.
            SystemProblem: If the row was found but the form ended up holding
                something else.
        """
        document_type, subtype = self.document_type, self.subtype
        terms = self._matrix_search_terms()
        last_error: SystemProblem | None = None
        last_missing: RowProblem | None = None

        for attempt, term in enumerate(terms, start=1):
            self._open_matrix_dialog(term)
            try:
                self._choose_matrix_row(use_double_click=(attempt > 1))
                self._press_matrix_select()
                self._confirm_type_subtype()
            except RowProblem as error:
                # Not under this search term. Another term may still find it,
                # so this is not yet grounds for calling the pair missing.
                last_missing = error
                logger.info(
                    "%s not among the %s row(s) the dialog showed for search %r. "
                    "Trying the next search term.",
                    self.pair, self._matrix_row_count(), term or "(cleared)",
                )
                self._close_matrix_dialog()
                continue
            except SystemProblem as error:
                last_error = error
                logger.warning(
                    "%s did not take on try %s of %s (%s). Reopening the matrix.",
                    self.pair, attempt, len(terms), error,
                )
                self._close_matrix_dialog()
                continue

            if attempt > 1:
                logger.info(
                    "Found %s by searching the dialog for %r, not %r",
                    self.pair, term, subtype,
                )
            self._pause(f"{self.pair} selected")
            logger.debug("Selected %s", self.pair)
            return

        # A row that was found but would not take is the more serious failure,
        # because something is wrong with the dialog rather than with the pair.
        if last_error is not None:
            raise last_error
        raise last_missing

    def _matrix_search_terms(self) -> tuple[str, ...]:
        """Search terms to try in the matrix dialog, in order.

        The subtype first, since that is the narrowest and works for most
        pairs. Then the type. Then nothing at all, which clears the filter and
        leaves every pair listed for the row scan to walk.
        """
        terms = [self.subtype]
        if self.document_type and self.document_type != self.subtype:
            terms.append(self.document_type)
        terms.append("")
        return tuple(terms)

    def _open_matrix_dialog(self, term: str) -> None:
        """Press Find and put `term` in the dialog's search box.

        An empty term clears the box, which leaves the grid unfiltered.
        """
        try:
            find_button = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, MATRIX_FIND_BUTTON_CSS))
            )
            self.driver.execute_script("arguments[0].click();", find_button)
        except TimeoutException as error:
            raise SystemProblem("Could not open the Type/Subtype dialog.") from error

        try:
            search_box = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, MATRIX_SEARCH_INPUT_CSS))
            )
            # Real key events, not clear(). Kendo filters on keyup, and a
            # clear() that sets the value directly leaves the grid showing the
            # previous term's results.
            search_box.send_keys(Keys.CONTROL, "a")
            search_box.send_keys(Keys.DELETE)
            if term:
                search_box.send_keys(term)
                # Enter as well. Typing alone did not filter anything: the
                # unfiltered alphabetical list came back every time, which is
                # how a search for ORDERS ended up on APPEALS/ORDERS.
                search_box.send_keys(Keys.ENTER)
        except TimeoutException as error:
            raise SystemProblem(
                "Could not find the search box in the Type/Subtype dialog."
            ) from error

        # Kendo rebuilds the rows after a filter.
        self._settle()
        self._flatten_matrix_grid()

    def _flatten_matrix_grid(self) -> None:
        """Drop the dialog's grouping and paging so every pair is on the page.

        Best effort. If the grid widget cannot be reached the row scan still
        runs against whatever is rendered, which is the behaviour this had
        before, so there is nothing to gain by failing here.
        """
        try:
            report = self.driver.execute_script(
                PREPARE_MATRIX_GRID_JS, MATRIX_DIALOG_CSS
            ) or {}
        except WebDriverException as error:
            logger.debug("Could not flatten the matrix grid: %s", error)
            return

        if not report.get("ok"):
            logger.debug("Did not flatten the matrix grid: %s", report.get("why"))
            return

        if report.get("grouped") or report.get("unpaged"):
            logger.debug(
                "Matrix dialog flattened: %s pairs, page size %s -> %s, %s rows rendered",
                report.get("total"),
                report.get("pageSizeWas"),
                report.get("pageSizeNow"),
                report.get("rowsRendered"),
            )
            # Re-rendering every row takes a moment longer than a filter.
            self._settle()

    def _close_matrix_dialog(self) -> None:
        """Get the dialog off the screen so the next try starts from Find again.

        Escape is enough when it is still open, and does nothing harmful when
        Select already closed it.
        """
        self._close_dropdown()
        self._settle()

    def _choose_matrix_row(self, use_double_click: bool = False) -> None:
        """Make the Type/Subtype row the grid's selected row.

        The row is found by its cell text, not by an XPath looking for exact
        <span> text. The XPath version could only see rows whose cells happen
        to be wrapped the way COURT/ORDERS is, and reported anything else as
        "not in STAC", which is a much more alarming thing to be wrong about.

        Raises:
            RowProblem: No row in the dialog holds this pair, with a list of
                what the dialog did show.
            SystemProblem: The row is there but the grid would not select it.
        """
        document_type, subtype = self.document_type, self.subtype

        # Find, mark and select in one pass, so the row is located by cell text
        # rather than by markup shape.
        result = self._run_matrix_pick()

        if result.get("why") == "no-row":
            raise RowProblem(
                f"STAC's Type/Subtype dialog is not showing a row for {self.pair}. "
                f"{self._describe_matrix_rows()}"
            )

        if result.get("matches", 0) > 1:
            logger.debug(
                "The dialog showed %s rows holding %s and %s; took the first",
                result["matches"], document_type, subtype,
            )

        # Now a real click on that exact row, so STAC's own row handler runs
        # and anything it does to the form happens. Scripted clicks skip it.
        self._click_marked_row(use_double_click)

        # And again, because the click may have moved the selection.
        result = self._run_matrix_pick()
        if result.get("ok"):
            return

        why = result.get("why")
        if why == "no-grid":
            # No widget to drive, so fall back to the class on the row. Weaker,
            # but it is what PCSO911 relied on throughout.
            logger.warning(
                "Could not reach Kendo's grid widget in the matrix dialog; falling "
                "back to checking the row's own selected class."
            )
            try:
                self.wait.until(
                    lambda d: KENDO_SELECTED_CLASS
                    in (
                        d.find_element(By.CSS_SELECTOR, MATRIX_PICKED_ROW_CSS)
                        .get_attribute("class") or ""
                    )
                )
            except (TimeoutException, WebDriverException) as error:
                raise SystemProblem(
                    f"The {self.pair} row never became the selected row, so Select "
                    "would have taken a different one."
                ) from error
            return

        raise SystemProblem(
            f"Could not make {self.pair} the selected matrix row ({why}; grid reports "
            f"{result.get('cells')}). Select would have taken whichever row Kendo had "
            "current, which is the first one in the list."
        )

    def _run_matrix_pick(self) -> dict:
        """Find, mark and select the matching row. Returns the script's report."""
        try:
            return self.driver.execute_script(
                SELECT_MATRIX_ROW_JS, self.document_type, self.subtype, MATRIX_DIALOG_CSS
            ) or {}
        except WebDriverException as error:
            raise SystemProblem(
                f"Could not read the Type/Subtype dialog's rows: {error}"
            ) from error

    def _click_marked_row(self, use_double_click: bool) -> None:
        """Click the row the pick script marked. Tolerant: the JS already chose it."""
        try:
            row = self.driver.find_element(By.CSS_SELECTOR, MATRIX_PICKED_ROW_CSS)
            self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", row)
            if use_double_click:
                ActionChains(self.driver).double_click(row).perform()
            else:
                row.click()
        except WebDriverException:
            # Something was over it, or it moved. The script's own selection
            # still stands, so this is not worth failing over.
            try:
                self.driver.execute_script(
                    f"var r = document.querySelector({MATRIX_PICKED_ROW_CSS!r});"
                    "if (r && r.cells.length) r.cells[0].click();"
                )
            except WebDriverException:
                pass

        self._settle()

    def _read_matrix_rows(self) -> dict:
        """What the dialog is listing, keyed to the pair being looked for."""
        try:
            return self.driver.execute_script(
                DESCRIBE_MATRIX_ROWS_JS,
                MATRIX_DIALOG_CSS,
                self.document_type,
                self.subtype,
            ) or {}
        except WebDriverException:
            return {}

    def _matrix_row_count(self) -> int:
        """How many rows the dialog is listing. 0 if it cannot be read.

        Printed on every miss, because the row count is what says whether the
        search box filters at all: the same number every time means it does not.
        """
        return self._read_matrix_rows().get("total", 0)

    def _describe_matrix_rows(self) -> str:
        """Why the wanted row is not there, for an error a person reads.

        "No row for PLS/RVW" is the same sentence whether the pair is absent
        from STAC, the search box did not match it, or the grid is showing one
        page of many. Those need completely different fixes, so the message
        says which of them it is: whether the type exists at all, whether the
        subtype exists under some other type, and how many pairs are on the
        page to have been looked at.
        """
        document_type, subtype = self.document_type, self.subtype
        report = self._read_matrix_rows()
        if not report:
            return "(could not read the dialog's rows)"

        total = report.get("total", 0)
        if not total:
            return (
                "no pairs at all. Either the search matched nothing or the grid had "
                "not loaded."
            )

        under_type = report.get("subtypesOfWantedType") or []
        over_sub = report.get("typesOfWantedSub") or []
        types = report.get("typeList") or []

        parts = [f"{total} pair(s) across {len(types)} type(s)"]

        near = report.get("nearPairs") or []
        if near:
            parts.append(
                "closest real pairs: " + ", ".join(sorted(set(near))[:MATRIX_ROW_SAMPLE])
            )

        if under_type:
            parts.append(
                f"{document_type} exists, with subtypes "
                f"{', '.join(sorted(set(under_type))[:MATRIX_ROW_SAMPLE])}"
            )
        else:
            parts.append(f"no type called {document_type} on the page")

        if over_sub:
            parts.append(
                f"subtype {subtype} exists under "
                f"{', '.join(sorted(set(over_sub))[:MATRIX_ROW_SAMPLE])}"
            )
        else:
            parts.append(f"no subtype called {subtype} on the page")

        if not under_type and not over_sub:
            parts.append(
                f"types listed: {', '.join(types[:MATRIX_ROW_SAMPLE])}"
                + (
                    f" and {len(types) - MATRIX_ROW_SAMPLE} more"
                    if len(types) > MATRIX_ROW_SAMPLE
                    else ""
                )
            )

        return ". ".join(parts)

    def _press_matrix_select(self) -> None:
        """Press Select, after letting Kendo finish rebuilding the dialog."""
        self._settle()
        try:
            select_button = self.wait.until(
                EC.element_to_be_clickable((By.ID, MATRIX_SELECT_BUTTON_ID))
            )
            self.driver.execute_script("arguments[0].click();", select_button)
        except TimeoutException as error:
            raise SystemProblem("Could not press Select in the Type/Subtype dialog.") from error

    def _read_kendo_value(self, selector: str):
        """Return a Kendo dropdown's current value, or None if it is not there."""
        try:
            return self.driver.execute_script(
                "var w = $(arguments[0]).data('kendoDropDownList');"
                "return w ? w.value() : null;",
                selector,
            )
        except WebDriverException:
            return None

    def _confirm_type_subtype(self) -> None:
        """Check the form now holds the Type and Subtype that were asked for.

        The subtype alone is not enough. Searching the matrix for ORDERS brings
        back APPEALS/ORDERS, FEL APP/ORDERS and COURT/ORDERS, so a subtype
        check passes while the document is filed under the wrong type entirely.

        Raises:
            SystemProblem: If either field holds something else, or if the Type
                field cannot be found at all. Not knowing is treated the same
                as being wrong.
        """
        document_type, subtype = self.document_type, self.subtype
        upload_wait = WebDriverWait(self.driver, self.upload_timeout)

        try:
            upload_wait.until(
                lambda d: self._read_kendo_value(SUBTYPE_INPUT_SELECTOR) == subtype
            )
        except TimeoutException as error:
            found = self._read_kendo_value(SUBTYPE_INPUT_SELECTOR)
            raise SystemProblem(
                f"Subtype never became {subtype}; the form holds {found!r}."
            ) from error

        # Read every candidate, not just the first one that answers. The field
        # names here are guesses at STAC's markup, and a guess that happens to
        # hit some other dropdown would otherwise fail a row that was about to
        # be filed correctly. Logged so the right name can be read off a run
        # instead of out of Inspect.
        found = {
            selector: self._read_kendo_value(selector)
            for selector in TYPE_INPUT_SELECTORS
        }
        answered = {s: v for s, v in found.items() if v is not None}

        if not answered:
            raise SystemProblem(
                "Could not read STAC's Type field, so there is no way to tell "
                f"{self.pair} from another type with the same subtype. "
                f"Tried: {', '.join(TYPE_INPUT_SELECTORS)}."
            )

        logger.debug("Type candidates after Select: %s", answered)

        if document_type in answered.values():
            return

        # Nothing holds the right type. Name every candidate and its value, so
        # the next run says whether the row was wrong or the selector was.
        readings = ", ".join(f"{s} = {v!r}" for s, v in answered.items())
        raise SystemProblem(
            f"No Type field holds {document_type!r} after selecting {self.pair}. "
            f"Found: {readings}. Either the matrix search for {subtype!r} took the "
            "wrong row, or none of those selectors is STAC's Type box."
        )

    def _save(self, ucn: str, names: str) -> None:
        """Press Save and wait for STAC to confirm.

        Raises:
            SaveMayHaveHappened: If Save was clicked but no confirmation came.
                Retrying would file the rap sheet twice, so this is the one
                failure that goes straight to a person.
            SystemProblem: If Save could not be pressed at all, in which case
                nothing was saved and a retry is safe.
        """
        upload_wait = WebDriverWait(self.driver, self.upload_timeout)

        # The subtype reaches the form through Kendo, a moment after the dialog
        # closes. Saving before it lands files the document under whatever was
        # there before.
        try:
            upload_wait.until(
                lambda d: self._read_kendo_value(SUBTYPE_INPUT_SELECTOR) == self.subtype
            )
        except TimeoutException as error:
            raise SystemProblem(
                f"The Subtype field never showed {self.subtype} before Save."
            ) from error

        try:
            upload_wait.until(EC.element_to_be_clickable((By.ID, SAVE_BUTTON_ID))).click()
        except TimeoutException as error:
            raise SystemProblem("Could not press Save.") from error

        # Past this line STAC may already have the document.
        try:
            upload_wait.until(
                EC.presence_of_element_located((By.XPATH, SAVED_NOTIFICATION_XPATH))
            )
        except TimeoutException as error:
            raise SaveMayHaveHappened(
                f"Pressed Save for {names} on {ucn} but STAC never confirmed. The rap "
                "sheet may or may not be on the case. Check it by hand; do not re-run "
                "this row."
            ) from error

        # Wait for the notification to clear, or it covers the next row's Save
        # button.
        try:
            upload_wait.until(
                EC.invisibility_of_element_located((By.XPATH, SAVED_NOTIFICATION_XPATH))
            )
        except TimeoutException:
            logger.warning("%s: the saved notification did not clear", ucn)

        logger.info("%s: saved %s under %s", ucn, names, self.pair)


def _name_tokens(name: str) -> set[str]:
    """Uppercase word set for a name, minus STAC's alert flags and initials."""
    without_brackets = re.sub(r"\(.*?\)", "", name or "")
    return {
        word
        for word in re.findall(r"[a-zA-Z]+", without_brackets.upper())
        if word not in EXCLUDED_NAME_TOKENS and len(word) > 1
    }


def names_match(stac_name: str, sheet_name: str) -> bool:
    """True when the two names are the same person.

    Deliberately loose, because the two sources disagree in harmless ways.
    STAC shows SMITH, JOHN A (ALERT) where the sheet reads Smith, John Allen,
    so one name's words being contained in the other's is enough: that accepts
    a missing middle name, a suffix, and STAC's extra flags.

    Then, because spreadsheets are typed by hand, words that are nearly the
    same count as the same. The threshold is set so that genuinely different
    surnames stay apart.

    This is the only guard that a row's case number and its name belong
    together, so it is less generous than DALYN's version needed to be: there
    the case number was already authoritative and the caption was OCR. Here a
    mismatch is the error, not noise.
    """
    stac_tokens = _name_tokens(stac_name)
    sheet_tokens = _name_tokens(sheet_name)

    if not stac_tokens or not sheet_tokens:
        return False

    if stac_tokens.issubset(sheet_tokens) or sheet_tokens.issubset(stac_tokens):
        return True

    # Whichever name has fewer words has to be fully accounted for in the
    # other. Going the other way would let a one-word name match anything.
    fewer, more = sorted((stac_tokens, sheet_tokens), key=len)
    if _every_word_has_a_near_match(fewer, more):
        logger.info(
            "Names matched allowing for spelling: %r and %r", stac_name, sheet_name
        )
        return True

    return False


def _every_word_has_a_near_match(fewer: set, more: set) -> bool:
    """True when each word in `fewer` has a close enough partner in `more`."""
    for word in fewer:
        best = max(
            (SequenceMatcher(None, word, other).ratio() for other in more), default=0.0
        )
        if best < NAME_TOKEN_SIMILARITY:
            return False
    return True


def _as_tuple(configured, default: tuple) -> tuple:
    """Normalise a config value that may be one label or a list of them.

    Absent falls back to `default`. A bare string becomes a one-item tuple, so
    `case_number_criteria: "Case No"` in config.yaml works as well as a list.
    """
    if configured is None or configured == "":
        return default
    if isinstance(configured, str):
        return (configured,)
    labels = tuple(str(item).strip() for item in configured if str(item).strip())
    return labels or default


def _xpath_literal(value: str) -> str:
    """Quote a string for XPath, including when it contains an apostrophe.

    XPath 1.0 has no escape character, so a value with both quote kinds has to
    be assembled with concat(). Criteria labels are configurable, so a label
    with an apostrophe in it would otherwise build a broken expression and
    report the option as absent.
    """
    if "'" not in value:
        return f"'{value}'"
    if '"' not in value:
        return f'"{value}"'
    parts = value.split("'")
    joined = ', "\'", '.join(f"'{part}'" for part in parts)
    return f"concat({joined})"


def _usable_driver_path(configured) -> str | None:
    """Return the configured chromedriver only if it is actually a file.

    A blank or wrong path used to reach Selenium as a directory, which fails
    with a message about not obtaining a driver and sends everyone looking in
    the wrong place. Better to say so here and fall through to the other routes.
    """
    if not configured:
        return None

    path = Path(str(configured))
    if path.is_file():
        return str(path)

    logger.warning(
        "paths.chromedriver is set to %s, which is not a file. Ignoring it and letting "
        "Selenium find a driver instead.",
        path,
    )
    return None


class StacRunner:
    """Owns one signed-in session for a whole run and files rows through it.

    Used as `with StacRunner(config) as runner:`. Chrome opens once, not once
    per row.

    fresh_browser defaults to FALSE here, the opposite of DALYN. DALYN closes
    Chrome after every upload box because its runs are fifty documents and the
    extra ten seconds buys a guaranteed clean page. A rap sheet run is a whole
    spreadsheet at one upload box per row, where that ten seconds is the
    largest single cost in the run. _leave_browser_clean still throws the
    browser away after any failure, so the state that fresh_browser guards
    against is handled where it actually arises.
    """

    def __init__(self, config: dict) -> None:
        self._config = config
        settings = StacSession(config)
        self.max_attempts = settings.max_attempts
        self.fresh_browser = bool(config["stac"].get("fresh_browser", False))
        self.session: StacSession | None = None

    def __enter__(self) -> "StacRunner":
        self._start_session()
        return self

    def __exit__(self, *_) -> None:
        if self.session:
            self.session.close()
            self.session = None

    def _start_session(self) -> None:
        self.session = StacSession(self._config)
        self.session.open()

    def _restart_session(self) -> None:
        """Throw the browser away and sign in again.

        A wedged Kendo dialog or a half-loaded page does not recover from
        another click. Starting over is what actually works the second time.
        """
        logger.debug("Restarting the STAC browser session")
        if self.session:
            self.session.close()
        self._start_session()

    def _leave_browser_clean(self) -> None:
        """Throw away a browser that failed, so the next row starts fresh.

        A failure leaves STAC wherever it broke: a matrix dialog open, an
        upload pending, a case page whose search box STAC has disabled. The
        next row then fails on the mess rather than on anything of its own,
        which is what "Element is not currently interactable" was: row 1's
        wreckage, charged to rows 2 and 3.

        Never allowed to raise. Whatever went wrong before this is the error
        worth reporting, not a browser that would not restart.
        """
        try:
            self._restart_session()
        except Exception as error:  # noqa: BLE001
            logger.warning("Could not restart the browser after a failure: %s", error)

    def enter_row(
        self, ucn: str, case_number: str, sheet_name: str, paths: list
    ) -> None:
        """File one row's PDF on its case, retrying once if STAC wedges.

        Args:
            ucn: UCN from the row. Tried first.
            case_number: Case number from the row, tried only if the UCN finds
                no case. Either may be "" as long as the other is not.
            sheet_name: Name from the same row, checked against STAC's
                defendant. "" skips the check, which StacSession refuses
                unless check_name is off.
            paths: The PDF to file. A list because a STAC upload box takes
                several and every row here shares one Type/Subtype.

        Raises:
            RowProblem: This row needs a person. Nothing was filed.
            SaveMayHaveHappened: Save was pressed with no confirmation. Never
                retried, since that could file the same rap sheet twice.
            SystemProblem: STAC is broken and the retries are used up. The run
                should stop.
        """
        identifier = ucn or case_number

        for attempt in range(1, self.max_attempts + 1):
            try:
                self.session.find_case(ucn, case_number, sheet_name)
                self.session.add_documents(identifier, paths)

                if self.fresh_browser:
                    # Close Chrome and sign in again, so the next row starts
                    # from a clean page rather than whatever STAC left behind.
                    self._restart_session()

                return

            except (RowProblem, SaveMayHaveHappened):
                # Both are this row's problem, and neither improves by trying
                # again. SaveMayHaveHappened especially: a retry is exactly
                # what would file it twice.
                self._leave_browser_clean()
                raise

            except SystemProblem as error:
                if attempt >= self.max_attempts:
                    self._leave_browser_clean()
                    raise

                logger.warning(
                    "%s: attempt %s of %s failed (%s). Restarting the browser.",
                    identifier, attempt, self.max_attempts, error,
                )
                self._restart_session()

        # The loop either returns or raises; this is unreachable.
        raise SystemProblem(f"{identifier}: ran out of attempts without a result.")