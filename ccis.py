"""Drives the CCIS web application with Selenium: sign in, search, capture a PDF.

A port of browser.py from venire_3.0, which has run this screen in production.
The automation itself is deliberately unchanged, element for element and click
for click: it works at four seconds a row and there is nothing to gain by
rewriting a proven sequence. What changes is the shape around it.

    One session object instead of eight free functions, so the driver and wait
    are not threaded through every call.

    Credentials and element locators come from config.yaml, not key.py and
    config.json.

    Failures raise RowProblem or SystemProblem, so main.py routes a row that
    CCIS cannot answer the same way it routes a row STAC cannot answer.

    The PDF comes back as bytes. The caller decides where it lands, because
    the same bytes have to reach STAC's upload box as a file on disk.

Carried over verbatim and worth knowing about: search() clicks the "exclude
attorneys" and "exclude judges" checkboxes on every search, and relies on
reset_search() having cleared them. If a reset ever stops clearing them, the
second click turns the exclusions back OFF rather than on, and the search
quietly widens to include attorneys and judges who share a defendant's name.
venire has this same behaviour and runs clean, so it is kept as-is rather than
"fixed" against an instance nobody can test here.
"""

import base64
import time

from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

from exceptions import RowProblem, SystemProblem
from logger import get_logger

logger = get_logger(__name__)

# Chrome DevTools Protocol, for a PDF with no print dialog.
PRINT_TO_PDF_COMMAND = "Page.printToPDF"
PDF_DATA_KEY = "data"
PRINT_BACKGROUND_KEY = "printBackground"

# Element locators, from venire_3.0's config.json. Defaults only; the ccis
# section of config.yaml overrides any of them, so a CCIS release that renames
# a field is a config edit rather than a code change.
DEFAULT_LOCATORS = {
    "username_field": "loginForm:username",
    "password_field": "loginForm:password",
    "submit_login_button": "loginForm:login",
    "exclude_attorneys_checkbox": "//label[@for='search_tab:personForm:nameTypes:0']",
    "exclude_judges_checkbox": "//label[@for='search_tab:personForm:nameTypes:1']",
    "last_name_field": "search_tab:personForm:lastname",
    "first_name_field": "search_tab:personForm:fname",
    "dob_field": "search_tab:personForm:dob_input",
    "search_button": "//button[.//span[text()='Search']]",
    "no_results_popup_message": (
        "//span[@class='ui-messages-info-detail' and text()='No matches found.']"
    ),
    "view_selection_button": "//button[.//span[text()='view selection']]",
    "back_button_from_pdf_page": "//button[contains(@id, 'caseSummary')]",
    "back_button_from_selection_page": (
        "//button[contains(@onclick, 'PrimeFaces.ab') and "
        "not(contains(@onclick, 'caseSummary'))]"
    ),
    "reset_button": "//button[.//span[text()='Reset']]",
}

# Headless flags, from venire's setup_browser. The window size is not
# cosmetic: the PDF is rendered from the page, so a narrow viewport changes
# what the captured case summary looks like.
HEADLESS_ARGUMENTS = (
    "--headless=new",
    "--no-sandbox",
    "--disable-dev-shm-usage",
    "--disable-gpu",
    "--window-size=1920,1080",
)


class CcisSession:
    """One signed-in CCIS session, reused for every row in a run.

    Used as `with CcisSession(config) as ccis:` so Chrome closes even when a
    row blows up partway through.
    """

    def __init__(self, config: dict) -> None:
        ccis = config["ccis"]
        self.url = str(ccis["url"]).rstrip("/")
        self._username = ccis["username"]
        self._password = ccis["password"]
        self.headless = bool(ccis.get("headless", False))
        self.wait_timeout = int(ccis.get("wait_timeout", 5))
        self.no_results_wait_timeout = int(ccis.get("no_results_wait_timeout", 1))
        self.pause_between_actions = float(ccis.get("pause_between_actions", 1) or 0)

        self.locators = dict(DEFAULT_LOCATORS)
        self.locators.update(ccis.get("locators") or {})

        self._chromedriver = ccis.get("chromedriver") or config["paths"].get("chromedriver")

        self.driver: webdriver.Chrome | None = None
        self.wait: WebDriverWait | None = None

    def __enter__(self) -> "CcisSession":
        self.open()
        return self

    def __exit__(self, *_) -> None:
        self.close()

    def _at(self, name: str) -> str:
        return self.locators[name]

    def open(self) -> None:
        """Launch Chrome, load CCIS, and sign in.

        Raises:
            SystemProblem: Chrome will not start, CCIS will not load, or the
                credentials are refused. None of these belong to any one row.
        """
        logger.info("Opening CCIS at %s (headless: %s)", self.url, self.headless)

        self.driver = self._start_chrome()
        self.wait = WebDriverWait(self.driver, self.wait_timeout)

        try:
            self.driver.get(self.url)
            if not self.headless:
                self.driver.maximize_window()
        except WebDriverException as error:
            self.close()
            raise SystemProblem(f"Could not load CCIS at {self.url}: {error}") from error

        try:
            self._sign_in()
        except Exception:
            self.close()
            raise

    def _start_chrome(self) -> webdriver.Chrome:
        """Start Chrome, trying each way of finding chromedriver in turn.

        Same order as the STAC session, and for the same reason: on a
        locked-down network the downloader is the route least likely to work,
        so it goes last.

        Raises:
            SystemProblem: If none of them produce a working Chrome.
        """
        options = webdriver.ChromeOptions()
        if self.headless:
            for argument in HEADLESS_ARGUMENTS:
                options.add_argument(argument)

        attempts = []
        if self._chromedriver:
            attempts.append(
                ("paths.chromedriver", lambda: Service(str(self._chromedriver)))
            )
        attempts.append(("Selenium's own driver lookup", lambda: None))
        attempts.append(
            ("downloading chromedriver", lambda: Service(ChromeDriverManager().install()))
        )

        failures = []
        for description, build_service in attempts:
            try:
                service = build_service()
                driver = (
                    webdriver.Chrome(service=service, options=options)
                    if service
                    else webdriver.Chrome(options=options)
                )
                logger.debug("CCIS Chrome started via %s", description)
                return driver
            except Exception as error:  # noqa: BLE001 - each route fails differently
                failures.append(
                    f"{description}: {type(error).__name__}: {str(error).splitlines()[0]}"
                )

        raise SystemProblem(
            "Could not start Chrome for CCIS. Set paths.chromedriver in config.yaml "
            "to a chromedriver.exe matching your installed Chrome version. Tried: "
            + " | ".join(failures)
        )

    def _sign_in(self) -> None:
        """Fill the sign-in form and wait for the search page.

        The last-name field is the proof of a successful sign-in, exactly as
        venire uses it: it only exists on the search page.

        Raises:
            SystemProblem: If sign-in does not complete.
        """
        try:
            self.wait.until(
                EC.element_to_be_clickable((By.ID, self._at("username_field")))
            ).send_keys(self._username)
            self.wait.until(
                EC.element_to_be_clickable((By.ID, self._at("password_field")))
            ).send_keys(self._password)
            self.driver.find_element(By.ID, self._at("submit_login_button")).click()
        except (TimeoutException, WebDriverException) as error:
            raise SystemProblem(f"Could not sign in to CCIS: {error}") from error

        try:
            self.wait.until(
                EC.presence_of_element_located((By.ID, self._at("last_name_field")))
            )
        except TimeoutException as error:
            raise SystemProblem(
                "Signed in to CCIS but the search page never appeared. Usually a "
                "wrong username or password in Credential Manager. Rerun store_secret.py."
            ) from error

        logger.info("Signed in to CCIS")

    def close(self) -> None:
        """Close Chrome. Safe to call twice, and never raises."""
        if self.driver is None:
            return
        try:
            self.driver.quit()
        except WebDriverException as error:
            logger.warning("CCIS's Chrome did not close cleanly: %s", error)
        finally:
            self.driver = None
            self.wait = None

    # --------------------------------------------------------------- one row

    def fetch_pdf(self, first_name: str, last_name: str, dob: str) -> bytes | None:
        """Search for one person and return their case summary as PDF bytes.

        The search page is always left reset, whether a record was found, no
        record was found, or something failed partway through. venire does this
        with a reset after each branch; here it is a finally, so a failure
        cannot leave a half-filled form for the next row to fail on.

        Args:
            first_name: ex 'John'
            last_name: ex 'Smith'
            dob: mm/dd/yyyy, ex '04/23/1985'

        Returns:
            The PDF bytes, or None when CCIS has no matching record. None is
            an ordinary answer, not a failure: plenty of people have no record.

        Raises:
            RowProblem: The search ran but CCIS did not behave as expected for
                this person, so the row needs a look.
            SystemProblem: The browser or the session is broken.
        """
        try:
            self._search(first_name, last_name, dob)

            if self._no_results():
                return None

            self._open_case_summary()
            pdf_bytes = self._print_to_pdf()
            self._back_to_search()
            return pdf_bytes

        except (RowProblem, SystemProblem):
            raise
        except TimeoutException as error:
            raise RowProblem(
                f"CCIS did not respond as expected for {last_name}, {first_name} "
                f"({dob}): {type(error).__name__}"
            ) from error
        except WebDriverException as error:
            # A dead driver is not this row's fault and the next row will fail
            # the same way, so it stops the run.
            raise SystemProblem(f"CCIS's browser failed: {error}") from error
        finally:
            self._reset_search()

    def _search(self, first_name: str, last_name: str, dob: str) -> None:
        """Fill the search form and submit it.

        The two checkbox clicks exclude attorneys and judges, who otherwise
        turn up alongside defendants of the same name. See the module docstring
        for why they are clicked unconditionally.
        """
        # Logged per search, because "the names look swapped" is otherwise
        # impossible to settle without watching the screen. This says exactly
        # which value went into which field, by the field's own element id, so
        # a real mix-up is visible in the log rather than inferred.
        logger.info(
            "CCIS search | %s <- %r | %s <- %r | dob %r",
            self._at("last_name_field"), last_name,
            self._at("first_name_field"), first_name,
            dob,
        )

        self.driver.find_element(By.XPATH, self._at("exclude_attorneys_checkbox")).click()
        self.driver.find_element(By.XPATH, self._at("exclude_judges_checkbox")).click()

        self.wait.until(
            EC.element_to_be_clickable((By.ID, self._at("last_name_field")))
        ).send_keys(last_name)
        self.wait.until(
            EC.element_to_be_clickable((By.ID, self._at("first_name_field")))
        ).send_keys(first_name)
        self.wait.until(
            EC.element_to_be_clickable((By.ID, self._at("dob_field")))
        ).send_keys(dob)
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self._at("search_button")))
        ).click()

    def _no_results(self) -> bool:
        """True when CCIS put up its 'No matches found.' banner.

        A short wait of its own, not the session's, because the common case is
        that the banner is absent and waiting the full timeout for something
        that is not coming would add that timeout to every row that has a
        record. Only TimeoutException is caught: a dead driver propagates.
        """
        brief = WebDriverWait(self.driver, self.no_results_wait_timeout)
        try:
            brief.until(
                EC.visibility_of_element_located(
                    (By.XPATH, self._at("no_results_popup_message"))
                )
            )
            return True
        except TimeoutException:
            return False

    def _open_case_summary(self) -> None:
        """Click through to the case summary page the PDF is taken from."""
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self._at("view_selection_button")))
        ).click()
        # The back button only exists on the summary page, so waiting for it
        # is how venire knows the page has actually arrived before printing.
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self._at("back_button_from_pdf_page")))
        )

    def _print_to_pdf(self) -> bytes:
        """Capture the current page as PDF bytes through Chrome DevTools.

        Silent: no print dialog, no file written by Chrome itself. Background
        graphics are included because the case summary uses shading to
        separate its sections and without them the capture is hard to read.
        """
        try:
            result = self.driver.execute_cdp_cmd(
                PRINT_TO_PDF_COMMAND, {PRINT_BACKGROUND_KEY: True}
            )
        except WebDriverException as error:
            raise SystemProblem(
                f"Chrome would not print the CCIS page to PDF: {error}"
            ) from error

        encoded = result.get(PDF_DATA_KEY)
        if not encoded:
            raise RowProblem("Chrome returned an empty PDF for this case summary.")

        pdf_bytes = base64.b64decode(encoded)
        if not pdf_bytes.startswith(b"%PDF-"):
            # Never seen, but a PDF that is not a PDF would reach STAC's
            # upload box and be filed as this person's criminal history.
            raise RowProblem(
                "What Chrome returned for this case summary is not a PDF "
                f"(starts {pdf_bytes[:8]!r})."
            )

        return pdf_bytes

    def _back_to_search(self) -> None:
        """Walk back from the summary page to the search page."""
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self._at("back_button_from_pdf_page")))
        ).click()
        self.wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, self._at("back_button_from_selection_page"))
            )
        ).click()

    def _reset_search(self) -> None:
        """Clear the search form, ready for the next row.

        Never raises. It runs in a finally, so an exception here would replace
        whatever real error sent the row down that path. A reset that fails is
        reported and left to the next row's search to trip over, where it will
        be attributed correctly.
        """
        try:
            self.wait.until(
                EC.element_to_be_clickable((By.XPATH, self._at("reset_button")))
            ).click()
            if self.pause_between_actions > 0:
                time.sleep(self.pause_between_actions)
        except (TimeoutException, WebDriverException) as error:
            logger.warning("Could not reset the CCIS search form: %s", error)