"""Unit tests for clearing a preview mask stranded over Save. No browser."""

import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import stac  # noqa: E402
from selenium.common.exceptions import WebDriverException  # noqa: E402

IN_PREVIEW_OVER_SAVE = {"chain": "div.k-loading-mask", "preview": True, "covers_save": True}
OUTSIDE_OVER_SAVE = {"chain": "div.k-loading-mask", "preview": False, "covers_save": True}
IN_PREVIEW_CLEAR = {"chain": "div.k-loading-mask", "preview": True, "covers_save": False}


class FakeDriver:
    """Answers the two scripts the mask check runs, and records the removal."""

    def __init__(self, spinners: list, fail: bool = False) -> None:
        self.spinners = spinners
        self.fail = fail
        self.cleared = False

    def execute_script(self, script: str, *args: object) -> object:
        if self.fail:
            raise WebDriverException("page went away")
        if script == stac.DESCRIBE_SPINNERS_JS:
            return self.spinners
        if script == stac.CLEAR_PREVIEW_MASK_JS:
            self.cleared = True
            return 1
        raise AssertionError("unexpected script")


def session_with(driver: FakeDriver) -> stac.StacSession:
    """A StacSession with only a driver, enough for the mask check."""
    session = stac.StacSession.__new__(stac.StacSession)
    session.driver = driver
    return session


class StrandedPreviewMaskTests(unittest.TestCase):
    def test_mask_only_in_preview_is_cleared(self):
        driver = FakeDriver([IN_PREVIEW_OVER_SAVE])
        self.assertTrue(session_with(driver)._clear_stranded_preview_mask("CF2600000AXXS"))
        self.assertTrue(driver.cleared)

    def test_mask_outside_preview_is_left(self):
        driver = FakeDriver([IN_PREVIEW_OVER_SAVE, OUTSIDE_OVER_SAVE])
        self.assertFalse(session_with(driver)._clear_stranded_preview_mask("CF2600000AXXS"))
        self.assertFalse(driver.cleared)

    def test_nothing_over_save_is_left(self):
        driver = FakeDriver([IN_PREVIEW_CLEAR])
        self.assertFalse(session_with(driver)._clear_stranded_preview_mask("CF2600000AXXS"))
        self.assertFalse(driver.cleared)

    def test_unreadable_page_does_not_raise(self):
        driver = FakeDriver([], fail=True)
        self.assertFalse(session_with(driver)._clear_stranded_preview_mask("CF2600000AXXS"))


if __name__ == "__main__":
    unittest.main()
