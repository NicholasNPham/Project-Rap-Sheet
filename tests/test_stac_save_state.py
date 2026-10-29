"""Unit tests for describing the page after an unconfirmed save. No browser."""

import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import stac  # noqa: E402
from selenium.common.exceptions import WebDriverException  # noqa: E402


class FakeDriver:
    """Answers the save-state script with a fixed reading."""

    def __init__(self, state: object, fail: bool = False) -> None:
        self.state = state
        self.fail = fail

    def execute_script(self, script: str, *args: object) -> object:
        if self.fail:
            raise WebDriverException("page went away")
        if script == stac.DESCRIBE_SAVE_STATE_JS:
            return self.state
        raise AssertionError("unexpected script")


def session_with(driver: FakeDriver) -> stac.StacSession:
    """A StacSession with only a driver, enough for the description."""
    session = stac.StacSession.__new__(stac.StacSession)
    session.driver = driver
    return session


class SaveStateTests(unittest.TestCase):
    def test_error_notification_and_open_requests_are_named(self):
        state = {
            "save": "enabled",
            "notifications": ["div.c-notification-error"],
            "validation": [],
            "dialogs": [],
            "requests": 2,
            "path": "/incidents/123456",
        }
        line = session_with(FakeDriver(state))._describe_save_state()
        self.assertIn("Save button enabled", line)
        self.assertIn("notifications: div.c-notification-error", line)
        self.assertIn("validation marks: none", line)
        self.assertIn("requests still open: 2", line)

    def test_digits_in_markup_and_path_are_masked(self):
        state = {
            "save": "disabled",
            "notifications": [],
            "validation": ["span#Field_9876.field-validation-error"],
            "dialogs": [],
            "requests": None,
            "path": "/incidents/123456",
        }
        line = session_with(FakeDriver(state))._describe_save_state()
        self.assertNotIn("9876", line)
        self.assertNotIn("123456", line)
        self.assertIn("span#Field_#.field-validation-error", line)
        self.assertIn("page: /incidents/#", line)
        self.assertIn("requests still open: unknown", line)

    def test_unreadable_page_does_not_raise(self):
        line = session_with(FakeDriver(None, fail=True))._describe_save_state()
        self.assertIn("could not read the page", line)

    def test_no_answer_does_not_raise(self):
        line = session_with(FakeDriver(None))._describe_save_state()
        self.assertIn("could not read the page", line)


if __name__ == "__main__":
    unittest.main()
