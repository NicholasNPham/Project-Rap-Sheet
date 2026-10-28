"""Unit tests for the log redactor. Made-up names and dates only."""

import logging
import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from logger import NAME_MASK, VALUE_MASK, Redactor  # noqa: E402


def record(msg: str, *args: object, exc_info: tuple | None = None) -> logging.LogRecord:
    """A log record as a handler would receive it."""
    return logging.LogRecord("RapSheet.test", logging.INFO, __file__, 1, msg, args, exc_info)


class RedactorTests(unittest.TestCase):
    def setUp(self):
        self.redactor = Redactor()
        self.redactor.hide_names(["DOE, John", "O'BRIEN-SMITH, Mary J"])
        self.redactor.hide_values(["01/02/1980", "1980-01-02"])

    def test_name_in_arguments_is_masked(self):
        line = record("%s: CCIS problem, %s", "row 4 (DOE, John / 282026CF000000CFAXMX)", "timeout")
        self.redactor.filter(line)
        message = line.getMessage()
        self.assertNotIn("DOE", message.upper())
        self.assertNotIn("JOHN", message.upper())
        self.assertIn("282026CF000000CFAXMX", message)
        self.assertIn(NAME_MASK, message)

    def test_name_in_pdf_filename_is_masked(self):
        self.assertEqual(
            self.redactor.redact("saved 12_DOE_John.pdf"),
            f"saved 12_{NAME_MASK}_{NAME_MASK}.pdf",
        )

    def test_case_and_order_do_not_matter(self):
        self.assertEqual(self.redactor.redact("john doe"), f"{NAME_MASK} {NAME_MASK}")

    def test_hyphenated_and_apostrophe_parts_are_masked(self):
        text = self.redactor.redact("O'BRIEN-SMITH, Mary and OBRIEN and BRIEN")
        for part in ("BRIEN", "SMITH", "MARY"):
            self.assertNotIn(part, text.upper().replace(NAME_MASK.upper(), ""))

    def test_middle_initial_alone_is_not_masked(self):
        self.assertEqual(self.redactor.redact("J is fine"), "J is fine")

    def test_name_inside_another_word_is_left(self):
        self.assertEqual(self.redactor.redact("DOES JOHNSON"), "DOES JOHNSON")

    def test_dates_of_birth_are_masked(self):
        self.assertEqual(
            self.redactor.redact("dob 01/02/1980 or 1980-01-02"),
            f"dob {VALUE_MASK} or {VALUE_MASK}",
        )

    def test_traceback_is_masked(self):
        try:
            raise ValueError("defendant DOE, John does not match")
        except ValueError:
            line = record("Unexpected error", exc_info=sys.exc_info())
        self.redactor.filter(line)
        self.assertNotIn("DOE", line.exc_text)
        self.assertIn(NAME_MASK, line.exc_text)

    def test_nothing_registered_changes_nothing(self):
        line = record("Clicked Excel export (try %s of %s)", 1, 3)
        Redactor().filter(line)
        self.assertEqual(line.getMessage(), "Clicked Excel export (try 1 of 3)")


if __name__ == "__main__":
    unittest.main()
