"""Tests for when --watch runs its passes. No browser, no config, no real data."""

import sys
import unittest
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from main import next_slot  # noqa: E402

# 2026-10-19 is a Monday.
MONDAY = (2026, 10, 19)
FRIDAY = (2026, 10, 23)
SATURDAY = (2026, 10, 24)
NEXT_MONDAY = (2026, 10, 26)


class NextSlotTests(unittest.TestCase):
    """next_slot picks the right half hour inside the working week."""

    def test_before_start_waits_for_first_slot(self) -> None:
        """Early morning goes to 7:30 the same day."""
        self.assertEqual(next_slot(datetime(*MONDAY, 6, 0)), datetime(*MONDAY, 7, 30))

    def test_exactly_on_a_slot_runs_it(self) -> None:
        """A slot that has not begun yet is still taken."""
        self.assertEqual(next_slot(datetime(*MONDAY, 9, 0)), datetime(*MONDAY, 9, 0))

    def test_just_after_a_slot_goes_to_the_next(self) -> None:
        """A pass that finished a few seconds into its slot does not run again."""
        self.assertEqual(next_slot(datetime(*MONDAY, 9, 0, 5)), datetime(*MONDAY, 9, 30))

    def test_long_pass_skips_missed_slots(self) -> None:
        """A pass ending at 10:47 goes to 11:00, not 10:00 or 10:30."""
        self.assertEqual(next_slot(datetime(*MONDAY, 10, 47)), datetime(*MONDAY, 11, 0))

    def test_six_pm_is_the_last_slot(self) -> None:
        """18:00 is included."""
        self.assertEqual(next_slot(datetime(*MONDAY, 17, 45)), datetime(*MONDAY, 18, 0))

    def test_after_six_goes_to_next_morning(self) -> None:
        """Evening goes to 7:30 the next weekday."""
        self.assertEqual(next_slot(datetime(*MONDAY, 18, 0, 1)), datetime(2026, 10, 20, 7, 30))

    def test_friday_evening_goes_to_monday(self) -> None:
        """No passes at the weekend."""
        self.assertEqual(next_slot(datetime(*FRIDAY, 19, 0)), datetime(*NEXT_MONDAY, 7, 30))

    def test_saturday_goes_to_monday(self) -> None:
        """Started at the weekend, it waits for Monday."""
        self.assertEqual(next_slot(datetime(*SATURDAY, 10, 0)), datetime(*NEXT_MONDAY, 7, 30))


if __name__ == "__main__":
    unittest.main()
