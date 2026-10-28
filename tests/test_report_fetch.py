"""Unit tests for report_fetch's download watching. No browser, made-up files only."""

import os
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import report_fetch as rf  # noqa: E402
from selenium.common.exceptions import InvalidSessionIdException  # noqa: E402
from exceptions import SystemProblem  # noqa: E402


class DownloadWatchTests(unittest.TestCase):
    def setUp(self):
        self._folder = tempfile.TemporaryDirectory()
        self.folder = Path(self._folder.name)
        # An export from an earlier run, which must never be picked up.
        (self.folder / "AutoCCIS.xls").write_bytes(b"old")
        self.before = rf._files_in(self.folder)

    def tearDown(self):
        self._folder.cleanup()

    def test_old_files_are_not_new(self):
        self.assertEqual(rf._new_files(self.folder, self.before), [])
        self.assertFalse(rf._wait_for_start(self.folder, self.before, 0.2))

    def test_partial_download_counts_as_started_but_not_finished(self):
        (self.folder / "AutoCCIS (1).xls.crdownload").write_bytes(b"half")
        self.assertTrue(rf._wait_for_start(self.folder, self.before, 0.2))
        with self.assertRaises(SystemProblem):
            rf._wait_for_finish(self.folder, self.before, 1.5)

    def test_finished_download_is_returned_once_size_holds(self):
        (self.folder / "AutoCCIS (1).xls").write_bytes(b"made-up export")
        found = rf._wait_for_finish(self.folder, self.before, 3)
        self.assertEqual(found.name, "AutoCCIS (1).xls")

    def test_waits_for_chrome_to_rename_the_partial(self):
        partial = self.folder / "AutoCCIS (1).xls.crdownload"
        partial.write_bytes(b"made-up export")

        def finish() -> None:
            time.sleep(0.6)
            partial.rename(self.folder / "AutoCCIS (1).xls")

        threading.Thread(target=finish).start()
        found = rf._wait_for_finish(self.folder, self.before, 5)
        self.assertEqual(found.name, "AutoCCIS (1).xls")

    def test_empty_file_is_not_finished(self):
        (self.folder / "AutoCCIS (1).xls").write_bytes(b"")
        with self.assertRaises(SystemProblem):
            rf._wait_for_finish(self.folder, self.before, 1.5)

    def _write_over_old_export(self) -> Path:
        """Write over the earlier run's AutoCCIS.xls, as Chrome does, and return it.

        The modified time is moved on explicitly so the test does not depend
        on the filesystem's clock resolution.
        """
        old = self.folder / "AutoCCIS.xls"
        old.write_bytes(b"made-up export")
        later = self.before["AutoCCIS.xls"] + 1_000_000_000
        os.utime(old, ns=(later, later))
        return old

    def test_export_written_over_old_file_counts_as_new(self):
        self._write_over_old_export()
        self.assertEqual(
            [path.name for path in rf._new_files(self.folder, self.before)],
            ["AutoCCIS.xls"],
        )
        self.assertTrue(rf._wait_for_start(self.folder, self.before, 0.2))

    def test_export_written_over_old_file_is_returned(self):
        self._write_over_old_export()
        found = rf._wait_for_finish(self.folder, self.before, 3)
        self.assertEqual(found.name, "AutoCCIS.xls")


class _ClosedSwitchTo:
    """switch_to for a browser that has closed: every switch fails."""

    def default_content(self) -> None:
        raise InvalidSessionIdException("browser has closed the connection")


class _ClosedDriver:
    """A WebDriver stand-in whose browser has closed under it."""

    switch_to = _ClosedSwitchTo()


class BrowserGoneTests(unittest.TestCase):
    def test_closed_browser_stops_the_wait_as_a_system_problem(self):
        with self.assertRaises(SystemProblem):
            rf._viewer_text_length(_ClosedDriver())


class EmptyIncomingTests(unittest.TestCase):
    def setUp(self):
        self._folder = tempfile.TemporaryDirectory()
        self.folder = Path(self._folder.name)

    def tearDown(self):
        self._folder.cleanup()

    def test_removes_leftover_downloads(self):
        (self.folder / "AutoCCIS.xls").write_bytes(b"made-up export")
        (self.folder / "downloads.htm").write_bytes(b"<html></html>")
        rf._empty_incoming(self.folder)
        self.assertEqual(list(self.folder.iterdir()), [])

    def test_leaves_subfolders_alone(self):
        (self.folder / "kept").mkdir()
        rf._empty_incoming(self.folder)
        self.assertTrue((self.folder / "kept").is_dir())

    def test_empty_folder_is_fine(self):
        rf._empty_incoming(self.folder)
        self.assertEqual(list(self.folder.iterdir()), [])


if __name__ == "__main__":
    unittest.main()
