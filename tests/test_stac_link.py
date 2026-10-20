"""Unit tests for opening a case from a report link, with a fake session. No browser."""

import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from exceptions import RowProblem, SystemProblem  # noqa: E402
from stac import SaveMayHaveHappened, StacRunner, StacSession  # noqa: E402

TEST_STAC = "https://stac-test.example.com"
OUR_LINK = "https://stac-test.example.com/incidents/00000"
LIVE_LINK = "https://stac-live.example.com/incidents/00000"
UCN = "282026CF000000CFAXMX"


def session_at(url: str) -> StacSession:
    """A StacSession with only its url set, enough for link_is_ours."""
    session = StacSession.__new__(StacSession)
    session.url = url
    return session


class FakeSession:
    """Records which calls the runner made; raises on request."""

    def __init__(self, fail_on: str = "", error: Exception | None = None):
        self.url = TEST_STAC
        self.calls: list[str] = []
        self._fail_on = fail_on
        self._error = error

    def _step(self, name: str) -> None:
        self.calls.append(name)
        if name == self._fail_on:
            raise self._error

    def link_is_ours(self, url: str) -> bool:
        return session_at(TEST_STAC).link_is_ours(url)

    def open_case_from_link(self, url: str, ucn: str, name: str) -> str:
        self._step("open_link")
        return name

    def find_case(self, ucn: str, case_number: str, name: str) -> str:
        self._step("find_case")
        return name

    def add_documents(self, ucn: str, paths: list) -> None:
        self._step("add")

    def close_case_tab(self) -> None:
        self.calls.append("close_tab")

    def close(self) -> None:
        self.calls.append("close_browser")


def runner_with(session: FakeSession, attempts: int = 2) -> StacRunner:
    """A StacRunner around a fake session, with restarts recorded instead of run."""
    runner = StacRunner.__new__(StacRunner)
    runner.session = session
    runner.max_attempts = attempts
    runner.fresh_browser = False
    runner.restarts = 0

    def restart() -> None:
        runner.restarts += 1

    runner._restart_session = restart
    return runner


class LinkIsOursTests(unittest.TestCase):
    def test_same_host(self):
        self.assertTrue(session_at(TEST_STAC).link_is_ours(OUR_LINK))

    def test_host_comparison_ignores_case(self):
        self.assertTrue(session_at(TEST_STAC).link_is_ours(OUR_LINK.upper().replace("HTTPS", "https")))

    def test_live_link_from_a_test_run_is_refused(self):
        self.assertFalse(session_at(TEST_STAC).link_is_ours(LIVE_LINK))

    def test_lookalike_host_is_refused(self):
        self.assertFalse(
            session_at(TEST_STAC).link_is_ours("https://stac-test.example.com.evil.test/x")
        )

    def test_non_http_and_blank_are_refused(self):
        session = session_at(TEST_STAC)
        for url in ("", "javascript:openCase(1)", "file:///C:/x", "/incidents/1"):
            self.assertFalse(session.link_is_ours(url), url)


class EnterRowTests(unittest.TestCase):
    def test_link_row_opens_tab_files_and_closes(self):
        session = FakeSession()
        runner_with(session).enter_row(UCN, "", "DOE, John", ["x.pdf"], OUR_LINK)
        self.assertEqual(session.calls, ["open_link", "add", "close_tab"])

    def test_no_link_searches_by_ucn(self):
        session = FakeSession()
        runner_with(session).enter_row(UCN, "", "DOE, John", ["x.pdf"], "")
        self.assertEqual(session.calls, ["find_case", "add"])

    def test_foreign_link_is_never_opened(self):
        session = FakeSession()
        runner_with(session).enter_row(UCN, "", "DOE, John", ["x.pdf"], LIVE_LINK)
        self.assertNotIn("open_link", session.calls)
        self.assertEqual(session.calls, ["find_case", "add"])

    def test_mismatch_closes_tab_without_restarting(self):
        session = FakeSession("open_link", RowProblem("Defendant mismatch"))
        runner = runner_with(session)
        with self.assertRaises(RowProblem):
            runner.enter_row(UCN, "", "DOE, John", ["x.pdf"], OUR_LINK)
        self.assertEqual(session.calls, ["open_link", "close_tab"])
        self.assertEqual(runner.restarts, 0)

    def test_unconfirmed_save_is_not_retried(self):
        session = FakeSession("add", SaveMayHaveHappened("no confirmation"))
        runner = runner_with(session)
        with self.assertRaises(SaveMayHaveHappened):
            runner.enter_row(UCN, "", "DOE, John", ["x.pdf"], OUR_LINK)
        self.assertEqual(session.calls.count("add"), 1)
        self.assertEqual(session.calls[-1], "close_tab")

    def test_system_problem_retries_then_gives_up(self):
        session = FakeSession("open_link", SystemProblem("page never loaded"))
        runner = runner_with(session, attempts=2)
        with self.assertRaises(SystemProblem):
            runner.enter_row(UCN, "", "DOE, John", ["x.pdf"], OUR_LINK)
        self.assertEqual(session.calls.count("open_link"), 2)
        self.assertEqual(session.calls.count("close_tab"), 2)
        # One restart between the attempts, one to leave the browser clean.
        self.assertEqual(runner.restarts, 2)


if __name__ == "__main__":
    unittest.main()
