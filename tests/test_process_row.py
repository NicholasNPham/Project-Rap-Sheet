"""Unit tests for main.process_row, using fake CCIS and STAC objects and made-up data."""

import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import main  # noqa: E402
from ccis import CcisResult  # noqa: E402
from exceptions import RowProblem  # noqa: E402
from excel_handler import Outcome  # noqa: E402

FAKE_PDF = b"%PDF-1.4 made-up bytes"


class FakeCcis:
    """Stands in for CcisSession; returns a canned result or raises."""

    def __init__(self, result: CcisResult | None = None, error: Exception | None = None):
        self._result = result
        self._error = error

    def fetch_pdf(self, first_name: str, last_name: str, dob: str) -> CcisResult:
        if self._error:
            raise self._error
        return self._result


class ScriptedCcis:
    """Stands in for CcisSession; finds a record only for the listed (first, last) pairs.

    Each result's PDF is numbered by search, so a test can tell which one was filed.
    """

    def __init__(self, hits: set = frozenset(), error_on: int | None = None):
        self._hits = hits
        self._error_on = error_on
        self.searches: list[tuple] = []

    def fetch_pdf(self, first_name: str, last_name: str, dob: str) -> CcisResult:
        self.searches.append((first_name, last_name))
        if len(self.searches) == self._error_on:
            raise RowProblem("odd")
        pdf = FAKE_PDF + str(len(self.searches)).encode()
        return CcisResult(pdf, (first_name, last_name) in self._hits)


class FakeRunner:
    """Stands in for StacRunner; records what it was asked to file."""

    def __init__(self, upload_enabled: bool = True, save_enabled: bool = True):
        self.session = SimpleNamespace(
            upload_enabled=upload_enabled, save_enabled=save_enabled
        )
        self.calls: list[tuple] = []

    def enter_row(
        self, ucn: str, case_number: str, name: str, files: list, case_url: str = ""
    ) -> None:
        self.calls.append((ucn, case_number, name, files, case_url))


def make_row(search_last: str = "DOE") -> dict:
    return {
        "row": 2, "person_id": "1", "first_name": "John", "last_name": "DOE",
        "search_last": search_last, "name": "DOE, John", "dob": "01/01/1980",
        "ucn": "282026CF000000CFAXMX", "case_number": "CF2600000AXXS",
    }


class ProcessRowTests(unittest.TestCase):
    def run_row(self, ccis, runner, row=None):
        with tempfile.TemporaryDirectory() as folder:
            outcome = main.process_row(row or make_row(), ccis, runner, Path(folder))
            written = sorted(p.name for p in Path(folder).iterdir())
        return outcome, written

    def test_record_is_filed(self):
        runner = FakeRunner()
        outcome, written = self.run_row(FakeCcis(CcisResult(FAKE_PDF, True)), runner)
        self.assertEqual(outcome, Outcome.FILED)
        self.assertEqual(len(runner.calls), 1)
        self.assertEqual(written, ["1_DOE_John.pdf"])

    def test_no_record_pdf_is_filed_too(self):
        runner = FakeRunner()
        outcome, written = self.run_row(FakeCcis(CcisResult(FAKE_PDF, False)), runner)
        self.assertEqual(outcome, Outcome.FILED_NO_RECORD)
        self.assertEqual(len(runner.calls), 1)
        self.assertEqual(written, ["1_DOE_John_NO_RECORD.pdf"])
        self.assertTrue(runner.calls[0][3][0].name.endswith("_NO_RECORD.pdf"))

    def compound_row(self) -> dict:
        row = make_row("DOE SMITH")
        row["last_name"] = "DOE SMITH"
        row["name"] = "DOE SMITH, John Paul"
        return row

    def test_compound_name_hit_on_a_part_is_filed_and_stops_searching(self):
        ccis, runner = ScriptedCcis(hits={("John", "SMITH")}), FakeRunner()
        outcome, written = self.run_row(ccis, runner, self.compound_row())
        self.assertEqual(outcome, Outcome.FILED_PARTIAL_NAME)
        self.assertEqual(
            ccis.searches, [("John", "DOE SMITH"), ("John", "DOE"), ("John", "SMITH")]
        )
        self.assertEqual(len(runner.calls), 1)
        # Named for the person on the sheet, not for the variant that hit.
        self.assertEqual(written, ["1_DOE_SMITH_John.pdf"])

    def test_compound_name_hit_on_the_full_name_is_an_ordinary_filing(self):
        ccis = ScriptedCcis(hits={("John", "DOE SMITH")})
        outcome, _ = self.run_row(ccis, FakeRunner(), self.compound_row())
        self.assertEqual(outcome, Outcome.FILED)
        self.assertEqual(len(ccis.searches), 1)

    def test_compound_name_all_misses_files_the_full_name_page(self):
        ccis, runner = ScriptedCcis(), FakeRunner()
        with tempfile.TemporaryDirectory() as folder:
            outcome = main.process_row(self.compound_row(), ccis, runner, Path(folder))
            filed = runner.calls[0][3][0].read_bytes()
        self.assertEqual(outcome, Outcome.FILED_NO_RECORD)
        self.assertEqual(len(ccis.searches), 6)
        self.assertEqual(filed, FAKE_PDF + b"1")

    def test_ccis_problem_on_a_later_variant_files_nothing(self):
        ccis, runner = ScriptedCcis(error_on=2), FakeRunner()
        outcome, written = self.run_row(ccis, runner, self.compound_row())
        self.assertEqual(outcome, Outcome.ERROR)
        self.assertEqual(runner.calls, [])
        self.assertEqual(written, [])

    def test_rehearsal_and_upload_only_outcomes_for_partial_name(self):
        ccis = ScriptedCcis(hits={("John", "DOE")})
        outcome, _ = self.run_row(
            ccis, FakeRunner(upload_enabled=False, save_enabled=False), self.compound_row()
        )
        self.assertEqual(outcome, Outcome.REHEARSED_PARTIAL_NAME)
        outcome, _ = self.run_row(
            ccis, FakeRunner(upload_enabled=True, save_enabled=False), self.compound_row()
        )
        self.assertEqual(outcome, Outcome.REACHED_SAVE_PARTIAL_NAME)

    def test_rehearsal_and_upload_only_outcomes_for_no_record(self):
        ccis = FakeCcis(CcisResult(FAKE_PDF, False))
        outcome, _ = self.run_row(ccis, FakeRunner(upload_enabled=False, save_enabled=False))
        self.assertEqual(outcome, Outcome.REHEARSED_NO_RECORD)
        outcome, _ = self.run_row(ccis, FakeRunner(upload_enabled=True, save_enabled=False))
        self.assertEqual(outcome, Outcome.REACHED_SAVE_NO_RECORD)

    def test_case_link_is_passed_to_stac(self):
        runner = FakeRunner()
        row = make_row()
        row["case_url"] = "https://stac-test.example.com/incidents/00000"
        self.run_row(FakeCcis(CcisResult(FAKE_PDF, True)), runner, row)
        self.assertEqual(runner.calls[0][4], row["case_url"])

    def test_report_session_can_never_upload_or_save(self):
        config = {
            "stac": {"url": "https://stac-test.example.com", "upload_enabled": True,
                     "save_enabled": True, "is_test_instance": True},
            "report": {"url": "https://stac-prod.example.org"},
        }
        prod = main.report_session_config(config)
        self.assertEqual(prod["stac"]["url"], "https://stac-prod.example.org")
        self.assertFalse(prod["stac"]["upload_enabled"])
        self.assertFalse(prod["stac"]["save_enabled"])
        # The original is untouched, so filing still goes to test.
        self.assertEqual(config["stac"]["url"], "https://stac-test.example.com")
        self.assertTrue(config["stac"]["save_enabled"])

    def test_problem_outcomes(self):
        self.assertEqual(main._problem_outcome("no UCN in this row"), Outcome.NO_UCN)
        self.assertEqual(
            main._problem_outcome("could not read the date of birth"), Outcome.BAD_FORMAT
        )

    def test_ccis_problem_files_nothing(self):
        runner = FakeRunner()
        outcome, written = self.run_row(FakeCcis(error=RowProblem("odd")), runner)
        self.assertEqual(outcome, Outcome.ERROR)
        self.assertEqual(runner.calls, [])
        self.assertEqual(written, [])


if __name__ == "__main__":
    unittest.main()
