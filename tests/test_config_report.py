"""Unit tests for the report and paths settings in config_loader. No config file, no keyring."""

import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import config_loader as cl  # noqa: E402
from exceptions import ConfigProblem  # noqa: E402


PROD_URL = "https://stac-prod.example.org/"


def merged(over: dict | None = None) -> dict:
    """DEFAULTS with `over` laid on top, the way load_config builds it.

    report.url is filled in unless `over` sets it, since it has no default.
    """
    over = dict(over or {})
    over["report"] = {"url": PROD_URL, **(over.get("report") or {})}
    return cl._merge(cl.DEFAULTS, over)


class ReportSettingsTests(unittest.TestCase):
    def test_defaults_pass(self):
        config = merged()
        cl._validate_report(config)
        self.assertEqual(
            config["report"],
            {"url": PROD_URL.rstrip("/"), "report_timeout": 180, "download_timeout": 120},
        )

    def test_report_url_is_required(self):
        for value in ("", "https://stac.example.com", None):
            config = merged({"report": {"url": value}})
            with self.assertRaises(ConfigProblem, msg=repr(value)):
                cl._validate_report(config)

    def test_numbers_as_text_are_accepted(self):
        config = merged({"report": {"report_timeout": "240"}})
        cl._validate_report(config)
        self.assertEqual(config["report"]["report_timeout"], 240)

    def test_bad_timeouts_are_refused(self):
        for value in ("soon", 0, -5, None):
            config = merged({"report": {"download_timeout": value}})
            with self.assertRaises(ConfigProblem, msg=repr(value)):
                cl._validate_report(config)

    def test_old_excel_block_is_dropped_not_refused(self):
        config = merged({"excel": {"file": "Path/to/Example_Cases.xlsx"}})
        with self.assertLogs("RapSheet", level="WARNING"):
            cl._validate_report(config)
        self.assertNotIn("excel", config)


class PathsTests(unittest.TestCase):
    def test_reports_folder_resolves_under_the_project(self):
        config = merged()
        cl._resolve_paths(config)
        self.assertEqual(config["paths"]["reports"], cl.PROJECT_ROOT / "reports")

    def test_absolute_reports_folder_is_kept(self):
        absolute = Path(PROJECT_ROOT.anchor) / "somewhere" / "reports"
        config = merged({"paths": {"reports": str(absolute)}})
        cl._resolve_paths(config)
        self.assertEqual(config["paths"]["reports"], absolute)


if __name__ == "__main__":
    unittest.main()
