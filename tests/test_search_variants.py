"""Unit tests for excel_handler.search_variants, using made-up names."""

import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from excel_handler import parse_name, search_variants, strip_suffix  # noqa: E402


class SearchVariantsTests(unittest.TestCase):
    def test_one_given_one_surname_is_one_search(self):
        self.assertEqual(search_variants("DOE, JOHN"), [("JOHN", "DOE")])

    def test_two_given_names(self):
        self.assertEqual(
            search_variants("DOE, JOHN PAUL"), [("JOHN", "DOE"), ("JOHN PAUL", "DOE")]
        )

    def test_two_surnames(self):
        self.assertEqual(
            search_variants("DOE SMITH, JOHN"),
            [("JOHN", "DOE SMITH"), ("JOHN", "DOE"), ("JOHN", "SMITH")],
        )

    def test_two_and_two_tries_every_surname_with_one_given_word_first(self):
        self.assertEqual(
            search_variants("DOE SMITH, JOHN PAUL"),
            [
                ("JOHN", "DOE SMITH"), ("JOHN", "DOE"), ("JOHN", "SMITH"),
                ("JOHN PAUL", "DOE SMITH"), ("JOHN PAUL", "DOE"), ("JOHN PAUL", "SMITH"),
            ],
        )

    def test_hyphen_splits_parts_but_full_surname_is_kept_as_written(self):
        self.assertEqual(
            search_variants("DOE-SMITH, JOHN"),
            [("JOHN", "DOE-SMITH"), ("JOHN", "DOE"), ("JOHN", "SMITH")],
        )

    def test_three_surnames_go_longest_run_first_in_order(self):
        self.assertEqual(
            [last for _, last in search_variants("DOE SMITH ROE, JOHN")],
            ["DOE SMITH ROE", "DOE SMITH", "SMITH ROE", "DOE", "SMITH", "ROE"],
        )

    def test_suffixes_are_never_searched(self):
        self.assertEqual(
            search_variants("DOE SMITH III, JOHN"),
            [("JOHN", "DOE SMITH"), ("JOHN", "DOE"), ("JOHN", "SMITH")],
        )
        self.assertEqual(search_variants("DOE, JOHN JR"), [("JOHN", "DOE")])

    def test_given_name_stops_at_two_words(self):
        self.assertEqual(
            search_variants("DOE, JOHN PAUL MARK"), [("JOHN", "DOE"), ("JOHN PAUL", "DOE")]
        )

    def test_first_variant_is_the_ordinary_search(self):
        for name in ("DOE, JOHN", "DOE SMITH III, JOHN PAUL", "DOE-SMITH, JOHN JR"):
            first_name, last_name = parse_name(name)
            self.assertEqual(search_variants(name)[0], (first_name, strip_suffix(last_name)))

    def test_no_repeats(self):
        for name in ("DOE DOE, JOHN", "DOE, JOHN JR", "DOE SMITH, JOHN PAUL"):
            variants = search_variants(name)
            self.assertEqual(len(variants), len(set(variants)), name)

    def test_unsplittable_names_give_nothing(self):
        for name in ("", "DOE", "DOE, JR., JOHN", ", JOHN"):
            self.assertEqual(search_variants(name), [], repr(name))


if __name__ == "__main__":
    unittest.main()
