import unittest

from crawler import description_matches_target, market_data_matches_target


class DescriptionMatchingTests(unittest.TestCase):
    def test_description_match_is_case_insensitive_and_exact(self):
        self.assertTrue(description_matches_target("A/M", "a/m"))
        self.assertFalse(description_matches_target("A/M: CAPA", "a/m"))
        self.assertTrue(description_matches_target("A/M: CAPA", "a/m: capa"))

    def test_market_data_requires_the_exact_description(self):
        market_data = [
            {"description": "A/M: CAPA"},
            {"description": "A/M"},
        ]

        self.assertTrue(market_data_matches_target(market_data, "a/m"))
        self.assertFalse(market_data_matches_target(market_data[:1], "a/m"))


if __name__ == "__main__":
    unittest.main()