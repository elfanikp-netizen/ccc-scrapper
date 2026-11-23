import unittest
from pathlib import Path
from unittest.mock import MagicMock

from crawler import CCCCrawler, description_matches_target, market_data_matches_target, vendor_output_rows


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

    def test_expanded_description_uses_grid_classification(self):
        market_data = [{"description": "A/M: CAPA", "description_expand": "A/M: CAPA Bumper Assy; Rear; LIMITED LIFETIME"}]
        self.assertTrue(market_data_matches_target(market_data, "a/m: capa"))
        self.assertFalse(market_data_matches_target(market_data, "a/m"))

    def test_legacy_expanded_description_is_recognized_for_resume(self):
        market_data = [{"description": "A/M: CAPA Bumper Assy; Rear; LIMITED LIFETIME"}]
        self.assertTrue(market_data_matches_target(market_data, "a/m: capa"))
        self.assertFalse(market_data_matches_target(market_data, "a/m"))
        self.assertFalse(market_data_matches_target(market_data, ""))

    def test_existing_description_type_is_supported(self):
        market_data = [{"description_type": "A/M: CAPA", "description": "A/M: CAPA Bumper"}]
        self.assertTrue(market_data_matches_target(market_data, "a/m: capa"))
        self.assertFalse(market_data_matches_target(market_data, "a/m"))

    def test_inventory_preserves_grid_and_expanded_descriptions(self):
        page = MagicMock()
        grid = MagicMock()
        detail = MagicMock()
        detail.count.return_value = 0
        detail.inner_text.return_value = "A/M: CAPA Bumper Assy; Rear"
        grid.evaluate_all.return_value = [{
            "rowIndex": 0, "description": "A/M: CAPA", "vendor": "iPart Automotive Inc.",
            "part-no": "123", "ext-price": "226.00",
        }]
        page.locator.side_effect = lambda selector: detail if selector == ".preview-section .display-text p:visible" else grid
        detail.last = detail
        crawler = CCCCrawler(page, "user", "password", Path("."))
        crawler.close_vendor_inventory = MagicMock()

        offers, market = crawler.read_vendor_inventory()

        self.assertEqual(offers[0].description, "A/M: CAPA")
        self.assertEqual(market[0]["description"], "A/M: CAPA")
        self.assertEqual(market[0]["description_expand"], "A/M: CAPA Bumper Assy; Rear")
        self.assertNotIn("description_type", market[0])
        output = vendor_output_rows({"VIN": "VIN", "OEM": "123"}, {"market_data": market}, "iPart*")
        self.assertEqual(output[0]["Description"], "A/M: CAPA Bumper Assy; Rear")


if __name__ == "__main__":
    unittest.main()