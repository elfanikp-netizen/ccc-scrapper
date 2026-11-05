import unittest
from unittest.mock import MagicMock, patch
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from crawler import (
    CCCCrawler,
    MODEL_MISMATCH_RE,
    NO_SOURCE_PARTS_RE,
    PartOffer,
    catalog_not_found_error,
    expand_group_result,
    is_no_catalog_results,
    prepare_result_slots,
    resume_state,
    should_retry_row_status,
)


class NoCatalogResultsDetectionTests(unittest.TestCase):
    def test_identical_vin_oem_rows_share_pending_lookup_and_keep_output_rows(self):
        records = [
            {"VIN": "VIN1", "OEM": "OEM1", "_row": "2"},
            {"VIN": "VIN1", "OEM": "OEM1", "_row": "3"},
        ]

        slots, pending = prepare_result_slots(records, [])
        result = {"VIN": "VIN1", "OEM": "OEM1", "Status": "OK", "Lowest CCC Price": 25}
        expanded = expand_group_result(result, records, list(pending.values())[0])

        self.assertEqual(list(pending.values()), [[0, 1]])
        self.assertEqual(len(expanded), 2)
        self.assertEqual([row["OEM"] for row in expanded], ["OEM1", "OEM1"])
        self.assertEqual([row["Lowest CCC Price"] for row in expanded], [25, 25])

    def test_source_parts_zero_counter_and_price_prompt_are_not_no_source(self):
        for text in (
            "Source parts 0\nEnter part price.",
            "0 Source parts\nEnter part price.",
            "0 source parts",
        ):
            with self.subTest(text=text):
                self.assertIsNone(NO_SOURCE_PARTS_RE.search(text))
        self.assertIsNotNone(NO_SOURCE_PARTS_RE.search("No part lines to source"))

    def test_detects_no_results_message(self):
        body = "No results found for search text: 312768310"
        self.assertTrue(is_no_catalog_results(body))

    def test_rejects_valid_catalog_pages(self):
        body = "A/C Compressor 1234 part list"
        self.assertFalse(is_no_catalog_results(body))

    def test_only_explicit_empty_message_is_catalog_not_found(self):
        error = catalog_not_found_error("No results found for search text: CJ5Z13008B", "CJ5Z13008B")
        self.assertEqual(str(error), "No catalog results found for OEM CJ5Z13008B")
        self.assertIsNone(
            catalog_not_found_error(
                "No results found for search text: CJ5Z13008B",
                "CJ5Z13008B",
                visible_category_count=3,
            )
        )

        self.assertIsNone(catalog_not_found_error("", "CJ5Z13008B"))
        self.assertIsNone(catalog_not_found_error("Loading catalog...", "CJ5Z13008B"))

    def test_resume_retries_failures_and_qualified_capa_market_rows(self):
        results = [
            {"VIN": "VIN1", "OEM": "OEM1", "Status": "No catalog results found for OEM OEM1; skipped"},
            {"VIN": "VIN2", "OEM": "OEM2", "Status": "ERROR: catalog did not load"},
            {"VIN": "VIN3", "OEM": "OEM3", "Status": "No offers matching description: A/M: CAPA"},
            {
                "VIN": "VIN4",
                "OEM": "OEM4",
                "Status": "No offers matching description: A/M: CAPA",
                "market_data": '[{"description": "A/M", "vendor": "Vendor", "ext-price": "20.00"}]',
            },
            {
                "VIN": "VIN5",
                "OEM": "OEM5",
                "Status": "No offers matching description: A/M: CAPA",
                "market_data": '[{"description": "A/M: DS, CAPA", "vendor": "Vendor", "ext-price": "267.00"}]',
            },
            {"VIN": "VIN6", "OEM": "OEM6", "Status": "OK"},
        ]

        completed, retry_rows = resume_state(results)

        self.assertEqual(list(retry_rows[("VIN1", "OEM1")]), [0])
        self.assertEqual(list(retry_rows[("VIN2", "OEM2")]), [1])
        self.assertEqual(list(retry_rows[("VIN3", "OEM3")]), [2])
        self.assertEqual(completed[("VIN4", "OEM4")], 1)
        self.assertEqual(list(retry_rows[("VIN5", "OEM5")]), [4])
        self.assertEqual(completed[("VIN6", "OEM6")], 1)

    def test_qualified_capa_description_is_included_in_summary(self):
        result = CCCCrawler.summarize(
            "VIN1",
            "OEM1",
            [
                PartOffer("iPart Automotive Inc.", "A/M: DS, CAPA", 267.0),
                PartOffer("Other supplier", "A/M", 100.0),
            ],
        )

        self.assertEqual(result["Status"], "OK")
        self.assertTrue(result["Ipart listed on CCC"])
        self.assertEqual(result["Ipart List Price CCC"], 267.0)
        self.assertEqual(result["Lowest CCC Price"], 267.0)

    def test_generic_aftermarket_is_not_capa(self):
        result = CCCCrawler.summarize(
            "VIN1",
            "OEM1",
            [PartOffer("Other supplier", "A/M: XC", 50.0)],
        )

        self.assertEqual(result["Status"], "No offers matching description: A/M: CAPA")
        self.assertIsNone(result["Lowest CCC Price"])

    def test_transient_row_statuses_are_retryable(self):
        for status in (
            "ERROR: catalog categories returned no part rows",
            "No catalog results found for OEM OEM1; skipped",
            "No sourceable parts available; skipped",
            "Vendor search failed after 3 Retry clicks; skipped",
        ):
            with self.subTest(status=status):
                self.assertTrue(should_retry_row_status(status))

        self.assertFalse(should_retry_row_status("No offers matching description: A/M: CAPA"))

    def test_non_sourceable_catalog_choice_retries_in_a_fresh_workfile(self):
        page = MagicMock()
        crawler = CCCCrawler(page, "user", "password", MagicMock())
        crawler.open_new_workfile = MagicMock()
        crawler.enter_vehicle = MagicMock(return_value=True)
        crawler.read_vendor_inventory = MagicMock(return_value=(
            [PartOffer("Vendor", "A/M: CAPA", 100.0)],
            [{"description": "A/M: CAPA", "vendor": "Vendor", "ext-price": "100.00"}],
        ))

        def find_next_part(_oem, attempted):
            if crawler.open_part_results.call_count == 1:
                attempted.add(("front lamps", 0))
                crawler.no_sourceable_candidates = True
                return None
            crawler.no_sourceable_candidates = False
            return 0

        crawler.open_part_results = MagicMock(side_effect=find_next_part)

        with patch("crawler.MAX_ROW_RETRIES", 0):
            result = crawler.crawl_row({"VIN": "VIN1", "OEM": "OEM1", "_row": "2"})

        self.assertEqual(result["Status"], "OK")
        self.assertEqual(crawler.open_new_workfile.call_count, 2)
        self.assertEqual(crawler.enter_vehicle.call_count, 2)
        self.assertEqual(crawler.open_part_results.call_count, 2)
        self.assertIs(
            crawler.open_part_results.call_args_list[0].args[1],
            crawler.open_part_results.call_args_list[1].args[1],
        )

    def test_catalog_retry_stops_when_no_new_candidate_appears(self):
        page = MagicMock()
        page.locator.return_value.inner_text.return_value = "No part lines to source"
        crawler = CCCCrawler(page, "user", "password", MagicMock())
        crawler.open_new_workfile = MagicMock()
        crawler.enter_vehicle = MagicMock(return_value=True)

        def find_same_candidate(_oem, attempted):
            attempted.add(("front lamps", 0))
            crawler.no_sourceable_candidates = True
            return None

        crawler.open_part_results = MagicMock(side_effect=find_same_candidate)

        with patch("crawler.MAX_ROW_RETRIES", 0):
            result = crawler.crawl_row({"VIN": "VIN1", "OEM": "OEM1", "_row": "2"})

        self.assertEqual(result["Status"], "No sourceable parts available; skipped")
        self.assertEqual(crawler.open_new_workfile.call_count, 2)
        self.assertEqual(crawler.open_part_results.call_count, 2)

    def test_enter_part_price_prompt_clicks_ok_in_inline_container(self):
        page = MagicMock()
        prompt = MagicMock()
        prompt_container = MagicMock()
        inline_ok = MagicMock()
        page.get_by_text.return_value.first = prompt
        prompt.locator.return_value = prompt_container
        overlay = MagicMock()
        page.locator.return_value.filter.return_value.last = overlay
        crawler = CCCCrawler(page, "user", "password", MagicMock())
        crawler.first_visible = MagicMock(side_effect=[None, None, inline_ok])

        acknowledged = crawler.dismiss_enter_part_price_prompt()

        self.assertTrue(acknowledged)
        prompt_pattern = page.get_by_text.call_args_list[0].args[0]
        self.assertRegex("“Enter part price.”", prompt_pattern)
        prompt.locator.assert_called_once_with(
            "xpath=ancestor::*[.//button[normalize-space(.)='OK']][1]"
        )
        inline_ok.click.assert_called_once_with(timeout=10000)
        prompt.wait_for.assert_any_call(state="hidden", timeout=10000)

    def test_engine_mismatch_prompt_clicks_visible_yes_in_inline_container(self):
        page = MagicMock()
        page.locator.return_value.inner_text.return_value = (
            "Part does not match the engine of the vehicle. Select anyway?"
        )
        mismatch_prompt = MagicMock()
        prompt_container = MagicMock()
        inline_yes = MagicMock()
        page.get_by_text.return_value.last = mismatch_prompt
        mismatch_prompt.locator.return_value = prompt_container
        page.wait_for_function.side_effect = [MagicMock(), PlaywrightTimeoutError("prompt dismissed")]
        crawler = CCCCrawler(page, "user", "password", MagicMock())
        crawler.first_visible = MagicMock(side_effect=[None, inline_yes])

        crawler.resolve_part_prompts()

        mismatch_prompt.locator.assert_called_once_with(
            "xpath=ancestor::*[.//button[normalize-space(.)='Yes']][1]"
        )
        inline_yes.click.assert_called_once_with(timeout=10000)
        mismatch_prompt.wait_for.assert_called_once_with(state="hidden", timeout=10000)

    def test_detects_engine_mismatch_prompt_wording(self):
        self.assertRegex(
            "Part does not match the engine of the vehicle. Select anyway?",
            MODEL_MISMATCH_RE,
        )


if __name__ == "__main__":
    unittest.main()
