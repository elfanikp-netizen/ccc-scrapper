import unittest
import json
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import MagicMock, patch

from openpyxl import Workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from crawler import (
    CCCCrawler, PART_TYPE_OUTPUT_COLUMNS, load_results, main,
    matching_part_type_categories, part_type_result_complete, read_input, run_part_type_groups, save_results,
)


class PartTypeScenarioTests(unittest.TestCase):
    def setUp(self):
        self.record = {"VIN": "VIN1", "Part Type Code": "101", "Part Type Description": "Bumper",
                       "Scenario": "front bumper", "_row": "2"}

    def test_input_and_failed_result_round_trip(self):
        with TemporaryDirectory() as directory:
            input_path = Path(directory) / "input.xlsx"
            output_path = Path(directory) / "output.xlsx"
            workbook = Workbook()
            workbook.active.append(["VIN", "Part Type Code", "Part Type Description", "Scenario"])
            workbook.active.append(["VIN1", "101", "Bumper", "front bumper"])
            workbook.save(input_path)
            workbook.close()
            records = read_input(input_path, "vin_part_type_scenario")
            self.assertEqual(records[0]["Scenario"], "front bumper")
            save_results(output_path, [{**records[0], "Status": "ERROR: retry"}], PART_TYPE_OUTPUT_COLUMNS)
            results = load_results(output_path, preserve_all=True)
            self.assertEqual(results[0]["Part Type Code"], "101")
            self.assertEqual(results[0]["Scenario"], "front bumper")

    def test_scenario_with_comma_remains_one_search_phrase(self):
        with TemporaryDirectory() as directory:
            input_path = Path(directory) / "input.xlsx"
            workbook = Workbook()
            workbook.active.append(["VIN", "Part Type Code", "Part Type Description", "Scenario"])
            workbook.active.append(["VIN1", "101", "Bumper", "bumper, front"])
            workbook.active.append(["VIN1", "101", "Bumper", "bumper, front"])
            workbook.save(input_path)
            workbook.close()
            records = read_input(input_path, "vin_part_type_scenario")
            self.assertEqual([row["Scenario"] for row in records], ["bumper, front", "bumper, front"])

    def test_category_matching_prefers_all_built_variants(self):
        self.assertEqual(matching_part_type_categories(" FRONT bumper ", [
            "Front Bumper", "FRONT BUMPER US BUILT", "Front Bumper Japan Built", "Rear Bumper"
        ]), ["FRONT BUMPER US BUILT", "Front Bumper Japan Built"])
        self.assertEqual(matching_part_type_categories("front bumper", ["Front Bumper", "Front Bumper Cover"]),
                         ["Front Bumper"])
        self.assertEqual(matching_part_type_categories("", ["Front Bumper"]), [])

    def test_part_type_catalog_opens_estimate_without_keyword_search(self):
        page = MagicMock()
        crawler = CCCCrawler(page, "user", "password", Path("."))
        crawler.apply_oem_catalog_filter = MagicMock()
        crawler.dismiss_jumpstart_prompt = MagicMock()
        crawler.open_part_type_catalog("FRONT BUMPER")
        crawler.apply_oem_catalog_filter.assert_not_called()
        page.locator.assert_called_with('a[aria-label="Estimate"]')
        page.locator.return_value.click.assert_called_once()

    def test_discovery_uses_both_build_variants_and_restores_scope(self):
        crawler = CCCCrawler(MagicMock(), "user", "password", Path("."))
        crawler.ensure_part_type_vehicle = MagicMock()
        crawler.open_part_type_catalog = MagicMock()
        crawler.apply_oem_catalog_filter = MagicMock()
        crawler.load_catalog_categories = MagicMock(side_effect=[
            (None, ["COOLING US BUILT"]),
            (None, ["COOLING US BUILT", "COOLING JAPAN BUILT", "FRONT BUMPER OLD"]),
            (None, ["FRONT BUMPER OLD", "FRONT BUMPER CURRENT"]),
            (None, ["FRONT BUMPER CURRENT"]),
            (None, ["FRONT BUMPER OLD", "FRONT BUMPER CURRENT"]),
        ])
        self.assertEqual(crawler.discover_part_type_categories({**self.record, "Scenario": "COOLING"}),
                         ["COOLING US BUILT", "COOLING JAPAN BUILT"])
        self.assertEqual(crawler.discover_part_type_categories(self.record), ["FRONT BUMPER CURRENT"])
        scopes = [call.kwargs["model_scope"] for call in crawler.apply_oem_catalog_filter.call_args_list]
        self.assertEqual(scopes, ["ALL MODELS", "THIS MODEL w/ OPTIONS", "ALL MODELS", "THIS MODEL w/ OPTIONS"])

    def test_no_sourceable_part_is_complete_on_resume(self):
        self.assertTrue(part_type_result_complete({"Status": "No sourceable parts available; skipped",
                                                   "Description Target": "A/M: CAPA"}, "A/M: CAPA"))

    def test_missing_scenario_is_skipped_before_next_scenario(self):
        with TemporaryDirectory() as directory:
            crawler = MagicMock()
            crawler.discover_part_type_categories.side_effect = [[], ["FENDER US BUILT"]]
            crawler.crawl_row.return_value = {"OEM": "OEM-1", "Status": "OK"}
            missing = {**self.record, "Scenario": "NOT FOUND"}
            following = {**self.record, "Part Type Code": "110", "Scenario": "FENDER"}
            results = run_part_type_groups(crawler, [missing, following], [], Path(directory) / "results.xlsx", "A/M: CAPA")
            self.assertEqual([row["Status"] for row in results], ["No matching catalog categories; skipped", "OK"])
            self.assertEqual(crawler.crawl_row.call_args.args[0]["Scenario"], "FENDER")

    def test_scenarios_are_processed_by_vin_before_next_vehicle(self):
        with TemporaryDirectory() as directory:
            crawler = MagicMock()
            crawler.discover_part_type_categories.side_effect = lambda record: [record["Scenario"]]
            crawler.crawl_row.return_value = {"Status": "OK"}
            records = [self.record, {**self.record, "VIN": "VIN2", "Scenario": "FENDER"},
                       {**self.record, "Scenario": "FENDER"}]
            run_part_type_groups(crawler, records, [], Path(directory) / "results.xlsx", "A/M: CAPA")
            self.assertEqual([(call.args[0]["VIN"], call.args[0]["Scenario"]) for call in crawler.crawl_row.call_args_list],
                             [("VIN1", "front bumper"), ("VIN1", "FENDER"), ("VIN2", "FENDER")])

    def test_vehicle_setup_is_reused_and_changed_on_next_vin(self):
        page = MagicMock()
        page.is_closed.return_value = False
        crawler = CCCCrawler(page, "user", "password", Path("."))
        crawler.recover_closed_page = MagicMock()
        crawler.open_new_workfile = MagicMock(side_effect=lambda: setattr(crawler, "_workfile_initialized", True))
        crawler.enter_vehicle = MagicMock(return_value=True)
        crawler.ensure_part_type_vehicle("VIN1")
        crawler.ensure_part_type_vehicle("VIN1")
        crawler.ensure_part_type_vehicle("VIN2")
        self.assertEqual(crawler.open_new_workfile.call_count, 2)
        self.assertEqual([call.args[0] for call in crawler.enter_vehicle.call_args_list], ["VIN1", "VIN2"])

    def test_catalog_flow_filters_and_reads_estimate_properties(self):
        page = MagicMock()
        parts = MagicMock()
        candidate = parts.first
        candidate.inner_text.return_value = "1\n\nBumper cover\n\n194.60\n\n3.1"
        page.locator.side_effect = lambda selector: parts if selector == "app-part-selector-part-row:visible" else MagicMock()
        crawler = CCCCrawler(page, "user", "password", Path("."))
        crawler.open_part_type_catalog = MagicMock()
        crawler.load_catalog_categories = MagicMock()
        crawler.select_catalog_category = MagicMock(return_value=True)
        crawler.wait_for_catalog_element = MagicMock(return_value=True)
        crawler.resolve_part_prompts = MagicMock()
        crawler.read_selected_part_properties = MagicMock(return_value=("OEM1", "Cover"))
        crawler.dismiss_enter_part_price_prompt = MagicMock()
        crawler.wait_for_vendor_inventory = MagicMock(return_value=0)
        self.assertEqual(crawler.open_part_type_scenario_results("front bumper", "Front Bumper US Built"),
                         (0, "OEM1", "Cover"))
        crawler.open_part_type_catalog.assert_called_once_with("front bumper")
        crawler.select_catalog_category.assert_called_once_with("Front Bumper US Built", "front bumper")
        candidate.dblclick.assert_called_once()
        crawler.read_selected_part_properties.assert_called_once_with("Bumper cover", select_added_line=True)
        page.locator.assert_any_call("app-estimate-row")

    def test_existing_estimate_line_without_source_button_is_not_market_listed(self):
        page = MagicMock()
        parts = MagicMock()
        parts.first.inner_text.return_value = "1\nCenter console\n663.83"
        source_row = MagicMock()
        source_button = source_row.filter.return_value.last.get_by_role.return_value
        source_button.wait_for.side_effect = PlaywrightTimeoutError("not present")
        source_button.count.return_value = 0
        page.locator.side_effect = lambda selector: parts if selector == "app-part-selector-part-row:visible" else source_row if selector == "app-estimate-row" else MagicMock()
        crawler = CCCCrawler(page, "user", "password", Path("."))
        crawler.open_part_type_catalog = MagicMock()
        crawler.load_catalog_categories = MagicMock()
        crawler.select_catalog_category = MagicMock(return_value=True)
        crawler.wait_for_catalog_element = MagicMock(return_value=True)
        crawler.resolve_part_prompts = MagicMock()
        crawler.read_selected_part_properties = MagicMock(return_value=("OEM1", "Center console"))
        crawler.dismiss_enter_part_price_prompt = MagicMock()
        self.assertEqual(crawler.open_part_type_scenario_results("CONSOLE", "CONSOLE"), (None, "OEM1", "Center console"))
        self.assertTrue(crawler.no_sourceable_candidates)

    def test_row_dispatch_uses_part_type_lookup_and_selected_properties(self):
        page = MagicMock()
        page.is_closed.return_value = False
        crawler = CCCCrawler(page, "user", "password", Path("."))
        crawler.recover_closed_page = MagicMock()
        crawler.open_new_workfile = MagicMock()
        crawler.enter_vehicle = MagicMock(return_value=True)
        crawler.open_part_type_scenario_results = MagicMock(return_value=(0, "OEM1", "Bumper cover"))
        crawler.read_vendor_inventory = MagicMock(return_value=([], []))
        result = crawler.crawl_row({**self.record, "_category": "Front Bumper US Built"})
        crawler.open_part_type_scenario_results.assert_called_once_with("front bumper", "Front Bumper US Built")
        self.assertEqual((result["OEM"], result["Part Description"]), ("OEM1", "Bumper cover"))

    def test_resume_after_interruption_keeps_each_category_and_duplicates(self):
        with TemporaryDirectory() as directory:
            output = Path(directory) / "results.xlsx"
            crawler = MagicMock()
            categories = ["Front Bumper US Built", "Front Bumper Japan Built"]
            crawler.discover_part_type_categories.return_value = categories
            crawler.crawl_row.side_effect = [{"OEM": "US-OEM", "Status": "OK"}, KeyboardInterrupt()]
            records = [self.record, {**self.record, "_row": "3"},
                       {**self.record, "Scenario": "rear bumper", "_row": "4"}]
            unrelated = {"VIN": "OTHER", "OEM": "KEEP", "Status": "OK"}
            with self.assertRaises(KeyboardInterrupt):
                run_part_type_groups(crawler, records, [unrelated], output, "A/M: CAPA")
            saved = load_results(output, preserve_all=True)
            self.assertEqual(len(saved), 3)
            self.assertEqual(json.loads(saved[1]["Scenario Categories"]), categories)
            crawler.reset_mock()
            crawler.discover_part_type_categories.return_value = ["Rear Bumper"]
            crawler.crawl_row.side_effect = [
                {"OEM": "JP-OEM", "Status": "OK"}, {"OEM": "REAR-OEM", "Status": "OK"}
            ]
            results = run_part_type_groups(crawler, records, saved, output, "A/M: CAPA")
            self.assertEqual([call.args[0]["_category"] for call in crawler.crawl_row.call_args_list],
                             ["Front Bumper Japan Built", "Rear Bumper"])
            self.assertEqual(crawler.discover_part_type_categories.call_count, 1)
            self.assertEqual(len(results), 6)
            self.assertEqual(len(load_results(output, preserve_all=True)), 6)
            self.assertEqual(sum(row.get("OEM") == "US-OEM" for row in results), 2)
            self.assertEqual(sum(row.get("OEM") == "JP-OEM" for row in results), 2)
            self.assertTrue(any(row.get("OEM") == "KEEP" for row in results))

    def test_failed_category_is_retried_without_repeating_success(self):
        with TemporaryDirectory() as directory:
            output = Path(directory) / "results.xlsx"
            crawler = MagicMock()
            crawler.discover_part_type_categories.return_value = ["Front Bumper US Built", "Front Bumper Japan Built"]
            crawler.crawl_row.side_effect = [{"OEM": "A", "Status": "OK"}, {"Status": "ERROR: failed"}]
            first = run_part_type_groups(crawler, [self.record], [], output, "A/M: CAPA")
            crawler.reset_mock()
            crawler.crawl_row.side_effect = [{"OEM": "B", "Status": "OK"}]
            second = run_part_type_groups(crawler, [self.record], first, output, "A/M: CAPA")
            crawler.discover_part_type_categories.assert_not_called()
            self.assertEqual(crawler.crawl_row.call_args.args[0]["_category"], "Front Bumper Japan Built")
            self.assertEqual({row["OEM"] for row in second}, {"A", "B"})

    def test_interrupted_retry_preserves_other_saved_failures(self):
        with TemporaryDirectory() as directory:
            output = Path(directory) / "results.xlsx"
            crawler = MagicMock()
            crawler.discover_part_type_categories.return_value = ["Front Bumper US Built", "Front Bumper Japan Built"]
            crawler.crawl_row.side_effect = [{"Status": "ERROR: first"}, {"Status": "ERROR: second"}]
            first = run_part_type_groups(crawler, [self.record], [], output, "A/M: CAPA")
            crawler.crawl_row.side_effect = [{"OEM": "A", "Status": "OK"}, KeyboardInterrupt()]
            with self.assertRaises(KeyboardInterrupt):
                run_part_type_groups(crawler, [self.record], first, output, "A/M: CAPA")
            saved = load_results(output, preserve_all=True)
            self.assertEqual({row["Status"] for row in saved}, {"OK", "ERROR: second"})

    def test_corrupt_market_data_is_not_treated_as_complete(self):
        row = {"Status": "No offers matching description: A/M: CAPA",
               "Description Target": "A/M: CAPA", "market_data": "invalid"}
        self.assertFalse(part_type_result_complete(row, "A/M: CAPA"))
        row["market_data"] = "[]"
        self.assertTrue(part_type_result_complete(row, "A/M: CAPA"))

    def test_existing_output_is_not_overwritten_without_resume(self):
        with TemporaryDirectory() as directory:
            input_path = Path(directory) / "input.xlsx"
            output_path = Path(directory) / "output.xlsx"
            workbook = Workbook()
            workbook.active.append(["VIN", "Part Type Code", "Part Type Description", "Scenario"])
            workbook.active.append(list(self.record.values())[:4])
            workbook.save(input_path)
            workbook.close()
            output_path.write_bytes(b"existing data")
            with patch.dict("os.environ", {"CCC_USERNAME": "user", "CCC_PASSWORD": "password"}), \
                 patch("sys.argv", ["crawler.py", "--search-action", "vin_part_type_scenario",
                                    "--input", str(input_path), "--output", str(output_path)]):
                self.assertEqual(main(), 2)
            self.assertEqual(output_path.read_bytes(), b"existing data")

    def test_incompatible_resume_workbook_is_not_overwritten(self):
        with TemporaryDirectory() as directory:
            input_path = Path(directory) / "input.xlsx"
            output_path = Path(directory) / "output.xlsx"
            workbook = Workbook()
            workbook.active.append(["VIN", "Part Type Code", "Part Type Description", "Scenario"])
            workbook.active.append(["VIN1", "101", "Bumper", "front bumper"])
            workbook.save(input_path)
            workbook.close()
            workbook = Workbook()
            workbook.active.append(["VIN", "OEM"])
            workbook.active.append(["OLD", "KEEP"])
            workbook.save(output_path)
            workbook.close()
            original = output_path.read_bytes()
            with patch.dict("os.environ", {"CCC_USERNAME": "user", "CCC_PASSWORD": "password"}), \
                 patch("sys.argv", ["crawler.py", "--search-action", "vin_part_type_scenario", "--resume",
                                    "--input", str(input_path), "--output", str(output_path)]):
                self.assertEqual(main(), 2)
            self.assertEqual(output_path.read_bytes(), original)

    def test_unexpected_error_does_not_replace_checkpoint_with_empty_results(self):
        with TemporaryDirectory() as directory:
            input_path = Path(directory) / "input.xlsx"
            output_path = Path(directory) / "output.xlsx"
            workbook = Workbook()
            workbook.active.append(["VIN", "Part Type Code", "Part Type Description", "Scenario"])
            workbook.active.append(["VIN1", "101", "Bumper", "front bumper"])
            workbook.save(input_path)
            workbook.close()

            def fail_after_checkpoint(_crawler, _records, _results, output, _target):
                save_results(output, [{"VIN": "VIN1", "OEM": "SAVED", "Status": "OK"}], PART_TYPE_OUTPUT_COLUMNS)
                raise RuntimeError("unexpected failure")

            with patch.dict("os.environ", {"CCC_USERNAME": "user", "CCC_PASSWORD": "password"}), \
                 patch("sys.argv", ["crawler.py", "--search-action", "vin_part_type_scenario",
                                    "--input", str(input_path), "--output", str(output_path)]), \
                 patch("crawler.AUTH_STATE", Path(directory) / "auth" / "state.json"), \
                 patch("crawler.sync_playwright"), patch("crawler.CCCCrawler"), \
                 patch("crawler.run_part_type_groups", side_effect=fail_after_checkpoint):
                self.assertEqual(main(), 1)
            self.assertEqual(load_results(output_path)[0]["OEM"], "SAVED")


if __name__ == "__main__":
    unittest.main()