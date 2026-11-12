import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import MagicMock, patch
from openpyxl import Workbook, load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from crawler import (
    CATALOG_LOADING_SELECTOR,
    CCCCrawler,
    FatalCrawlerError,
    CatalogSearchTimeoutError,
    MODEL_MISMATCH_RE,
    NO_SOURCE_PARTS_RE,
    OUTPUT_COLUMNS,
    PartOffer,
    catalog_not_found_error,
    choose_vin_vehicle_option,
    expand_group_result,
    is_no_catalog_results,
    load_results,
    remaining_catalog_categories,
    prepare_result_slots,
    resume_state,
    read_input,
    scenario_path_parts,
    save_results,
    should_retry_row_status,
    vin_model_year_matches,
    validate_selected_vehicle_year,
)


class NoCatalogResultsDetectionTests(unittest.TestCase):
    def test_result_columns_put_partslink_before_oem_without_og_oem(self):
        self.assertIn("Partslink", OUTPUT_COLUMNS)
        self.assertIn("OEM", OUTPUT_COLUMNS)
        self.assertIn("Part Description", OUTPUT_COLUMNS)
        self.assertNotIn("OG OEM", OUTPUT_COLUMNS)

        with TemporaryDirectory() as directory:
            output_path = Path(directory) / "scenario-output.xlsx"
            save_results(output_path, [{"VIN": "VIN123", "Partslink": "MI3030181", "OEM": "92110W040P"}])
            workbook = load_workbook(output_path, read_only=True, data_only=True)
            headers = next(workbook.active.iter_rows(values_only=True))
            workbook.close()

        self.assertIn("Partslink", headers)
        self.assertIn("OEM", headers)
        self.assertNotIn("OG OEM", headers)

    def test_resume_converts_legacy_og_oem_output_columns(self):
        with TemporaryDirectory() as directory:
            output_path = Path(directory) / "legacy-output.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.append(["VIN", "OEM", "OG OEM", "Status"])
            worksheet.append(["VIN123", "MI3030181", "92110W040P", "OK"])
            workbook.save(output_path)
            workbook.close()

            results = load_results(output_path)

        self.assertEqual(results[0]["Partslink"], "MI3030181")
        self.assertEqual(results[0]["OEM"], "92110W040P")
        self.assertNotIn("OG OEM", results[0])

    def test_scenario_path_splits_nonempty_catalog_levels(self):
        self.assertEqual(
            scenario_path_parts(" AIR CONDITIONER & HEATER > CONDENSER, COMPRESSOR & LINES > CONDENSER "),
            ["AIR CONDITIONER & HEATER", "CONDENSER, COMPRESSOR & LINES", "CONDENSER"],
        )

    def test_scenario_input_reads_enriched_workbook_and_keeps_partslink_separate(self):
        with TemporaryDirectory() as directory:
            input_path = Path(directory) / "scenario-input.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.append(["VIN", "PARTSLINK", "Scenario"])
            worksheet.append([
                "VIN123",
                "MI3030181",
                "AIR CONDITIONER & HEATER > CONDENSER, COMPRESSOR & LINES > CONDENSER",
            ])
            workbook.save(input_path)
            workbook.close()

            records = read_input(input_path, "vin_partslink_scenario")

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["PARTSLINK"], "MI3030181")
        self.assertNotIn("OEM", records[0])
        self.assertTrue(records[0]["Scenario"].endswith("> CONDENSER"))

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

    def test_vin_model_year_code_selects_matching_ccc_option(self):
        vin = "2G1FB1E39D9148021"

        self.assertTrue(vin_model_year_matches(vin, 2013))
        self.assertFalse(vin_model_year_matches(vin, 2011))
        self.assertEqual(
            choose_vin_vehicle_option(vin, [
                "2011 Chevrolet Equinox LT2 FWD",
                "2013 Chevrolet Camaro LT1",
            ]),
            1,
        )

    def test_vehicle_option_falls_back_when_ccc_options_have_no_year(self):
        self.assertEqual(
            choose_vin_vehicle_option("2G1FB1E39D9148021", ["Camaro LT1", "Camaro SS"]),
            0,
        )

    def test_incompatible_ccc_vehicle_year_is_retryable_state_error(self):
        with self.assertRaisesRegex(FatalCrawlerError, "model year 2011.*conflicts with VIN"):
            validate_selected_vehicle_year("2G1FB1E39D9148021", "2011")
        validate_selected_vehicle_year("2G1FB1E39D9148021", "2013")

    def test_rejects_valid_catalog_pages(self):
        body = "A/C Compressor 1234 part list"
        self.assertFalse(is_no_catalog_results(body))

    def test_only_explicit_empty_message_is_catalog_not_found(self):
        error = catalog_not_found_error("No results found for search text: CJ5Z13008B", "CJ5Z13008B")
        self.assertEqual(str(error), "No catalog results found for OEM CJ5Z13008B")
        self.assertEqual(
            str(catalog_not_found_error(
                'No results found for search text: "CJ5Z13008B"',
                "CJ5Z13008B",
                visible_category_count=47,
            )),
            "No catalog results found for OEM CJ5Z13008B",
        )
        self.assertIsNone(
            catalog_not_found_error(
                "No results found for search text: DIFFERENTOEM",
                "CJ5Z13008B",
                visible_category_count=47,
            )
        )

        self.assertIsNone(catalog_not_found_error("", "CJ5Z13008B"))
        self.assertIsNone(catalog_not_found_error("Loading catalog...", "CJ5Z13008B"))
        self.assertIsNone(
            catalog_not_found_error(
                "Source parts 0; generic vehicle categories are visible",
                "CJ5Z13008B",
                visible_category_count=47,
            )
        )

    def test_resume_retries_failures_and_excludes_nonexact_capa_market_rows(self):
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

        self.assertEqual(completed[("VIN1", "OEM1")], 1)
        self.assertNotIn(("VIN1", "OEM1"), retry_rows)
        self.assertEqual(list(retry_rows[("VIN2", "OEM2")]), [1])
        self.assertEqual(list(retry_rows[("VIN3", "OEM3")]), [2])
        self.assertEqual(completed[("VIN4", "OEM4")], 1)
        self.assertEqual(completed[("VIN5", "OEM5")], 1)
        self.assertNotIn(("VIN5", "OEM5"), retry_rows)
        self.assertEqual(completed[("VIN6", "OEM6")], 1)

    def test_exact_capa_description_is_included_in_summary(self):
        result = CCCCrawler.summarize(
            "VIN1",
            "OEM1",
            [
                PartOffer("iPart Automotive Inc.", "A/M: CAPA", 267.0),
                PartOffer("Other supplier", "A/M", 100.0),
            ],
        )

        self.assertEqual(result["Status"], "OK")
        self.assertTrue(result["Ipart listed on CCC"])
        self.assertEqual(result["Ipart List Price CCC"], 267.0)
        self.assertEqual(result["Lowest CCC Price"], 267.0)

    def test_summary_writes_selected_estimate_part_number_as_oem(self):
        result = CCCCrawler.summarize(
            "VIN1",
            "92110W040P",
            [PartOffer("iPart Automotive Inc.", "A/M: CAPA", 267.0)],
            partslink="MI3030181",
            part_description="Condenser",
        )

        self.assertEqual(result["OEM"], "92110W040P")
        self.assertEqual(result["Partslink"], "MI3030181")
        self.assertEqual(result["Part Description"], "Condenser")
        self.assertNotIn("OG OEM", result)

    def test_edit_properties_waits_for_part_number_and_description_before_reading(self):
        page = MagicMock()
        edit_button = MagicMock()
        edit_button.is_visible.return_value = True
        edit_buttons = MagicMock()
        edit_buttons.count.return_value = 1
        edit_buttons.nth.return_value = edit_button
        close_buttons = MagicMock()
        close_buttons.count.return_value = 0
        page.get_by_role.side_effect = [edit_buttons, close_buttons]

        part_number_label = MagicMock()
        description_label = MagicMock()
        part_number_label.last = part_number_label
        description_label.last = description_label
        page.get_by_text.side_effect = [part_number_label, description_label]
        overlay_dialog = MagicMock()
        overlay_dialog.count.return_value = 1
        overlay_locator = MagicMock()
        overlay_locator.filter.return_value.last = overlay_dialog
        page.locator.return_value = overlay_locator

        events = []
        part_number_label.wait_for.side_effect = lambda **_: events.append("part-number-ready")
        description_label.wait_for.side_effect = lambda **_: events.append("description-ready")
        page.evaluate.side_effect = lambda _: events.append("read-values") or {
            "partNumber": "92110W040P",
            "description": "Condenser",
        }
        crawler = CCCCrawler(page, "user", "password", Path("."))

        values = crawler.read_selected_part_properties("Condenser")

        self.assertEqual(values, ("92110W040P", "Condenser"))
        self.assertEqual(events, ["part-number-ready", "description-ready", "read-values"])

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
            "No sourceable parts available; skipped",
            "Vendor search failed after 3 Retry clicks; skipped",
        ):
            with self.subTest(status=status):
                self.assertTrue(should_retry_row_status(status))

            self.assertFalse(should_retry_row_status("No catalog results found for OEM OEM1; skipped"))
        self.assertFalse(should_retry_row_status("No offers matching description: A/M: CAPA"))

    def test_wait_for_catalog_element_extends_wait_while_ccc_is_loading(self):
        page = MagicMock()
        body = MagicMock()
        body.inner_text.return_value = "Loading catalog"
        loading = MagicMock()
        loading.count.return_value = 1
        page.locator.side_effect = lambda selector: loading if selector == CATALOG_LOADING_SELECTOR else body
        locator = MagicMock()
        locator.first.wait_for.side_effect = [PlaywrightTimeoutError("quick wait"), None]
        crawler = CCCCrawler(page, "user", "password", MagicMock())

        with patch("crawler.CATALOG_QUICK_WAIT_MS", 10), patch(
            "crawler.CATALOG_LOADING_WAIT_MS", 25
        ):
            self.assertTrue(crawler.wait_for_catalog_element(locator, "OEM1", "checking a category"))

        waits = locator.first.wait_for.call_args_list
        self.assertEqual(len(waits), 2)
        self.assertEqual(waits[0].kwargs["timeout"], 10)
        self.assertEqual(waits[1].kwargs["timeout"], 25)

    def test_wait_for_catalog_element_does_not_extend_without_loading_indicator(self):
        page = MagicMock()
        body = MagicMock()
        body.inner_text.return_value = "Catalog ready"
        loading = MagicMock()
        loading.count.return_value = 0
        page.locator.side_effect = lambda selector: loading if selector == CATALOG_LOADING_SELECTOR else body
        locator = MagicMock()
        locator.first.wait_for.side_effect = PlaywrightTimeoutError("quick wait")
        locator.count.return_value = 0
        crawler = CCCCrawler(page, "user", "password", MagicMock())

        with patch("crawler.CATALOG_QUICK_WAIT_MS", 10):
            self.assertFalse(crawler.wait_for_catalog_element(locator, "OEM1", "checking a category"))

        locator.first.wait_for.assert_called_once()

    def test_load_catalog_categories_opens_selected_group_menu(self):
        page = MagicMock()
        categories = MagicMock()
        categories.all_inner_texts.return_value = ["FRONT BUMPER", "FRONT LAMPS"]
        selected_group = MagicMock()
        selected_group.count.return_value = 1
        page.locator.side_effect = lambda selector: {
            "button.group-list-item:visible": categories,
            "button.selected-group-button:visible": selected_group,
        }[selector]
        crawler = CCCCrawler(page, "user", "password", MagicMock())

        with patch.object(crawler, "wait_for_catalog_element", side_effect=[False, True]):
            returned_categories, category_names = crawler.load_catalog_categories("OEM1")

        self.assertIs(returned_categories, categories)
        self.assertEqual(category_names, ["FRONT BUMPER", "FRONT LAMPS"])
        selected_group.first.click.assert_called_once_with(force=True)

    def test_load_catalog_categories_does_not_treat_page_text_as_controls(self):
        page = MagicMock()
        categories = MagicMock()
        categories.all_inner_texts.return_value = []
        selected_group = MagicMock()
        selected_group.count.return_value = 0
        body = MagicMock()
        body.inner_text.return_value = "FRONT BUMPER\nFRONT LAMPS\nTIRE"
        page.locator.side_effect = lambda selector: {
            "button.group-list-item:visible": categories,
            "button.selected-group-button:visible": selected_group,
            "button.selected-group-button": selected_group,
            "body": body,
        }[selector]
        crawler = CCCCrawler(page, "user", "password", MagicMock())

        with patch.object(crawler, "wait_for_catalog_element", return_value=False):
            with self.assertRaisesRegex(CatalogSearchTimeoutError, "category controls did not load"):
                crawler.load_catalog_categories("OEM1")

        categories.all_inner_texts.assert_called_once_with()
        selected_group.first.click.assert_not_called()

    def test_select_catalog_category_reopens_menu_before_clicking_category(self):
        page = MagicMock()
        categories = MagicMock()
        categories.count.side_effect = [2, 2, 2, 2]
        first_category = MagicMock()
        first_category.inner_text.return_value = "FRONT BUMPER"
        subgroup = MagicMock()
        subgroup.inner_text.return_value = "BUMPER COVER"
        requested_category = MagicMock()
        requested_category.inner_text.return_value = "GRILLE"
        categories.nth.side_effect = [first_category, subgroup, first_category, requested_category]
        selected_group = MagicMock()
        selected_group.count.return_value = 1
        page.locator.side_effect = lambda selector: {
            "button.group-list-item:visible": categories,
            "button.selected-group-button:visible": selected_group,
        }[selector]
        crawler = CCCCrawler(page, "user", "password", MagicMock())

        with patch.object(crawler, "wait_for_catalog_element", return_value=True) as wait_for:
            crawler.select_catalog_category("GRILLE", "OEM1")

        selected_group.first.click.assert_called_once_with(force=True)
        wait_for.assert_called_once_with(categories, "OEM1", "reopening catalog categories")
        first_category.click.assert_not_called()
        requested_category.click.assert_called_once_with(force=True)

    def test_select_catalog_category_uses_attached_group_control_when_not_visible(self):
        page = MagicMock()
        categories = MagicMock()
        categories.count.side_effect = [0, 1]
        requested_category = MagicMock()
        requested_category.inner_text.return_value = "GRILLE"
        categories.nth.return_value = requested_category
        hidden_group = MagicMock()
        hidden_group.count.return_value = 0
        attached_group = MagicMock()
        attached_group.count.return_value = 1
        page.locator.side_effect = lambda selector: {
            "button.group-list-item:visible": categories,
            "button.selected-group-button:visible": hidden_group,
            "button.selected-group-button": attached_group,
        }[selector]
        crawler = CCCCrawler(page, "user", "password", MagicMock())

        with patch.object(crawler, "wait_for_catalog_element", return_value=True):
            crawler.select_catalog_category("GRILLE", "OEM1")

        attached_group.first.click.assert_called_once_with(force=True)
        requested_category.click.assert_called_once_with(force=True)

    def test_select_catalog_category_reapplies_oem_filter_when_group_control_is_absent(self):
        page = MagicMock()
        categories = MagicMock()
        categories.count.side_effect = [0, 1]
        refreshed_category = MagicMock()
        refreshed_category.inner_text.return_value = "FRONT LAMPS"
        categories.nth.return_value = refreshed_category
        no_group = MagicMock()
        no_group.count.return_value = 0
        page.locator.side_effect = lambda selector: {
            "button.group-list-item:visible": categories,
            "button.selected-group-button:visible": no_group,
            "button.selected-group-button": no_group,
        }[selector]
        crawler = CCCCrawler(page, "user", "password", MagicMock())
        crawler.apply_oem_catalog_filter = MagicMock()

        with patch.object(crawler, "wait_for_catalog_element", return_value=True):
            selected = crawler.select_catalog_category("GRILLE", "OEM1")

        crawler.apply_oem_catalog_filter.assert_called_once_with("OEM1", open_estimate=False)
        self.assertFalse(selected)
        refreshed_category.click.assert_not_called()

    def test_refreshed_category_queue_skips_categories_already_checked(self):
        pending = remaining_catalog_categories(
            ["FRONT BUMPER", "Grille", "Front Lamps"],
            {"front bumper", "front lamps"},
        )

        self.assertEqual(list(pending), ["Grille"])

    def test_catalog_scan_deadline_requests_retry(self):
        with patch("crawler.CATALOG_SCAN_TIMEOUT_SECONDS", 60), patch(
            "crawler.time.monotonic", return_value=61
        ):
            with self.assertRaisesRegex(RuntimeError, "exceeded 60 seconds.*retry required"):
                CCCCrawler.check_catalog_scan_deadline(0, "OEM1")

    def test_catalog_search_timeout_is_retried_without_stopping_the_row_loop(self):
        crawler = CCCCrawler(MagicMock(), "user", "password", MagicMock())
        crawler.open_new_workfile = MagicMock()
        crawler.enter_vehicle = MagicMock(return_value=True)
        crawler.open_part_results = MagicMock(
            side_effect=CatalogSearchTimeoutError("CCC remained busy; retry required")
        )

        with patch("crawler.MAX_ROW_RETRIES", 1):
            result = crawler.crawl_row({"VIN": "VIN1", "OEM": "OEM1", "_row": "2"})

        self.assertTrue(str(result["Status"]).startswith("ERROR: catalog search did not complete:"))
        self.assertIn("retry on resume", result["Status"])
        self.assertEqual(crawler.open_part_results.call_count, 2)
        self.assertFalse(crawler._workfile_initialized)

    def test_fatal_row_state_is_recorded_and_retried(self):
        crawler = CCCCrawler(MagicMock(), "user", "password", MagicMock())
        crawler.open_new_workfile = MagicMock()
        crawler.enter_vehicle = MagicMock(return_value=True)
        crawler.open_part_results = MagicMock(
            side_effect=FatalCrawlerError("vehicle selection state was unexpected")
        )

        with patch("crawler.MAX_ROW_RETRIES", 1):
            result = crawler.crawl_row({"VIN": "VIN1", "OEM": "OEM1", "_row": "2"})

        self.assertTrue(str(result["Status"]).startswith("ERROR: CCC row state could not be recovered:"))
        self.assertIn("retry on resume", result["Status"])
        self.assertEqual(crawler.open_part_results.call_count, 2)

    def test_closed_browser_tab_is_replaced_and_session_restored(self):
        page = MagicMock()
        page.is_closed.return_value = True
        context = MagicMock()
        replacement_page = MagicMock()
        context.new_page.return_value = replacement_page
        crawler = CCCCrawler(page, "user", "password", MagicMock(), context=context)
        crawler.sign_in_if_needed = MagicMock()
        crawler.open_workfiles = MagicMock()

        crawler.recover_closed_page()

        self.assertIs(crawler.page, replacement_page)
        replacement_page.route.assert_called_once()
        crawler.sign_in_if_needed.assert_called_once_with()
        context.storage_state.assert_called_once()
        crawler.open_workfiles.assert_called_once_with()
        self.assertFalse(crawler._workfile_initialized)

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
