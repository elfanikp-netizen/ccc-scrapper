"""CCC workfile crawler driven by an input Excel workbook."""

from __future__ import annotations

import argparse
from collections import Counter, defaultdict, deque
import logging
import os
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from playwright.sync_api import BrowserContext, Locator, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright


BASE_URL = "https://www.cccone.com/login/login-idp"
AUTH_STATE = Path("playwright/.auth/ccc-state.json")
OUTPUT_COLUMNS = [
    "VIN",
    "OEM",
    "Ipart listed on CCC",
    "Ipart List Price CCC",
    "Lowest CCC Price",
    "Lowest CCC Supplier",
    "Next Lowest CCC Price",
    "Next Lowest CCC Supplier",
    "Lowest CCC Price Supplier",
    "Status",
    "date_time",
]
IPART_VENDOR = "iPart Automotive Inc."
TARGET_DESCRIPTION = "A/M: CAPA"
MAX_VENDOR_SEARCH_RETRIES = 3
MODEL_MISMATCH_RE = re.compile(
    r"part does not match (?:(?:the )?(?:year[, ]+)?model(?: of the vehicle)?|the year of the vehicle)",
    re.I,
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger("ccc-crawler")


@dataclass
class PartOffer:
    vendor: str
    description: str
    total: float


class CatalogPartNotFoundError(RuntimeError):
    pass


def normalize(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip().casefold()


def parse_price(value: str) -> float | None:
    cleaned = re.sub(r"[^0-9.\-]", "", value.replace(",", ""))
    try:
        return float(cleaned)
    except ValueError:
        return None


def read_input(path: Path) -> list[dict[str, str]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    rows = sheet.iter_rows(values_only=True)
    headers = [normalize(str(value or "")) for value in next(rows, ())]
    try:
        vin_index = headers.index("vin")
        oem_index = headers.index("oem")
    except ValueError as error:
        raise ValueError("Input workbook must have VIN and OEM column headers.") from error

    records = []
    for row_number, row in enumerate(rows, start=2):
        vin = str(row[vin_index] or "").strip() if len(row) > vin_index else ""
        oem_cell = str(row[oem_index] or "").strip() if len(row) > oem_index else ""
        oems = [value.strip() for value in oem_cell.split(",") if value.strip()]
        if vin and oems:
            for oem_index_in_cell, oem in enumerate(oems, start=1):
                row_label = str(row_number) if len(oems) == 1 else f"{row_number}.{oem_index_in_cell}"
                records.append({"VIN": vin, "OEM": oem, "_row": row_label})
        elif vin or oem_cell:
            log.warning("Skipping input row %d: VIN or OEM is blank", row_number)
    workbook.close()
    return records


def save_results(path: Path, results: list[dict[str, Any]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "CCC Results"
    sheet.append(OUTPUT_COLUMNS)
    for result in results:
        sheet.append([result.get(column) for column in OUTPUT_COLUMNS])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    workbook.save(path)


def load_results(path: Path) -> list[dict[str, Any]]:
    if not path.is_file():
        return []
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.active
        rows = sheet.iter_rows(values_only=True)
        headers = [str(value or "").strip() for value in next(rows, ())]
        header_indexes = {header: index for index, header in enumerate(headers)}
        if "VIN" not in header_indexes or "OEM" not in header_indexes:
            log.warning("Existing output has no VIN/OEM columns; starting without prior results")
            return []
        results = []
        for row in rows:
            result = {
                column: row[index] if index < len(row) else None
                for column, index in header_indexes.items()
                if column in OUTPUT_COLUMNS
            }
            if result.get("VIN") and result.get("OEM"):
                results.append(result)
        return results
    finally:
        workbook.close()


def resume_state(results: list[dict[str, Any]]) -> tuple[Counter, dict[tuple[str, str], deque[int]]]:
    completed: Counter = Counter()
    retry_rows: dict[tuple[str, str], deque[int]] = defaultdict(deque)
    for index, result in enumerate(results):
        key = (str(result.get("VIN") or "").strip(), str(result.get("OEM") or "").strip())
        if not all(key):
            continue
        status = str(result.get("Status") or "").strip()
        if status.upper().startswith("ERROR:"):
            retry_rows[key].append(index)
        else:
            completed[key] += 1
    return completed, retry_rows


class CCCCrawler:
    def __init__(
        self,
        page: Page,
        username: str,
        password: str,
        artifacts: Path,
        description_target: str = TARGET_DESCRIPTION,
    ):
        self.page = page
        self.username = username
        self.password = password
        self.artifacts = artifacts
        self.description_target = description_target
        self.workfiles_url: str | None = None
        self.artifacts.mkdir(parents=True, exist_ok=True)

    def checkpoint(self, label: str) -> None:
        log.info("%s", label)

    @staticmethod
    def first_visible(locator: Locator) -> Locator | None:
        for index in range(locator.count()):
            candidate = locator.nth(index)
            try:
                if candidate.is_visible():
                    return candidate
            except Exception:
                continue
        return None

    def find_textbox(self, hints: tuple[str, ...]) -> Locator:
        for hint in hints:
            candidates = [
                self.page.get_by_label(re.compile(hint, re.I)),
                self.page.get_by_placeholder(re.compile(hint, re.I)),
                self.page.locator(f'input[aria-label*="{hint}" i]'),
                self.page.locator(f'input[name*="{hint}" i]'),
            ]
            for candidate in candidates:
                visible = self.first_visible(candidate)
                if visible is not None:
                    return visible
        visible_inputs = self.page.locator('input:visible:not([type="hidden"]):not([type="checkbox"])')
        if visible_inputs.count() == 1:
            return visible_inputs.first
        raise RuntimeError(f"Could not identify a visible text box using hints: {hints}")

    def click_text(self, label: str, *, exact: bool = True, timeout: int = 8000) -> None:
        candidates = [
            self.page.get_by_role("button", name=re.compile(re.escape(label), re.I)),
            self.page.get_by_role("link", name=re.compile(re.escape(label), re.I)),
            self.page.get_by_text(label, exact=exact),
        ]
        for candidate in candidates:
            visible = self.first_visible(candidate)
            if visible is not None:
                try:
                    visible.click(timeout=timeout)
                    return
                except Exception:
                    continue
        raise RuntimeError(f"Could not find clickable control: {label}")

    def wait_for_textbox(self, hints: tuple[str, ...], timeout: int = 30000) -> Locator:
        try:
            self.page.wait_for_function(
                """hints => [...document.querySelectorAll('input')].some(input => {
                    if (!input.offsetWidth && !input.offsetHeight) return false;
                    const text = [input.placeholder, input.getAttribute('aria-label'), input.name, input.id].join(' ');
                    return hints.some(hint => new RegExp(hint, 'i').test(text));
                })""",
                list(hints),
                timeout=timeout,
            )
        except PlaywrightTimeoutError:
            pass
        return self.find_textbox(hints)

    def capture_failure(self, label: str) -> None:
        safe_label = re.sub(r"[^A-Za-z0-9_-]+", "_", label)[:60]
        try:
            self.page.screenshot(path=self.artifacts / f"{safe_label}.png", full_page=True)
            (self.artifacts / f"{safe_label}.txt").write_text(self.page.locator("body").inner_text()[:20000], encoding="utf-8")
            log.error("Saved failure evidence under %s", self.artifacts)
        except Exception as error:
            log.error("Could not save failure evidence: %s", error)

    def sign_in_if_needed(self) -> None:
        self.checkpoint("Opening CCC and checking sign-in state")
        self.page.goto(BASE_URL, wait_until="domcontentloaded")
        login_button = self.page.get_by_role("button", name=re.compile("^Login$", re.I))
        try:
            login_button.wait_for(state="visible", timeout=15000)
            self.checkpoint("CCCOne login landing page found; clicking Login")
            login_button.click()
        except PlaywrightTimeoutError:
            log.info("CCCOne login landing button not shown; checking the current session page")

        self.checkpoint(f"Identity page reached: {self.page.url}")
        self.checkpoint("Waiting for the sign-in form or authenticated navigation to load")
        try:
            self.page.wait_for_function(
                r"""() => {
                    const visible = element => !!element && !!(element.offsetWidth || element.offsetHeight);
                    const username = visible(document.querySelector('#idp-discovery-username, input[name="username"]'));
                    const password = [...document.querySelectorAll('input[type="password"]')].some(visible);
                    const body = (document.body?.innerText || '').toLowerCase();
                    const authenticatedNavigation = (location.hostname === 'www.cccone.com' && !location.pathname.startsWith('/login/'))
                        || (body.includes('workfiles') && !body.includes('sign in'));
                    return username || password || authenticatedNavigation;
                }""",
                timeout=30000,
            )
        except PlaywrightTimeoutError as error:
            self.capture_failure("ccc_entry_surface_not_loaded")
            raise RuntimeError("CCC did not show a login form or authenticated navigation within 30 seconds") from error

        password_box = self.page.locator('input[type="password"]:visible').first
        if not password_box.count():
            username_box = self.page.locator('#idp-discovery-username:visible, input[name="username"]:visible').first
            if not username_box.count():
                if self.page.url.startswith("https://www.cccone.com/") and "/login/" not in self.page.url:
                    log.info("Authenticated CCCOne application route detected; skipping sign-in")
                    return
                raise RuntimeError("CCC loaded, but neither a username field nor authenticated application route was found")

            username_box.fill(self.username)
            for attempt in range(5):
                if self.page.locator('input[type="password"]:visible').count():
                    break
                self.checkpoint(f"Advancing username sign-in step ({attempt + 1}/5)")
                self.click_text("Next")
                try:
                    self.page.locator('input[type="password"]').wait_for(state="visible", timeout=8000)
                except PlaywrightTimeoutError:
                    continue
            password_box = self.page.locator('input[type="password"]:visible').first
        if not password_box.count():
            raise RuntimeError("Password field did not appear after advancing the sign-in flow")
        if password_box.count():
            password_box.fill(self.password)
            self.click_text("Sign in")
            self.page.wait_for_load_state("domcontentloaded")
            try:
                self.page.locator('input[type="password"]:visible').wait_for(state="hidden", timeout=30000)
            except PlaywrightTimeoutError:
                log.warning("Password field remains visible; sign-in may require an additional prompt")
        log.info("Sign-in flow completed or an existing session was detected")

    def open_workfiles(self) -> None:
        self.checkpoint("Opening CCCOne navigation menu")
        self.page.get_by_role("button", name="Nav Menu").click()
        workfiles = self.page.get_by_text("Workfiles", exact=True)
        workfiles.first.wait_for(state="visible", timeout=15000)
        workfiles.first.click()
        self.page.wait_for_function(
            """() => [...document.querySelectorAll('*')].filter(element => {
                const rect = element.getBoundingClientRect();
                return rect.width && rect.height && element.children.length === 0 && element.textContent.trim() === 'Workfiles';
            }).length >= 2""",
            timeout=10000,
        )
        self.page.get_by_text("Workfiles", exact=True).last.click()
        self.page.wait_for_url("**/workfile-search/workfile", timeout=30000)
        self.workfiles_url = self.page.url
        self.checkpoint(f"Workfiles list opened: {self.workfiles_url}")

    def open_new_workfile(self) -> None:
        if not self.workfiles_url:
            raise RuntimeError("Workfiles list URL is not initialized")
        if self.page.url != self.workfiles_url:
            self.page.goto(self.workfiles_url, wait_until="domcontentloaded")
        self.checkpoint("Creating a new workfile for this input row")
        self.page.get_by_text("New workfile", exact=True).click(timeout=15000)
        self.page.locator('a[aria-label="Vehicle"]').wait_for(state="visible", timeout=30000)
        log.info("New workfile opened: %s", self.page.url)

    def return_to_workfiles(self) -> None:
        if self.workfiles_url:
            self.page.goto(self.workfiles_url, wait_until="domcontentloaded")

    def enter_vehicle(self, vin: str) -> bool:
        self.checkpoint(f"Vehicle: searching VIN {vin}")
        self.page.locator('a[aria-label="Vehicle"]').click()
        self.page.get_by_role("button", name=re.compile("^Select vehicle$", re.I)).click()
        vin_box = self.page.locator(".vin-input input.input-box:visible").last
        vin_box.wait_for(state="visible", timeout=10000)
        vin_box.fill(vin)
        self.page.get_by_text("Valid VIN", exact=True).wait_for(state="visible", timeout=20000)
        self.checkpoint("CCC validated the VIN; advancing to model selection")
        self.click_text("Next")
        options = self.page.locator('[role="option"]:visible')
        try:
            options.first.wait_for(state="visible", timeout=20000)
        except PlaywrightTimeoutError:
            done = self.page.get_by_role("button", name=re.compile("^DONE$", re.I))
            if self.first_visible(done) is not None:
                self.checkpoint("CCC auto-selected the only model result and advanced to Finish")
                self.finish_vehicle_selection()
                log.info("Vehicle selected for VIN %s", vin)
                return True
            log.warning("No vehicle model result found for VIN %s; skipping this row", vin)
            self.capture_failure(f"vehicle_not_found_{vin}")
            try:
                self.click_text("Cancel")
            except RuntimeError:
                pass
            return False
        options.first.click()
        self.click_text("Next")
        self.finish_vehicle_selection()
        log.info("Vehicle selected for VIN %s", vin)
        return True

    def finish_vehicle_selection(self) -> None:
        done = self.page.get_by_role("button", name=re.compile("^DONE$", re.I))
        done.wait_for(state="visible", timeout=30000)
        done.click()
        self.handle_vehicle_recall_prompts()

    def handle_vehicle_recall_prompts(self) -> None:
        recall_notice = self.page.get_by_text(re.compile("vehicle recall notice", re.I)).first
        try:
            recall_notice.wait_for(state="visible", timeout=2500)
        except PlaywrightTimeoutError:
            return

        self.checkpoint("Vehicle recall notice shown; acknowledging with OK")
        self.click_text("OK")

        print_notice = self.page.get_by_text(re.compile("print recall notice", re.I)).first
        try:
            print_notice.wait_for(state="visible", timeout=10000)
        except PlaywrightTimeoutError:
            log.info("No print-recall prompt appeared after acknowledging the recall notice")
            return

        self.checkpoint("Print recall notice prompt shown; choosing No")
        no_pattern = re.compile(r"^\s*NO\s*$", re.I)
        no_control = self.page.locator(".cdk-overlay-pane .dialog-actions button").filter(has_text=no_pattern)
        visible_control = self.first_visible(no_control)
        if visible_control is None:
            no_control = self.page.locator(".cdk-overlay-pane .dialog-actions").get_by_text(no_pattern)
            visible_control = self.first_visible(no_control)
        if visible_control is None:
            self.capture_failure("print_recall_no_control_not_found")
            raise RuntimeError("Print recall prompt appeared, but its No control was not found")
        visible_control.click()
        self.page.get_by_text(re.compile("Print recall notice", re.I)).wait_for(state="hidden", timeout=10000)

    def open_part_results(self, oem: str) -> int | None:
        self.checkpoint(f"Estimate: searching OEM {oem}")
        self.page.locator('a[aria-label="Estimate"]').click()
        filter_button = self.page.locator("button.filter-button:visible").first
        filter_button.wait_for(state="visible", timeout=20000)
        filter_button.click()
        keyword_box = self.page.locator("input.search-input:visible").last
        keyword_box.wait_for(state="visible", timeout=10000)
        keyword_box.fill(oem)
        self.click_text("ALL MODELS")
        self.click_text("OK")
        categories = self.page.locator("button.group-list-item:visible")
        try:
            categories.first.wait_for(state="visible", timeout=5000)
        except PlaywrightTimeoutError as error:
            category_labels = self.page.locator(".part-group-name:visible, .group-list-item:visible, [class*=group-list-item]:visible")
            if not category_labels.count():
                raise CatalogPartNotFoundError(f"No catalog results found for OEM {oem}") from error

        parts = self.page.locator("app-part-selector-part-row:visible")
        category_names = categories.all_inner_texts() if categories.count() else []
        if not category_names:
            category_names = [
                text.strip()
                for text in self.page.locator("body").inner_text().splitlines()
                if text.strip() and text.strip().isupper() and len(text.strip()) < 80
            ]
        selected_parts = False
        for category_name in category_names:
            self.page.wait_for_timeout(500)
            visible_categories = self.page.locator("button.group-list-item:visible")
            matching_category = None
            for index in range(visible_categories.count()):
                candidate = visible_categories.nth(index)
                if normalize(candidate.inner_text()) == normalize(category_name):
                    matching_category = candidate
                    break
            self.checkpoint(f"Checking catalog category {category_name.strip()}")
            if matching_category is not None:
                matching_category.click(force=True)
            else:
                category_text = self.page.get_by_text(category_name.strip(), exact=True)
                visible_text = [
                    category_text.nth(index)
                    for index in range(category_text.count())
                    if category_text.nth(index).is_visible()
                ]
                if not visible_text:
                    continue
                visible_text[-1].click(force=True)
            try:
                parts.first.wait_for(state="visible", timeout=5000)
                selected_parts = True
            except PlaywrightTimeoutError:
                pass
            for _ in range(10):
                if selected_parts or parts.count():
                    selected_parts = True
                    break
                subgroups = self.page.locator("button.subgroup-button:visible")
                collapsed_subgroups = [
                    subgroups.nth(index)
                    for index in range(subgroups.count())
                    if "expanded" not in (subgroups.nth(index).get_attribute("class") or "")
                ]
                if not collapsed_subgroups:
                    break
                self.checkpoint("Expanding a catalog subgroup")
                collapsed_subgroups[0].click()
                try:
                    parts.first.wait_for(state="visible", timeout=5000)
                    selected_parts = True
                    break
                except PlaywrightTimeoutError:
                    continue
            if selected_parts:
                break
            selected_group = self.page.locator("button.selected-group-button:visible")
            if selected_group.count():
                selected_group.click(force=True)
                self.page.locator("button.group-list-item:visible").first.wait_for(timeout=10000)
        if not selected_parts or not parts.count():
            raise CatalogPartNotFoundError(f"No catalog results found for OEM {oem} in any returned category")

        self.checkpoint("Double-clicking the first filtered catalog part")
        parts.first.dblclick()
        self.resolve_part_prompts()
        source_tab = self.page.locator("[role=tab].source-parts-tab")
        source_tab.wait_for(state="visible", timeout=10000)
        try:
            self.page.wait_for_function(
                """() => document.querySelector('[role="tab"].source-parts-tab')?.getAttribute('aria-disabled') === 'false'""",
                timeout=15000,
            )
        except PlaywrightTimeoutError as error:
            raise RuntimeError(f"Part selection did not complete for OEM {oem}; Source parts stayed disabled") from error
        source_tab.click()
        source_tag = self.page.get_by_role("button", name="Show sourced parts").first
        try:
            source_tag.wait_for(state="visible", timeout=10000)
        except PlaywrightTimeoutError as error:
            raise RuntimeError(f"CCC added no sourceable estimate line for OEM {oem}") from error
        source_tag.click()
        return self.wait_for_vendor_inventory()

    def wait_for_vendor_inventory(self) -> int | None:
        retry_count = 0
        self.vendor_search_retry_count = retry_count
        offer_rows = self.page.locator(".cdk-overlay-pane .table-row")
        retry_button = self.page.get_by_role("button", name=re.compile(r"retry|try again", re.I))
        searching_text = self.page.get_by_text("Searching vendor inventory", exact=False)

        while True:
            self.checkpoint("Waiting for CCC vendor inventory search")
            try:
                searching_text.wait_for(state="hidden", timeout=60000)
            except PlaywrightTimeoutError:
                log.warning("CCC vendor search is still showing its loading state")

            if self.first_visible(offer_rows) is not None:
                return retry_count
            retry_control = self.first_visible(retry_button)
            if retry_control is None:
                try:
                    self.page.wait_for_function(
                        """() => {
                            const visible = element => !!element && !!(element.offsetWidth || element.offsetHeight);
                            const offers = [...document.querySelectorAll('.cdk-overlay-pane .table-row')].some(visible);
                            const retry = [...document.querySelectorAll('button')].some(button => visible(button) && /retry|try again/i.test(button.innerText));
                            return offers || retry;
                        }""",
                        timeout=5000,
                    )
                except PlaywrightTimeoutError:
                    pass
                if self.first_visible(offer_rows) is not None:
                    return retry_count
                retry_control = self.first_visible(retry_button)
                if retry_control is None:
                    log.error("CCC vendor search ended without offers or a Retry control")
                    return None

            if retry_control is None:
                log.error("CCC vendor search ended without offers or a Retry control")
                return None
            if retry_count >= MAX_VENDOR_SEARCH_RETRIES:
                log.error("CCC vendor search still has no offers after %d retries", retry_count)
                return None

            retry_count += 1
            self.vendor_search_retry_count = retry_count
            self.checkpoint(f"Retrying CCC vendor inventory search ({retry_count}/{MAX_VENDOR_SEARCH_RETRIES})")
            retry_control.click()
            try:
                searching_text.wait_for(state="visible", timeout=10000)
            except PlaywrightTimeoutError:
                pass

    def resolve_part_prompts(self) -> None:
        for _ in range(4):
            try:
                self.page.wait_for_function(
                    """() => {
                        const text = (document.body?.innerText || '').toLowerCase();
                        return (text.includes('part does not match') && (text.includes('model') || text.includes('year')))
                            || text.includes('refinish this part')
                            || text.includes('select additional labor')
                            || text.includes('removing duplicate part');
                    }""",
                    timeout=3000,
                )
            except PlaywrightTimeoutError:
                return
            prompt = normalize(self.page.locator("body").inner_text())
            if MODEL_MISMATCH_RE.search(prompt):
                self.page.get_by_role("button", name=re.compile("^Yes$", re.I)).last.click()
                self.checkpoint("Accepted CCC model-mismatch confirmation for the selected part")
                self.page.get_by_text(MODEL_MISMATCH_RE).wait_for(state="hidden", timeout=10000)
            elif "refinish this part" in prompt:
                self.page.get_by_role("button", name=re.compile("^No$", re.I)).last.click()
                self.page.get_by_text(re.compile("Refinish this part", re.I)).wait_for(state="hidden", timeout=10000)
            elif "select additional labor" in prompt:
                self.page.get_by_role("button", name=re.compile("^OK$", re.I)).last.click()
                self.page.get_by_text(re.compile("Select additional labor", re.I)).wait_for(state="hidden", timeout=10000)
            elif "removing duplicate part" in prompt:
                self.page.get_by_role("button", name=re.compile("^OK$", re.I)).last.click()
                self.checkpoint("Acknowledged CCC duplicate-part notice")
                self.page.get_by_text(re.compile("Removing duplicate part", re.I)).wait_for(state="hidden", timeout=10000)
            else:
                return

    def read_offers(self) -> list[PartOffer]:
        offers: list[PartOffer] = []
        rows = self.page.locator(".cdk-overlay-pane .table-row")
        for row_index in range(rows.count()):
            row = rows.nth(row_index)
            description = row.locator(".column-description").inner_text().strip() if row.locator(".column-description").count() else ""
            vendor = row.locator(".column-vendor").inner_text().strip() if row.locator(".column-vendor").count() else ""
            ext_price_text = row.locator(".column-ext-price").inner_text().strip() if row.locator(".column-ext-price").count() else ""
            ext_price = parse_price(ext_price_text)
            if description and vendor and ext_price is not None:
                offers.append(PartOffer(vendor, description, ext_price))
        if not offers:
            raise RuntimeError("CCC vendor result grid contained no readable Vendor/Description/EXT. PRICE rows")
        return offers

    @staticmethod
    def summarize(
        vin: str,
        oem: str,
        offers: list[PartOffer],
        description_target: str = TARGET_DESCRIPTION,
    ) -> dict[str, Any]:
        matching_offers = [offer for offer in offers if normalize(description_target) in normalize(offer.description)]
        ipart = [offer for offer in matching_offers if normalize(offer.vendor) == normalize(IPART_VENDOR)]
        ipart_available = bool(ipart)
        sorted_all = sorted(matching_offers, key=lambda offer: offer.total)
        lowest_overall = sorted_all[0] if sorted_all else None
        second_lowest = sorted_all[1] if len(sorted_all) > 1 else None
        ipart_lowest = min(ipart, key=lambda offer: offer.total) if ipart else None
        return {
            "VIN": vin,
            "OEM": oem,
            "Ipart listed on CCC": ipart_available,
            "Ipart List Price CCC": ipart_lowest.total if ipart_lowest else None,
            "Lowest CCC Price": lowest_overall.total if lowest_overall else None,
            "Lowest CCC Supplier": lowest_overall.vendor if lowest_overall else None,
            "Next Lowest CCC Price": second_lowest.total if second_lowest else None,
            "Next Lowest CCC Supplier": second_lowest.vendor if second_lowest else None,
            "Lowest CCC Price Supplier": lowest_overall.vendor if lowest_overall else None,
            "Status": "OK" if matching_offers else f"No offers matching description: {description_target}",
        }

    def crawl_row(self, record: dict[str, str]) -> dict[str, Any]:
        vin, oem = record["VIN"], record["OEM"]
        self.checkpoint(f"Input row {record['_row']}: VIN={vin}, OEM={oem}")
        result: dict[str, Any] = {"VIN": vin, "OEM": oem}
        try:
            self.open_new_workfile()
            if not self.enter_vehicle(vin):
                result["Status"] = "Vehicle/model not found; skipped"
            else:
                retry_count = self.open_part_results(oem)
                if retry_count is None:
                    retries = getattr(self, "vendor_search_retry_count", 0)
                    result["Status"] = f"Vendor search failed after {retries} Retry clicks; skipped"
                else:
                    result = self.summarize(vin, oem, self.read_offers(), self.description_target)
                    log.info("Completed VIN=%s OEM=%s; iPart listed=%s", vin, oem, result["Ipart listed on CCC"])
        except CatalogPartNotFoundError as error:
            log.warning("Skipping VIN=%s OEM=%s: %s", vin, oem, error)
            result["Status"] = f"{error}; skipped"
        except Exception as error:
            log.exception("Row failed: VIN=%s OEM=%s: %s", vin, oem, error)
            self.capture_failure(f"row_{record['_row']}_{vin}")
            result["Status"] = f"ERROR: {error}"
        finally:
            try:
                self.return_to_workfiles()
            except Exception as error:
                log.error("Could not return to the Workfiles list after VIN=%s: %s", vin, error)
            result["date_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        return result


def main() -> int:
    load_dotenv()
    parser = argparse.ArgumentParser(description="Crawl CCC vendor prices for VIN/OEM rows in an Excel file.")
    parser.add_argument("--input", type=Path, default=Path(os.getenv("CCC_INPUT", "input.xlsx")))
    parser.add_argument("--output", type=Path, default=Path(os.getenv("CCC_OUTPUT", "ccc-results.xlsx")))
    parser.add_argument(
        "--description",
        default=os.getenv("CCC_DESCRIPTION", TARGET_DESCRIPTION),
        help=f"Description substring to match when ranking offers (default: {TARGET_DESCRIPTION}).",
    )
    parser.add_argument("--headed", action="store_true", help="Show the browser window (recommended for first run).")
    args = parser.parse_args()

    username = os.getenv("CCC_USERNAME")
    password = os.getenv("CCC_PASSWORD")
    if not username or not password:
        log.error("Set CCC_USERNAME and CCC_PASSWORD in your environment or .env file.")
        return 2
    if not args.input.is_file():
        log.error("Input workbook not found: %s", args.input)
        return 2

    try:
        records = read_input(args.input)
    except Exception as error:
        log.error("Could not read input workbook: %s", error)
        return 2
    if not records:
        log.error("No rows with both VIN and OEM were found in %s", args.input)
        return 2

    results = load_results(args.output)
    completed, retry_rows = resume_state(results)
    if results:
        log.info("Loaded %d saved result(s) from %s", len(results), args.output)
    saved_remaining = completed.copy()
    retry_remaining = Counter({key: len(indexes) for key, indexes in retry_rows.items()})
    unfinished = 0
    for record in records:
        key = (record["VIN"], record["OEM"])
        if saved_remaining[key] > 0:
            saved_remaining[key] -= 1
        elif retry_remaining[key] > 0:
            retry_remaining[key] -= 1
        else:
            unfinished += 1
    if unfinished == 0:
        log.info("All %d input rows already have saved results", len(records))
        return 0

    AUTH_STATE.parent.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=not args.headed)
        context_options: dict[str, Any] = {}
        if AUTH_STATE.is_file():
            context_options["storage_state"] = str(AUTH_STATE)
        context: BrowserContext = browser.new_context(**context_options)
        page = context.new_page()
        crawler = CCCCrawler(page, username, password, Path("artifacts"), args.description)
        try:
            crawler.sign_in_if_needed()
            context.storage_state(path=str(AUTH_STATE))
            crawler.open_workfiles()
            for index, record in enumerate(records, start=1):
                log.info("Progress: %d/%d", index, len(records))
                key = (record["VIN"], record["OEM"])
                if completed[key] > 0:
                    completed[key] -= 1
                    log.info("Skipping row %d; result is already saved", index)
                    continue
                result = crawler.crawl_row(record)
                if retry_rows[key]:
                    results[retry_rows[key].popleft()] = result
                else:
                    results.append(result)
                save_results(args.output, results)
                log.info("Saved progress to %s", args.output)
        except Exception as error:
            log.exception("Crawler stopped before completing all rows: %s", error)
            crawler.capture_failure("crawler_stopped")
            save_results(args.output, results)
            return 1
        finally:
            context.close()
            browser.close()
    log.info("Crawl complete: %d/%d rows written to %s", len(results), len(records), args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())