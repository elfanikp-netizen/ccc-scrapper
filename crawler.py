"""CCC workfile crawler driven by an input Excel workbook."""

from __future__ import annotations

import argparse
from collections import Counter, defaultdict, deque
import json
import logging
import os
import re
import sys
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from playwright.sync_api import BrowserContext, Error as PlaywrightError, Locator, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright


BASE_URL = "https://www.cccone.com/login/login-idp"
AUTH_STATE = Path("playwright/.auth/ccc-state.json")
OUTPUT_COLUMNS = [
    "Year",
    "Make",
    "Model",
    "VIN",
    "Partslink",
    "OEM",
    "Part Description",
    "Ipart listed on CCC",
    "Ipart List Price CCC",
    "Lowest CCC Price",
    "Lowest CCC Supplier",
    "Next Lowest CCC Price",
    "Next Lowest CCC Supplier",
    "Lowest CCC Price Supplier",
    "Status",
    "date_time",
    "market_data",
    "Description Target",
]
IPART_VENDOR = "iPart Automotive Inc."
TARGET_DESCRIPTION = "A/M: CAPA"
MAX_VENDOR_SEARCH_RETRIES = 3
MAX_ROW_RETRIES = 1
MAX_CATALOG_CANDIDATE_RETRIES = 20
RESULT_SAVE_INTERVAL = 1
RECALL_NOTICE_DETECT_TIMEOUT_MS = 1000
CATALOG_QUICK_WAIT_MS = 1500
CATALOG_LOADING_WAIT_MS = 15000
CATALOG_SCAN_TIMEOUT_SECONDS = 300
CATALOG_LOADING_SELECTOR = (
    "[aria-busy='true']:visible, [role='progressbar']:visible, "
    "mat-progress-spinner:visible, mat-spinner:visible, "
    ".loading-spinner:visible, .spinner:visible"
)
MODEL_MISMATCH_RE = re.compile(
    r"part does not match (?:(?:the )?(?:year[, ]+)?(?:model|engine)(?: of the vehicle)?|the year of the vehicle)",
    re.I,
)
NO_SOURCE_PARTS_RE = re.compile(
    r"no part lines to source",
    re.I,
)
VEHICLE_TITLE_RE = re.compile(
    r"(?:^|\n)\s*New workfile,\s*(\d{4})\s+([A-Za-z][A-Za-z-]*)\s+(.+?)\s*(?:\n|$)",
    re.I,
)
MULTIWORD_MAKES = ("Alfa Romeo", "Aston Martin", "Land Rover", "Rolls Royce")
VIN_MODEL_YEAR_CODES = "ABCDEFGHJKLMNPRSTVWXY123456789"
MARKET_DATA_FIELDS = {
    "description": "description",
    "vendor": "vendor",
    "partno": "part-no",
    "partnumber": "part-no",
    "delivery": "delivery",
    "extprice": "ext-price",
    "labor": "labor",
    "paint": "paint",
    "total": "total",
    "diff": "diff",
    "margin": "margin",
}

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger("ccc-crawler")


@dataclass
class PartOffer:
    vendor: str
    description: str
    total: float


class CatalogPartNotFoundError(RuntimeError):
    pass


class CatalogSearchTimeoutError(RuntimeError):
    pass


class FatalCrawlerError(RuntimeError):
    pass


def block_nonessential_resources(route: Any) -> None:
    if route.request.resource_type in {"image", "media"}:
        route.abort()
    else:
        route.continue_()


def normalize(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip().casefold()


def vin_model_year_matches(vin: str, year: str | int) -> bool:
    rendered_vin = (vin or "").strip().upper()
    try:
        model_year = int(year)
    except (TypeError, ValueError):
        return False
    if len(rendered_vin) < 10 or not 1980 <= model_year <= 2100:
        return False
    return rendered_vin[9] == VIN_MODEL_YEAR_CODES[(model_year - 1980) % len(VIN_MODEL_YEAR_CODES)]


def validate_selected_vehicle_year(vin: str, year: str) -> None:
    if year and not vin_model_year_matches(vin, year):
        raise FatalCrawlerError(
            f"CCC selected model year {year}, which conflicts with VIN {vin}'s model-year code; retry required"
        )


def choose_vin_vehicle_option(vin: str, option_texts: list[str]) -> int:
    for index, option_text in enumerate(option_texts):
        years = re.findall(r"(?<!\d)(?:19|20)\d{2}(?!\d)", option_text)
        if any(vin_model_year_matches(vin, year) for year in years):
            return index
    return 0


def remaining_catalog_categories(category_names: list[str], checked: set[str]) -> deque[str]:
    return deque(name for name in category_names if normalize(name) not in checked)


def scenario_path_parts(scenario: str) -> list[str]:
    return [part.strip() for part in scenario.split(">") if part.strip()]


def description_matches_target(description: str, target: str) -> bool:
    return normalize(description) == normalize(target)


def market_data_matches_target(market_data: Any, target: str) -> bool:
    if isinstance(market_data, str):
        try:
            market_data = json.loads(market_data)
        except json.JSONDecodeError:
            return False
    return isinstance(market_data, list) and any(
        isinstance(row, dict)
        and description_matches_target(str(row.get("description") or ""), target)
        for row in market_data
    )


def should_retry_row_status(status: str) -> bool:
    normalized_status = (status or "").strip().casefold()
    return normalized_status.startswith((
        "error:",
        "no sourceable parts available",
        "vendor search failed",
    ))


def is_no_catalog_results(body_text: str) -> bool:
    text = normalize(body_text)
    return "no results found for search text:" in text


def catalog_not_found_error(
    body_text: str,
    oem: str,
    visible_category_count: int = 0,
) -> CatalogPartNotFoundError | None:
    if not is_no_catalog_results(body_text):
        return None
    result = re.search(
        r"no results found for search text:\s*[\"']?([^\s\"';]+)",
        normalize(body_text),
    )
    if result and normalize(result.group(1).rstrip(".,:)!?")) == normalize(oem):
        return CatalogPartNotFoundError(f"No catalog results found for OEM {oem}")
    if visible_category_count == 0 and result is None:
        return CatalogPartNotFoundError(f"No catalog results found for OEM {oem}")
    return None


def parse_price(value: str) -> float | None:
    cleaned = re.sub(r"[^0-9.\-]", "", value.replace(",", ""))
    try:
        return float(cleaned)
    except ValueError:
        return None


def extract_vehicle_details(page_text: str) -> dict[str, str]:
    match = VEHICLE_TITLE_RE.search(page_text)
    if match is None:
        return {"Year": "", "Make": "", "Model": ""}

    year, make, model = match.groups()
    title = f"{make} {model}"
    for multiword_make in MULTIWORD_MAKES:
        if title.casefold().startswith(f"{multiword_make.casefold()} "):
            make = title[:len(multiword_make)]
            model = title[len(multiword_make):].strip()
            break
    return {"Year": year, "Make": make, "Model": model.strip()}


def read_input(path: Path, search_action: str = "vin_oem") -> list[dict[str, str]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    rows = sheet.iter_rows(values_only=True)
    headers = [normalize(str(value or "")) for value in next(rows, ())]
    if search_action == "vin_partslink_scenario":
        try:
            vin_index = headers.index("vin")
            partslink_index = headers.index("partslink")
            scenario_index = headers.index("scenario")
        except ValueError as error:
            raise ValueError("Scenario input workbook must have VIN, PARTSLINK, and Scenario columns.") from error

        records = []
        for row_number, row in enumerate(rows, start=2):
            vin = str(row[vin_index] or "").strip() if len(row) > vin_index else ""
            partslink = str(row[partslink_index] or "").strip() if len(row) > partslink_index else ""
            scenario = str(row[scenario_index] or "").strip() if len(row) > scenario_index else ""
            if vin and partslink and scenario:
                records.append({
                    "VIN": vin,
                    "PARTSLINK": partslink,
                    "Scenario": scenario,
                    "_row": str(row_number),
                })
            elif vin or partslink or scenario:
                log.warning("Skipping input row %d: VIN, PARTSLINK, or Scenario is blank", row_number)
        workbook.close()
        return records

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
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "CCC Results"
    sheet.append(OUTPUT_COLUMNS)
    for result in results:
        sheet.append([result.get(column) for column in OUTPUT_COLUMNS])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=path.parent,
            prefix=f".{path.stem}.",
            suffix=path.suffix,
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)
        workbook.save(temporary_path)
        with temporary_path.open("rb+") as saved_file:
            os.fsync(saved_file.fileno())
        os.replace(temporary_path, path)
    finally:
        workbook.close()
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


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
                if column in OUTPUT_COLUMNS or column == "OG OEM"
            }
            if result.get("OG OEM"):
                result["Partslink"] = result.get("Partslink") or result.get("OEM")
                result["OEM"] = result["OG OEM"]
            result.pop("OG OEM", None)
            if result.get("VIN") and (result.get("Partslink") or result.get("OEM")):
                results.append(result)
        return results
    finally:
        workbook.close()


def resume_state(
    results: list[dict[str, Any]],
    description_target: str = TARGET_DESCRIPTION,
) -> tuple[Counter, dict[tuple[str, str], deque[int]]]:
    completed: Counter = Counter()
    retry_rows: dict[tuple[str, str], deque[int]] = defaultdict(deque)
    for index, result in enumerate(results):
        key = (
            str(result.get("VIN") or "").strip(),
            str(result.get("Partslink") or result.get("OEM") or "").strip(),
        )
        if not all(key):
            continue
        status = str(result.get("Status") or "").strip()
        saved_market_has_target = market_data_matches_target(
            result.get("market_data"), description_target
        )
        saved_target = normalize(str(result.get("Description Target") or ""))
        target_matches = not saved_target or saved_target == normalize(description_target)
        completed_no_match = (
            status.casefold().startswith("no catalog results found for oem ")
            or (
                status.casefold().startswith("no offers matching description:")
                and target_matches
                and bool(result.get("market_data"))
                and not saved_market_has_target
            )
        )
        if status.casefold() == "ok" or completed_no_match:
            completed[key] += 1
        else:
            retry_rows[key].append(index)
    return completed, retry_rows


def prepare_result_slots(
    records: list[dict[str, str]],
    saved_results: list[dict[str, Any]],
    description_target: str = TARGET_DESCRIPTION,
) -> tuple[list[dict[str, Any] | None], dict[tuple[str, str], list[int]]]:
    """Align saved row occurrences with input order and collect rows needing lookup."""
    saved_by_key: dict[tuple[str, str], deque[tuple[int, dict[str, Any]]]] = defaultdict(deque)
    for index, result in enumerate(saved_results):
        key = (
            str(result.get("VIN") or "").strip(),
            str(result.get("Partslink") or result.get("OEM") or "").strip(),
        )
        if all(key):
            saved_by_key[key].append((index, result))

    slots: list[dict[str, Any] | None] = [None] * len(records)
    pending: dict[tuple[str, str, str], list[int]] = defaultdict(list)
    input_groups: dict[tuple[str, str], list[int]] = defaultdict(list)
    for index, record in enumerate(records):
        row_key = (record["VIN"], record.get("PARTSLINK") or record.get("OEM", ""))
        input_groups[row_key].append(index)
    for row_key, indexes in input_groups.items():
        group_key = (*row_key, normalize(description_target))
        saved_rows = list(saved_by_key[row_key])
        valid_saved = next(
            (result for _, result in saved_rows
             if resume_state([result], description_target)[0][row_key] > 0),
            None,
        )
        if valid_saved is not None:
            for index in indexes:
                slots[index] = valid_saved
        else:
            pending[group_key].extend(indexes)
    return slots, pending


def expand_group_result(
    result: dict[str, Any],
    records: list[dict[str, str]],
    indexes: list[int],
) -> list[dict[str, Any]]:
    expanded = []
    for index in indexes:
        row = dict(result)
        row["VIN"] = records[index]["VIN"]
        if records[index].get("PARTSLINK"):
            row["Partslink"] = records[index]["PARTSLINK"]
        else:
            row["OEM"] = records[index]["OEM"]
        expanded.append(row)
    return expanded


class CCCCrawler:
    def __init__(
        self,
        page: Page,
        username: str,
        password: str,
        artifacts: Path,
        description_target: str = TARGET_DESCRIPTION,
        context: BrowserContext | None = None,
    ):
        self.page = page
        self.context = context or page.context
        self.username = username
        self.password = password
        self.artifacts = artifacts
        self.description_target = description_target
        self.workfiles_url: str | None = None
        self._workfile_initialized = False
        self._vehicle_details: dict[str, str] = {"Year": "", "Make": "", "Model": ""}
        self.no_sourceable_candidates = False
        self.artifacts.mkdir(parents=True, exist_ok=True)

    def checkpoint(self, label: str) -> None:
        log.info("%s", label)

    def wait_for_catalog_element(self, locator: Locator, oem: str, action: str) -> bool:
        try:
            locator.first.wait_for(state="visible", timeout=CATALOG_QUICK_WAIT_MS)
            return True
        except PlaywrightTimeoutError:
            body_text = normalize(self.page.locator("body").inner_text())
            not_found_error = catalog_not_found_error(body_text, oem)
            if not_found_error is not None:
                raise not_found_error

            loading = self.page.locator(CATALOG_LOADING_SELECTOR)
            if not loading.count():
                return bool(locator.count())

            self.checkpoint(f"CCC is still loading while {action}; extending the wait")
            try:
                locator.first.wait_for(state="visible", timeout=CATALOG_LOADING_WAIT_MS)
                return True
            except PlaywrightTimeoutError as error:
                body_text = normalize(self.page.locator("body").inner_text())
                not_found_error = catalog_not_found_error(body_text, oem)
                if not_found_error is not None:
                    raise not_found_error from error
                raise CatalogSearchTimeoutError(
                    f"CCC remained busy while {action} for OEM {oem}; retry required"
                ) from error

    @staticmethod
    def check_catalog_scan_deadline(started_at: float, oem: str) -> None:
        elapsed = time.monotonic() - started_at
        if elapsed >= CATALOG_SCAN_TIMEOUT_SECONDS:
            raise CatalogSearchTimeoutError(
                f"CCC catalog scan exceeded {CATALOG_SCAN_TIMEOUT_SECONDS} seconds "
                f"for OEM {oem}; retry required"
            )

    def selected_catalog_group_control(self) -> Locator | None:
        visible = self.page.locator("button.selected-group-button:visible")
        if visible.count():
            return visible.first
        attached = self.page.locator("button.selected-group-button")
        return attached.first if attached.count() else None

    def apply_oem_catalog_filter(self, oem: str, open_estimate: bool = True) -> None:
        self.dismiss_jumpstart_prompt()
        if open_estimate:
            self.page.locator('a[aria-label="Estimate"]').click()
        self.dismiss_jumpstart_prompt()
        filter_button = self.page.locator("button.filter-button:visible").first
        filter_button.wait_for(state="visible", timeout=20000)
        filter_button.click()
        keyword_box = self.page.locator("input.search-input:visible").last
        keyword_box.wait_for(state="visible", timeout=10000)
        keyword_box.fill(oem)
        self.click_text("ALL MODELS")
        self.dismiss_jumpstart_prompt()
        filter_dialog = self.page.locator(".cdk-overlay-pane").filter(
            has_text=re.compile(r"Filter parts database", re.I)
        ).last
        try:
            filter_dialog.wait_for(state="visible", timeout=5000)
            ok_button = filter_dialog.locator("button").filter(
                has_text=re.compile(r"^\s*OK\s*$", re.I)
            )
            ok_control = self.first_visible(ok_button)
            if ok_control is None:
                ok_control = self.first_visible(
                    filter_dialog.get_by_text(re.compile(r"^\s*OK\s*$", re.I))
                )
            if ok_control is None:
                raise RuntimeError("Visible exact OK button was not found")
            ok_control.click(timeout=10000)
            filter_dialog.wait_for(state="hidden", timeout=10000)
        except (PlaywrightError, RuntimeError) as error:
            cancel = self.first_visible(
                filter_dialog.get_by_text(re.compile(r"^\s*CANCEL\s*$", re.I))
            )
            if cancel is None:
                cancel = self.first_visible(self.page.get_by_text("CANCEL", exact=True))
            if cancel is not None:
                cancel.click(timeout=5000)
            raise RuntimeError(f"CCC OEM filter could not confirm {oem} with its OK button: {error}") from error

    def load_catalog_categories(self, oem: str) -> tuple[Locator, list[str]]:
        categories = self.page.locator("button.group-list-item:visible")
        if not self.wait_for_catalog_element(categories, oem, "loading catalog categories"):
            selected_group = self.selected_catalog_group_control()
            if selected_group is not None:
                selected_group.click(force=True)
                self.wait_for_catalog_element(categories, oem, "opening catalog categories")

        category_names = categories.all_inner_texts()
        if category_names:
            return categories, category_names

        body_text = normalize(self.page.locator("body").inner_text())
        not_found_error = catalog_not_found_error(body_text, oem)
        if not_found_error is not None:
            raise not_found_error
        raise CatalogSearchTimeoutError(
            f"CCC catalog category controls did not load for OEM {oem}; retry required"
        )

    def select_catalog_category(self, category_name: str, oem: str) -> bool:
        categories = self.page.locator("button.group-list-item:visible")
        matching_category = None
        for index in range(categories.count()):
            candidate = categories.nth(index)
            if normalize(candidate.inner_text()) == normalize(category_name):
                matching_category = candidate
                break

        if matching_category is None:
            selected_group = self.selected_catalog_group_control()
            if selected_group is None:
                self.checkpoint(f"CCC category control is missing; resetting OEM search for {oem}")
                self.apply_oem_catalog_filter(oem, open_estimate=False)
            else:
                selected_group.click(force=True)
            if not self.wait_for_catalog_element(
                categories, oem, "reopening catalog categories"
            ):
                raise CatalogSearchTimeoutError(
                    f"CCC category menu did not reopen for OEM {oem}; retry required"
                )
            for index in range(categories.count()):
                candidate = categories.nth(index)
                if normalize(candidate.inner_text()) == normalize(category_name):
                    matching_category = candidate
                    break

        if matching_category is None:
            self.checkpoint(
                f"CCC category list changed after reopening for OEM {oem}; refreshing remaining categories"
            )
            return False
        matching_category.click(force=True)
        return True

    def find_catalog_text_control(self, selectors: tuple[str, ...], text: str) -> Locator | None:
        wanted = normalize(text)
        for selector in selectors:
            controls = self.page.locator(selector)
            for index in range(controls.count()):
                candidate = controls.nth(index)
                try:
                    if candidate.is_visible() and wanted in normalize(candidate.inner_text()):
                        return candidate
                except PlaywrightError:
                    continue
        return None

    def read_selected_part_properties(self, selected_part_text: str) -> tuple[str, str]:
        edit_properties = self.first_visible(
            self.page.get_by_role("button", name=re.compile(r"edit properties", re.I))
        )
        if edit_properties is None:
            edit_properties = self.first_visible(
                self.page.get_by_text(re.compile(r"^\s*EDIT PROPERTIES\s*$", re.I))
            )
        if edit_properties is None:
            estimate_part = self.first_visible(
                self.page.get_by_text(re.compile(re.escape(selected_part_text), re.I)).last
            )
            if estimate_part is not None:
                estimate_part.click(timeout=10000)
                edit_properties = self.first_visible(
                    self.page.get_by_role("button", name=re.compile(r"edit properties", re.I))
                )
        if edit_properties is None:
            raise RuntimeError("Selected estimate part has no visible Edit Properties control")
        edit_properties.click(timeout=10000)
        part_number_label = self.page.get_by_text(re.compile(r"^\s*Part Number\s*$", re.I)).last
        description_label = self.page.get_by_text(re.compile(r"^\s*Description\s*$", re.I)).last
        part_number_label.wait_for(state="visible", timeout=15000)
        description_label.wait_for(state="visible", timeout=15000)

        properties = self.page.evaluate(
            """() => {
                    const visible = element => {
                        const rect = element.getBoundingClientRect();
                        const style = getComputedStyle(element);
                        return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
                    };
                    const overlays = [...document.querySelectorAll('.cdk-overlay-pane')].filter(visible);
                    const root = overlays[overlays.length - 1] || document.body;
                    const controls = 'input, textarea, [contenteditable="true"]';
                    const controlValue = element => (element.value ?? element.innerText ?? element.textContent ?? '').trim();

                    const readLabel = label => {
                        const wanted = label.toLowerCase();
                        const exactLabels = [...root.querySelectorAll('*')].filter(element =>
                            visible(element)
                            && (element.innerText || element.textContent || '').trim().toLowerCase() === wanted
                        );
                        for (const labelElement of exactLabels) {
                            let ancestor = labelElement;
                            for (let depth = 0; ancestor && depth < 8; depth++, ancestor = ancestor.parentElement) {
                                const nearbyControls = [...ancestor.querySelectorAll(controls)].filter(visible);
                                if (nearbyControls.length === 1) return controlValue(nearbyControls[0]);
                            }
                        }

                        let best = null;
                        const otherLabel = wanted === 'description' ? /part\\s*number/i : /description/i;
                        for (const control of [...root.querySelectorAll(controls)].filter(visible)) {
                            let ancestor = control;
                            for (let depth = 0; ancestor && depth < 8; depth++, ancestor = ancestor.parentElement) {
                                const text = (ancestor.innerText || '').toLowerCase();
                                if (text.includes(wanted) && !otherLabel.test(text)) {
                                    const candidate = { depth, value: controlValue(control) };
                                    if (!best || candidate.depth < best.depth) best = candidate;
                                    break;
                                }
                            }
                        }
                        return best?.value || '';
                    };

                    const values = {
                        partNumber: readLabel('Part Number'),
                        description: readLabel('Description'),
                    };
                    if (!values.partNumber || !values.description) {
                        values.visibleControls = [...root.querySelectorAll(controls)].filter(visible).map(control => ({
                            value: controlValue(control),
                            labels: [...(control.labels || [])].map(label => (label.innerText || '').trim()),
                            ancestors: [control.parentElement, control.parentElement?.parentElement,
                                control.parentElement?.parentElement?.parentElement]
                                .map(element => (element?.innerText || '').trim()),
                        }));
                    }
                    return values;
                }"""
        )
        part_number = str(properties.get("partNumber") or "").strip()
        description = str(properties.get("description") or "").strip()
        if not part_number:
            self.capture_failure("scenario_part_number_not_found")
            log.error("Visible Edit Properties controls: %s", properties.get("visibleControls", []))
            raise RuntimeError("Edit Properties opened, but its Part Number value was empty or inaccessible")
        dialog = self.page.locator(".cdk-overlay-pane").filter(has_text=re.compile(r"part number", re.I)).last
        close_control = self.first_visible(
            dialog.get_by_role("button", name=re.compile(r"^\s*(?:DONE|CANCEL|CLOSE)\s*$", re.I))
        ) if dialog.count() else None
        if close_control is not None:
            close_control.click(timeout=5000)
        else:
            self.page.keyboard.press("Escape")
        return part_number, description

    def open_scenario_part_results(self, scenario: str, partslink: str) -> tuple[int | None, str, str]:
        path_parts = scenario_path_parts(scenario)
        if not path_parts:
            raise CatalogPartNotFoundError(f"No catalog path is configured for PARTSLINK {partslink}")
        self.checkpoint(f"Estimate: following CCC Scenario for PARTSLINK {partslink}")
        self.dismiss_jumpstart_prompt()
        self.page.locator('a[aria-label="Estimate"]').click()
        categories, _ = self.load_catalog_categories(partslink)
        category_selectors = ("button.group-list-item:visible", "button.subgroup-button:visible")

        for segment in path_parts[:-1]:
            control = self.find_catalog_text_control(category_selectors, segment)
            if control is None:
                raise CatalogPartNotFoundError(
                    f"CCC catalog path segment '{segment}' was not found for PARTSLINK {partslink}"
                )
            self.checkpoint(f"CCC Scenario catalog: selecting {segment}")
            control.click(force=True)
            self.wait_for_catalog_element(
                categories, partslink, f"opening CCC Scenario catalog segment {segment}"
            )

        final_segment = path_parts[-1]
        part_rows = self.page.locator("app-part-selector-part-row:visible")
        candidate = self.find_catalog_text_control(("app-part-selector-part-row:visible",), final_segment)
        if candidate is None:
            final_category = self.find_catalog_text_control(category_selectors, final_segment)
            if final_category is not None:
                self.checkpoint(f"CCC Scenario catalog: selecting final category {final_segment}")
                final_category.click(force=True)
                if not self.wait_for_catalog_element(
                    part_rows, partslink, f"loading CCC Scenario part {final_segment}"
                ):
                    raise CatalogPartNotFoundError(
                        f"CCC catalog has no part containing '{final_segment}' for PARTSLINK {partslink}"
                    )
                candidate = self.find_catalog_text_control(("app-part-selector-part-row:visible",), final_segment)
        if candidate is None:
            raise CatalogPartNotFoundError(
                f"CCC catalog has no part containing '{final_segment}' for PARTSLINK {partslink}"
            )

        self.checkpoint(f"Double-clicking CCC Scenario part {final_segment} for PARTSLINK {partslink}")
        candidate.dblclick(timeout=15000)
        self.resolve_part_prompts()
        part_number, part_description = self.read_selected_part_properties(final_segment)
        price_prompt_acknowledged = self.dismiss_enter_part_price_prompt()
        source_tab = self.page.locator("[role=tab].source-parts-tab")
        source_tag = self.page.get_by_role("button", name="Show sourced parts").first
        if price_prompt_acknowledged:
            source_tab.wait_for(state="visible", timeout=10000)
            source_tab.click(timeout=10000)
            source_tag.wait_for(state="visible", timeout=10000)
            source_tag.click()
            return self.wait_for_vendor_inventory(), part_number, part_description
        source_tab.wait_for(state="visible", timeout=30000)
        source_tab.click()
        source_tag.wait_for(state="visible", timeout=10000)
        source_tag.click()
        return self.wait_for_vendor_inventory(), part_number, part_description

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

        last_error: Exception | None = None
        for attempt in range(1, 4):
            try:
                if attempt > 1:
                    body_text = ""
                    try:
                        body_text = normalize(self.page.locator("body").inner_text())
                    except PlaywrightError:
                        pass
                    if not body_text:
                        log.warning("CCC workfiles page is blank; rebuilding the browser session (attempt %d/3)", attempt)
                        self.sign_in_if_needed()
                        self.workfiles_url = None
                        self.open_workfiles()

                if self.page.url != self.workfiles_url:
                    self.page.goto(self.workfiles_url, wait_until="domcontentloaded")
                else:
                    current_url = self.page.url
                    self.page.wait_for_load_state("domcontentloaded", timeout=15000)
                    if self.page.url != current_url:
                        self.page.goto(self.workfiles_url, wait_until="domcontentloaded")

                self.dismiss_onboarding_guide()
                control = self.page.get_by_text("New workfile", exact=True)
                try:
                    control.first.wait_for(state="visible", timeout=30000)
                except PlaywrightTimeoutError:
                    pass
                visible_control = self.first_visible(control)
                if visible_control is None:
                    log.warning(
                        "New workfile control not visible (attempt %d/3); reloading workfiles list",
                        attempt,
                    )
                    self.page.goto(self.workfiles_url, wait_until="domcontentloaded", timeout=30000)
                    try:
                        self.page.wait_for_function(
                            "() => document.body && document.body.innerText.trim().length > 100",
                            timeout=30000,
                        )
                    except PlaywrightTimeoutError:
                        pass
                    self.page.wait_for_load_state("domcontentloaded", timeout=15000)
                    self.dismiss_onboarding_guide()
                    control = self.page.get_by_text("New workfile", exact=True)
                    try:
                        control.first.wait_for(state="visible", timeout=30000)
                    except PlaywrightTimeoutError:
                        pass
                    visible_control = self.first_visible(control)
                if visible_control is None:
                    raise PlaywrightTimeoutError(
                        'waiting for get_by_text("New workfile", exact=True)'
                    )

                self.checkpoint("Creating a new workfile for this input row")
                visible_control.click(timeout=15000, no_wait_after=True)
                try:
                    self.page.wait_for_url("**/edit/contacts", wait_until="commit", timeout=45000)
                    self.page.locator('a[aria-label="Vehicle"]').wait_for(state="visible", timeout=45000)
                except PlaywrightTimeoutError as error:
                    self.capture_failure("new_workfile_surface_not_loaded")
                    raise RuntimeError(
                        f"CCC did not finish opening the workfile Vehicle panel; current URL: {self.page.url}"
                    ) from error
                self.dismiss_jumpstart_prompt()
                self._workfile_initialized = True
                log.info("New workfile opened: %s", self.page.url)
                return
            except (PlaywrightError, RuntimeError) as error:
                last_error = error
                log.warning(
                    "New workfile recovery attempt %d/3 failed: %s", attempt, str(error)
                )
                self.capture_failure(f"new_workfile_recovery_{attempt}")
                try:
                    self.page.goto(self.workfiles_url, wait_until="domcontentloaded", timeout=30000)
                    self.page.wait_for_load_state("domcontentloaded", timeout=15000)
                except PlaywrightError:
                    pass

        assert last_error is not None
        raise PlaywrightTimeoutError(
            'New workfile control remained unavailable after 3 workfiles-list reloads'
        ) from last_error

    def dismiss_jumpstart_prompt(self) -> None:
        prompt = self.page.locator(".cdk-overlay-pane").filter(
            has_text=re.compile(r"What would you like to do first\?", re.I)
        ).last
        if not prompt.count():
            return
        jumpstart = prompt.get_by_text("Jumpstart estimate", exact=True)
        control = self.first_visible(jumpstart)
        if control is None:
            raise RuntimeError("CCC first-estimate prompt appeared without a Jumpstart estimate control")
        self.checkpoint("CCC first-estimate prompt shown; choosing Jumpstart estimate")
        control.click(timeout=10000)
        prompt.wait_for(state="hidden", timeout=15000)

    def dismiss_onboarding_guide(self) -> None:
        guide = self.page.locator("#pendo-base")
        if not guide.count():
            return
        maybe_later = guide.get_by_role("button", name=re.compile(r"^MAYBE LATER$", re.I))
        control = self.first_visible(maybe_later)
        if control is not None:
            control.click(timeout=2000)

    def return_to_workfiles(self) -> None:
        if self.workfiles_url:
            self.page.goto(self.workfiles_url, wait_until="domcontentloaded")

    def enter_vehicle(self, vin: str) -> bool:
        self._vehicle_details = {"Year": "", "Make": "", "Model": ""}
        self.checkpoint(f"Vehicle: searching VIN {vin}")
        self.dismiss_jumpstart_prompt()
        self.page.locator('a[aria-label="Vehicle"]').click()
        self.open_vehicle_selector()
        vin_box = self.page.locator(".vin-input input.input-box:visible").last
        vin_box.wait_for(state="visible", timeout=10000)
        vin_box.fill(vin)
        try:
            self.page.wait_for_function(
                """() => {
                    const visible = element => {
                        const rect = element.getBoundingClientRect();
                        const style = getComputedStyle(element);
                        return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
                    };
                    const text = document.body?.innerText || '';
                    if (/Unable to decode\\. Please verify VIN or proceed\\./i.test(text)) return 'decode-warning';
                    if ([...document.querySelectorAll('*')].some(element =>
                        visible(element) && element.children.length === 0 && (element.textContent || '').trim() === 'Valid VIN'
                    )) return 'valid';
                    return false;
                }""",
                timeout=30000,
            )
        except PlaywrightTimeoutError as error:
            raise RuntimeError(f"CCC did not validate or present a proceed option for VIN {vin}") from error
        decode_warning = self.page.get_by_text(
            re.compile(r"Unable to decode\. Please verify VIN or proceed\.", re.I)
        ).first
        if self.first_visible(decode_warning) is not None:
            log.warning("CCC could not decode VIN %s but offers a proceed option; continuing", vin)
        else:
            self.checkpoint("CCC validated the VIN; advancing to model selection")
        self.click_text("Next")
        selection_state = self.resolve_vehicle_change_state(self.page.url)
        if selection_state == "unknown-vehicle":
            self.checkpoint(f"CCC could not identify VIN {vin}; manual vehicle details are required")
            return False
        while selection_state == "options":
            options = self.page.locator('[role="option"]:visible')
            option_texts = [options.nth(index).inner_text().strip() for index in range(options.count())]
            selected_option = choose_vin_vehicle_option(vin, option_texts)
            self.checkpoint(
                f"Selecting CCC vehicle option {selected_option + 1}/{len(option_texts)} "
                f"for VIN {vin}: {option_texts[selected_option]}"
            )
            options.nth(selected_option).click()
            self.checkpoint("CCC vehicle model selected; advancing to Finish")
            self.click_text("Next")
            selection_state = self.resolve_vehicle_change_state(self.page.url)
            if selection_state == "unknown-vehicle":
                self.checkpoint(f"CCC could not identify VIN {vin}; manual vehicle details are required")
                return False
        if selection_state != "done":
            raise FatalCrawlerError(f"Unexpected CCC vehicle selection state: {selection_state}")
        self.checkpoint("CCC vehicle selection complete; clicking DONE")
        self.finish_vehicle_selection()
        self._vehicle_details = extract_vehicle_details(self.page.locator("body").inner_text())
        selected_year = self._vehicle_details["Year"]
        validate_selected_vehicle_year(vin, selected_year)
        if self._vehicle_details["Year"]:
            log.info(
                "Selected vehicle: Year=%s Make=%s Model=%s",
                self._vehicle_details["Year"],
                self._vehicle_details["Make"],
                self._vehicle_details["Model"],
            )
        log.info("Vehicle selected for VIN %s", vin)
        return True

    def resolve_vehicle_change_state(self, workfile_url: str) -> str:
        try:
            state = self.wait_for_vehicle_selection_state()
        except PlaywrightError as error:
            if self.page.is_closed() or "closed" in str(error).lower() or "target" in str(error).lower():
                self.reopen_workfile_page(workfile_url)
                state = self.wait_for_vehicle_selection_state()
            else:
                raise
        while state == "change-confirmation":
            self.confirm_change_vehicle()
            try:
                state = self.wait_for_vehicle_selection_state()
            except PlaywrightError as error:
                if self.page.is_closed() or "closed" in str(error).lower() or "target" in str(error).lower():
                    self.reopen_workfile_page(workfile_url)
                    state = self.wait_for_vehicle_selection_state()
                else:
                    raise
        return state

    def wait_for_vehicle_selection_state(self) -> str:
        return self.page.wait_for_function(
            """() => {
                const visible = element => {
                    const rect = element.getBoundingClientRect();
                    const style = getComputedStyle(element);
                    return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
                };
                const changeConfirmation = [...document.querySelectorAll('.cdk-overlay-pane')].some(pane =>
                    visible(pane) && /Any estimate lines written will be erased[\\s\\S]*Change vehicle anyway\\?/i.test(pane.innerText || '')
                );
                if (changeConfirmation) return 'change-confirmation';
                const text = document.body?.innerText || '';
                if (/Unknown vehicle/i.test(text) && /Year\\*/i.test(text) && /Make\\*/i.test(text)) return 'unknown-vehicle';
                const options = [...document.querySelectorAll('[role="option"]')].some(option =>
                    visible(option) && (option.innerText || option.textContent || '').trim().length > 0
                );
                if (options) return 'options';
                const finish = [...document.querySelectorAll('button')].some(button =>
                    visible(button) && /^(?:DONE|FINISH)$/i.test((button.innerText || button.textContent || '').trim()) && !button.disabled
                );
                return finish ? 'done' : false;
            }""",
            timeout=60000,
        ).json_value()

    def confirm_change_vehicle(self) -> None:
        try:
            self.page.get_by_text(
                re.compile(r"Any estimate lines written will be erased.*Change vehicle anyway\?", re.I | re.S)
            ).wait_for(state="visible", timeout=15000)
            prompt = self.page.locator(".cdk-overlay-pane").filter(
                has_text=re.compile(r"Any estimate lines written will be erased.*Change vehicle anyway\?", re.I | re.S)
            ).last
            prompt.wait_for(state="visible", timeout=10000)
            buttons = prompt.locator("button")
            change_button = buttons.filter(has_text=re.compile(r"^\s*CHANGE VEHICLE\s*$", re.I)).last
            change_button.wait_for(state="visible", timeout=10000)
            self.checkpoint("CCC warned that changing vehicle erases estimate lines; clicking CHANGE VEHICLE as requested")
            change_button.click(timeout=15000)
            self.checkpoint("CCC CHANGE VEHICLE confirmation clicked")
        except PlaywrightError as error:
            raise FatalCrawlerError(f"CCC Change Vehicle confirmation could not complete: {error}") from error

    def reopen_workfile_page(self, workfile_url: str) -> None:
        self.checkpoint("CCC closed the workfile tab during vehicle change; reopening the same workfile")
        pages = [page for page in self.context.pages if not page.is_closed()]
        self.page = pages[-1] if pages else self.context.new_page()
        if self.page.url != workfile_url:
            self.page.goto(workfile_url, wait_until="domcontentloaded", timeout=60000)
        try:
            self.page.locator('a[aria-label="Vehicle"]').wait_for(state="visible", timeout=45000)
        except PlaywrightTimeoutError as error:
            raise RuntimeError("CCC did not restore the existing workfile after vehicle change") from error

    def open_vehicle_selector(self) -> None:
        self.page.wait_for_function(
            """() => [...document.querySelectorAll('button')].some(button => {
                const rect = button.getBoundingClientRect();
                const style = getComputedStyle(button);
                const visible = rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
                const label = (button.innerText || button.textContent || button.getAttribute('aria-label') || '').trim().toLowerCase();
                return visible && (label === 'select vehicle' || label === 'change vehicle');
            })""",
            timeout=30000,
        )
        buttons = self.page.locator("button")
        for index in range(buttons.count()):
            button = buttons.nth(index)
            label = normalize(button.inner_text() or button.get_attribute("aria-label") or "")
            if button.is_visible() and label in {"select vehicle", "change vehicle"}:
                self.checkpoint(f"Opening CCC vehicle selector with {label.title()}")
                button.click(timeout=15000)
                return
        raise RuntimeError("CCC Vehicle panel showed neither Select vehicle nor Change Vehicle")

    def finish_vehicle_selection(self) -> None:
        self.page.wait_for_function(
            """() => [...document.querySelectorAll('button')].some(button => {
                const rect = button.getBoundingClientRect();
                const style = getComputedStyle(button);
                const visible = rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
                const label = (button.innerText || button.textContent || '').trim();
                return visible && /^(?:DONE|FINISH)$/i.test(label) && !button.disabled;
            })""",
            timeout=60000,
        )
        button_index = self.page.locator("button").evaluate_all(
            """buttons => buttons.findIndex(button => {
                const rect = button.getBoundingClientRect();
                const style = getComputedStyle(button);
                const visible = rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
                const label = (button.innerText || button.textContent || '').trim();
                return visible && /^(?:DONE|FINISH)$/i.test(label) && !button.disabled;
            })"""
        )
        if button_index < 0:
            raise RuntimeError("CCC's visible DONE/FINISH button disappeared before it could be clicked")
        self.page.locator("button").nth(button_index).click(timeout=15000)
        self.page.wait_for_function(
            """() => ![...document.querySelectorAll('button')].some(button => {
                const rect = button.getBoundingClientRect();
                const style = getComputedStyle(button);
                const visible = rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
                const label = (button.innerText || button.textContent || '').trim();
                return visible && /^(?:DONE|FINISH)$/i.test(label);
            })""",
            timeout=30000,
        )
        self.handle_vehicle_recall_prompts()

    def handle_vehicle_recall_prompts(self) -> None:
        recall_notice = self.page.get_by_text(re.compile("vehicle recall notice", re.I)).first
        try:
            recall_notice.wait_for(state="visible", timeout=RECALL_NOTICE_DETECT_TIMEOUT_MS)
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

    def open_part_results(
        self,
        oem: str,
        attempted_catalog_parts: set[tuple[str, int]] | None = None,
    ) -> int | None:
        if attempted_catalog_parts is None:
            attempted_catalog_parts = set()
            self.no_sourceable_candidates = False
        self.checkpoint(f"Estimate: searching OEM {oem}")
        self.apply_oem_catalog_filter(oem)
        source_tab = self.page.locator("[role=tab].source-parts-tab")
        source_tag = self.page.get_by_role("button", name="Show sourced parts").first
        if (
            source_tab.count()
            and source_tab.get_attribute("aria-disabled") == "false"
            and source_tag.is_visible()
        ):
            self.checkpoint(f"Source parts are ready after OEM search for {oem}; opening them directly")
            source_tab.click()
            source_tag.click()
            return self.wait_for_vendor_inventory()
        categories, category_names = self.load_catalog_categories(oem)

        parts = self.page.locator("app-part-selector-part-row:visible")
        body_text = normalize(self.page.locator("body").inner_text())
        not_found_error = catalog_not_found_error(body_text, oem, categories.count())
        if not_found_error is not None:
            raise not_found_error
        selected_parts = False
        selected_category_name = ""
        scan_started_at = time.monotonic()
        categories_checked = 0
        subgroups_expanded = 0
        pending_categories = deque(category_names)
        checked_categories: set[str] = set()
        while pending_categories:
            self.check_catalog_scan_deadline(scan_started_at, oem)
            category_name = pending_categories.popleft()
            category_key = normalize(category_name)
            if category_key in checked_categories:
                continue
            self.checkpoint(f"Checking catalog category {category_name.strip()}")
            category_selected = self.select_catalog_category(category_name, oem)
            checked_categories.add(category_key)
            if not category_selected:
                refreshed_names = categories.all_inner_texts()
                pending_categories = remaining_catalog_categories(
                    refreshed_names, checked_categories
                )
                continue
            categories_checked += 1
            selected_parts = self.wait_for_catalog_element(
                parts, oem, f"checking catalog category {category_name.strip()}"
            )
            while True:
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
                self.check_catalog_scan_deadline(scan_started_at, oem)
                self.checkpoint("Expanding a catalog subgroup")
                collapsed_subgroups[0].click()
                subgroups_expanded += 1
                if self.wait_for_catalog_element(
                    parts, oem, f"expanding a catalog subgroup for {category_name.strip()}"
                ):
                    selected_parts = True
                    break
            if selected_parts:
                if any(
                    (normalize(category_name), index) not in attempted_catalog_parts
                    for index in range(parts.count())
                ):
                    selected_category_name = category_name
                    break
                selected_parts = False
            selected_group = self.page.locator("button.selected-group-button:visible")
            if selected_group.count():
                selected_group.click(force=True)
                try:
                    state = self.page.wait_for_function(
                        """() => {
                            const visible = element => !!element && !!(element.offsetWidth || element.offsetHeight);
                            if ([...document.querySelectorAll('button.group-list-item')].some(visible)) return 'categories';
                            if (/No results found for search text:/i.test(document.body?.innerText || '')) return 'empty';
                            return false;
                        }""",
                        timeout=10000,
                    ).json_value()
                except PlaywrightTimeoutError:
                    state = None
        self.checkpoint(
            f"Catalog scan for OEM {oem}: checked {categories_checked} categories, "
            f"expanded {subgroups_expanded} subgroups in "
            f"{time.monotonic() - scan_started_at:.1f}s"
        )
        if not selected_parts or not parts.count():
            if attempted_catalog_parts:
                self.no_sourceable_candidates = True
                return None
            body_text = normalize(self.page.locator("body").inner_text())
            not_found_error = catalog_not_found_error(body_text, oem)
            if not_found_error is not None:
                raise not_found_error
            raise CatalogSearchTimeoutError(
                f"CCC showed no part rows for OEM {oem} after checking all categories, "
                "but did not confirm that no results were found; retry required"
            )

        candidate_limit = parts.count()
        for part_index in range(candidate_limit):
            candidate_key = (normalize(selected_category_name), part_index)
            if candidate_key in attempted_catalog_parts:
                continue
            candidate = parts.nth(part_index)
            if not candidate.count():
                break
            self.checkpoint(f"Double-clicking catalog result {part_index + 1} for OEM {oem}")
            try:
                candidate.dblclick(timeout=15000)
            except PlaywrightError:
                if part_index + 1 >= candidate_limit:
                    raise
                continue
            self.resolve_part_prompts()
            price_prompt_acknowledged = self.dismiss_enter_part_price_prompt()
            if price_prompt_acknowledged:
                source_tab = self.page.locator("[role=tab].source-parts-tab")
                source_tab.wait_for(state="visible", timeout=10000)
                self.checkpoint(f"Opening Source parts after accepting the price prompt for OEM {oem}")
                source_tab.click(timeout=10000)
                source_tag = self.page.get_by_role("button", name="Show sourced parts").first
                source_tag.wait_for(state="visible", timeout=10000)
                source_tag.click()
                return self.wait_for_vendor_inventory()
            body_text = normalize(self.page.locator("body").inner_text())
            no_source_parts = bool(NO_SOURCE_PARTS_RE.search(body_text))
            source_tab = self.page.locator("[role=tab].source-parts-tab")
            if not source_tab.count() and no_source_parts:
                log.info("Catalog part %s for OEM %s is non-sourceable; trying the next catalog match", part_index + 1, oem)
                self.no_sourceable_candidates = True
                attempted_catalog_parts.add(candidate_key)
                return None
            try:
                source_tab.wait_for(state="visible", timeout=10000)
            except PlaywrightTimeoutError:
                if no_source_parts:
                    log.info("Catalog part %s for OEM %s is non-sourceable; trying the next catalog match", part_index + 1, oem)
                    self.no_sourceable_candidates = True
                    attempted_catalog_parts.add(candidate_key)
                    return None
                if part_index + 1 < candidate_limit:
                    attempted_catalog_parts.add(candidate_key)
                    return None
                raise RuntimeError(f"Part selection did not complete for OEM {oem}; Source parts tab never appeared") from None
            try:
                enabled = self.page.wait_for_function(
                    """() => {
                        const tab = document.querySelector('[role="tab"].source-parts-tab');
                        if (!tab) return 'missing';
                        return tab.getAttribute('aria-disabled') === 'false' ? 'enabled' : 'disabled';
                    }""",
                    timeout=30000,
                ).json_value()
            except PlaywrightTimeoutError:
                enabled = "disabled-timeout"
            if enabled != "enabled":
                body_text = normalize(self.page.locator("body").inner_text())
                if no_source_parts or NO_SOURCE_PARTS_RE.search(body_text):
                    log.info("Catalog part %s for OEM %s has no sourceable estimate line; trying the next catalog match", part_index + 1, oem)
                    self.no_sourceable_candidates = True
                    attempted_catalog_parts.add(candidate_key)
                    return None
                if part_index + 1 < candidate_limit:
                    attempted_catalog_parts.add(candidate_key)
                    return None
                log.warning("CCC Source parts tab is %s for OEM %s; retrying the estimate line", enabled, oem)
                return self.retry_source_parts(oem)
            source_tab.click()
            source_tag = self.page.get_by_role("button", name="Show sourced parts").first
            try:
                source_tag.wait_for(state="visible", timeout=10000)
            except PlaywrightTimeoutError as error:
                body_text = normalize(self.page.locator("body").inner_text())
                if NO_SOURCE_PARTS_RE.search(body_text):
                    log.info("Catalog part %s for OEM %s has no sourceable estimate line; trying the next catalog match", part_index + 1, oem)
                    self.no_sourceable_candidates = True
                    attempted_catalog_parts.add(candidate_key)
                    return None
                if part_index + 1 < candidate_limit:
                    attempted_catalog_parts.add(candidate_key)
                    return None
                raise RuntimeError(f"CCC added no sourceable estimate line for OEM {oem}") from error
            source_tag.click()
            result = self.wait_for_vendor_inventory()
            if result is not None:
                return result
            if part_index + 1 < candidate_limit:
                attempted_catalog_parts.add(candidate_key)
                return None
            return None

        if attempted_catalog_parts:
            self.no_sourceable_candidates = True
        return None

    def dismiss_enter_part_price_prompt(self) -> bool:
        prompt = self.page.get_by_text(
            re.compile(r"^\s*['\"“”‘’]?\s*Enter part price\.?\s*['\"“”‘’]?\s*$", re.I)
        ).first
        try:
            prompt.wait_for(state="visible", timeout=1500)
        except PlaywrightTimeoutError:
            return False
        dialog = self.page.locator(".cdk-overlay-pane").filter(has_text=re.compile(r"Enter part price", re.I)).last
        ok_pattern = re.compile(r"^\s*OK\s*$", re.I)
        prompt_container = prompt.locator("xpath=ancestor::*[.//button[normalize-space(.)='OK']][1]")
        button = None
        for candidate in (
            dialog.get_by_role("button", name=ok_pattern),
            dialog.get_by_text(ok_pattern),
            prompt_container.get_by_role("button", name=ok_pattern),
            prompt_container.get_by_text(ok_pattern),
            self.page.get_by_role("button", name=ok_pattern),
            self.page.get_by_text(ok_pattern),
        ):
            button = self.first_visible(candidate)
            if button is not None:
                break
        if button is None:
            self.capture_failure("enter_part_price_ok_not_found")
            raise RuntimeError("CCC's Enter part price prompt appeared without a visible OK control")
        self.checkpoint("CCC Enter part price prompt shown; acknowledging with OK")
        button.click(timeout=10000)
        prompt.wait_for(state="hidden", timeout=10000)
        return True

    def retry_source_parts(self, oem: str) -> int | None:
        retry_button = self.first_visible(
            self.page.get_by_role("button", name=re.compile(r"^\s*RETRY\s*$", re.I))
        )
        if retry_button is None:
            raise RuntimeError(f"CCC kept Source parts disabled for OEM {oem}; retry button not found")
        self.checkpoint(f"Retrying CCC source-parts preparation for OEM {oem}")
        retry_button.click(timeout=10000)
        try:
            outcome = self.page.wait_for_function(
                """() => {
                    const visible = element => !!element && !!(element.offsetWidth || element.offsetHeight);
                    if ([...document.querySelectorAll('.cdk-overlay-pane .table-row')].some(visible)) return 'offers';
                    if ([...document.querySelectorAll('button')].some(button => visible(button) && /^\\s*RETRY\\s*$/i.test(button.innerText || ''))) return 'retry';
                    const body = document.body?.innerText || '';
                    if (/No part lines to source/i.test(body)) return 'no-source';
                    const tab = document.querySelector('[role="tab"].source-parts-tab');
                    if (tab && tab.getAttribute('aria-disabled') === 'false') return 'enabled';
                    return false;
                }""",
                timeout=60000,
            ).json_value()
        except PlaywrightTimeoutError:
            outcome = "timeout"
        if outcome == "offers":
            return 0
        if outcome == "no-source":
            return None
        if outcome == "enabled":
            self.page.locator('[role="tab"].source-parts-tab').click(timeout=10000)
            self.page.get_by_role("button", name="Show sourced parts").first.click(timeout=10000)
            return self.wait_for_vendor_inventory()
        if outcome == "retry":
            return None
        raise RuntimeError(f"CCC Source parts preparation did not recover for OEM {oem} ({outcome})")

    def wait_for_vendor_inventory(self) -> int | None:
        retry_count = 0
        self.vendor_search_retry_count = retry_count
        retry_button = self.page.get_by_role("button", name=re.compile(r"retry|try again", re.I))

        while True:
            self.checkpoint("Waiting for CCC vendor inventory search")
            try:
                outcome = self.page.wait_for_function(
                    """() => {
                        const visible = element => !!element && !!(element.offsetWidth || element.offsetHeight);
                        if ([...document.querySelectorAll('.cdk-overlay-pane .table-row')].some(visible)) return 'offers';
                        if ([...document.querySelectorAll('button')].some(button => visible(button) && /retry|try again/i.test(button.innerText))) return 'retry';
                        return false;
                    }""",
                    timeout=60000,
                ).json_value()
            except PlaywrightTimeoutError:
                log.error("CCC vendor search produced no offers or Retry control within 60 seconds")
                return None

            if outcome == "offers":
                self.page.evaluate("() => { window.__cccVendorInventoryStability = null; }")
                try:
                    self.page.wait_for_function(
                        """() => {
                            const visible = element => !!element && !!(element.offsetWidth || element.offsetHeight);
                            const rows = [...document.querySelectorAll('.cdk-overlay-pane .table-row')].filter(visible);
                            if (!rows.length) return false;
                            const signature = rows.map(row => row.innerText).join('\\n');
                            const now = performance.now();
                            const previous = window.__cccVendorInventoryStability;
                            if (!previous || previous.signature !== signature) {
                                window.__cccVendorInventoryStability = { signature, startedAt: now };
                                return false;
                            }
                            return now - previous.startedAt >= 1500;
                        }""",
                        timeout=8000,
                        polling=250,
                    )
                except PlaywrightTimeoutError:
                    log.warning("CCC vendor inventory kept changing; reading currently visible rows")
                return retry_count
            retry_control = self.first_visible(retry_button)
            if retry_control is None:
                log.error("CCC exposed a Retry outcome but its button was not accessible")
                return None
            if retry_count >= MAX_VENDOR_SEARCH_RETRIES:
                log.error("CCC vendor search still has no offers after %d retries", retry_count)
                return None

            retry_count += 1
            self.vendor_search_retry_count = retry_count
            self.checkpoint(f"Retrying CCC vendor inventory search ({retry_count}/{MAX_VENDOR_SEARCH_RETRIES})")
            retry_control.click()

    def resolve_part_prompts(self) -> None:
        for _ in range(4):
            try:
                self.page.wait_for_function(
                    """() => {
                        const text = (document.body?.innerText || '').toLowerCase();
                        return (text.includes('part does not match') && (text.includes('model') || text.includes('engine') || text.includes('year')))
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
                mismatch_prompt = self.page.get_by_text(MODEL_MISMATCH_RE).last
                prompt_container = mismatch_prompt.locator(
                    "xpath=ancestor::*[.//button[normalize-space(.)='Yes']][1]"
                )
                yes_pattern = re.compile(r"^\s*YES\s*$", re.I)
                yes_button = None
                for candidate in (
                    prompt_container.get_by_role("button", name=yes_pattern),
                    prompt_container.get_by_text(yes_pattern),
                    self.page.get_by_role("button", name=yes_pattern),
                    self.page.get_by_text(yes_pattern),
                ):
                    yes_button = self.first_visible(candidate)
                    if yes_button is not None:
                        break
                if yes_button is None:
                    self.capture_failure("model_mismatch_yes_not_found")
                    raise RuntimeError("CCC's model-mismatch prompt appeared without a visible Yes control")
                yes_button.click(timeout=10000)
                self.checkpoint("Accepted CCC model-mismatch confirmation for the selected part")
                mismatch_prompt.wait_for(state="hidden", timeout=10000)
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

    def read_vendor_inventory(self) -> tuple[list[PartOffer], list[dict[str, str]]]:
        rows = self.page.locator(".cdk-overlay-pane .table-row").evaluate_all(
            """rows => {
                const fields = {
                    description: 'description', vendor: 'vendor', partno: 'part-no',
                    partnumber: 'part-no', delivery: 'delivery', extprice: 'ext-price',
                    labor: 'labor', paint: 'paint', total: 'total', diff: 'diff', margin: 'margin'
                };
                const outputFields = [...new Set(Object.values(fields))];
                const attributes = ['data-full-value', 'data-original-value', 'data-value', 'title', 'aria-label'];
                return rows.map(row => {
                    const data = Object.fromEntries(outputFields.map(field => [field, '']));
                    for (const cell of row.querySelectorAll('[class*="column-"]')) {
                        const columnClass = [...cell.classList].find(name => name.startsWith('column-')) || '';
                        const key = columnClass.slice(7).toLowerCase().replace(/[^a-z0-9]/g, '');
                        const field = fields[key];
                        if (!field) continue;
                        let value = (cell.innerText || cell.textContent || '').trim();
                        if (field === 'part-no') {
                            const candidates = [cell, ...cell.querySelectorAll('*')].flatMap(node =>
                                attributes.map(name => node.getAttribute(name)).filter(candidate => candidate && candidate.trim())
                            ).map(candidate => candidate.trim()).filter(candidate =>
                                candidate.length > value.length && !/(?:\\.\\.\\.|…)/.test(candidate)
                            );
                            if (candidates.length) value = candidates.reduce((best, candidate) => candidate.length > best.length ? candidate : best, value);
                        }
                        data[field] = value;
                    }
                    return data;
                }).filter(data => Object.values(data).some(Boolean));
            }"""
        )
        market_data = sorted(rows, key=lambda item: normalize(item.get("description", "")))
        offers = []
        for row in market_data:
            total = parse_price(row["ext-price"])
            if row["description"] and row["vendor"] and total is not None:
                offers.append(PartOffer(row["vendor"], row["description"], total))
        if not offers:
            raise RuntimeError("CCC vendor result grid contained no readable Vendor/Description/EXT. PRICE rows")
        self.close_vendor_inventory()
        return offers, market_data

    def close_vendor_inventory(self) -> None:
        self.checkpoint("Closing CCC vendor inventory with Done")
        done = self.page.locator(".cdk-overlay-pane button").filter(
            has_text=re.compile(r"^\s*DONE\s*$", re.I)
        )
        try:
            done.first.wait_for(state="visible", timeout=10000)
            done.first.click(timeout=10000)
        except PlaywrightTimeoutError:
            raise RuntimeError("CCC vendor inventory results loaded, but the modal Done button was not found") from None
        self.page.locator(".cdk-overlay-pane").wait_for(state="hidden", timeout=15000)
        self.checkpoint("CCC vendor inventory modal closed")

    @staticmethod
    def summarize(
        vin: str,
        oem: str,
        offers: list[PartOffer],
        description_target: str = TARGET_DESCRIPTION,
        partslink: str = "",
        part_description: str = "",
    ) -> dict[str, Any]:
        matching_offers = [
            offer for offer in offers
            if description_matches_target(offer.description, description_target)
        ]
        ipart = [offer for offer in matching_offers if normalize(offer.vendor) == normalize(IPART_VENDOR)]
        ipart_available = bool(ipart)
        sorted_all = sorted(matching_offers, key=lambda offer: offer.total)
        lowest_overall = sorted_all[0] if sorted_all else None
        second_lowest = sorted_all[1] if len(sorted_all) > 1 else None
        ipart_lowest = min(ipart, key=lambda offer: offer.total) if ipart else None
        return {
            "VIN": vin,
            "OEM": oem,
            "Partslink": partslink,
            "Part Description": part_description,
            "Ipart listed on CCC": ipart_available,
            "Ipart List Price CCC": ipart_lowest.total if ipart_lowest else None,
            "Lowest CCC Price": lowest_overall.total if lowest_overall else None,
            "Lowest CCC Supplier": lowest_overall.vendor if lowest_overall else None,
            "Next Lowest CCC Price": second_lowest.total if second_lowest else None,
            "Next Lowest CCC Supplier": second_lowest.vendor if second_lowest else None,
            "Lowest CCC Price Supplier": lowest_overall.vendor if lowest_overall else None,
            "Status": "OK" if matching_offers else f"No offers matching description: {description_target}",
        }

    def recover_closed_page(self) -> None:
        if self.page.is_closed() is not True:
            return
        self.checkpoint("CCC browser tab closed; opening a new tab and restoring the session")
        self.page = self.context.new_page()
        self.page.route("**/*", block_nonessential_resources)
        self.sign_in_if_needed()
        self.context.storage_state(path=str(AUTH_STATE))
        self.open_workfiles()
        self._workfile_initialized = False

    def crawl_row(
        self,
        record: dict[str, str],
        retry_attempt: int = 0,
    ) -> dict[str, Any]:
        vin = record["VIN"]
        scenario_mode = bool(record.get("Scenario"))
        partslink = record.get("PARTSLINK", "")
        oem = record.get("OEM", "")
        lookup_value = partslink if scenario_mode else oem
        self.checkpoint(f"Input row {record['_row']}: VIN={vin}, {'PARTSLINK' if scenario_mode else 'OEM'}={lookup_value}")
        result: dict[str, Any] = {
            "Year": "", "Make": "", "Model": "", "VIN": vin,
            "OEM": oem, "Partslink": partslink, "Part Description": "",
        }
        self.no_sourceable_candidates = False
        try:
            self.recover_closed_page()
            if not self._workfile_initialized:
                self.open_new_workfile()
            elif not self.page.locator('a[aria-label="Vehicle"]').count():
                if self.workfiles_url:
                    self.page.goto(self.workfiles_url, wait_until="domcontentloaded")
                self.open_new_workfile()
            if not self.enter_vehicle(vin):
                result["Status"] = "Vehicle/model not found; skipped"
            else:
                result.update(self._vehicle_details)
                attempted_catalog_parts: set[tuple[str, int]] = set()
                selected_oem = ""
                part_description = ""

                def open_catalog_lookup() -> int | None:
                    nonlocal selected_oem, part_description
                    if scenario_mode:
                        retry_count_value, selected_oem, part_description = self.open_scenario_part_results(
                            record["Scenario"], record["PARTSLINK"]
                        )
                        return retry_count_value
                    return self.open_part_results(oem, attempted_catalog_parts)

                retry_count = open_catalog_lookup()
                if scenario_mode:
                    result["OEM"] = selected_oem
                    result["Part Description"] = part_description
                for candidate_attempt in range(1, MAX_CATALOG_CANDIDATE_RETRIES):
                    if scenario_mode or retry_count is not None or not self.no_sourceable_candidates:
                        break
                    log.info(
                        "Trying the next untested catalog option for VIN=%s OEM=%s (%d/%d)",
                        vin,
                        oem,
                        candidate_attempt,
                        MAX_CATALOG_CANDIDATE_RETRIES - 1,
                    )
                    self._workfile_initialized = False
                    self.open_new_workfile()
                    if not self.enter_vehicle(vin):
                        retry_count = None
                        break
                    result.update(self._vehicle_details)
                    previous_candidate_count = len(attempted_catalog_parts)
                    retry_count = open_catalog_lookup()
                    if scenario_mode:
                        result["OEM"] = selected_oem
                        result["Part Description"] = part_description
                    if len(attempted_catalog_parts) == previous_candidate_count:
                        log.info(
                            "No new catalog options appeared for VIN=%s OEM=%s; stopping candidate retries",
                            vin,
                            oem,
                        )
                        break
                if retry_count is None:
                    body_text = normalize(self.page.locator("body").inner_text())
                    if self.no_sourceable_candidates or NO_SOURCE_PARTS_RE.search(body_text):
                        result["Status"] = "No sourceable parts available; skipped"
                        log.warning("Skipping VIN=%s OEM=%s: CCC added the estimate line but no sourceable parts were available", vin, oem)
                    else:
                        retries = getattr(self, "vendor_search_retry_count", 0)
                        result["Status"] = f"Vendor search failed after {retries} Retry clicks; skipped"
                else:
                    offers, market_data = self.read_vendor_inventory()
                    result = self.summarize(
                        vin,
                        selected_oem if scenario_mode else oem,
                        offers,
                        self.description_target,
                        partslink,
                        part_description,
                    )
                    result.update(self._vehicle_details)
                    result["market_data"] = json.dumps(market_data, ensure_ascii=False)
                    log.info("Completed VIN=%s OEM=%s; iPart listed=%s", vin, oem, result["Ipart listed on CCC"])
        except CatalogPartNotFoundError as error:
            log.warning("Skipping VIN=%s OEM=%s: %s", vin, oem, error)
            result["Status"] = f"{error}; skipped"
        except CatalogSearchTimeoutError as error:
            self._workfile_initialized = False
            log.warning(
                "Catalog search did not complete for VIN=%s OEM=%s; continuing after row retry: %s",
                vin,
                oem,
                error,
            )
            result["Status"] = f"ERROR: catalog search did not complete: {error}; retry on resume"
        except FatalCrawlerError as error:
            self._workfile_initialized = False
            log.warning(
                "Unrecoverable CCC row state for VIN=%s OEM=%s; continuing with the next row after retry: %s",
                vin,
                oem,
                error,
            )
            result["Status"] = f"ERROR: CCC row state could not be recovered: {error}; retry on resume"
        except Exception as error:
            error_text = str(error).lower()
            browser_closed = self.page.is_closed() or (
                isinstance(error, PlaywrightError) and "closed" in error_text
            )
            overlay_blocked = isinstance(error, PlaywrightError) and "intercepts pointer events" in error_text
            if browser_closed:
                try:
                    self.capture_failure(f"browser_closed_{record['_row']}_{vin}")
                except PlaywrightError:
                    pass
                log.warning(
                    "CCC browser tab closed while processing VIN=%s OEM=%s; the row will be retried and the run will continue",
                    vin,
                    oem,
                )
                result["Status"] = f"ERROR: CCC browser tab closed: {error}; retry on resume"
                self._workfile_initialized = False
            elif overlay_blocked:
                try:
                    self.capture_failure(f"overlay_blocked_{record['_row']}_{vin}")
                except PlaywrightError:
                    pass
                log.warning(
                    "Skipping VIN=%s OEM=%s: CCC overlay blocked interaction; continuing with a fresh workfile",
                    vin,
                    oem,
                )
                result["Status"] = "ERROR: CCC overlay blocked interaction; retry on resume"
                self._workfile_initialized = False
            else:
                log.exception("Row failed: VIN=%s OEM=%s: %s", vin, oem, error)
                self.capture_failure(f"row_{record['_row']}_{vin}")
                result["Status"] = f"ERROR: {error}"
                self._workfile_initialized = False
        finally:
            result["date_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        result["Description Target"] = self.description_target
        if should_retry_row_status(str(result.get("Status") or "")):
            self._workfile_initialized = False
            if retry_attempt < MAX_ROW_RETRIES:
                log.warning(
                    "Retrying VIN=%s OEM=%s in a fresh workfile (%d/%d) after: %s",
                    vin,
                    oem,
                    retry_attempt + 1,
                    MAX_ROW_RETRIES,
                    result["Status"],
                )
                return self.crawl_row(record, retry_attempt + 1)
        return result


def main() -> int:
    load_dotenv()
    parser = argparse.ArgumentParser(description="Crawl CCC vendor prices from an Excel input workbook.")
    parser.add_argument("--input", type=Path, default=Path(os.getenv("CCC_INPUT", "input.xlsx")))
    parser.add_argument("--output", type=Path, default=Path(os.getenv("CCC_OUTPUT", "ccc-results.xlsx")))
    parser.add_argument(
        "--description",
        default=os.getenv("CCC_DESCRIPTION", TARGET_DESCRIPTION),
        help=f"Description substring to match when ranking offers (default: {TARGET_DESCRIPTION}).",
    )
    parser.add_argument("--headed", action="store_true", help="Show the browser window (recommended for first run).")
    parser.add_argument("--resume", action="store_true", help="Continue from completed rows in the existing output workbook.")
    parser.add_argument(
        "--search-action",
        choices=("vin_oem", "vin_partslink_scenario"),
        default="vin_oem",
    )
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
        records = read_input(args.input, args.search_action)
    except Exception as error:
        log.error("Could not read input workbook: %s", error)
        return 2
    if not records:
        log.error("No valid CCC input rows were found in %s", args.input)
        return 2

    results = load_results(args.output) if args.resume else []
    if args.resume:
        log.info("Resume enabled: loaded %d saved result row(s) from %s", len(results), args.output)
    if results:
        log.info("Loaded %d saved result(s) from %s", len(results), args.output)
    result_slots, pending_groups = prepare_result_slots(records, results, args.description)
    if not pending_groups:
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
        page.route("**/*", block_nonessential_resources)
        crawler = CCCCrawler(page, username, password, Path("artifacts"), args.description, context)
        try:
            crawler.sign_in_if_needed()
            context.storage_state(path=str(AUTH_STATE))
            crawler.open_workfiles()
            for group_number, (key, indexes) in enumerate(pending_groups.items(), start=1):
                first_index = indexes[0]
                log.info(
                    "Progress: lookup %d/%d for %d matching input row(s)",
                    group_number,
                    len(pending_groups),
                    len(indexes),
                )
                result = crawler.crawl_row(records[first_index])
                for index, expanded in zip(indexes, expand_group_result(result, records, indexes)):
                    result_slots[index] = expanded
                save_results(args.output, [row for row in result_slots if row is not None])
                log.info("Saved progress to %s", args.output)
        except Exception as error:
            log.exception("Crawler stopped before completing all rows: %s", error)
            try:
                save_results(args.output, [row for row in result_slots if row is not None])
            except Exception:
                log.exception("Could not save partial results to %s", args.output)
            try:
                crawler.capture_failure("crawler_stopped")
            except Exception:
                log.exception("Could not capture failure evidence")
            return 1
        finally:
            try:
                context.close()
            except PlaywrightError as error:
                log.warning("CCC browser context was already closed during cleanup: %s", error)
            try:
                browser.close()
            except PlaywrightError as error:
                log.warning("CCC browser was already closed during cleanup: %s", error)
    completed_rows = sum(result is not None for result in result_slots)
    log.info("Crawl complete: %d/%d rows written to %s", completed_rows, len(records), args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())