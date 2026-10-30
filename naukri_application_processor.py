from __future__ import annotations

import argparse
import json
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse

from openpyxl import load_workbook
from playwright.sync_api import (
    BrowserContext,
    Page,
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

# Main Excel created by your job extractor
MAIN_EXCEL = BASE_DIR / "naukri_job_list.xlsx"

# Separate workbook for jobs that require applying on the company website.
EXTERNAL_EXCEL = BASE_DIR / "naukri_external_applications.xlsx"

# Permanent URL ledger used by the extractor to avoid processing jobs again.
PROCESSED_REQUESTS_EXCEL = BASE_DIR / "processed_requests.xlsx"

# Durable record of confirmed application outcomes whose Excel row cleanup failed.
# This prevents the same confirmed application from being submitted again on a later run.
CONFIRMED_OUTCOMES_FILE = BASE_DIR / "naukri_confirmed_outcomes.json"

# Use your existing Naukri profile/session
PROFILE_DIR = BASE_DIR / "naukri_browser_profile"

# How long to wait for the application to be completed manually
APPLICATION_TIMEOUT_MINUTES = 30

# Time to wait after opening a job page
JOB_PAGE_LOAD_WAIT_MS = 2000

# Time to wait after opening the application panel
APPLICATION_PANEL_SETTLE_WAIT_MS = 2500

# Main Excel column names
COMPANY_COLUMN = "Company Name"
TITLE_COLUMN = "Job Title"
URL_COLUMN = "Job URL"

# Columns for the separate external-applications workbook
EXTERNAL_COMPANY_COLUMN = "Company Name"
EXTERNAL_TITLE_COLUMN = "Job Title"
EXTERNAL_URL_COLUMN = "Company Application URL"


# ============================================================
# GENERAL HELPERS
# ============================================================

def log(message: str) -> None:
    """Print a timestamped message and persist it to a debug log file."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    print(line, flush=True)

    # Keep a persistent log so row-removal problems can be diagnosed after
    # the terminal window closes. Logging must never stop the automation.
    try:
        base_dir = globals().get("BASE_DIR", Path(__file__).resolve().parent)
        with (base_dir / "naukri_processor_debug.log").open(
                "a", encoding="utf-8"
        ) as log_file:
            log_file.write(line + "\n")
    except Exception:
        pass


def normalize_text(value: object) -> str:
    """Convert a value to a clean string."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def safe_filename_part(value: str) -> str:
    """Make a string safe for use in a filename."""
    value = re.sub(r'[<>:"/\\|?*]', "_", value)
    value = re.sub(r"\s+", "_", value).strip("_")
    return value[:80] or "jobs"


def is_valid_http_url(url: str) -> bool:
    """Return True when the value looks like a valid HTTP(S) URL."""
    try:
        parsed = urlparse(url)
        return parsed.scheme in {"http", "https"} and bool(parsed.netloc)
    except Exception:
        return False


def is_naukri_job_url(url: str) -> bool:
    """Check whether a URL appears to be a Naukri job URL."""
    try:
        parsed = urlparse(url)
        return (
                parsed.scheme in {"http", "https"}
                and "naukri.com" in parsed.netloc.lower()
        )
    except Exception:
        return False


# ============================================================
# EXCEL HELPERS
# ============================================================

def load_main_jobs() -> list[dict]:
    """
    Load jobs from the main Excel file.

    Expected columns:
      Company Name
      Job Title
      Job URL

    Extra columns, if present, are preserved when saving the file.
    """
    if not MAIN_EXCEL.exists():
        raise FileNotFoundError(
            f"Main Excel file was not found:\n{MAIN_EXCEL}"
        )

    workbook = load_workbook(MAIN_EXCEL)
    worksheet = workbook.active

    if worksheet.max_row < 2:
        workbook.close()
        return []

    headers = [
        normalize_text(cell.value)
        for cell in worksheet[1]
    ]

    header_map = {
        header.lower(): index
        for index, header in enumerate(headers)
        if header
    }

    required_columns = [
        COMPANY_COLUMN,
        TITLE_COLUMN,
        URL_COLUMN,
    ]

    missing_columns = [
        column
        for column in required_columns
        if column.lower() not in header_map
    ]

    if missing_columns:
        workbook.close()
        raise ValueError(
            "The main Excel file is missing these columns: "
            + ", ".join(missing_columns)
            + f"\nFound columns: {headers}"
        )

    company_index = header_map[COMPANY_COLUMN.lower()]
    title_index = header_map[TITLE_COLUMN.lower()]
    url_index = header_map[URL_COLUMN.lower()]

    jobs = []
    seen_urls = set()

    for row_number in range(2, worksheet.max_row + 1):
        company = normalize_text(
            worksheet.cell(row_number, company_index + 1).value
        )
        title = normalize_text(
            worksheet.cell(row_number, title_index + 1).value
        )
        url = normalize_text(
            worksheet.cell(row_number, url_index + 1).value
        )

        # Ignore completely empty rows and rows without a job URL.
        if not url:
            continue

        # Do not treat stale/orphan URL-only rows as jobs. These can be left
        # behind by earlier workbook edits and must not open blank-company jobs.
        if not company and not title:
            log(
                f"[EXCEL-LOAD] Skipping orphan URL-only row {row_number}: {url!r}"
            )
            continue

        # Excel tables can retain stale rows after deletion. Deduplicate
        # URLs in this in-memory snapshot so one job is never reopened
        # repeatedly during the same run.
        if url in seen_urls:
            log(
                f"Skipping duplicate URL in Excel row {row_number}: {url}"
            )
            continue

        seen_urls.add(url)
        jobs.append(
            {
                "company": company,
                "title": title,
                "url": url,
                "row_number": row_number,
            }
        )

    workbook.close()
    return jobs


def _last_nonempty_row(worksheet) -> int:
    """Find the last row containing at least one actual cell value."""
    for row_number in range(worksheet.max_row, 1, -1):
        if any(
                cell.value is not None and normalize_text(cell.value) != ""
                for cell in worksheet[row_number]
        ):
            return row_number
    return 1


def _refresh_excel_table_ranges(worksheet) -> None:
    """
    Keep Excel table/filter ranges aligned after row deletion.
    openpyxl does not automatically update table references when rows
    are deleted; a stale table range can make old URLs appear again.
    """
    last_row = _last_nonempty_row(worksheet)

    for table in worksheet.tables.values():
        try:
            from openpyxl.utils.cell import range_boundaries

            min_col, min_row, max_col, _old_max_row = range_boundaries(table.ref)
            if last_row < min_row:
                last_row = min_row
            table.ref = (
                f"{worksheet.cell(min_row, min_col).coordinate}:"
                f"{worksheet.cell(last_row, max_col).coordinate}"
            )
            if getattr(table, "autoFilter", None) is not None:
                table.autoFilter.ref = table.ref
        except Exception as exc:
            log(f"WARNING: Could not refresh an Excel table range: {exc}")

    if worksheet.auto_filter.ref:
        try:
            from openpyxl.utils.cell import range_boundaries

            min_col, min_row, max_col, _old_max_row = range_boundaries(
                worksheet.auto_filter.ref
            )
            worksheet.auto_filter.ref = (
                f"{worksheet.cell(min_row, min_col).coordinate}:"
                f"{worksheet.cell(last_row, max_col).coordinate}"
            )
        except Exception:
            pass


def remove_job_from_main_excel(job_url: str) -> bool:
    """
    Remove every row whose Job URL cell value matches job_url, repair
    URL hyperlinks after row deletion, and remove previously-created
    orphan rows that have a URL but no Company Name and Job Title.
    """
    log("[EXCEL-REMOVE] ===== START =====")
    log(f"[EXCEL-REMOVE] Script path: {Path(__file__).resolve()}")
    log(f"[EXCEL-REMOVE] Main workbook path: {MAIN_EXCEL.resolve()}")
    log(f"[EXCEL-REMOVE] Incoming job URL (raw): {job_url!r}")

    if not MAIN_EXCEL.exists():
        log("[EXCEL-REMOVE] ERROR: Main Excel does not exist.")
        return False

    job_url = normalize_text(job_url)
    if not job_url:
        log("[EXCEL-REMOVE] ERROR: Empty job URL; refusing to remove a row.")
        return False

    workbook = None
    try:
        stat = MAIN_EXCEL.stat()
        log(
            f"[EXCEL-REMOVE] Workbook before edit: size={stat.st_size} bytes, "
            f"modified={datetime.fromtimestamp(stat.st_mtime).isoformat(timespec='seconds')}"
        )

        workbook = load_workbook(MAIN_EXCEL, data_only=False)
        worksheet = workbook.active
        headers = [normalize_text(cell.value) for cell in worksheet[1]]
        log(
            f"[EXCEL-REMOVE] Loaded sheet={worksheet.title!r}; "
            f"max_row={worksheet.max_row}; max_column={worksheet.max_column}; "
            f"tables={list(worksheet.tables.keys())}"
        )
        log(f"[EXCEL-REMOVE] Headers: {headers!r}")

        column_map = {
            header.lower(): index
            for index, header in enumerate(headers, start=1)
            if header
        }
        company_col = column_map.get(COMPANY_COLUMN.lower())
        title_col = column_map.get(TITLE_COLUMN.lower())
        url_col = column_map.get(URL_COLUMN.lower())
        if not company_col or not title_col or not url_col:
            raise ValueError(
                "Required columns not found. "
                f"Expected {COMPANY_COLUMN!r}, {TITLE_COLUMN!r}, {URL_COLUMN!r}; "
                f"found {headers!r}"
            )

        matching_rows = []
        for row_number in range(2, worksheet.max_row + 1):
            url_cell = worksheet.cell(row_number, url_col)
            value = normalize_text(url_cell.value)
            target = normalize_text(
                url_cell.hyperlink.target if url_cell.hyperlink else ""
            )
            if value == job_url:
                matching_rows.append(row_number)
                log(
                    f"[EXCEL-REMOVE] VALUE MATCH row={row_number}: "
                    f"company={worksheet.cell(row_number, company_col).value!r}; "
                    f"title={worksheet.cell(row_number, title_col).value!r}; "
                    f"url_value={value!r}; hyperlink_target={target!r}"
                )
            elif target == job_url:
                # A stale hyperlink target must not cause a different job row
                # to be deleted. It will be repaired to match the cell value.
                log(
                    f"[EXCEL-REMOVE] STALE HYPERLINK row={row_number}: "
                    f"displayed_url={value!r}; hyperlink_target={target!r}. "
                    "Will repair hyperlink without deleting this different job."
                )

        if not matching_rows:
            log(
                f"[EXCEL-REMOVE] ERROR: No cell-value row matched {job_url!r}. "
                "No job row will be deleted."
            )
            return False

        log(f"[EXCEL-REMOVE] Deleting complete matching row(s): {matching_rows}")
        for row_number in reversed(matching_rows):
            before = [
                worksheet.cell(row_number, c).value
                for c in range(1, worksheet.max_column + 1)
            ]
            log(f"[EXCEL-REMOVE] Deleting row {row_number}: {before!r}")
            worksheet.delete_rows(row_number, 1)

        # Remove corrupt rows left by earlier runs: a row with no company and
        # no title is not a usable job record, even if column C still has a URL.
        orphan_rows = []
        for row_number in range(2, worksheet.max_row + 1):
            company = normalize_text(worksheet.cell(row_number, company_col).value)
            title = normalize_text(worksheet.cell(row_number, title_col).value)
            url_cell = worksheet.cell(row_number, url_col)
            url_value = normalize_text(url_cell.value)
            hyperlink_target = normalize_text(
                url_cell.hyperlink.target if url_cell.hyperlink else ""
            )
            row_has_content = any(
                normalize_text(worksheet.cell(row_number, column).value)
                for column in range(1, worksheet.max_column + 1)
            ) or bool(url_cell.hyperlink)

            # User rule: if BOTH identifying fields are blank, the row is useless.
            # Remove any populated stale row, not only rows containing a URL.
            if row_has_content and not company and not title:
                orphan_rows.append(row_number)
                log(
                    f"[EXCEL-REMOVE] Removing stale row {row_number}: "
                    f"Company Name and Job Title are both blank; "
                    f"url_value={url_value!r}; hyperlink_target={hyperlink_target!r}"
                )

        for row_number in reversed(orphan_rows):
            worksheet.delete_rows(row_number, 1)

        # openpyxl does not reliably shift hyperlink relationships when rows
        # are deleted. Do not trust the old hyperlink object, even if its target
        # looks correct in memory: clear every URL-column hyperlink and rebuild
        # it from that row's displayed URL value before saving.
        rebuilt_links = 0
        cleared_links = 0
        for row_number in range(2, worksheet.max_row + 1):
            url_cell = worksheet.cell(row_number, url_col)
            url_value = normalize_text(url_cell.value)

            # Explicitly clear the underlying hyperlink object first. This also
            # discards stale relationship IDs left behind by row deletion.
            url_cell._hyperlink = None
            if not url_value:
                cleared_links += 1
                continue

            if is_valid_http_url(url_value):
                from openpyxl.worksheet.hyperlink import Hyperlink
                url_cell._hyperlink = Hyperlink(
                    ref=url_cell.coordinate,
                    target=url_value,
                    display=url_value,
                )
                rebuilt_links += 1

        log(
            f"[EXCEL-REMOVE] Deleted {len(matching_rows)} matching row(s) "
            f"and {len(orphan_rows)} orphan row(s). "
            f"Rebuilt {rebuilt_links} URL hyperlink(s) from cell values; "
            f"cleared {cleared_links} blank URL cell(s)."
        )

        _refresh_excel_table_ranges(worksheet)
        log(
            "[EXCEL-REMOVE] Table refs after refresh: "
            + repr({name: table.ref for name, table in worksheet.tables.items()})
        )

        workbook.save(MAIN_EXCEL)
        workbook.close()
        workbook = None
        log("[EXCEL-REMOVE] Save completed; reopening the physical workbook for verification.")

        verify_workbook = load_workbook(MAIN_EXCEL, data_only=False)
        try:
            verify_sheet = verify_workbook.active
            verify_headers = [normalize_text(cell.value) for cell in verify_sheet[1]]
            verify_url_col = next(
                (i for i, header in enumerate(verify_headers, start=1)
                 if header.lower() == URL_COLUMN.lower()),
                None,
            )
            verify_company_col = next(
                (i for i, header in enumerate(verify_headers, start=1)
                 if header.lower() == COMPANY_COLUMN.lower()),
                None,
            )
            verify_title_col = next(
                (i for i, header in enumerate(verify_headers, start=1)
                 if header.lower() == TITLE_COLUMN.lower()),
                None,
            )
            if not verify_url_col or not verify_company_col or not verify_title_col:
                raise ValueError("Required columns are missing after save.")

            remaining_values = []
            remaining_links = []
            remaining_orphans = []
            hyperlink_mismatches = []

            for row_number in range(2, verify_sheet.max_row + 1):
                url_cell = verify_sheet.cell(row_number, verify_url_col)
                value = normalize_text(url_cell.value)
                target = normalize_text(
                    url_cell.hyperlink.target if url_cell.hyperlink else ""
                )
                company = normalize_text(verify_sheet.cell(row_number, verify_company_col).value)
                title = normalize_text(verify_sheet.cell(row_number, verify_title_col).value)

                if value == job_url:
                    remaining_values.append(row_number)
                if target == job_url:
                    remaining_links.append(row_number)
                if (value or target) and not company and not title:
                    remaining_orphans.append(row_number)
                if value and is_valid_http_url(value) and target != value:
                    hyperlink_mismatches.append((row_number, value, target))

            log(
                f"[EXCEL-REMOVE] Verification: value_matches={remaining_values}; "
                f"hyperlink_matches={remaining_links}; orphan_rows={remaining_orphans}; "
                f"hyperlink_mismatches={len(hyperlink_mismatches)}"
            )
            for mismatch in hyperlink_mismatches[:20]:
                log(
                    f"[EXCEL-REMOVE] HYPERLINK MISMATCH row={mismatch[0]}: "
                    f"value={mismatch[1]!r}; target={mismatch[2]!r}"
                )

            if remaining_values or remaining_links or remaining_orphans or hyperlink_mismatches:
                log("[EXCEL-REMOVE] ERROR: Verification failed; check the rows reported above.")
                return False
        finally:
            verify_workbook.close()

        stat = MAIN_EXCEL.stat()
        log(
            f"[EXCEL-REMOVE] SUCCESS: Verified removal of {job_url!r}. "
            f"Workbook size={stat.st_size}; "
            f"modified={datetime.fromtimestamp(stat.st_mtime).isoformat(timespec='seconds')}"
        )
        log("[EXCEL-REMOVE] ===== END: SUCCESS =====")
        return True

    except Exception as exc:
        log(f"[EXCEL-REMOVE] EXCEPTION: {type(exc).__name__}: {exc}")
        return False
    finally:
        if workbook is not None:
            try:
                workbook.close()
            except Exception:
                pass

def append_processed_request(
        company: str,
        title: str,
        naukri_url: str,
        processing_type: str,
) -> bool:
    """
    Persist a job's Naukri URL after its processing outcome is confirmed.
    The original Naukri URL is the deduplication key; external company URLs
    remain in the separate external-applications workbook.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    naukri_url = normalize_text(naukri_url)
    if not is_valid_http_url(naukri_url) or not is_naukri_job_url(naukri_url):
        log(f"[PROCESSED-REQUESTS] Refusing invalid Naukri job URL: {naukri_url!r}")
        return False

    headers = [
        "Company Name",
        "Job Title",
        "Job URL",
        "Processing Type",
        "Processed At",
    ]
    temporary_path = PROCESSED_REQUESTS_EXCEL.with_name(
        f"{PROCESSED_REQUESTS_EXCEL.stem}_temporary{PROCESSED_REQUESTS_EXCEL.suffix}"
    )
    workbook = None

    try:
        if PROCESSED_REQUESTS_EXCEL.exists():
            workbook = load_workbook(PROCESSED_REQUESTS_EXCEL)
            worksheet = workbook.active
            existing_headers = [normalize_text(cell.value) for cell in worksheet[1]]
            if [h.lower() for h in existing_headers[:len(headers)]] != [h.lower() for h in headers]:
                raise ValueError(
                    f"Unexpected headers in {PROCESSED_REQUESTS_EXCEL.name}: "
                    f"{existing_headers!r}"
                )
        else:
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Processed Requests"
            worksheet.append(headers)
            for cell in worksheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="2457C5")
                cell.alignment = Alignment(vertical="center")
            worksheet.freeze_panes = "A2"
            worksheet.column_dimensions["A"].width = 32
            worksheet.column_dimensions["B"].width = 48
            worksheet.column_dimensions["C"].width = 100
            worksheet.column_dimensions["D"].width = 26
            worksheet.column_dimensions["E"].width = 22

        url_col = headers.index("Job URL") + 1
        for row_number in range(2, worksheet.max_row + 1):
            existing_url = normalize_text(worksheet.cell(row_number, url_col).value)
            if existing_url == naukri_url:
                log(
                    "[PROCESSED-REQUESTS] URL already recorded; "
                    f"not adding a duplicate: {naukri_url}"
                )
                workbook.close()
                workbook = None
                return True

        new_row = worksheet.max_row + 1
        values = [
            normalize_text(company),
            normalize_text(title),
            naukri_url,
            normalize_text(processing_type),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ]
        for column, value in enumerate(values, start=1):
            worksheet.cell(row=new_row, column=column, value=value)

        url_cell = worksheet.cell(row=new_row, column=url_col)
        url_cell.hyperlink = naukri_url
        url_cell.style = "Hyperlink"

        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.row_dimensions[1].height = 24
        workbook.save(temporary_path)
        workbook.close()
        workbook = None
        temporary_path.replace(PROCESSED_REQUESTS_EXCEL)

        log(
            f"[PROCESSED-REQUESTS] Saved {processing_type}: "
            f"{naukri_url} -> {PROCESSED_REQUESTS_EXCEL.name}"
        )
        return True

    except Exception as exc:
        log(f"[PROCESSED-REQUESTS] ERROR saving processed URL: {exc}")
        return False
    finally:
        if workbook is not None:
            try:
                workbook.close()
            except Exception:
                pass
        if temporary_path.exists():
            try:
                temporary_path.unlink()
            except Exception:
                pass


def append_external_application(
        company: str,
        title: str,
        naukri_url: str,
        company_url: str,
) -> bool:
    """Save company name and external URL in a separate workbook.

    The job is removed from the main workbook only after this function
    successfully saves the external URL. Duplicate external URLs are not
    added more than once.
    """
    workbook = None
    temporary_path = EXTERNAL_EXCEL.with_name(
        f"{EXTERNAL_EXCEL.stem}_temporary{EXTERNAL_EXCEL.suffix}"
    )

    try:
        if not is_valid_http_url(company_url):
            log(f"ERROR: Invalid company application URL: {company_url!r}")
            return False

        if EXTERNAL_EXCEL.exists():
            workbook = load_workbook(EXTERNAL_EXCEL)
            worksheet = workbook.active
            headers = [normalize_text(cell.value) for cell in worksheet[1]]

            required_headers = [
                EXTERNAL_COMPANY_COLUMN,
                EXTERNAL_TITLE_COLUMN,
                EXTERNAL_URL_COLUMN,
            ]
            header_lookup = {
                header.lower(): index + 1
                for index, header in enumerate(headers)
                if header
            }

            # Add missing columns to an existing external workbook without
            # discarding previously saved rows.
            for required_header in required_headers:
                if required_header.lower() not in header_lookup:
                    new_column = worksheet.max_column + 1
                    worksheet.cell(row=1, column=new_column, value=required_header)
                    header_lookup[required_header.lower()] = new_column
        else:
            from openpyxl import Workbook

            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "External Applications"
            worksheet.append([
                EXTERNAL_COMPANY_COLUMN,
                EXTERNAL_TITLE_COLUMN,
                EXTERNAL_URL_COLUMN,
            ])
            header_lookup = {
                EXTERNAL_COMPANY_COLUMN.lower(): 1,
                EXTERNAL_TITLE_COLUMN.lower(): 2,
                EXTERNAL_URL_COLUMN.lower(): 3,
            }

        company_col = header_lookup[EXTERNAL_COMPANY_COLUMN.lower()]
        title_col = header_lookup[EXTERNAL_TITLE_COLUMN.lower()]
        url_col = header_lookup[EXTERNAL_URL_COLUMN.lower()]

        # If this exact external URL is already present, treat it as saved.
        for row_number in range(2, worksheet.max_row + 1):
            existing_url = normalize_text(worksheet.cell(row_number, url_col).value)
            if existing_url == company_url:
                log(
                    "External URL already exists in the separate workbook; "
                    "it will not be duplicated."
                )
                workbook.close()
                workbook = None
                return True

        new_row = worksheet.max_row + 1
        worksheet.cell(row=new_row, column=company_col, value=company)
        worksheet.cell(row=new_row, column=title_col, value=title)
        worksheet.cell(row=new_row, column=url_col, value=company_url)
        workbook.save(temporary_path)
        workbook.close()
        workbook = None
        temporary_path.replace(EXTERNAL_EXCEL)

        log(f"External application saved to separate Excel: {EXTERNAL_EXCEL.name}")
        log(f"Saved company: {company}")
        log(f"Saved job title: {title}")
        log(f"Saved external URL: {company_url}")
        return True

    except Exception as exc:
        log(f"ERROR saving external application to separate Excel: {exc}")
        return False
    finally:
        if workbook is not None:
            try:
                workbook.close()
            except Exception:
                pass
        if temporary_path.exists():
            try:
                temporary_path.unlink()
            except Exception:
                pass


# ============================================================
# PAGE AND BUTTON DETECTION
# ============================================================

def visible_text(page: Page) -> str:
    """Get the visible page text, returning an empty string on failure."""
    try:
        return page.locator("body").inner_text(timeout=5000)
    except Exception:
        return ""


def find_visible_button(
        page: Page,
        pattern: re.Pattern,
):
    """
    Find a visible button or button-like element whose accessible
    name or text matches the supplied regular expression.
    """
    selectors = [
        page.get_by_role("button", name=pattern),
        page.locator("button").filter(has_text=pattern),
        page.locator('[role="button"]').filter(has_text=pattern),
    ]

    for locator in selectors:
        try:
            count = min(locator.count(), 20)

            for index in range(count):
                candidate = locator.nth(index)

                if candidate.is_visible():
                    return candidate

        except Exception:
            continue

    return None



def dump_button_diagnostics(page: Page, reason: str) -> None:
    """Save page DOM and print visible interactive elements for debugging."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_reason = re.sub(r"[^a-zA-Z0-9_-]+", "_", reason).strip("_") or "unknown"
    log("----- BUTTON DETECTION DOM DIAGNOSTICS -----")
    try:
        log(f"Diagnostic reason: {reason}")
        log(f"Page URL: {page.url}")
        log(f"Page title: {page.title()}")
    except Exception as exc:
        log(f"Could not read page metadata: {exc}")

    try:
        elements = page.locator("body").evaluate(
            """body => {
                const visible = el => {
                    const s = window.getComputedStyle(el);
                    const r = el.getBoundingClientRect();
                    return s && s.display !== 'none' && s.visibility !== 'hidden' &&
                           Number(s.opacity || 1) > 0 && r.width > 0 && r.height > 0;
                };
                const norm = s => (s || '').replace(/\\s+/g, ' ').trim();
                const all = Array.from(body.querySelectorAll('button, a, [role="button"], [role="link"], input[type="button"], input[type="submit"], [onclick]'));
                return all.filter(visible).map((el, i) => ({
                    index: i,
                    tag: el.tagName,
                    role: el.getAttribute('role') || '',
                    text: norm(el.innerText || el.value || el.getAttribute('aria-label') || el.getAttribute('title') || ''),
                    ariaLabel: el.getAttribute('aria-label') || '',
                    title: el.getAttribute('title') || '',
                    href: el.href || el.getAttribute('href') || '',
                    className: typeof el.className === 'string' ? el.className : '',
                    testId: el.getAttribute('data-testid') || '',
                    outerHTML: el.outerHTML.slice(0, 1200)
                })).slice(0, 120);
            }"""
        )
        log(f"Visible interactive elements found: {len(elements)}")
        for item in elements:
            log(
                "DOM ELEMENT " + str(item.get("index")) + ": " +
                f"tag={item.get('tag')} role={item.get('role')} " +
                f"text={item.get('text')!r} href={item.get('href')!r} " +
                f"aria-label={item.get('ariaLabel')!r} " +
                f"class={item.get('className')!r}"
            )
            if re.search(r"apply|company|site|applied|save", item.get("text", ""), re.I):
                log("MATCHING ELEMENT HTML: " + item.get("outerHTML", ""))
    except Exception as exc:
        log(f"Could not enumerate interactive DOM elements: {exc}")

    try:
        dom_path = BASE_DIR / f"naukri_dom_{safe_reason}_{timestamp}.html"
        dom_path.write_text(page.content(), encoding="utf-8")
        log(f"Full DOM saved to: {dom_path.name}")
    except Exception as exc:
        log(f"Could not save full DOM: {exc}")

    try:
        screenshot_path = BASE_DIR / f"naukri_dom_{safe_reason}_{timestamp}.png"
        page.screenshot(path=str(screenshot_path), full_page=True)
        log(f"DOM screenshot saved to: {screenshot_path.name}")
    except Exception as exc:
        log(f"Could not save DOM screenshot: {exc}")

    log("----- END BUTTON DETECTION DIAGNOSTICS -----")

def get_keyskills_green_status(page: Page) -> Optional[bool]:
    """Inspect and log the Keyskills match indicator before any application action.

    Returns True for green, False for a confidently found non-green indicator,
    and None when the indicator cannot be identified confidently. Diagnostic
    details are written to naukri_processor_debug.log to make DOM changes easier
    to investigate without changing the conservative application policy.
    """
    try:
        try:
            page_title = page.title()
        except Exception as title_exc:
            page_title = f"<title unavailable: {title_exc}>"
        log(
            f"[KEYSKILLS] Starting detection on URL={page.url!r}; "
            f"page_title={page_title!r}"
        )

        result = page.locator("body").evaluate(
            r"""body => {
                const norm = s => (s || '').replace(/\s+/g, ' ').trim().toLowerCase();
                const visible = el => {
                    const s = window.getComputedStyle(el);
                    const r = el.getBoundingClientRect();
                    return s && s.display !== 'none' && s.visibility !== 'hidden' &&
                           Number(s.opacity || 1) > 0 && r.width > 0 && r.height > 0;
                };
                const colorInfo = el => {
                    const st = window.getComputedStyle(el);
                    return {
                        color: st.color,
                        backgroundColor: st.backgroundColor,
                        fill: st.fill,
                        stroke: st.stroke,
                        borderColor: st.borderColor
                    };
                };
                const isGreenColor = value => {
                    if (!value) return false;
                    const v = value.toLowerCase();
                    if (/green|success|text-green|bg-green/i.test(v)) return true;
                    const hex = v.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
                    if (hex) {
                        let h = hex[1];
                        if (h.length === 3) h = h.split('').map(c => c + c).join('');
                        const r = parseInt(h.slice(0, 2), 16);
                        const g = parseInt(h.slice(2, 4), 16);
                        const b = parseInt(h.slice(4, 6), 16);
                        return g >= 100 && g > r * 1.18 && g > b * 1.08;
                    }
                    const m = v.match(/rgba?\((\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
                    if (!m) return false;
                    const r = Number(m[1]), g = Number(m[2]), b = Number(m[3]);
                    return g >= 100 && g > r * 1.18 && g > b * 1.08;
                };
                const isGreenElement = el => {
                    const cls = typeof el.className === 'string' ? el.className :
                        (el.className && el.className.baseVal) || '';
                    const attrs = [cls, el.getAttribute('style') || '', el.getAttribute('aria-label') || '',
                                   el.getAttribute('title') || '', el.getAttribute('data-testid') || '',
                                   el.getAttribute('fill') || '', el.getAttribute('stroke') || ''].join(' ');
                    if (/green|success|text-green|bg-green|check-circle/i.test(attrs)) return true;
                    const colors = colorInfo(el);
                    return isGreenColor(colors.color) || isGreenColor(colors.backgroundColor) ||
                           isGreenColor(colors.fill) || isGreenColor(colors.stroke) ||
                           isGreenColor(el.getAttribute('fill')) || isGreenColor(el.getAttribute('stroke'));
                };
                const describe = el => {
                    const cls = typeof el.className === 'string' ? el.className :
                        (el.className && el.className.baseVal) || '';
                    return {
                        tag: el.tagName,
                        text: norm(el.innerText || el.textContent).slice(0, 100),
                        className: String(cls).slice(0, 220),
                        ariaLabel: el.getAttribute('aria-label') || '',
                        title: el.getAttribute('title') || '',
                        testId: el.getAttribute('data-testid') || '',
                        colors: colorInfo(el),
                        outerHTML: (el.outerHTML || '').slice(0, 900)
                    };
                };

                const labels = Array.from(body.querySelectorAll('body *')).filter(el =>
                    visible(el) && el.children.length === 0 && norm(el.textContent) === 'keyskills'
                );
                const diagnostics = {
                    labelCount: labels.length,
                    labels: [],
                    inspectedItems: [],
                    visibleTextHasJobMatchScore: norm(body.innerText).includes('job match score')
                };
                if (!labels.length) {
                    return {status: 'unknown', reason: 'Keyskills label not found', diagnostics};
                }

                for (const label of labels.slice(0, 5)) {
                    const labelInfo = describe(label);
                    let node = label;
                    const ancestors = [];
                    for (let depth = 0; node && depth < 5; depth++, node = node.parentElement) {
                        const nodeText = norm(node.innerText || node.textContent);
                        ancestors.push(describe(node));
                        // Restrict inspection to the small Keyskills label item.
                        // Do not inspect the whole score card, where other green checks exist.
                        if (!nodeText.includes('keyskills') || nodeText.length > 80) continue;
                        const candidates = [node, ...Array.from(node.querySelectorAll('svg, path, i, span, div'))].slice(0, 30);
                        const greenCandidates = candidates.filter(isGreenElement);
                        diagnostics.inspectedItems.push({
                            item: describe(node),
                            candidateCount: candidates.length,
                            greenCandidateCount: greenCandidates.length,
                            candidates: candidates.slice(0, 15).map(describe)
                        });
                        if (greenCandidates.length) {
                            diagnostics.labels.push({label: labelInfo, ancestors});
                            return {
                                status: 'green',
                                reason: 'green indicator found in the Keyskills label item',
                                diagnostics
                            };
                        }
                    }
                    diagnostics.labels.push({label: labelInfo, ancestors});
                }
                return {
                    status: 'not_green',
                    reason: 'Keyskills label found but no green indicator found nearby',
                    diagnostics
                };
            }"""
        )
        status = (result or {}).get("status")
        reason = (result or {}).get("reason", "")
        diagnostics = (result or {}).get("diagnostics", {})
        log(
            f"[KEYSKILLS] Detection result: status={status!r}; reason={reason}; "
            f"label_count={diagnostics.get('labelCount', 'n/a')}; "
            f"job_match_score_text_found={diagnostics.get('visibleTextHasJobMatchScore', 'n/a')}"
        )

        # Log DOM evidence in a readable form. Truncate the number of entries so
        # a changed page layout does not create an excessively large log.
        for label_info in diagnostics.get("labels", [])[:3]:
            log(f"[KEYSKILLS][DIAGNOSTIC] Label/ancestor evidence: {json.dumps(label_info, ensure_ascii=False)}")
        for item_info in diagnostics.get("inspectedItems", [])[:5]:
            log(f"[KEYSKILLS][DIAGNOSTIC] Inspected Keyskills item: {json.dumps(item_info, ensure_ascii=False)}")

        if status == "green":
            log("[KEYSKILLS] PASS: indicator detected as green; job is eligible for the existing application flow.")
            return True
        if status == "not_green":
            log("[KEYSKILLS] FAIL-CLOSED: Keyskills label found, but no green indicator was detected; job will be skipped.")
            return False

        log("[KEYSKILLS] UNKNOWN: unable to identify the Keyskills label; no application action will be taken.")
        # Preserve diagnostic artifacts only when detection is unknown, so the
        # next run can be compared against the actual page layout.
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            dom_path = BASE_DIR / f"naukri_keyskills_unknown_{timestamp}.html"
            dom_path.write_text(page.content(), encoding="utf-8")
            log(f"[KEYSKILLS] Unknown-state DOM saved to: {dom_path.name}")
        except Exception as artifact_exc:
            log(f"[KEYSKILLS] Could not save unknown-state DOM: {artifact_exc}")
        try:
            screenshot_path = BASE_DIR / f"naukri_keyskills_unknown_{timestamp}.png"
            page.screenshot(path=str(screenshot_path), full_page=True)
            log(f"[KEYSKILLS] Unknown-state screenshot saved to: {screenshot_path.name}")
        except Exception as artifact_exc:
            log(f"[KEYSKILLS] Could not save unknown-state screenshot: {artifact_exc}")
        return None
    except Exception as exc:
        log(f"[KEYSKILLS] ERROR while inspecting indicator: {type(exc).__name__}: {exc}")
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            dom_path = BASE_DIR / f"naukri_keyskills_error_{timestamp}.html"
            dom_path.write_text(page.content(), encoding="utf-8")
            log(f"[KEYSKILLS] Error-state DOM saved to: {dom_path.name}")
        except Exception as artifact_exc:
            log(f"[KEYSKILLS] Could not save error-state DOM: {artifact_exc}")
        return None


def find_applied_state(page: Page) -> bool:
    """
    Return True if the page clearly indicates that the job
    has already been applied to.
    """
    applied_pattern = re.compile(
        r"^\s*Applied\s*$",
        re.IGNORECASE,
    )

    candidate = find_visible_button(page, applied_pattern)

    if candidate is not None:
        return True

    # Some Naukri layouts use non-button elements for the status.
    try:
        exact_applied = page.get_by_text(
            re.compile(r"^\s*Applied\s*$", re.IGNORECASE)
        )

        for index in range(min(exact_applied.count(), 10)):
            element = exact_applied.nth(index)

            if element.is_visible():
                return True

    except Exception:
        pass

    return False


def find_external_apply_link(page: Page) -> Optional[str]:
    """
    Detect 'Apply on company site' and resolve its destination.

    Some Naukri layouts expose a direct href; others use a button that
    opens a new tab or navigates the current tab. In the latter case,
    click only the external-application button and capture the resulting
    URL. This does not fill or submit the company application form.
    """
    pattern = re.compile(
        r"apply\s+on\s+(the\s+)?company\s+site",
        re.IGNORECASE,
    )

    candidates = [
        page.get_by_role("link", name=pattern),
        page.get_by_role("button", name=pattern),
        page.locator("a").filter(has_text=pattern),
        page.locator("button").filter(has_text=pattern),
        page.locator('[role="button"]').filter(has_text=pattern),
    ]

    element = None
    for locator in candidates:
        try:
            for index in range(min(locator.count(), 20)):
                candidate = locator.nth(index)
                if candidate.is_visible():
                    element = candidate
                    break
            if element is not None:
                break
        except Exception:
            continue

    if element is None:
        log("No candidate matched the accessible-name/text selectors for 'Apply on company site'.")
        return None

    # First try direct links and common Naukri redirect attributes.
    try:
        direct_url = element.evaluate(
            """el => {
                const attrs = [
                    'href', 'data-href', 'data-url',
                    'data-redirect-url', 'data-apply-url',
                    'data-company-url'
                ];
                for (const attr of attrs) {
                    const value = el.getAttribute(attr);
                    if (value && value.trim()) return value.trim();
                }
                const anchor = el.closest('a') || el.querySelector('a');
                if (anchor && anchor.href) return anchor.href;
                const onclick = el.getAttribute('onclick') || '';
                const match = onclick.match(/https?:[^'"]+/i);
                if (match) return match[0];
                return '';
            }"""
        )
        if direct_url:
            if direct_url.startswith("//"):
                direct_url = "https:" + direct_url
            if direct_url.startswith("/"):
                from urllib.parse import urljoin
                direct_url = urljoin(page.url, direct_url)
            if is_valid_http_url(direct_url):
                log(f"External application URL found in button/link: {direct_url}")
                return direct_url
    except Exception as exc:
        log(f"Could not read external button URL attributes: {exc}")

    # Some buttons have no href. Click it and capture a popup or redirect.
    original_url = page.url
    popup_page = None

    try:
        with page.expect_popup(timeout=5000) as popup_info:
            element.click(timeout=10000)
        popup_page = popup_info.value
        try:
            popup_page.wait_for_load_state("domcontentloaded", timeout=10000)
        except PlaywrightTimeoutError:
            pass
        popup_url = popup_page.url
        log(f"Captured external application popup URL: {popup_url}")
        # Give Naukri a few seconds to register the external link opening.
        page.wait_for_timeout(3000)
        if is_valid_http_url(popup_url) and popup_url != original_url:
            try:
                popup_page.close()
            except Exception:
                pass
            return popup_url
    except PlaywrightTimeoutError:
        # No popup may mean the button navigated the current page instead.
        pass
    except Exception as exc:
        log(f"External application button click/capture failed: {exc}")

    try:
        page.wait_for_timeout(1500)
        current_url = page.url
        if (
                current_url != original_url
                and is_valid_http_url(current_url)
                and not is_naukri_job_url(current_url)
        ):
            log(f"Captured external application redirect URL: {current_url}")
            return current_url
    except Exception:
        pass
    finally:
        if popup_page is not None:
            try:
                popup_page.close()
            except Exception:
                pass

    log(
        "The 'Apply on company site' button was visible, but its destination "
        "URL could not be captured. The job will be kept for review."
    )
    return None

def find_apply_button(page: Page):
    """Find the normal Naukri Apply button, excluding 'Applied'."""
    apply_pattern = re.compile(
        r"^\s*Apply(\s+now)?\s*$",
        re.IGNORECASE,
    )

    candidates = [
        page.get_by_role("button", name=apply_pattern),
        page.get_by_role("link", name=apply_pattern),
        page.locator("button").filter(has_text=apply_pattern),
        page.locator('[role="button"]').filter(has_text=apply_pattern),
    ]

    for locator in candidates:
        try:
            for index in range(min(locator.count(), 20)):
                element = locator.nth(index)

                if element.is_visible() and element.is_enabled():
                    text = normalize_text(element.inner_text())

                    if re.fullmatch(
                            r"Apply(\s+now)?",
                            text,
                            flags=re.IGNORECASE,
                    ):
                        return element

        except Exception:
            continue

    return None


# ============================================================
# APPLICATION PANEL DETECTION
# ============================================================

def find_visible_application_panel(page: Page):
    """
    Try to locate the application drawer/modal.

    Important: this does not consider arbitrary page inputs sufficient
    evidence that the application panel has loaded.
    """
    panel_selectors = [
        '[role="dialog"]',
        '[aria-modal="true"]',
        '[class*="sidePanel"]',
        '[class*="side-panel"]',
        '[class*="drawer"]',
        '[class*="modal"]',
        '[class*="applyForm"]',
        '[class*="apply-form"]',
    ]

    for selector in panel_selectors:
        try:
            locator = page.locator(selector)
            count = min(locator.count(), 20)

            for index in range(count):
                panel = locator.nth(index)

                if not panel.is_visible():
                    continue

                try:
                    text = normalize_text(panel.inner_text(timeout=1500))
                except Exception:
                    text = ""

                # A panel should contain application-related content.
                if re.search(
                        r"save|submit|response|resume|question|"
                        r"application|thank you|relocat|notice period",
                        text,
                        re.IGNORECASE,
                ):
                    return panel

        except Exception:
            continue

    return None


def find_save_button(page: Page, panel=None):
    """
    Find a visible Save button inside the application panel.

    If a panel was found, prioritize buttons inside it.
    This function only detects the button; it does not click it.
    """
    save_pattern = re.compile(r"^\s*Save\s*$", re.IGNORECASE)

    search_roots = []

    if panel is not None:
        search_roots.append(panel)

    search_roots.append(page)

    for root in search_roots:
        try:
            locator = root.get_by_role(
                "button",
                name=save_pattern,
            )

            for index in range(min(locator.count(), 20)):
                candidate = locator.nth(index)

                if candidate.is_visible():
                    return candidate

        except Exception:
            pass

        try:
            locator = root.locator("button").filter(
                has_text=save_pattern
            )

            for index in range(min(locator.count(), 20)):
                candidate = locator.nth(index)

                if not candidate.is_visible():
                    continue

                button_text = normalize_text(candidate.inner_text())

                if re.fullmatch(
                        r"Save",
                        button_text,
                        re.IGNORECASE,
                ):
                    return candidate

        except Exception:
            pass

    return None


def print_panel_diagnostics(page: Page, panel=None) -> None:
    """
    Print useful diagnostics when the application panel is open
    but its Save button is not visible.
    """
    log("----- Application panel diagnostics -----")

    try:
        log(f"Current page URL: {page.url}")
    except Exception:
        pass

    try:
        log(f"Browser viewport: {page.evaluate('() => ({width: window.innerWidth, height: window.innerHeight})')}")
    except Exception:
        pass

    if panel is not None:
        try:
            text = normalize_text(panel.inner_text(timeout=3000))
            log(f"Panel text: {text[:1800]}")
        except Exception:
            log("Could not read panel text.")
    else:
        log("No recognized application panel was detected.")

    save_button = find_save_button(page, panel)

    if save_button is not None:
        log("A visible Save button was detected.")
    else:
        log("No visible Save button was detected.")

        try:
            page.screenshot(
                path=str(
                    BASE_DIR / "naukri_application_panel_debug.png"
                ),
                full_page=True,
            )
            log(
                "Diagnostic screenshot saved as "
                "'naukri_application_panel_debug.png'."
            )
        except Exception as exc:
            log(f"Could not save diagnostic screenshot: {exc}")

    log("-----------------------------------------")


# ============================================================
# APPLICATION QUESTION EXTRACTION
# ============================================================

def collect_visible_questions(page: Page, panel=None) -> list[str]:
    """
    Collect visible application question text.

    Questions are printed only. The script does not select answers,
    fill responses, upload a resume, or click Save/Submit.
    """
    questions = []

    roots = []

    if panel is not None:
        roots.append(panel)

    # Fallback to the whole page only if no panel was found.
    if not roots:
        roots.append(page.locator("body"))

    question_patterns = [
        r"\?",
        r"are you currently",
        r"do you have",
        r"how many years",
        r"what is your",
        r"what are your",
        r"when can you",
        r"notice period",
        r"current ctc",
        r"expected ctc",
        r"willing to relocate",
        r"currently residing",
        r"upload.*resume",
        r"upload.*cv",
    ]

    compiled_patterns = [
        re.compile(pattern, re.IGNORECASE)
        for pattern in question_patterns
    ]

    for root in roots:
        try:
            text = root.inner_text(timeout=3000)
        except Exception:
            continue

        for raw_line in text.splitlines():
            line = normalize_text(raw_line)

            if len(line) < 8 or len(line) > 400:
                continue

            if any(pattern.search(line) for pattern in compiled_patterns):
                if line not in questions:
                    questions.append(line)

    return questions


def print_application_questions(page: Page, panel=None) -> None:
    """Print the questions currently visible in the application panel."""
    questions = collect_visible_questions(page, panel)

    log("")
    log("========== APPLICATION QUESTIONS ==========")

    if questions:
        for index, question in enumerate(questions, start=1):
            print(f"{index}. {question}", flush=True)
    else:
        log(
            "No question text was detected automatically. "
            "Please inspect the browser panel."
        )

    log("============================================")
    log(
        "Answer the questions and upload any required files "
        "directly in the browser."
    )
    log(
        "The script will not fill answers or click Save/Submit."
    )
    log("")


# ============================================================
# APPLICATION COMPLETION
# ============================================================

def confirmation_visible(page: Page) -> bool:
    """
    Detect only Naukri's final applied confirmation.

    "Thank you for your responses" is an intermediate screen and must
    NOT be treated as completion because Naukri may still be saving the
    application. Keep the tab open until the final "Applied to ..."
    confirmation is visible.
    """
    text = visible_text(page)

    final_confirmation_pattern = r"applied\s+to\s+[\"“][^\"”\n]+[\"”]"
    return bool(re.search(final_confirmation_pattern, text, re.IGNORECASE))


def wait_for_manual_application_completion(
        page: Page,
        panel=None,
) -> bool:
    """
    Wait while the user completes the application manually.

    The script periodically checks for the success confirmation
    and prints newly visible questions when they change.

    Returns True only after the success phrase is detected.
    """
    deadline = time.time() + (
            APPLICATION_TIMEOUT_MINUTES * 60
    )

    last_questions = None
    last_reminder = 0.0

    log(
        f"Waiting up to {APPLICATION_TIMEOUT_MINUTES} minutes "
        "for the application confirmation."
    )
    log(
        "Complete the form in the browser. When finished, the page "
        "should display the final 'Applied to \"Job Title\"' confirmation. The intermediate 'Thank you for your responses' screen does not count."
    )

    while time.time() < deadline:
        try:
            if page.is_closed():
                log(
                    "The application tab was closed before confirmation. "
                    "The job will remain in the main Excel."
                )
                return False
        except Exception:
            return False

        if confirmation_visible(page):
            log("Final Naukri 'Applied to ...' confirmation detected.")
            return True

        current_panel = find_visible_application_panel(page)

        if current_panel is not None:
            panel = current_panel

        questions = collect_visible_questions(page, panel)
        questions_signature = tuple(questions)

        if questions and questions_signature != last_questions:
            log("Application questions visible now:")
            for index, question in enumerate(questions, start=1):
                print(f"{index}. {question}", flush=True)
            last_questions = questions_signature

        # Give a periodic reminder without flooding the terminal.
        if time.time() - last_reminder >= 60:
            log(
                "Still waiting for completion. Continue in the browser; "
                "do not enter answers in the terminal."
            )
            last_reminder = time.time()

        page.wait_for_timeout(1000)

    log(
        "Timed out waiting for the success confirmation. "
        "The job will remain in the main Excel."
    )
    return False


# ============================================================
# PROCESS ONE JOB
# ============================================================

def process_one_job(
        context: BrowserContext,
        job: dict,
) -> str:
    """
    Process one job in a new tab.

    Possible outcomes:
      APPLIED_ALREADY
      EXTERNAL_APPLICATION (saved to separate external workbook)
      APPLICATION_COMPLETED
      KEEP
    """
    company = job["company"]
    title = job["title"]
    url = job["url"]

    log("")
    log("=" * 70)
    log(f"Company: {company}")
    log(f"Job title: {title}")
    log(f"Job URL: {url}")
    log("=" * 70)

    if not is_valid_http_url(url):
        log("Invalid or empty job URL. Keeping this row.")
        return "KEEP"

    page = None

    try:
        # The original tab is never navigated or reused.
        page = context.new_page()

        log("Opening job in a new tab...")
        page.goto(
            url,
            wait_until="domcontentloaded",
            timeout=60000,
        )

        page.wait_for_timeout(JOB_PAGE_LOAD_WAIT_MS)

        # Wait for a typical Naukri page to settle.
        try:
            page.wait_for_load_state(
                "networkidle",
                timeout=7000,
            )
        except PlaywrightTimeoutError:
            # Naukri may keep network requests running.
            pass

        page.wait_for_timeout(1000)

        if not is_naukri_job_url(page.url):
            log(
                "WARNING: The page redirected away from Naukri. "
                "Inspect the tab manually. Keeping the job row."
            )
            return "KEEP"

        # Case 1: Already applied
        if find_applied_state(page):
            log("This job already shows 'Applied'.")
            return "APPLIED_ALREADY"

        # Check the Keyskills match indicator before any application action.
        # A non-green indicator means this job is discarded without applying
        # or saving an external application URL. Unknown detection is fail-safe.
        keyskills_green = get_keyskills_green_status(page)
        if keyskills_green is False:
            log(
                "Keyskills match indicator is NOT green. "
                "Skipping application and removing this job from the main Excel."
            )
            return "KEYSKILLS_NOT_GREEN"
        if keyskills_green is None:
            log(
                "WARNING: Could not confidently determine the Keyskills indicator color. "
                "No application action will be taken; keeping the job for review."
            )
            return "KEEP"

        log("Keyskills match indicator is green; eligible to continue.")

        # Case 2: Apply on company site
        external_url = find_external_apply_link(page)

        if external_url:
            if not is_valid_http_url(external_url):
                log(
                    "External application was detected, but its URL "
                    "could not be validated. Keeping the job row."
                )
                return "KEEP"

            log("Detected 'Apply on company site'.")
            log(f"Company application URL: {external_url}")

            saved = append_external_application(
                company=company,
                title=title,
                naukri_url=url,
                company_url=external_url,
            )

            if saved:
                return "EXTERNAL_APPLICATION"

            log(
                "Could not safely save the external application "
                "to its separate Excel. Keeping the job row in the main Excel."
            )
            return "KEEP"

        # Case 3: Normal Apply
        apply_button = find_apply_button(page)

        if apply_button is None:
            log(
                "Could not identify 'Applied', 'Apply on company site', "
                "or a normal 'Apply' button."
            )
            log("Capturing the live DOM and visible interactive elements before closing the tab.")
            dump_button_diagnostics(page, "all_apply_buttons_not_found")
            log("Keeping this job in the main Excel.")
            return "KEEP"

        log("Clicking the normal Apply button...")

        try:
            apply_button.click(timeout=10000)
        except Exception as exc:
            log(f"Could not click Apply: {exc}")
            return "KEEP"

        # Wait for the actual application panel, not arbitrary inputs.
        panel = None
        panel_deadline = time.time() + 15

        while time.time() < panel_deadline:
            if confirmation_visible(page):
                log("Final Naukri 'Applied to ...' confirmation detected immediately.")
                return "APPLICATION_COMPLETED"

            panel = find_visible_application_panel(page)

            if panel is not None:
                break

            page.wait_for_timeout(500)

        # Allow panel content and footer to render.
        page.wait_for_timeout(APPLICATION_PANEL_SETTLE_WAIT_MS)

        # Re-check after waiting for the panel.
        panel = find_visible_application_panel(page)

        if confirmation_visible(page):
            log("Final Naukri 'Applied to ...' confirmation detected.")
            return "APPLICATION_COMPLETED"

        if panel is None:
            log(
                "WARNING: The application panel was not detected by the selectors. "
                "The tab will remain open and the script will keep checking for "
                "the success confirmation for up to 30 minutes."
            )
            print_panel_diagnostics(page, panel)
            print_application_questions(page, panel)
            completed = wait_for_manual_application_completion(page, panel)
            return "APPLICATION_COMPLETED" if completed else "KEEP"

        save_button = find_save_button(page, panel)

        if save_button is not None:
            log(
                "A visible Save button was detected. "
                "You must click it yourself after answering the questions."
            )
        else:
            log(
                "No visible Save button was detected yet. "
                "Please inspect the application panel in the browser."
            )

        # Print the current questions for the user.
        print_application_questions(page, panel)

        # Print panel details if the Save button is missing.
        if save_button is None:
            print_panel_diagnostics(page, panel)

        # Wait for the user to complete the application in the browser.
        completed = wait_for_manual_application_completion(
            page,
            panel,
        )

        if completed:
            return "APPLICATION_COMPLETED"

        return "KEEP"

    except PlaywrightTimeoutError as exc:
        log(f"Timeout while processing job: {exc}")
        return "KEEP"

    except Exception as exc:
        log(f"Unexpected error while processing job: {exc}")
        return "KEEP"

    finally:
        # Close only the tab created for this job.
        # The original tab remains untouched.
        if page is not None:
            try:
                if not page.is_closed():
                    page.close()
                    log("Job tab closed.")
            except Exception as exc:
                log(f"Could not close job tab cleanly: {exc}")


# ============================================================
# CONFIRMED-OUTCOME LEDGER
# ============================================================

def load_confirmed_outcomes() -> dict[str, str]:
    """Load confirmed applications awaiting successful Excel row cleanup."""
    if not CONFIRMED_OUTCOMES_FILE.exists():
        return {}
    try:
        raw = json.loads(CONFIRMED_OUTCOMES_FILE.read_text(encoding="utf-8"))
        if not isinstance(raw, dict):
            raise ValueError("Expected a JSON object")
        return {
            normalize_text(url): normalize_text(outcome)
            for url, outcome in raw.items()
            if normalize_text(url) and normalize_text(outcome)
        }
    except Exception as exc:
        log(f"WARNING: Could not read confirmed-outcome ledger: {exc}. "
            "No ledger entries will be assumed.")
        return {}


def save_confirmed_outcomes(outcomes: dict[str, str]) -> bool:
    """Atomically persist confirmed outcomes so failed cleanup never triggers reapplication."""
    temporary_path = CONFIRMED_OUTCOMES_FILE.with_suffix(".json.tmp")
    try:
        temporary_path.write_text(
            json.dumps(outcomes, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        temporary_path.replace(CONFIRMED_OUTCOMES_FILE)
        return True
    except Exception as exc:
        log(f"ERROR: Could not persist confirmed-outcome ledger: {exc}")
        try:
            if temporary_path.exists():
                temporary_path.unlink()
        except Exception:
            pass
        return False


# ============================================================
# MAIN
# ============================================================

def main(job_urls: Optional[set[str]] = None, wait_for_enter: bool = True) -> None:
    log("Starting Naukri application processor...")
    log(f"[STARTUP] Running script: {Path(__file__).resolve()}")
    log(f"[STARTUP] Debug log: {(BASE_DIR / 'naukri_processor_debug.log').resolve()}")

    log(f"Project folder: {BASE_DIR}")
    log(f"Main Excel: {MAIN_EXCEL}")
    log(f"External applications Excel: {EXTERNAL_EXCEL}")
    log(f"Processed requests Excel: {PROCESSED_REQUESTS_EXCEL}")
    log(f"Browser profile: {PROFILE_DIR}")

    if not MAIN_EXCEL.exists():
        log("ERROR: Main Excel file was not found.")
        log("Please ensure the Excel file is beside this script.")
        return

    if not PROFILE_DIR.exists():
        log(
            "ERROR: The existing Naukri browser profile directory "
            "was not found."
        )
        log(
            "Expected profile directory: "
            f"{PROFILE_DIR}"
        )
        log(
            "Make sure the profile directory is beside this script. "
            "This script will not create a fresh login profile."
        )
        return

    try:
        jobs = load_main_jobs()
    except Exception as exc:
        log(f"ERROR reading main Excel: {exc}")
        return

    if job_urls is not None:
        requested_urls = {normalize_text(url) for url in job_urls if normalize_text(url)}
        jobs = [job for job in jobs if normalize_text(job.get("url")) in requested_urls]
        log(
            f"[BATCH] Page-scoped processing enabled: requested_urls={len(requested_urls)}, "
            f"matching_excel_jobs={len(jobs)}"
        )
    else:
        log("[BATCH] No URL filter supplied; processing all jobs currently in the main Excel.")

    if not jobs:
        log("No matching jobs found in the main Excel for this batch. Nothing to process.")
        return

    log(f"Found {len(jobs)} job(s) to process in this batch.")
    confirmed_outcomes = load_confirmed_outcomes()
    if confirmed_outcomes:
        log(
            f"Loaded {len(confirmed_outcomes)} confirmed application(s) "
            "awaiting Excel cleanup; these will NOT be applied again."
        )
    log(
        "IMPORTANT: Close other Chrome/Playwright instances that are "
        "using this same profile before starting the script."
    )

    completed_count = 0
    external_saved_count = 0
    cleanup_failure_count = 0
    failed_count = 0
    kept_count = 0
    keyskills_skipped_count = 0
    keyskills_skip_cleanup_failures = 0

    try:
        with sync_playwright() as playwright:
            # IMPORTANT FIX:
            # no_viewport=True avoids forcing a small fixed viewport,
            # allowing the browser window to use its actual dimensions.
            context = playwright.chromium.launch_persistent_context(
                user_data_dir=str(PROFILE_DIR),
                headless=False,
                no_viewport=True,
            )

            # Preserve the existing original tab, if one exists.
            # If no page exists, create a blank original tab.
            if context.pages:
                original_page = context.pages[0]
                log(
                    "Original tab preserved. "
                    f"Current URL: {original_page.url}"
                )
            else:
                original_page = context.new_page()
                original_page.goto("about:blank")
                log("Created a blank original tab.")

            # Process a snapshot of the job list. Each job is removed
            # by URL only after its outcome has been safely recorded.
            for index, job in enumerate(jobs, start=1):
                log("")
                log(f"Processing job {index} of {len(jobs)}")

                saved_outcome = confirmed_outcomes.get(job["url"])
                if saved_outcome:
                    log(
                        "This URL already has a confirmed application outcome "
                        f"({saved_outcome}) from a previous run. Retrying Excel cleanup "
                        "only; the application flow will not be opened again."
                    )
                    processing_type = (
                        "EXTERNAL_APPLICATION"
                        if saved_outcome == "EXTERNAL_APPLICATION"
                        else "NORMAL_APPLY"
                    )
                    processed_saved = append_processed_request(
                        company=job.get("company", ""),
                        title=job.get("title", ""),
                        naukri_url=job["url"],
                        processing_type=processing_type,
                    )
                    if not processed_saved:
                        cleanup_failure_count += 1
                        log(
                            "WARNING: Confirmed outcome is in the recovery ledger, but "
                            "processed_requests.xlsx could not be updated. Main Excel row "
                            "will be retained for a safe retry."
                        )
                        continue

                    processing_type = (
                        "EXTERNAL_APPLICATION"
                        if outcome == "EXTERNAL_APPLICATION"
                        else "NORMAL_APPLY"
                    )
                    processed_saved = append_processed_request(
                        company=job.get("company", ""),
                        title=job.get("title", ""),
                        naukri_url=job["url"],
                        processing_type=processing_type,
                    )
                    if not processed_saved:
                        cleanup_failure_count += 1
                        log(
                            "WARNING: Application outcome is confirmed and protected by "
                            "the recovery ledger, but processed_requests.xlsx could not be "
                            "updated. Main Excel row will be retained; no further application "
                            "action will be taken for this URL on the next run."
                        )
                        continue

                    removed = remove_job_from_main_excel(job["url"])
                    if removed:
                        completed_count += 1
                        confirmed_outcomes.pop(job["url"], None)
                        save_confirmed_outcomes(confirmed_outcomes)
                        log("SUCCESS: Processed URL recorded and main Excel cleanup verified.")
                    else:
                        cleanup_failure_count += 1
                        log("WARNING: Excel cleanup still failed. Confirmed outcome remains "
                            "in the ledger for a cleanup-only retry on the next run.")
                    continue

                outcome = process_one_job(context, job)

                if outcome == "KEYSKILLS_NOT_GREEN":
                    keyskills_skipped_count += 1
                    removed = remove_job_from_main_excel(job["url"])
                    if removed:
                        log(
                            "SKIPPED: Keyskills was not green. Job row removed from the main Excel; "
                            "no application was submitted and no external URL was saved."
                        )
                    else:
                        keyskills_skip_cleanup_failures += 1
                        log(
                            "WARNING: Keyskills was not green, so no application was submitted, "
                            "but removing the job row from Excel failed. It remains for cleanup."
                        )
                    time.sleep(1.5)
                    continue

                if outcome in {
                    "EXTERNAL_APPLICATION",
                    "APPLIED_ALREADY",
                    "APPLICATION_COMPLETED",
                }:
                    # For external applications, the separate workbook is saved
                    # inside process_one_job before this removal is attempted.
                    if outcome == "EXTERNAL_APPLICATION":
                        external_saved_count += 1
                        log("External application URL was saved to the separate workbook.")
                    else:
                        log(f"Application outcome confirmed: {outcome}")

                    # Persist the confirmed outcome before touching the main workbook.
                    # If cleanup fails or the process stops, the next run retries cleanup
                    # only instead of submitting the application a second time.
                    confirmed_outcomes[job["url"]] = outcome
                    ledger_saved = save_confirmed_outcomes(confirmed_outcomes)
                    if not ledger_saved:
                        log(
                            "WARNING: Could not persist the confirmed-outcome ledger. "
                            "Excel cleanup will still be attempted, but duplicate-application "
                            "protection cannot be guaranteed if cleanup also fails."
                        )

                    removed = remove_job_from_main_excel(job["url"])
                    if removed:
                        completed_count += 1
                        confirmed_outcomes.pop(job["url"], None)
                        save_confirmed_outcomes(confirmed_outcomes)
                        log(
                            "SUCCESS: Application outcome confirmed and main Excel "
                            "row cleanup verified."
                        )
                    else:
                        cleanup_failure_count += 1
                        log(
                            "WARNING: Application outcome is confirmed, but Excel cleanup "
                            "failed. This remains an application success; the URL is recorded "
                            "for cleanup-only retry and will not be applied again."
                        )

                else:
                    kept_count += 1
                    log(
                        "Job not confirmed as successfully processed. "
                        "Its row will remain in the main Excel for review/retry."
                    )

                # Small delay between jobs to avoid rapid navigation.
                time.sleep(1.5)

            log("")
            log("=" * 70)
            log("PROCESSING SUMMARY")
            log("=" * 70)
            log(f"Confirmed application outcomes with row cleanup verified: {completed_count}")
            log(f"External application URLs saved: {external_saved_count}")
            log(f"Jobs skipped because Keyskills was not green: {keyskills_skipped_count}")
            log(f"Keyskills-skip Excel cleanup failures: {keyskills_skip_cleanup_failures}")
            log(f"Jobs kept for review/retry (application not confirmed): {kept_count}")
            log(f"Excel cleanup failures (application outcome remains confirmed): {cleanup_failure_count}")
            log(f"Confirmed outcomes awaiting cleanup retry: {len(confirmed_outcomes)}")
            log(f"Jobs originally loaded: {len(jobs)}")
            log(f"Main Excel: {MAIN_EXCEL}")
            log(f"External applications Excel: {EXTERNAL_EXCEL}")
            log(f"Processed requests Excel: {PROCESSED_REQUESTS_EXCEL}")
            log("=" * 70)

            if wait_for_enter:
                input(
                    "\nProcessing finished. Press Enter to close the "
                    "automation browser, or Ctrl+C to stop."
                )

            # Closing the persistent context closes its browser window.
            # The original tab was never navigated by this script.
            context.close()

    except KeyboardInterrupt:
        log("Stopped by user.")
        # When orchestrated by naukri_login.py, propagate Ctrl+C so the
        # page-by-page loop stops instead of starting the next page.
        if not wait_for_enter:
            raise

    except Exception as exc:
        log(f"FATAL ERROR: {exc}")
        log(
            "If the browser profile is already in use, close the other "
            "browser instance and try again."
        )
        if not wait_for_enter:
            raise


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process Naukri job applications.")
    parser.add_argument(
        "--batch-file",
        type=Path,
        help="JSON file containing a list of Job URLs to process in this batch.",
    )
    args = parser.parse_args()

    batch_urls = None
    if args.batch_file:
        try:
            payload = json.loads(args.batch_file.read_text(encoding="utf-8"))
            if not isinstance(payload, list) or not all(isinstance(item, str) for item in payload):
                raise ValueError("Batch file must contain a JSON list of URL strings.")
            batch_urls = set(payload)
        except Exception as exc:
            log(f"FATAL ERROR: Could not read batch file {args.batch_file}: {exc}")
            sys.exit(2)

    try:
        main(job_urls=batch_urls, wait_for_enter=False if args.batch_file else True)
    except KeyboardInterrupt:
        log("Stopped by user.")
        sys.exit(130)
