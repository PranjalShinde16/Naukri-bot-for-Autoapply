from __future__ import annotations

import re
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
MAIN_EXCEL = BASE_DIR / "naukri_software_jobs_first_page.xlsx"

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

# Additional columns stored in the same main Excel workbook
STATUS_COLUMN = "Processing Status"
EXTERNAL_URL_COLUMN = "Company Application URL"


# ============================================================
# GENERAL HELPERS
# ============================================================

def log(message: str) -> None:
    """Print a timestamped message to the terminal."""
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] {message}", flush=True)


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

        # Ignore completely empty rows.
        if not company and not title and not url:
            continue

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


def remove_job_from_main_excel(job_url: str) -> bool:
    """Remove a successfully processed job from the single main Excel."""
    if not MAIN_EXCEL.exists():
        log("ERROR: Main Excel does not exist. Cannot remove a row.")
        return False

    job_url = normalize_text(job_url)
    if not job_url:
        log("ERROR: Empty job URL. Refusing to remove a row.")
        return False

    workbook = None
    try:
        workbook = load_workbook(MAIN_EXCEL)
        worksheet = workbook.active
        headers = [normalize_text(cell.value) for cell in worksheet[1]]
        url_column_index = next(
            (i for i, header in enumerate(headers, start=1)
             if header.lower() == URL_COLUMN.lower()),
            None,
        )
        if url_column_index is None:
            raise ValueError(f"Column '{URL_COLUMN}' was not found in the main Excel.")

        matching_rows = [
            row_number
            for row_number in range(2, worksheet.max_row + 1)
            if normalize_text(worksheet.cell(row_number, url_column_index).value) == job_url
        ]
        if not matching_rows:
            log("WARNING: Job URL was not found. No row was removed.")
            return False

        for row_number in reversed(matching_rows):
            worksheet.delete_rows(row_number, 1)

        # Save directly to the existing workbook: no backup or extra Excel files.
        workbook.save(MAIN_EXCEL)
        log(f"Removed {len(matching_rows)} row(s) from the main Excel.")
        return True
    except Exception as exc:
        log(f"ERROR removing job from main Excel: {exc}")
        return False
    finally:
        if workbook is not None:
            try:
                workbook.close()
            except Exception:
                pass

def append_external_application(
    company: str,
    title: str,
    naukri_url: str,
    company_url: str,
) -> bool:
    """Record an external application URL in the same main Excel row."""
    workbook = None
    try:
        workbook = load_workbook(MAIN_EXCEL)
        worksheet = workbook.active
        headers = [normalize_text(cell.value) for cell in worksheet[1]]
        header_lookup = {header.lower(): i for i, header in enumerate(headers) if header}

        # Add the two tracking columns to the existing workbook only when missing.
        if STATUS_COLUMN.lower() not in header_lookup:
            worksheet.cell(1, worksheet.max_column + 1, STATUS_COLUMN)
            header_lookup[STATUS_COLUMN.lower()] = worksheet.max_column - 1
        if EXTERNAL_URL_COLUMN.lower() not in header_lookup:
            worksheet.cell(1, worksheet.max_column + 1, EXTERNAL_URL_COLUMN)
            header_lookup[EXTERNAL_URL_COLUMN.lower()] = worksheet.max_column - 1

        url_index = header_lookup.get(URL_COLUMN.lower())
        if url_index is None:
            raise ValueError(f"Column '{URL_COLUMN}' was not found in the main Excel.")

        matching_row = None
        for row_number in range(2, worksheet.max_row + 1):
            if normalize_text(worksheet.cell(row_number, url_index + 1).value) == naukri_url:
                matching_row = row_number
                break

        if matching_row is None:
            log("ERROR: Could not find external-application job in the main Excel.")
            return False

        worksheet.cell(matching_row, header_lookup[STATUS_COLUMN.lower()] + 1).value = "Apply on company site"
        worksheet.cell(matching_row, header_lookup[EXTERNAL_URL_COLUMN.lower()] + 1).value = company_url
        workbook.save(MAIN_EXCEL)
        log("External application URL recorded in the main Excel; the row was kept.")
        return True
    except Exception as exc:
        log(f"ERROR recording external application in main Excel: {exc}")
        return False
    finally:
        if workbook is not None:
            try:
                workbook.close()
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
    Detect 'Apply on company site' and return its destination URL.

    The script does not automatically navigate to or submit the
    company website application.
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

    for locator in candidates:
        try:
            for index in range(min(locator.count(), 20)):
                element = locator.nth(index)

                if not element.is_visible():
                    continue

                href = element.get_attribute("href")

                if href:
                    try:
                        return element.evaluate(
                            "(element) => element.href"
                        )
                    except Exception:
                        return href

                # Some buttons may open an external page via JavaScript.
                # We cannot safely infer the destination if no href exists.

        except Exception:
            continue

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
    Detect the requested success confirmation.

    The exact phrase 'Thank you for your responses' is the main
    success signal. Other generic text is not enough to delete a row.
    """
    text = visible_text(page)

    patterns = [
        r"thank\s+you\s+for\s+your\s+responses",
        r"applied\s+to\s+[\"“][^\"”\n]+[\"”]",
    ]
    return any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns)


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
        "should display 'Thank you for your responses'."
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
            log("Application confirmation detected.")
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
      EXTERNAL_APPLICATION (recorded in main Excel, row retained)
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
                "to its Excel. Keeping the job row."
            )
            return "KEEP"

        # Case 3: Normal Apply
        apply_button = find_apply_button(page)

        if apply_button is None:
            log(
                "Could not identify 'Applied', 'Apply on company site', "
                "or a normal 'Apply' button."
            )
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
                log("Application confirmation detected immediately.")
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
            log("Application confirmation detected.")
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
# MAIN
# ============================================================

def main() -> None:
    log("Starting Naukri application processor...")

    log(f"Project folder: {BASE_DIR}")
    log(f"Main Excel: {MAIN_EXCEL}")
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

    if not jobs:
        log("No jobs found in the main Excel. Nothing to process.")
        return

    log(f"Found {len(jobs)} job(s) in the main Excel.")
    log(
        "IMPORTANT: Close other Chrome/Playwright instances that are "
        "using this same profile before starting the script."
    )

    completed_count = 0
    skipped_count = 0
    kept_count = 0

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

                outcome = process_one_job(context, job)

                if outcome == "EXTERNAL_APPLICATION":
                    # External applications stay in the same workbook with a status and URL.
                    completed_count += 1
                    log("External application recorded in the main Excel; row retained.")

                elif outcome in {"APPLIED_ALREADY", "APPLICATION_COMPLETED"}:
                    removed = remove_job_from_main_excel(job["url"])
                    if removed:
                        completed_count += 1
                        log(f"Job processed successfully: {outcome}")
                    else:
                        kept_count += 1
                        log("Outcome detected, but the row could not be removed safely.")

                else:
                    kept_count += 1
                    log(
                        "Job not confirmed as successfully processed. "
                        "Its row will remain in the main Excel."
                    )

                # Small delay between jobs to avoid rapid navigation.
                time.sleep(1.5)

            log("")
            log("=" * 70)
            log("PROCESSING SUMMARY")
            log("=" * 70)
            log(f"Jobs successfully processed: {completed_count}")
            log(f"Jobs kept for review/retry: {kept_count}")
            log(f"Jobs originally loaded: {len(jobs)}")
            log(f"Main Excel: {MAIN_EXCEL}")
            log("All job records are kept in the main Excel workbook.")
            log("=" * 70)

            input(
                "\nProcessing finished. Press Enter to close the "
                "automation browser, or Ctrl+C to stop."
            )

            # Closing the persistent context closes its browser window.
            # The original tab was never navigated by this script.
            context.close()

    except KeyboardInterrupt:
        log("Stopped by user.")

    except Exception as exc:
        log(f"FATAL ERROR: {exc}")
        log(
            "If the browser profile is already in use, close the other "
            "browser instance and try again."
        )


if __name__ == "__main__":
    main()