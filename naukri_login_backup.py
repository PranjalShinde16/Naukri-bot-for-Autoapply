
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import os
import re

load_dotenv()

EMAIL = os.getenv("NAUKRI_EMAIL")
PASSWORD = os.getenv("NAUKRI_PASSWORD")

HOME_URL = "https://www.naukri.com/mnjuser/homepage"
PROFILE_DIR = Path(__file__).resolve().parent / "naukri_browser_profile"

SALARY_RANGES = [
    "15-25 Lakhs",
    "25-50 Lakhs",
    "50-75 Lakhs",
    "75-100 Lakhs",
    "1-5 Cr",
    "5+ Cr",
    "10-15 Lakhs",
]


def visible(locator):
    try:
        return locator.is_visible()
    except Exception:
        return False


def login_if_required(page):
    email_field = page.locator("#usernameField")
    password_field = page.locator("#passwordField")

    if not (visible(email_field) and visible(password_field)):
        print("Login form not visible; reusing saved session.", flush=True)
        return

    if not EMAIL or not PASSWORD:
        raise RuntimeError("Login required, but credentials are missing in .env.")

    email_field.fill(EMAIL)
    password_field.fill(PASSWORD)

    button = page.locator("#loginButton")
    if visible(button):
        button.click()
    else:
        password_field.press("Enter")

    input("Complete any CAPTCHA/OTP, verify login, then press Enter...")

    if visible(email_field) and visible(password_field):
        raise RuntimeError("Login form is still visible.")


def search_software(page):
    print("Opening the search bar...", flush=True)

    page.locator("button.nI-gNb-sb__expand").wait_for(
        state="visible", timeout=15000
    )
    page.locator("button.nI-gNb-sb__expand").click()

    search_input = page.locator(
        'input[placeholder="Enter keyword / designation / companies"]'
    )
    search_input.wait_for(state="visible", timeout=10000)
    search_input.fill("software")

    print("Entered keyword: software", flush=True)

    page.get_by_role("button", name="Search", exact=True).click()
    print("Search submitted.", flush=True)

    page.wait_for_load_state("domcontentloaded", timeout=30000)
    page.wait_for_timeout(2000)

    print(f"Search results page: {page.url}", flush=True)


def click_salary_view_more(page):
    print("Finding Salary filter...", flush=True)

    # Find the Salary heading and the View More link in its section.
    salary_heading = page.get_by_text("Salary", exact=True).first
    salary_heading.wait_for(state="visible", timeout=15000)

    # Find the smallest ancestor containing the Salary heading and
    # a View More control. Inspect ancestors instead of selecting an
    # arbitrary View More elsewhere in the sidebar.
    section = salary_heading.locator("xpath=..")

    for _ in range(7):
        try:
            if (
                section.get_by_text("View More", exact=True).count() > 0
                and section.get_by_text("Lakhs", exact=False).count() > 0
            ):
                view_more = section.get_by_text("View More", exact=True).last

                if visible(view_more):
                    view_more.click()
                    print("Clicked View More under Salary.", flush=True)
                    page.wait_for_timeout(700)
                    return
        except Exception:
            pass

        section = section.locator("xpath=..")

    raise RuntimeError(
        "Could not locate View More within the Salary filter. "
        "No salary options were changed."
    )

def select_salary_ranges(page):
    import re

    salary_ranges = [
        ("15-25 Lakhs", r"15\s*-\s*25\s*Lakhs"),
        ("25-50 Lakhs", r"25\s*-\s*50\s*Lakhs"),
        ("50-75 Lakhs", r"50\s*-\s*75\s*Lakhs"),
        ("75-100 Lakhs", r"75\s*-\s*100\s*Lakhs"),
        ("1-5 Cr", r"1\s*-\s*5\s*Cr"),
        ("5+ Cr", r"5\s*\+\s*Cr"),
        ("10-15 Lakhs", r"10\s*-\s*15\s*Lakhs"),
    ]

    print("Locating the open Salary popup...", flush=True)

    apply_button = page.get_by_role(
        "button", name="Apply", exact=True
    )
    apply_button.wait_for(state="visible", timeout=10000)

    # Find the smallest visible container containing the Apply button,
    # Salary heading, and all seven requested options.
    popup = None

    for depth in range(1, 12):
        candidate = apply_button.locator(
            "xpath=" + "/".join([".."] * depth)
        )

        try:
            text = candidate.inner_text(timeout=2000)
            if (
                "Salary" in text
                and all(re.search(pattern, text, re.I)
                        for _, pattern in salary_ranges)
            ):
                popup = candidate
                break
        except Exception:
            continue

    if popup is None:
        raise RuntimeError(
            "Could not identify a single Salary popup containing "
            "all seven ranges. No selections were made."
        )

    # Verify the popup remains open and select options only within it.
    for name, pattern in salary_ranges:
        print(f"Selecting {name}...", flush=True)

        # Reacquire the popup and option after every UI update.
        apply_button.wait_for(state="visible", timeout=5000)

        option = popup.get_by_text(
            re.compile(pattern, re.I)
        ).last

        option.wait_for(state="visible", timeout=5000)

        # Prefer the actual checkbox if present.
        row = option.locator("xpath=..")
        checkbox = row.locator('input[type="checkbox"]')

        if checkbox.count() > 0:
            if not checkbox.first.is_checked():
                checkbox.first.check()
        else:
            option.click()

        # Do not proceed if the popup disappeared.
        if not apply_button.is_visible():
            raise RuntimeError(
                f"Salary popup closed while selecting {name}. "
                "Stopped without clicking Apply."
            )

        print(f"Selected {name}", flush=True)

    print("All seven ranges selected. Applying once...", flush=True)
    apply_button.click()
    page.wait_for_timeout(1000)

    print("Salary filter applied.", flush=True)

def select_freshness_last_1_day(page):
    print("Finding Freshness filter...", flush=True)

    freshness_heading = page.get_by_text("Freshness", exact=True).first
    freshness_heading.wait_for(state="visible", timeout=15000)

    def visible_exact_option():
        # Naukri may render dropdown menus outside the filter section's DOM
        # subtree, so search globally but click only a visible exact match.
        options = page.get_by_text("Last 1 day", exact=True)
        for index in range(options.count()):
            candidate = options.nth(index)
            if visible(candidate):
                return candidate
        return None

    option = visible_exact_option()

    if option is None:
        print("Freshness options are not open; opening the control...", flush=True)

        # Try the heading and nearby ancestors, stopping at the first
        # interaction that makes the requested option visible.
        control_candidates = [freshness_heading]
        ancestor = freshness_heading.locator("xpath=..")
        for _ in range(4):
            control_candidates.append(ancestor)
            ancestor = ancestor.locator("xpath=..")

        opened = False
        for control in control_candidates:
            try:
                if not visible(control):
                    continue
                control.click(timeout=1500)
                page.wait_for_timeout(350)
                option = visible_exact_option()
                if option is not None:
                    opened = True
                    break
            except Exception:
                continue

        if not opened and option is None:
            raise RuntimeError(
                "Could not open the Freshness dropdown or find a visible "
                "'Last 1 day' option. The browser is left open for inspection."
            )

    print("Selecting Freshness: Last 1 day...", flush=True)
    option.click()
    page.wait_for_timeout(1200)

    # Verify the selection by checking the URL for Naukri's freshness filter
    # parameter where present, or by confirming the selected label in the UI.
    from urllib.parse import urlparse, parse_qs
    query = parse_qs(urlparse(page.url).query)
    freshness_values = query.get("f", []) + query.get("freshness", [])

    selected_label = page.get_by_text("Last 1 day", exact=True)
    label_visible = any(
        visible(selected_label.nth(i))
        for i in range(selected_label.count())
    )

    if freshness_values and not any(
        value.lower() in {"1", "1day", "last1day", "last_1_day"}
        for value in freshness_values
    ):
        print(
            f"Warning: freshness query parameter is {freshness_values}; "
            "please verify the filter in the browser.",
            flush=True,
        )
    elif label_visible:
        print("Freshness option clicked; 'Last 1 day' is visible.", flush=True)
    else:
        print(
            "Freshness option clicked. Please verify the applied filter "
            "in the browser.",
            flush=True,
        )

def main():
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1366, "height": 900},
        )

        page = context.pages[0] if context.pages else context.new_page()

        try:
            print("\n=== Naukri Automation Bot ===", flush=True)

            page.goto(
                HOME_URL,
                wait_until="domcontentloaded",
                timeout=60000,
            )
            print(f"Opened: {page.url}", flush=True)

            login_if_required(page)

            if "/software-jobs" not in page.url:
                search_software(page)

            click_salary_view_more(page)
            select_salary_ranges(page)
            select_freshness_last_1_day(page)

            input(
                "\nVerify the applied salary filters in the browser. "
                "Press Enter to close..."
            )

        except Exception as error:
            print(
                f"ERROR: {type(error).__name__}: {error}",
                flush=True,
            )
            input("Inspect the browser, then press Enter to close...")
        finally:
            context.close()
            print("Browser closed; persistent profile retained.", flush=True)


if __name__ == "__main__":
    main()
