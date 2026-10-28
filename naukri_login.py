
from pathlib import Path
from naukri_job_extractor import scrape_first_page_jobs_to_excel
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import os
import re
import json
from datetime import datetime

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

def save_freshness_diagnostics(page):
    """Save the live DOM and a focused report when Freshness cannot be located."""

    output_dir = Path(__file__).resolve().parent
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    html_path = output_dir / f"naukri_freshness_dom_{stamp}.html"
    png_path = output_dir / f"naukri_freshness_screen_{stamp}.png"
    txt_path = output_dir / f"naukri_freshness_report_{stamp}.txt"

    try:
        html_path.write_text(page.content(), encoding="utf-8")
    except Exception as exc:
        html_path.write_text(f"Could not capture page HTML: {exc}", encoding="utf-8")

    try:
        page.screenshot(path=str(png_path), full_page=True)
        screenshot_error = "none"
    except Exception as exc:
        png_path = None
        screenshot_error = str(exc)

    report = page.evaluate("""() => {
      const isVisible = (el) => {
        if (!el) return false;
        const s = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return s.display !== 'none' && s.visibility !== 'hidden' &&
               Number(s.opacity) !== 0 && r.width > 0 && r.height > 0;
      };
      const describe = (el) => ({
        tag: el.tagName,
        id: el.id || '',
        className: typeof el.className === 'string' ? el.className : '',
        role: el.getAttribute('role') || '',
        ariaLabel: el.getAttribute('aria-label') || '',
        title: el.getAttribute('title') || '',
        text: (el.innerText || el.textContent || '').trim().replace(/\\s+/g, ' ').slice(0, 500),
        visible: isVisible(el),
        expanded: el.getAttribute('aria-expanded'),
        html: el.outerHTML.slice(0, 1800)
      });
      const all = [...document.querySelectorAll('body *')];
      const exactFreshness = all.filter(e => e.children.length === 0 && (e.textContent || '').trim() === 'Freshness');
      const exactLastDay = all.filter(e => e.children.length === 0 && (e.textContent || '').trim() === 'Last 1 day');
      const related = all.filter(e => {
        const t = (e.innerText || '').trim();
        return t && t.length < 250 && /freshness|last 1 day/i.test(t);
      });
      const ancestry = (els) => els.map(el => {
        const chain = []; let n = el;
        for (let i = 0; n && i < 6; i++, n = n.parentElement) chain.push(describe(n));
        return chain;
      });
      const controls = all.filter(e => /^(BUTTON|A|INPUT|LABEL|SUMMARY)$/.test(e.tagName) || e.getAttribute('role') === 'button' || e.getAttribute('role') === 'checkbox')
        .filter(e => {
          const t = [e.innerText, e.getAttribute('aria-label'), e.getAttribute('title'), e.getAttribute('placeholder')].filter(Boolean).join(' ');
          return /freshness|last 1 day|last 3 days|last 7 days|last 15 days|last 30 days|filter/i.test(t) || exactFreshness.some(f => e.contains(f) || f.contains(e));
        }).map(describe);
      return {
        url: location.href,
        title: document.title,
        freshnessNodes: exactFreshness.map(describe),
        freshnessAncestors: ancestry(exactFreshness),
        last1DayNodes: exactLastDay.map(describe),
        last1DayAncestors: ancestry(exactLastDay),
        relatedSmallNodes: related.slice(0, 80).map(describe),
        likelyControls: controls.slice(0, 100).map(describe)
      };
    }""")

    report_text = (
        f"URL: {page.url}\nHTML snapshot: {html_path}\n"
        f"Screenshot: {png_path if png_path else 'failed: ' + screenshot_error}\n\n"
        + json.dumps(report, indent=2, ensure_ascii=False)
    )
    txt_path.write_text(report_text, encoding="utf-8")
    print("Freshness diagnostics saved:", flush=True)
    print(f"  DOM: {html_path}", flush=True)
    print(f"  Screenshot: {png_path if png_path else screenshot_error}", flush=True)
    print(f"  Focused report: {txt_path}", flush=True)
    print("Focused DOM findings:", flush=True)
    print(json.dumps({
        "url": report.get("url"),
        "freshnessNodes": report.get("freshnessNodes"),
        "last1DayNodes": report.get("last1DayNodes"),
        "likelyControls": report.get("likelyControls"),
    }, indent=2, ensure_ascii=False), flush=True)


def select_freshness_last_1_day(page):
    """Select Last 1 day through the Freshness dropdown's DOM controls.

    This intentionally avoids Playwright's viewport/visibility click checks:
    Naukri's menu item can be present and enabled in the DOM while positioned
    outside the current viewport. DOM click still dispatches the element's
    click event when the browser window is in the background.
    """
    print("Finding Freshness filter by its DOM id...", flush=True)

    freshness_button = page.locator("button#filter-freshness")
    freshness_button.wait_for(state="attached", timeout=15000)

    print("Opening Freshness dropdown via DOM...", flush=True)
    expanded = freshness_button.get_attribute("aria-expanded")
    if expanded != "true":
        freshness_button.evaluate("el => el.click()")

    menu = page.locator("#filter-freshness-menu")
    menu.wait_for(state="attached", timeout=5000)

    # Wait for the dropdown's actual DOM state, not Playwright's viewport
    # visibility calculation. Naukri toggles visibility on this listbox.
    page.wait_for_function("""() => {
        const button = document.querySelector('#filter-freshness');
        const menu = document.querySelector('#filter-freshness-menu');
        if (!button || !menu) return false;
        const style = getComputedStyle(menu);
        return button.getAttribute('aria-expanded') === 'true' &&
               style.visibility !== 'hidden' && style.display !== 'none';
    }""", timeout=5000)

    option = menu.locator('a[data-id="filter-freshness-1"]')
    option.wait_for(state="attached", timeout=5000)

    # Validate the exact option exists and is enabled, then dispatch its native
    # click event directly. This avoids the "outside of the viewport" timeout.
    option_state = option.evaluate("""el => ({
        text: (el.innerText || el.textContent || '').trim(),
        disabled: el.getAttribute('aria-disabled') === 'true' ||
                 el.closest('[aria-disabled="true"]') !== null,
        menuVisible: (() => {
            const menu = document.querySelector('#filter-freshness-menu');
            if (!menu) return false;
            const style = getComputedStyle(menu);
            return style.visibility !== 'hidden' && style.display !== 'none';
        })()
    })""")

    if option_state["text"] != "Last 1 day":
        raise RuntimeError(
            f"Unexpected Freshness option text: {option_state['text']!r}"
        )
    if option_state["disabled"]:
        raise RuntimeError("Freshness option 'Last 1 day' is disabled in the DOM.")
    if not option_state["menuVisible"]:
        raise RuntimeError("Freshness menu did not open in the DOM.")

    print("Selecting Freshness: Last 1 day via DOM click...", flush=True)
    option.evaluate("el => el.click()")

    # Wait for the dropdown's selected label or the site URL to update.
    try:
        page.wait_for_function("""() => {
            const button = document.querySelector('#filter-freshness');
            const label = button ? (button.innerText || button.textContent || '').trim() : '';
            return label === 'Last 1 day' || /(?:jobAge|freshness|job-age|jobAgeFilter)[^&]*1/i.test(location.href);
        }""", timeout=10000)
    except PlaywrightTimeoutError:
        pass

    selected_text = freshness_button.inner_text().strip()
    print(f"Freshness dropdown now displays: {selected_text!r}", flush=True)
    print(f"Results URL after freshness selection: {page.url}", flush=True)

    url_has_freshness = bool(
        re.search(r"(?:jobAge|freshness|job-age|jobAgeFilter)[^&]*1", page.url, re.I)
    )
    if selected_text != "Last 1 day" and not url_has_freshness:
        raise RuntimeError(
            "The DOM click was dispatched, but the Freshness selection could not "
            "be confirmed from the dropdown label or URL. The browser remains open."
        )

    print("Freshness filter selection confirmed.", flush=True)




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
            scrape_first_page_jobs_to_excel(page)

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
