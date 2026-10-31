
from pathlib import Path
from naukri_job_extractor import scrape_first_page_jobs_to_excel
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import os
import re
import json
import sys
import subprocess
from datetime import datetime

load_dotenv()

EMAIL = os.getenv("NAUKRI_EMAIL")
PASSWORD = os.getenv("NAUKRI_PASSWORD")

HOME_URL = "https://www.naukri.com/mnjuser/homepage"
PROFILE_DIR = Path(__file__).resolve().parent / "naukri_browser_profile"


PAGINATION_LOG = Path(__file__).resolve().parent / "naukri_pagination_debug.log"


def pagination_log(message):
    """Print and persist pagination diagnostics without interrupting the run."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    print(line, flush=True)
    try:
        with PAGINATION_LOG.open("a", encoding="utf-8") as log_file:
            log_file.write(line + "\n")
    except Exception:
        pass

SALARY_RANGES = [
    "15-25 Lakhs",
    "25-50 Lakhs",
    "50-75 Lakhs",
    "75-100 Lakhs",
    "1-5 Cr",
    "5+ Cr",
    "10-15 Lakhs",
]


def ensure_required_search_filters(url: str) -> str:
    """
    Ensure every results-page URL contains all required salary, freshness,
    and additional filters while preserving pagination and unrelated parameters.
    """
    from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode

    parts = urlsplit(url)
    params = parse_qsl(parts.query, keep_blank_values=True)

    # These filter keys are managed explicitly. Rebuild them to avoid missing
    # repeated values or retaining conflicting values from a Next-page href.
    managed_keys = {
        "jobAge",
        "ctcFilter",
        "glbl_qcrc",
        "experience",
        "jobPostType",
    }
    params = [(key, value) for key, value in params if key not in managed_keys]

    # Preserve all seven salary ranges and the Last 3 day freshness filter.
    params.append(("jobAge", "3"))
    for value in (
        "15to25",
        "25to50",
        "50to75",
        "75to100",
        "100to500",
        "501",
        "10to15",
    ):
        params.append(("ctcFilter", value))

    # glbl_qcrc is intentionally repeated: Naukri expects one parameter per value.
    for value in ("1026", "1028", "1060"):
        params.append(("glbl_qcrc", value))

    params.extend([("experience", "4"), ("jobPostType", "1")])

    return urlunsplit((
        parts.scheme,
        parts.netloc,
        parts.path,
        urlencode(params, doseq=True),
        parts.fragment,
    ))


def preserve_search_filters(current_url: str, next_url: str) -> str:
    """
    Preserve page-specific parameters from the Next URL, then explicitly
    restore all configured filters, including repeated ctcFilter/glbl_qcrc values.
    """
    from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode

    current = urlsplit(current_url)
    target = urlsplit(next_url)
    current_params = parse_qsl(current.query, keep_blank_values=True)
    target_params = parse_qsl(target.query, keep_blank_values=True)

    # Keep the target page's pagination parameters. Add current parameters only
    # when the target does not already contain that key; required filters are
    # rebuilt below so repeated values cannot be accidentally dropped.
    target_keys = {key for key, _ in target_params}
    merged_params = list(target_params)
    for key, value in current_params:
        if key not in target_keys:
            merged_params.append((key, value))

    merged_url = urlunsplit((
        target.scheme or current.scheme,
        target.netloc or current.netloc,
        target.path,
        urlencode(merged_params, doseq=True),
        target.fragment,
    ))
    return ensure_required_search_filters(merged_url)


def verify_search_filters_in_url(url: str) -> tuple[bool, list[str]]:
    """Fail closed unless salary, freshness, and extra URL filters remain."""
    from urllib.parse import urlsplit, parse_qsl

    params = parse_qsl(urlsplit(url).query, keep_blank_values=True)
    param_map = {}
    for key, value in params:
        param_map.setdefault(key, []).append(value)

    problems = []
    if "3" not in param_map.get("jobAge", []):
        problems.append("jobAge=3 (Freshness: Last 3 day) is missing")

    expected_salary = {
        "15to25", "25to50", "50to75", "75to100",
        "100to500", "501", "10to15",
    }
    missing_salary = expected_salary - set(param_map.get("ctcFilter", []))
    if missing_salary:
        problems.append("salary filter(s) missing: " + ", ".join(sorted(missing_salary)))

    expected_glbl_qcrc = {"1026", "1028", "1060"}
    missing_glbl_qcrc = expected_glbl_qcrc - set(param_map.get("glbl_qcrc", []))
    if missing_glbl_qcrc:
        problems.append("glbl_qcrc filter(s) missing: " + ", ".join(sorted(missing_glbl_qcrc)))

    if "4" not in param_map.get("experience", []):
        problems.append("experience=4 is missing")
    if "1" not in param_map.get("jobPostType", []):
        problems.append("jobPostType=1 is missing")

    return not problems, problems


def append_additional_url_filters(page):
    """Ensure the initial results URL has all required filters, then navigate."""
    current_url = page.url
    updated_url = ensure_required_search_filters(current_url)

    print("Navigating to results URL with all required filters:", flush=True)
    print(updated_url, flush=True)
    page.goto(updated_url, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(1500)

    filters_ok, problems = verify_search_filters_in_url(page.url)
    if not filters_ok:
        raise RuntimeError(
            "Search filters could not be verified after navigation: "
            + "; ".join(problems)
        )

    print("All required search filters verified.", flush=True)
    print(f"Updated results URL: {page.url}", flush=True)


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
    search_input.fill("Java, Spring Boot, ")

    print("Entered keyword: Java, Spring Boot, ", flush=True)

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
      const exactLastDay = all.filter(e => e.children.length === 0 && (e.textContent || '').trim() === 'Last 3 days');
      const related = all.filter(e => {
        const t = (e.innerText || '').trim();
        return t && t.length < 250 && /freshness|last 3 days/i.test(t);
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
    """Select Last 3 day through the Freshness dropdown's DOM controls.

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

    option = menu.locator('a[data-id="filter-freshness-3"]')
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

    if option_state["text"] != "Last 3 days":
        raise RuntimeError(
            f"Unexpected Freshness option text: {option_state['text']!r}"
        )
    if option_state["disabled"]:
        raise RuntimeError("Freshness option 'Last 3 days' is disabled in the DOM.")
    if not option_state["menuVisible"]:
        raise RuntimeError("Freshness menu did not open in the DOM.")

    print("Selecting Freshness: Last 3 days via DOM click...", flush=True)
    option.evaluate("el => el.click()")

    # Wait for the dropdown's selected label or the site URL to update.
    try:
        page.wait_for_function("""() => {
            const button = document.querySelector('#filter-freshness');
            const label = button ? (button.innerText || button.textContent || '').trim() : '';
            return label === 'Last 3 days' || /(?:jobAge|freshness|job-age|jobAgeFilter)[^&]*3/i.test(location.href);
        }""", timeout=10000)
    except PlaywrightTimeoutError:
        pass

    selected_text = freshness_button.inner_text().strip()
    print(f"Freshness dropdown now displays: {selected_text!r}", flush=True)
    print(f"Results URL after freshness selection: {page.url}", flush=True)

    url_has_freshness = bool(
        re.search(r"(?:jobAge|freshness|job-age|jobAgeFilter)[^&]*3", page.url, re.I)
    )
    if selected_text != "Last 3 days" and not url_has_freshness:
        raise RuntimeError(
            "The DOM click was dispatched, but the Freshness selection could not "
            "be confirmed from the dropdown label or URL. The browser remains open."
        )

    print("Freshness filter selection confirmed.", flush=True)




def first_job_href(page):
    try:
        href = page.locator('a[href*="/job-listings-"]').first.get_attribute("href", timeout=3000)
        if not href:
            return ""
        return page.evaluate("(href) => new URL(href, location.origin).href", href)
    except Exception:
        return ""


def get_next_results_page_url(page):
    """
    Discover the next results page URL. Prefer reading an actual href so the
    current page does not change until its batch has been processed.
    """
    current_url = page.url
    pagination_log(f"[PAGINATION][NEXT] Looking for Next control on {current_url}")
    selectors = [
        'a[rel="next"]',
        'a[aria-label*="Next" i]',
        'a[title*="Next" i]',
        'a:has-text("Next")',
        'button[aria-label*="Next" i]',
        'button:has-text("Next")',
        '[role="button"][aria-label*="Next" i]',
        '[role="button"]:has-text("Next")',
    ]

    for selector in selectors:
        try:
            candidates = page.locator(selector)
            count = min(candidates.count(), 10)
            for index in range(count):
                candidate = candidates.nth(index)
                if not candidate.is_visible():
                    continue
                details = candidate.evaluate("""el => ({
                    tag: el.tagName,
                    text: (el.innerText || el.textContent || '').trim(),
                    href: el.href || el.getAttribute('href') || '',
                    ariaLabel: el.getAttribute('aria-label') || '',
                    disabled: el.disabled === true ||
                        el.getAttribute('aria-disabled') === 'true' ||
                        el.classList.contains('disabled'),
                    html: el.outerHTML.slice(0, 900)
                })""")
                pagination_log(
                    f"[PAGINATION][NEXT][CANDIDATE] selector={selector!r}; "
                    f"details={json.dumps(details, ensure_ascii=False)}"
                )
                if details.get("disabled"):
                    continue
                href = (details.get("href") or "").strip()
                if href and not href.lower().startswith(("javascript:", "#")):
                    absolute = page.evaluate(
                        "(href) => new URL(href, location.origin).href", href
                    )
                    if absolute != current_url:
                        filtered_url = preserve_search_filters(current_url, absolute)
                        ok, problems = verify_search_filters_in_url(filtered_url)
                        if not ok:
                            pagination_log("[PAGINATION][NEXT][ERROR] Refusing unfiltered URL: " + "; ".join(problems))
                            continue
                        pagination_log(f"[PAGINATION][NEXT] Found next-page href: {absolute}")
                        pagination_log(f"[PAGINATION][NEXT] Preserved active search filters: {filtered_url}")
                        return filtered_url
        except Exception as exc:
            pagination_log(
                f"[PAGINATION][NEXT][DEBUG] Selector {selector!r} failed: "
                f"{type(exc).__name__}: {exc}"
            )

    # Some SPA versions use a button without a navigable href.
    for selector in [
        'a[aria-label*="Next" i]',
        'a:has-text("Next")',
        'button[aria-label*="Next" i]',
        'button:has-text("Next")',
        '[role="button"][aria-label*="Next" i]',
        '[role="button"]:has-text("Next")',
    ]:
        try:
            candidates = page.locator(selector)
            for index in range(min(candidates.count(), 10)):
                candidate = candidates.nth(index)
                if not candidate.is_visible() or not candidate.is_enabled():
                    continue
                old_url = page.url
                old_first_job = first_job_href(page)
                pagination_log(
                    f"[PAGINATION][NEXT] Clicking Next button via {selector!r}; "
                    f"old_first_job={old_first_job!r}"
                )
                candidate.click(timeout=8000)
                try:
                    page.wait_for_function(
                        """({oldUrl, oldFirst}) => {
                            const first = document.querySelector('a[href*="/job-listings-"]');
                            const firstHref = first ? new URL(first.getAttribute('href'), location.origin).href : '';
                            return location.href !== oldUrl || (firstHref && firstHref !== oldFirst);
                        }""",
                        arg={"oldUrl": old_url, "oldFirst": old_first_job},
                        timeout=15000,
                    )
                except PlaywrightTimeoutError:
                    pagination_log(
                        "[PAGINATION][NEXT][ERROR] Click did not produce a detectable "
                        "URL or first-job change."
                    )
                    continue
                new_url = page.url
                pagination_log(
                    f"[PAGINATION][NEXT] After click: url={new_url}; "
                    f"first_job={first_job_href(page)!r}"
                )
                if new_url != old_url:
                    filtered_url = preserve_search_filters(old_url, new_url)
                    ok, problems = verify_search_filters_in_url(filtered_url)
                    if not ok:
                        pagination_log("[PAGINATION][NEXT][ERROR] Filter verification failed: " + "; ".join(problems))
                        return None
                    return filtered_url
                pagination_log(
                    "[PAGINATION][NEXT][ERROR] Results changed but URL did not. "
                    "Cannot safely resume this page after processor hand-off."
                )
                return None
        except Exception as exc:
            pagination_log(
                f"[PAGINATION][NEXT][DEBUG] Button click path failed: "
                f"{type(exc).__name__}: {exc}"
            )

    # Capture a compact DOM artifact when no usable Next control is found.
    try:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        html_path = Path(__file__).resolve().parent / f"naukri_next_missing_{stamp}.html"
        png_path = Path(__file__).resolve().parent / f"naukri_next_missing_{stamp}.png"
        html_path.write_text(page.content(), encoding="utf-8")
        page.screenshot(path=str(png_path), full_page=True)
        pagination_log(
            f"[PAGINATION][NEXT] No further page/control found (or selector changed). "
            f"DOM={html_path.name}; screenshot={png_path.name}"
        )
    except Exception as exc:
        pagination_log(f"[PAGINATION][NEXT][DIAGNOSTIC] Could not save artifacts: {exc}")
    return None


def launch_scraping_context(playwright):
    context = playwright.chromium.launch_persistent_context(
        user_data_dir=str(PROFILE_DIR),
        headless=False,
        viewport={"width": 1366, "height": 900},
    )
    page = context.pages[0] if context.pages else context.new_page()
    return context, page


def main():
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    pagination_log("=" * 72)
    pagination_log("Naukri page-by-page automation starting.")
    pagination_log(f"Script folder: {Path(__file__).resolve().parent}")
    pagination_log(f"Pagination log: {PAGINATION_LOG.resolve()}")

    visited_pages = set()
    page_number = 1
    append_to_workbook = False
    next_url = None
    first_page_setup_complete = False

    with sync_playwright() as p:
        context = None
        page = None
        try:
            context, page = launch_scraping_context(p)
            pagination_log("[STARTUP] Opened persistent Naukri browser profile.")
            page.goto(HOME_URL, wait_until="domcontentloaded", timeout=60000)
            pagination_log(f"[STARTUP] Opened: {page.url}")
            login_if_required(page)

            if "/software-jobs" not in page.url:
                search_software(page)

            click_salary_view_more(page)
            select_salary_ranges(page)
            select_freshness_last_1_day(page)
            append_additional_url_filters(page)
            first_page_setup_complete = True
            filters_ok, filter_problems = verify_search_filters_in_url(page.url)
            if not filters_ok:
                raise RuntimeError("Initial filters could not be verified: " + "; ".join(filter_problems) + ". Refusing to scrape unfiltered results.")
            pagination_log(f"[FILTERS] Search and filters configured and URL-verified. Results URL={page.url}")

            while True:
                current_url = page.url
                if current_url in visited_pages:
                    pagination_log(
                        f"[PAGINATION][STOP] Page URL already visited: {current_url}. "
                        "Stopping to prevent an infinite loop."
                    )
                    break
                visited_pages.add(current_url)

                pagination_log(
                    f"[PAGINATION] ===== PAGE {page_number} START ====="
                )
                pagination_log(f"[PAGINATION] Current page URL: {current_url}")

                # Scrape current page. First page resets the workbook; later
                # pages append only URLs that have not appeared earlier.
                page_urls = scrape_first_page_jobs_to_excel(
                    page,
                    append=append_to_workbook,
                )
                page_urls = list(dict.fromkeys(
                    url.strip() for url in page_urls if url and url.strip()
                ))
                pagination_log(
                    f"[PAGINATION][SCRAPE] Page {page_number}: "
                    f"{len(page_urls)} newly added URL(s) queued."
                )

                # Resolve next page before closing the scraping context.
                next_url = get_next_results_page_url(page)
                if next_url and next_url in visited_pages:
                    pagination_log(
                        f"[PAGINATION][STOP] Next URL was already visited: {next_url}"
                    )
                    next_url = None

                # Release the persistent profile before processor opens it.
                context.close()
                context = None
                pagination_log(
                    f"[PAGINATION][HANDOFF] Scraping browser closed for page {page_number}; "
                    f"processor batch size={len(page_urls)}."
                )

                if page_urls:
                    # Run the synchronous Playwright processor in a separate OS process.
                    # This avoids nested Playwright Sync API runtimes in this pagination process.
                    processor_script = Path(__file__).resolve().parent / "naukri_application_processor.py"
                    if not processor_script.exists():
                        raise FileNotFoundError(
                            f"Application processor not found: {processor_script}"
                        )

                    batch_file = Path(__file__).resolve().parent / f".naukri_batch_page_{page_number}.json"
                    try:
                        batch_file.write_text(
                            json.dumps(page_urls, ensure_ascii=False, indent=2),
                            encoding="utf-8",
                        )
                        pagination_log(
                            f"[PAGINATION][HANDOFF] Starting processor subprocess for "
                            f"page {page_number}; URLs={len(page_urls)}."
                        )
                        result = subprocess.run(
                            [sys.executable, str(processor_script), "--batch-file", str(batch_file)],
                            cwd=str(Path(__file__).resolve().parent),
                            check=False,
                        )
                        pagination_log(
                            f"[PAGINATION][HANDOFF] Processor subprocess exit code={result.returncode} "
                            f"for page {page_number}."
                        )
                        if result.returncode != 0:
                            raise RuntimeError(
                                f"Processor failed on page {page_number} with exit code "
                                f"{result.returncode}. Pagination stopped to avoid silently "
                                "skipping jobs. Fix the processor error and rerun."
                            )
                        pagination_log(
                            f"[PAGINATION][HANDOFF] Processor completed successfully for page {page_number}."
                        )
                    finally:
                        try:
                            batch_file.unlink(missing_ok=True)
                        except Exception as cleanup_exc:
                            pagination_log(
                                f"[PAGINATION][WARNING] Could not remove temporary batch file: {cleanup_exc}"
                            )
                else:
                    pagination_log(
                        f"[PAGINATION][HANDOFF] No new eligible URLs on page {page_number}; "
                        "processor was not started for this page."
                    )

                if not next_url:
                    pagination_log(
                        f"[PAGINATION][STOP] No safe next-page URL found after page {page_number}. "
                        "Automation finished."
                    )
                    break

                page_number += 1
                append_to_workbook = True
                pagination_log(
                    f"[PAGINATION] Preparing page {page_number}: {next_url}"
                )
                context, page = launch_scraping_context(p)
                page.goto(next_url, wait_until="domcontentloaded", timeout=60000)
                page.wait_for_timeout(1500)
                loaded_ok, loaded_problems = verify_search_filters_in_url(page.url)
                if not loaded_ok:
                    pagination_log("[PAGINATION][FILTERS][WARNING] Loaded URL lost filters: " + "; ".join(loaded_problems) + ". Retrying filtered URL.")
                    page.goto(next_url, wait_until="domcontentloaded", timeout=60000)
                    page.wait_for_timeout(1500)
                    loaded_ok, loaded_problems = verify_search_filters_in_url(page.url)
                if not loaded_ok:
                    raise RuntimeError(f"Page {page_number} loaded without required filters: " + "; ".join(loaded_problems) + ". Stopping before scraping irrelevant jobs.")
                pagination_log(f"[PAGINATION] Page {page_number} loaded with filters verified: {page.url}")
        except KeyboardInterrupt:
            pagination_log("[STOP] Keyboard interrupt received. Stopping safely.")
        except Exception as error:
            pagination_log(
                f"[FATAL] {type(error).__name__}: {error}"
            )
            try:
                if page is not None:
                    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    html_path = Path(__file__).resolve().parent / f"naukri_pagination_error_{stamp}.html"
                    png_path = Path(__file__).resolve().parent / f"naukri_pagination_error_{stamp}.png"
                    html_path.write_text(page.content(), encoding="utf-8")
                    page.screenshot(path=str(png_path), full_page=True)
                    pagination_log(
                        f"[FATAL][DIAGNOSTIC] HTML={html_path.name}; screenshot={png_path.name}"
                    )
            except Exception as diagnostic_exc:
                pagination_log(f"[FATAL][DIAGNOSTIC] Capture failed: {diagnostic_exc}")
            input("Inspect the browser and logs, then press Enter to close...")
        finally:
            if context is not None:
                try:
                    context.close()
                except Exception as exc:
                    pagination_log(f"[SHUTDOWN] Browser close error: {exc}")
            pagination_log(
                f"[SHUTDOWN] Automation ended. Pages visited={len(visited_pages)}"
            )


if __name__ == "__main__":
    main()
