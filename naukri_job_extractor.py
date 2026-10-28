
"""Extract first-page Naukri job cards to Excel without opening job URLs."""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


# Companies whose job listings should be excluded from the Excel output.
EXCLUDED_COMPANIES = {
    "infosys",
    "tata consultancy services",
    "capgemini",
    "wipro",
    "cgi",
    "cognizant",
    "epam systems",
    "infosys bpm",
}


def filter_excluded_companies(jobs):
    """Remove jobs whose company name exactly matches the exclusion list."""

    filtered_jobs = []
    excluded_jobs = []

    for job in jobs:
        company = (job.get("company") or "").strip()
        normalized_company = " ".join(company.casefold().split())

        if normalized_company in EXCLUDED_COMPANIES:
            excluded_jobs.append(job)
        else:
            filtered_jobs.append(job)

    print(
        f"\nExcluded-company filter: removed {len(excluded_jobs)} job(s).",
        flush=True,
    )

    for job in excluded_jobs:
        print(
            f"  Excluded: {job.get('company')} — {job.get('title')}",
            flush=True,
        )

    print(
        f"Jobs remaining after exclusion: {len(filtered_jobs)}",
        flush=True,
    )

    return filtered_jobs


def scrape_first_page_jobs_to_excel(page):
    """
    Extract first-page job titles, company names, and URLs.

    Individual job URLs are never opened. Jobs belonging to excluded
    companies are removed before the Excel workbook is created.
    """

    output_path = (
        Path(__file__).resolve().parent
        / "naukri_software_jobs_first_page.xlsx"
    )

    print(
        "\nScraping job cards from the current results page "
        "(without opening job URLs)...",
        flush=True,
    )

    # Wait for actual job-listing anchors to appear after filtering.
    page.wait_for_function(
        """() => document.querySelectorAll(
            'a[href*="/job-listings-"]'
        ).length > 0""",
        timeout=25000,
    )

    # Read job information directly from the current page DOM.
    jobs = page.evaluate("""() => {
        const clean = value =>
            (value || '').replace(/\\s+/g, ' ').trim();

        const jobAnchors = [
            ...document.querySelectorAll('a[href*="/job-listings-"]')
        ];

        const seen = new Set();
        const rows = [];

        const companySelectors = [
            'a.comp-name',
            '.comp-name',
            '[class*="comp-name"]',
            '[class*="companyName"]',
            '[class*="company-name"]',
            '[data-testid*="company"]',
            'a[href*="/company/"]',
            'a[href*="/companies/"]'
        ];

        const titleSelectors = [
            'a.title',
            '[class*="jobTitle"]',
            '[class*="title"]',
            '[data-testid*="title"]'
        ];

        const getText = element => clean(
            element && (
                element.innerText ||
                element.textContent ||
                element.getAttribute('title')
            )
        );

        const getCompany = card => {
            if (!card) return '';

            for (const selector of companySelectors) {
                for (const element of card.querySelectorAll(selector)) {
                    const text = getText(element);

                    if (
                        text &&
                        text.length < 160 &&
                        !/^(view details|reviews?|save)$/i.test(text)
                    ) {
                        return text;
                    }
                }
            }

            return '';
        };

        const getTitle = (anchor, card) => {
            let title = getText(anchor);

            if (!title && card) {
                for (const selector of titleSelectors) {
                    const element = card.querySelector(selector);
                    title = getText(element);

                    if (title) break;
                }
            }

            return title;
        };

        for (const anchor of jobAnchors) {
            let url = '';

            try {
                url = new URL(
                    anchor.getAttribute('href'),
                    location.origin
                ).href;
            } catch (_) {
                continue;
            }

            // Only accept actual job-listing URLs.
            // Cards without a genuine job URL are skipped.
            if (!url.includes('/job-listings-')) continue;

            // Keep extraction scoped to the relevant job card.
            let card = anchor.closest(
                '.srp-jobtuple-wrapper, ' +
                '.cust-job-tuple, [data-job-id], article'
            );

            if (!card) {
                let node = anchor.parentElement;

                for (
                    let depth = 0;
                    node && depth < 9;
                    depth++, node = node.parentElement
                ) {
                    const textLength = clean(node.innerText || '').length;

                    if (
                        textLength > 0 &&
                        textLength < 2200 &&
                        node.querySelector('a[href*="/job-listings-"]')
                    ) {
                        card = node;
                        break;
                    }
                }
            }

            const title = getTitle(anchor, card);

            // Do not mark a URL as seen until usable data is available.
            if (!title || seen.has(url)) continue;

            const company = getCompany(card);

            seen.add(url);

            rows.push({
                company,
                title,
                url
            });
        }

        return rows;
    }""")

    # Deduplicate the extracted jobs by URL.
    unique_jobs = []
    seen_urls = set()

    for job in jobs:
        url = (job.get("url") or "").strip()

        if not url or url in seen_urls:
            continue

        seen_urls.add(url)

        unique_jobs.append({
            "company": (job.get("company") or "").strip(),
            "title": (job.get("title") or "").strip(),
            "url": url,
        })

    extracted_count = len(unique_jobs)

    # Exclude unwanted companies before writing the workbook.
    unique_jobs = filter_excluded_companies(unique_jobs)

    if not unique_jobs:
        raise RuntimeError(
            "No jobs remain after extraction and company filtering. "
            "The existing Excel file has not been overwritten."
        )

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "First Page Jobs"

    sheet.append(["Company Name", "Job Title", "Job URL"])

    # Style the header.
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2457C5")
        cell.alignment = Alignment(vertical="center")

    missing_company_count = 0

    for job in unique_jobs:
        company = job["company"]

        if not company:
            missing_company_count += 1

        sheet.append([
            company,
            job["title"],
            job["url"],
        ])

        url_cell = sheet.cell(row=sheet.max_row, column=3)
        url_cell.hyperlink = job["url"]
        url_cell.style = "Hyperlink"

    # Excel usability and formatting.
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 24

    sheet.column_dimensions["A"].width = 32
    sheet.column_dimensions["B"].width = 48
    sheet.column_dimensions["C"].width = 100

    # Each run replaces the previous workbook.
    workbook.save(output_path)

    print(
        f"\nUnique jobs extracted before exclusion: {extracted_count}",
        flush=True,
    )
    print(
        f"Jobs saved after exclusion: {len(unique_jobs)}",
        flush=True,
    )
    print(f"Excel saved: {output_path}", flush=True)

    if missing_company_count:
        print(
            f"WARNING: {missing_company_count} job(s) had no company name. "
            "Their titles and URLs were retained for manual review.",
            flush=True,
        )

    print(
        "No individual job URLs were opened; "
        "only the current results-page DOM was read.",
        flush=True,
    )

    return output_path
