# Naukri Job Search and Application Automation

A Python automation workflow that searches Naukri for software jobs, applies configured search filters, collects job listings into Excel, and processes the collected jobs through a browser-based application workflow.

## Features

- Opens Naukri using a persistent Playwright browser profile.
- Reuses an existing login session where available; if login is required, it can use credentials from a local `.env` file and pause for CAPTCHA/OTP completion.
- Searches for the keyword **`Java, Spring Boot, `**.
- Applies the configured search filters:
  - Freshness: **Last 3 day**
  - Salary ranges: **10–15 Lakhs**, **15–25 Lakhs**, **25–50 Lakhs**, **50–75 Lakhs**, **75–100 Lakhs**, **1–5 Cr**, and **5+ Cr**
  - Additional URL filters: `glbl_qcrc=1026`, `glbl_qcrc=1028`, `glbl_qcrc=1060`, `experience=4`, and `jobPostType=1`
- Verifies required filter parameters on result-page URLs and preserves them while paginating.
- Extracts job title, company name, and job URL from each results page without opening each job during extraction.
- Excludes listings from a configured list of companies.
- Deduplicates job URLs before adding them to the main workbook.
- Hands each page's newly collected jobs to the application processor.
- Records diagnostics in log files to help troubleshoot failures.

> **Important:** Naukri's page structure and URL parameters can change. The script verifies the configured URL parameters, but this does not guarantee that Naukri has applied every filter exactly as intended. Review results and applications yourself.

## Project files

Keep the following three Python files in the **same directory**, using these exact filenames because the scripts import or locate one another by name:

| File | Purpose |
|---|---|
| `naukri_login.py` | Main entry point. Opens Naukri, configures filters, extracts each results page, and coordinates pagination and processing. |
| `naukri_job_extractor.py` | Extracts job cards and saves new job URLs to the Excel workbook. |
| `naukri_application_processor.py` | Processes the jobs stored in the workbook and manages application outcomes and workbook cleanup. |

Other files are created as the application runs.

## Requirements

- Python **3.10 or newer** recommended.
- A working internet connection.
- A Naukri account.
- Google Chrome/Chromium support through Playwright.
- The Python packages listed below.

## Setup

### 1. Create and activate a virtual environment

**Windows (PowerShell):**

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
```

**macOS/Linux:**

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 2. Install dependencies

```bash
python -m pip install --upgrade pip
pip install playwright openpyxl python-dotenv
python -m playwright install chromium
```

### 3. Configure credentials

Create a `.env` file in the project directory:

```dotenv
NAUKRI_EMAIL=your_naukri_email@example.com
NAUKRI_PASSWORD=your_naukri_password
```

The login script reads these values if a login form appears. If a saved session is already active, the credentials may not be needed for that run.

**Security:** Do not commit `.env` to Git, share it, or hard-code credentials into the Python files. Add `.env`, browser-profile data, logs, and generated Excel/JSON files to `.gitignore`.

### 4. Ensure the project filenames match

The scripts expect these names in the same directory:

```text
naukri_login.py
naukri_job_extractor.py
naukri_application_processor.py
```

If your downloaded copies have timestamps or suffixes in their filenames, rename them to the names above before running the application.

## Run the application

From the project directory, with the virtual environment activated:

```bash
python naukri_login.py
```

The main script will:

1. Open the persistent browser profile and navigate to Naukri.
2. Reuse the existing login session, or pause if manual CAPTCHA/OTP completion is needed.
3. Search for `software` if the browser is not already on the expected results page.
4. configure salary, freshness, and additional URL filters.
5. Extract the current results page and write eligible new jobs to `naukri_job_list.xlsx`.
6. Find the next results page, preserve the required filters, and verify the loaded URL.
7. Close the scraping browser before starting the application processor for that page's batch.
8. Continue page by page until no safe next-page URL is found or a failure requires the workflow to stop.

Keep the terminal open while it runs. Follow any prompts shown by the script.

## How the workflow works

```text
naukri_login.py
    |
    |-- Open/reuse Naukri browser session
    |-- Apply search filters and verify result-page URL
    |-- Extract current page
    v
naukri_job_extractor.py
    |
    |-- Read job cards from the current page
    |-- Exclude configured companies
    |-- Deduplicate job URLs
    |-- Append new jobs to Excel
    v
naukri_job_list.xlsx
    |
    v
naukri_application_processor.py
    |
    |-- Process the current page's batch
    |-- Record confirmed outcomes
    |-- Save external application URLs separately
    |-- Remove successfully handled rows when cleanup is verified
    v
Next results page (filters verified) -> repeat
```

## Generated files

| File | Purpose |
|---|---|
| `naukri_job_list.xlsx` | Main job queue. Contains `Company Name`, `Job Title`, and `Job URL`. |
| `naukri_external_applications.xlsx` | Stores company name, job title, and company application URL for external application flows. |
| `naukri_confirmed_outcomes.json` | Temporary/durable record of confirmed outcomes awaiting successful Excel cleanup, used to reduce duplicate processing after cleanup failures. |
| `naukri_pagination_debug.log` | Pagination, filter-verification, and page-handoff diagnostics. |
| `naukri_processor_debug.log` | Application-processing and Excel-cleanup diagnostics. |
| `naukri_browser_profile/` | Persistent browser profile used to retain the Naukri session. |
| `naukri_freshness_dom_*.html` / `naukri_freshness_screen_*.png` / `naukri_freshness_report_*.txt` | Diagnostic artifacts created if the Freshness control needs investigation. |
| `naukri_scrape_failure_*.html` / `naukri_scrape_failure_*.png` | Page snapshots created when job-listing anchors cannot be found. |
| `naukri_pagination_error_*.html` / `naukri_pagination_error_*.png` | Page snapshots created after a pagination-level failure. |

Temporary page-batch JSON files may also be created and removed automatically.

## Excel workbook details

### Main workbook: `naukri_job_list.xlsx`

The extractor creates this workbook on the first results page of a run and appends new jobs from subsequent pages. Duplicate URLs already in the workbook are skipped.

Expected columns:

- `Company Name`
- `Job Title`
- `Job URL`

The processor uses these columns to identify and process jobs. Do not rename them unless you also update the Python code.

### External application workbook

`naukri_external_applications.xlsx` stores jobs whose application flow leads to a company website, with these columns:

- `Company Name`
- `Job Title`
- `Company Application URL`

The processor avoids adding the same external application URL more than once.

## Configuring excluded companies

Edit `EXCLUDED_COMPANIES` in `naukri_job_extractor.py` to change the company exclusion list. Company names are normalized for case and whitespace before matching.

The current exclusions are:

- Infosys
- Tata Consultancy Services
- Capgemini
- Wipro
- CGI
- Cognizant
- EPAM Systems
- Infosys BPM

## Troubleshooting

### Login is required or the session has expired

- Confirm the `.env` file is in the project directory and the variable names are exactly `NAUKRI_EMAIL` and `NAUKRI_PASSWORD`.
- Complete CAPTCHA/OTP prompts manually when requested.
- Rerun the script if login cannot be completed.

### The browser profile is already in use

The scraping script and processor share `naukri_browser_profile/`. The main workflow closes the scraping browser before starting the processor. If a previous run crashed, close any remaining browser/Playwright processes using that profile and try again. Do not run two copies of the automation simultaneously with the same profile.

### No job cards are found

- Check that Naukri results loaded and that the session is still authenticated.
- Inspect the generated `naukri_scrape_failure_*.html` and `.png` files.
- Naukri may have changed its page markup; selectors may need updating.

### Pagination stops because filters cannot be verified

- Open `naukri_pagination_debug.log`.
- Check the URL printed for the current and next pages.
- The script intentionally stops rather than scrape a page when required filter parameters are missing. This is a safety check, not necessarily proof that the site itself applied those filters.

### An application outcome is not confirmed

The processor is designed to leave unconfirmed jobs in the main workbook for review or retry. Check `naukri_processor_debug.log` and the browser before rerunning. A row may remain if the application outcome could not be confirmed or Excel cleanup failed.

### Excel file cannot be saved

- Close the workbook in Excel or another application if it is open.
- Check file permissions and ensure the project directory is writable.
- Review `naukri_processor_debug.log` for cleanup and save diagnostics.

## Important operational notes

- Review each application carefully. Automation can misread page content or behave differently when Naukri changes its interface.
- The application processor uses a browser-based workflow and may require manual review or interaction for some application steps.
- Do not run multiple instances against the same browser profile or workbook.
- Back up the Excel workbooks before making manual edits or running recovery operations.
- Use the automation in accordance with Naukri's terms, applicable laws, and the requirements of each employer's application process.
- Keep passwords, session data, logs, and personal information private.

## Disclaimer

This is a personal automation project, not an official Naukri product. It is provided without guarantees of uninterrupted operation or successful applications. You are responsible for reviewing the data collected, the actions taken, and the applications submitted.
