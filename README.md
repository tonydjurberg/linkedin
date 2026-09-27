# ProspectHunter

Windows desktop prospecting tool focused on LinkedIn search result collection.

## What it does
- Open a normal Windows GUI (no terminal window in the EXE build)
- Build a LinkedIn people-search URL from keywords/title/location, or use an exact search URL
- Launches Google Chrome with a dedicated ProspectHunter profile
- Waits for manual LinkedIn sign-in/security checks when needed
- Collects visible search-result data: name, headline, location, profile URL and visible snippet
- Optional profile-visit mode for extra visible details
- Deduplicates records
- Exports CSV and XLSX
- Start / Pause / Stop controls and progress log

## Run
Download the Windows artifact from GitHub Actions after the workflow completes, extract it, and start `ProspectHunter.exe`.

Chrome must be installed. The program does not bundle Chrome.

## Important
The program only collects information that is actually exposed to the logged-in browser session. It does not attempt to bypass CAPTCHA, security checks, access controls, or other anti-automation mechanisms.
