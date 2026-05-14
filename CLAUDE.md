# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview
Playwright-based RPA that auto-logs into [qualityauditsuite.r1rcm.com](https://qualityauditsuite.r1rcm.com/), sets a month-to-date date range, runs the Audit Report, saves the download, analyses quality scores for four target facilities, shows a live dashboard, and emails results via Outlook. Entry point is `app_ui.py`.

## Running the Project

```bash
# Install dependencies (once)
pip install playwright dash pandas openpyxl apscheduler pytz
playwright install chromium

# Start the Dash web UI (port 8052, headless browser in background thread)
python app_ui.py
# → open http://localhost:8052

# Debug login / page structure interactively
python main.py          # visible browser, no UI
python debug_modal.py   # captures debug_output.html + modal_screenshot.png

# Kill stale processes before restarting (Windows)
taskkill //F //IM python.exe
```

`viewer.py` is a standalone file browser on port 8052 — don't run alongside `app_ui.py`.

## Architecture (`app_ui.py`)

### Shared state
RPA thread ↔ Dash callbacks communicate through a module-level dict guarded by a lock:
```python
rpa = {
    "status": "idle",        # idle | running | otp_needed | done | error
    "log": [],               # rolling 60-line log shown in UI
    "otp_value": None,       # OTP entered via UI fallback
    "auth_number": None,     # number shown for MS Authenticator tap
    "download_path": None,
    "report_run_time": None, # datetime when file was saved
    "analysis": None,        # result dict from analyze_report()
    "stop_requested": False, # set True by Reset button to signal thread
}
```
Dash polls every 1.5 s via `dcc.Interval(id="ticker")`.

### Dashboard visibility
Uses `dcc.Store(id="dash-active", storage_type="memory")` — cleared on browser refresh. Dashboard is hidden until a report completes in the current browser session. Reset button sets it back to `False` and clears `rpa["analysis"]`.

### MFA detection (`handle_ms_login`)
Scans up to 40 s for MFA screen types in priority order: `auth_number` (MS Authenticator number-match) → `otp_input` → `email_mfa` → `stay_signed_in` → `app_loaded`. If none detected, logs "No MFA required" and continues. OTP auto-read uses `win32com.client` Outlook MAPI; falls back to manual UI entry.

### Date setting (`set_audit_dates`)
Primary: JS React native-setter on hidden `input[name='daterange_from'/'daterange_to']`.  
Fallback: open `.react-calendar`, click tiles (skips `neighboringMonth`/`otherMonth` classes).

### Download handling
1. `page.expect_download(timeout=20s)` on "Run Report" click — direct download.
2. On timeout: dismiss confirmation modal, then poll up to 90 s using JS bounding-box coordinates of `table tbody tr:first-child td:last-child` to click the blue SVG download icon (not a `<button>` or `<a>` — coordinate click required).

Before saving, `purge_old_downloads()` deletes all existing files in `downloads/`.

### Analysis (`analyze_report`)
Reads the Excel with `header=0, skiprows=[1]` (row 0 = headers, row 1 = blank).  
Filters to `FACILITY_TARGETS` (4 CBOS facilities). Parses `Total Quality Score` as float, checks `Correction Made` for NaN (blank). Returns summary dict + pending-corrections list with hours elapsed from `Audited Date` to report run time.

### Email (`send_dashboard_email` / `build_email_html`)
Called automatically after each successful analysis. Uses `win32com.client` Outlook to send inline-CSS HTML email to `EMAIL_TO` list. Logs success or failure — requires Outlook to be running.

### Scheduler (APScheduler)
`setup_scheduler()` starts on app launch:
- **Daily 9:00 PM IST** and **2:00 AM IST** — recurring cron jobs.
- **7:00 AM IST** — one-time DateTrigger (test run); shifts to next day if already past.

All jobs call `scheduled_run()` which skips if RPA is already running.

## Key Constants
| Name | Value |
|---|---|
| `SITE_URL` | `https://qualityauditsuite.r1rcm.com/` |
| `DOWNLOAD_DIR` | `downloads/` (relative to script) |
| `PORT` | `8052` |
| `FACILITY_TARGETS` | 4 CBOS facilities (Seneca, Van Diest, Memorial, Gila Regional) |
| `EMAIL_TO` | `nkumar06@r1rcm.com`, `kwilson10@r1rcm.com` |

## Credentials
`EMAIL` and `PASSWORD` are hardcoded at the top of `app_ui.py` and `main.py`. Update both if credentials change.

## Debug artifacts
- `debug_download.html` / `debug_download.png` — saved when download button search fails; includes JS-evaluated list of visible buttons on the page.
- `debug_output.html` / `modal_screenshot.png` — from `debug_modal.py` for calendar/modal structure changes.
