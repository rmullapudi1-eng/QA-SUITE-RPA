# CLAUDE.md

Guidance for Claude Code when working in this repository.

## Overview
Playwright-based RPA + Dash web UI (`app_ui.py`) that auto-logs into qualityauditsuite.r1rcm.com, runs Audit Reports for **4 LOBs**, downloads Excel files, analyses quality scores, shows a live dashboard, and emails results via Outlook.

**LOBs (run in order):**
1. CashPosting-CBOS
2. CBOS AR
3. Claim Processing -CBOS
4. Credit Balance-CBOS

## Running the Project

```bash
# Install dependencies (once)
pip install playwright dash pandas openpyxl apscheduler pytz
playwright install chromium

# Start the app (port 8052, headless browser in background thread)
python app_ui.py
# open http://localhost:8052

# Kill & restart (Windows)
taskkill //F //IM python.exe && python app_ui.py
```

## Architecture (`app_ui.py`)

### Shared state (`rpa` dict, guarded by `rpa_lock`)
```python
rpa = {
    "status":        "idle",   # idle | running | otp_needed | done | error
    "log":           [],       # rolling 60-line log shown in UI
    "otp_value":     None,     # OTP entered via UI fallback
    "auth_number":   None,     # number shown for MS Authenticator tap
    "download_path": None,
    "report_run_time": None,
    "analysis":      None,     # most recent LOB's analysis dict
    "analyses":      {},       # {lob_name: analysis} — all LOBs this run
    "current_lob":   None,
    "send_email":    True,     # controlled by UI toggle
    "stop_requested": False,
}
```
Dash polls every 1.5 s via `dcc.Interval(id="ticker")`.

### LOB switching (`select_lob`)
Uses `page.mouse.click(x, y)` — React ignores JS `.click()`. Baseline count of LOB names in DOM is taken before opening dropdown to avoid false positives. Returns `True/False`; caller aborts if `False`.

### MFA detection (`handle_ms_login`)
Scans up to 40 s in priority order: `auth_number` → `otp_input` → `email_mfa` → `stay_signed_in` → `app_loaded`. OTP auto-read via `win32com.client` Outlook MAPI; falls back to manual UI entry.

### Date setting (`set_audit_dates`)
Primary: JS React native-setter on `input[name='daterange_from'/'daterange_to']`.
Fallback: open `.react-calendar`, click tiles (skip `neighboringMonth`/`otherMonth` classes).

### Download handling (`finish_download`)
- Fixed filename per LOB: `report_{ABBREV}.xlsx` (e.g. `report_AR.xlsx`) — prevents duplicates when old file is open in Excel.
- Deletes stale `*_{ABBREV}.xlsx` files before saving.
- Falls back to timestamped name only if fixed file is locked.
- Download triggered by `page.expect_download(timeout=20s)` on "Run Report". On timeout: dismiss modal, then poll 90 s clicking the blue SVG download icon via JS bounding-box coordinates.

### Analysis (`analyze_report`)
Reads Excel with `header=0, skiprows=[1]`. Filters to `LOB_FACILITY_TARGETS[lob_name]` (or all rows if `None`). Returns `{summary, pending, run_time, processes}`.
- `pct` = average of `Total Quality Score` column per facility.
- `pending` = rows with score < 100 AND `Correction Made` is blank.
- `processes` = distinct values from `Process` column (shown in UI banner).

### Email (`send_dashboard_email`)
Auto-called after each LOB analysis, gated by `rpa["send_email"]`.
- **CBOS AR**: per-facility emails via `FACILITY_EMAILS` dict. Skips only facilities with `total == 0`.
- **Other LOBs**: single email to `LOB_EMAIL_TO[lob_name]`.
- **No corrections found**: sends brief "No Outstanding Errors" email instead of skipping.
- **Corrections found**: sends full rework table email.
- Requires Outlook running. Uses `win32com.client` + `pythoncom.CoInitialize`.

### Scheduler (APScheduler)
- **9:00 PM IST** and **2:00 AM IST** — daily cron, runs all LOBs.
- **7:00 AM IST** — one-time DateTrigger on startup (shifts to next day if past).
- All jobs call `scheduled_run()` → skips if already running, always sets `send_email=True`.

## Key Constants
| Name | Purpose |
|---|---|
| `LOB_LIST` | Ordered list of 4 LOBs |
| `LOB_ABBREV` | Short code per LOB (e.g. `"CBOS AR" → "AR"`) |
| `LOB_FACILITY_TARGETS` | Per-LOB facility filter dict (CBOS AR has 8 facilities; others have fewer) |
| `LOB_EMAIL_TO` | Recipients for CashPosting, Claim Processing, Credit Balance LOBs |
| `FACILITY_EMAILS` | Per-facility recipient for CBOS AR |
| `DOWNLOAD_DIR` | `downloads/` (relative to script) |
| `PORT` | `8052` |

## Important Notes
- Facility names differ per LOB — do NOT assume same name across LOBs. `LOB_FACILITY_TARGETS` is keyed per LOB.
- CBOS AR Lincoln facility is `CBOS_Lincoln Health` (not `CBOS_Lincoln`).
- `sys.stdout`/`sys.stderr` reconfigured to UTF-8 at startup to handle Unicode in Windows console.
- `EMAIL` and `PASSWORD` are hardcoded at top of `app_ui.py`. Update if credentials change.
