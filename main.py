"""
QA Suite Audit Report RPA
=========================
Auto-login (O365 + email MFA), navigate to Audit Report, set month-to-date
range via the calendar picker, click Run Report, save download.

Usage:   python main.py
Setup:   pip install playwright && playwright install chromium
"""

import os
import re
import time
from datetime import date, datetime
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ── Credentials & Settings ─────────────────────────────────────────────────
EMAIL        = "Mrajesh5255@r1rcm.com"
PASSWORD     = "Tevos@1111"
URL          = "https://qualityauditsuite.r1rcm.com/"
DOWNLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads")
LOGIN_WAIT   = 180_000   # 3 min for login steps
NAV_WAIT     = 30_000    # 30 s for page actions

os.makedirs(DOWNLOAD_DIR, exist_ok=True)


# ──────────────────────────────────────────────────────────────────────────
# OTP reader (uses local Outlook via win32com — no extra auth needed)
# ──────────────────────────────────────────────────────────────────────────
def get_otp_from_outlook(wait_sec=90):
    try:
        import win32com.client
        print("  Reading OTP from Outlook inbox...")
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = ns.GetDefaultFolder(6)  # 6 = olFolderInbox
        deadline = time.time() + wait_sec
        while time.time() < deadline:
            items = inbox.Items
            items.Sort("[ReceivedTime]", True)
            for i in range(1, min(11, items.Count + 1)):
                try:
                    msg = items.Item(i)
                    recv = msg.ReceivedTime
                    recv_dt = datetime(recv.year, recv.month, recv.day,
                                       recv.hour, recv.minute, recv.second)
                    if (datetime.now() - recv_dt).total_seconds() > 300:
                        break   # emails are sorted newest-first; rest are older
                    combined = f"{msg.Subject} {msg.Body}"
                    m = re.search(r'\b(\d{6})\b', combined)
                    if m:
                        print(f"  OTP: {m.group(1)}")
                        return m.group(1)
                except Exception:
                    continue
            print("  OTP not yet in inbox — waiting 5 s...")
            time.sleep(5)
        print("  OTP not found within timeout.")
        return None
    except Exception as e:
        print(f"  Outlook error: {e}")
        return None


# ──────────────────────────────────────────────────────────────────────────
# MS O365 login handler
# ──────────────────────────────────────────────────────────────────────────
def handle_ms_login(page):
    print("Handling MS O365 login...")

    # ── Step 1: enter email ────────────────────────────────────────────────
    try:
        page.wait_for_selector("input[type='email']", timeout=15_000)
        page.fill("input[type='email']", EMAIL)
        page.click("#idSIButton9, input[type='submit'], button[type='submit']")
        print("  Email submitted.")
    except PWTimeout:
        print("  No email input found — skipping.")

    # ── Step 2: enter password ─────────────────────────────────────────────
    try:
        page.wait_for_selector("input[type='password']", timeout=15_000)
        page.fill("input[type='password']", PASSWORD)
        page.click("#idSIButton9, input[type='submit'], button[type='submit']")
        print("  Password submitted.")
    except PWTimeout:
        print("  No password input found — skipping.")

    # ── Step 3: MFA — select email option ─────────────────────────────────
    # Wait a moment for whatever MFA screen appears
    page.wait_for_timeout(3000)
    mfa_selected = False
    for sel in [
        "[data-value='EmailOtpContextData']",
        "input[value='EmailOtpContextData']",
        "text=Send code",
        "text=Email code",
        "text=Email",
    ]:
        try:
            el = page.locator(sel)
            if el.count() > 0 and el.first.is_visible():
                el.first.click()
                mfa_selected = True
                print(f"  Email MFA selected ({sel}).")
                break
        except Exception:
            continue

    if mfa_selected:
        page.wait_for_timeout(1000)
        # Confirm / Next after selecting email method
        for btn_sel in ["#idSubmit_ProofUp_Redirect", "#idSIButton9", "button[type='submit']"]:
            try:
                btn = page.locator(btn_sel)
                if btn.count() > 0 and btn.first.is_visible():
                    btn.first.click()
                    break
            except Exception:
                continue
    else:
        print("  MFA method selection not found — may already be on OTP screen.")

    # ── Step 4: enter OTP ─────────────────────────────────────────────────
    otp_inp = None
    for otp_sel in ["input[name='otc']", "input[autocomplete='one-time-code']",
                    "input[placeholder*='ode']", "input[placeholder*='code']"]:
        try:
            page.wait_for_selector(otp_sel, timeout=30_000)
            otp_inp = page.locator(otp_sel).first
            break
        except PWTimeout:
            continue

    if otp_inp:
        otp = get_otp_from_outlook()
        if otp:
            otp_inp.fill(otp)
            for btn_sel in ["#idSubmit_SAOTCC_Continue", "#idSIButton9",
                            "input[type='submit']", "button[type='submit']"]:
                try:
                    btn = page.locator(btn_sel)
                    if btn.count() > 0 and btn.first.is_visible():
                        btn.first.click()
                        break
                except Exception:
                    continue
            print("  OTP submitted.")
        else:
            print("  Enter OTP manually in the browser.")
            page.wait_for_selector("text=Reports", timeout=300_000)
            return
    else:
        print("  No OTP input detected — continuing.")

    # ── Step 5: "Stay signed in?" ──────────────────────────────────────────
    try:
        page.wait_for_selector("#idSIButton9", timeout=10_000)
        page.click("#idSIButton9")
        print("  'Stay signed in' accepted.")
    except PWTimeout:
        pass


# ──────────────────────────────────────────────────────────────────────────
# Calendar date selector
# ──────────────────────────────────────────────────────────────────────────
def click_calendar_day(page, day_number):
    """Click a calendar tile by its day number, skipping neighbouring-month tiles."""
    tiles = page.locator("button.react-calendar__tile")
    count = tiles.count()
    for i in range(count):
        tile = tiles.nth(i)
        cls  = tile.get_attribute("class") or ""
        if "neighboringMonth" in cls or "otherMonth" in cls:
            continue
        text = tile.inner_text().strip()
        if text == str(day_number):
            tile.click()
            page.wait_for_timeout(400)
            return True
    print(f"  WARNING: Could not find calendar tile for day {day_number}")
    return False


def set_audit_dates(page, today, month_start):
    """Close any open calendar, reopen it cleanly, click start then end date."""
    # If calendar is open, close it first by clicking the modal title area
    if page.locator(".react-calendar").count() > 0:
        page.locator("text=Audit Report").nth(0).click()
        page.wait_for_timeout(500)

    # Open the date picker cleanly
    wrapper = page.locator(".react-daterange-picker__wrapper").first
    wrapper.click()
    page.wait_for_selector(".react-calendar", timeout=NAV_WAIT)
    page.wait_for_timeout(400)

    print(f"  Clicking start day: {month_start.day}")
    click_calendar_day(page, month_start.day)

    print(f"  Clicking end day: {today.day}")
    click_calendar_day(page, today.day)

    page.wait_for_timeout(500)


# ──────────────────────────────────────────────────────────────────────────
# Main flow
# ──────────────────────────────────────────────────────────────────────────
def run():
    today       = date.today()
    month_start = today.replace(day=1)
    print(f"\nAudit date range: {month_start.strftime('%m/%d/%Y')} to {today.strftime('%m/%d/%Y')}")
    print("=" * 55)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=500)
        context = browser.new_context(
            viewport={"width": 1400, "height": 900},
            accept_downloads=True,
        )
        page = context.new_page()

        # ── 1. Open site ───────────────────────────────────────────
        print("Opening QA Suite...")
        page.goto(URL, wait_until="domcontentloaded")
        page.wait_for_timeout(2000)

        # ── 2. Auto-login if needed ────────────────────────────────
        # If on MS login page, handle it; otherwise already logged in
        try:
            page.wait_for_selector(
                "input[type='email'], input[type='password']",
                timeout=8_000
            )
            handle_ms_login(page)
        except PWTimeout:
            print("Login page not detected — may already be authenticated.")

        # ── 3. Wait for main app ───────────────────────────────────
        print("Waiting for app dashboard...")
        try:
            page.wait_for_selector("text=Reports", timeout=LOGIN_WAIT)
            print("App loaded.\n")
        except PWTimeout:
            print("ERROR: Dashboard not reached. Please check credentials / MFA.")
            browser.close()
            return

        # ── 4. Reports → Audit Report ──────────────────────────────
        print("Clicking Reports...")
        page.click("text=Reports")
        page.wait_for_load_state("networkidle", timeout=NAV_WAIT)

        print("Clicking Audit Report...")
        page.wait_for_selector("text=Audit Report", timeout=NAV_WAIT)
        page.click("text=Audit Report")
        page.wait_for_load_state("networkidle", timeout=NAV_WAIT)

        # ── 5. Request Report ──────────────────────────────────────
        print("Clicking Request Report...")
        page.wait_for_selector("text=Request Report", timeout=NAV_WAIT)
        page.click("text=Request Report")
        page.wait_for_selector("text=Audited Date", timeout=NAV_WAIT)
        page.wait_for_timeout(800)

        # ── 6. Set date range via calendar ─────────────────────────
        print("Setting Audited Date range...")
        set_audit_dates(page, today, month_start)
        print(f"  Date set: {month_start.strftime('%m/%d/%Y')} to {today.strftime('%m/%d/%Y')}")

        # ── 7. Click Run Report ────────────────────────────────────
        print("Clicking Run Report...")
        page.wait_for_selector("text=Run Report", timeout=NAV_WAIT)

        # Watch for a file download triggered by Run Report
        try:
            with page.expect_download(timeout=15_000) as dl_info:
                page.click("text=Run Report")
            dl = dl_info.value
            fname = dl.suggested_filename or f"audit_report_{today.strftime('%Y%m%d')}.xlsx"
            save_path = os.path.join(DOWNLOAD_DIR, fname)
            dl.save_as(save_path)
            print(f"\nReport downloaded: {save_path}")
        except PWTimeout:
            # Report is generated server-side — click Run Report and wait for list update
            page.click("text=Run Report")
            print("\nReport queued (generated server-side). Waiting for it to appear...")
            page.wait_for_timeout(5000)
            # Try to download from the report list
            try:
                with page.expect_download(timeout=60_000) as dl_info:
                    # Click the newest row in the report list
                    page.locator("table tr, .report-row, [class*='row']").first.click()
                dl = dl_info.value
                fname = dl.suggested_filename or f"audit_report_{today.strftime('%Y%m%d')}.xlsx"
                save_path = os.path.join(DOWNLOAD_DIR, fname)
                dl.save_as(save_path)
                print(f"Report downloaded: {save_path}")
            except Exception:
                print("Download not captured automatically.")
                print(f"Check the Audit Report page and save manually to: {DOWNLOAD_DIR}")

        print("\n" + "=" * 55)
        print("Done. Close the browser window when finished.")

        # Keep browser open until user closes it
        try:
            page.wait_for_event("close", timeout=0)
        except Exception:
            pass
        browser.close()


if __name__ == "__main__":
    run()
