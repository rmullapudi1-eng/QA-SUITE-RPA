"""
QA Suite Audit Report - UI
===========================
Run:  python app_ui.py
Open: http://localhost:8052

Browser automation runs headless in the background.
All interaction happens through this UI.
"""

import os
import re
import sys
import io
import time
import threading

# Force stdout/stderr to UTF-8 on Windows so Unicode chars in log messages
# (em-dash, arrows, box-drawing) don't crash with cp1252 UnicodeEncodeError.
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace", line_buffering=True)
if hasattr(sys.stderr, "buffer"):
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace", line_buffering=True)
import subprocess
from datetime import date, datetime

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
import dash
from dash import dcc, html, dash_table, Input, Output, State, ctx, no_update, ALL

# ── Config ──────────────────────────────────────────────────────────────
EMAIL        = "Mrajesh5255@r1rcm.com"
PASSWORD     = "Tevos@1111"
SITE_URL     = "https://qualityauditsuite.r1rcm.com/"
DOWNLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads")
PORT         = 8052

os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# ── Shared state (main thread <-> RPA thread) ───────────────────────────
FACILITY_TARGETS = [
    "CBOS_Seneca Healthcare District",
    "CBOS_Van Diest Medical Center",
    "CBOS_Memorial Health",
    "CBOS_Gila Regional Medical Center",
    "CBOS_Arkansas Valley",
    "CBOS_Hardtner Medical Center",
    "CBOS_Lincoln Health",
    "CBOS_Modoc Medical Center",
]

# ── Multi-LOB config ─────────────────────────────────────────────────────
LOB_LIST = ["CBOS AR", "CashPosting-CBOS", "Claim Processing -CBOS", "Credit Balance-CBOS"]

# Run-All order: CashPosting first (page default), then the rest
RUN_ALL_ORDER = ["CashPosting-CBOS", "CBOS AR", "Claim Processing -CBOS", "Credit Balance-CBOS"]

# Short abbreviations used in saved filenames
LOB_ABBREV = {
    "CBOS AR":                "AR",
    "CashPosting-CBOS":       "CASH",
    "Claim Processing -CBOS": "CLAIMS",
    "Credit Balance-CBOS":    "CREDIT",
}

# facility_targets per LOB — None means use ALL facilities found in the report
LOB_FACILITY_TARGETS = {
    "CBOS AR":                 FACILITY_TARGETS,
    "CashPosting-CBOS":        None,
    "Claim Processing -CBOS":  None,
    "Credit Balance-CBOS":     None,
}

# Email enabled per LOB (CBOS AR excluded per user request)
LOB_SEND_EMAIL = {
    "CBOS AR":                 False,
    "CashPosting-CBOS":        True,
    "Claim Processing -CBOS":  True,
    "Credit Balance-CBOS":     True,
}

# LOB-level email recipients for new LOBs  ← update these addresses as needed
LOB_EMAIL_TO = {
    "CashPosting-CBOS":        ["rkumar205@r1rcm.com", "kpragada@r1rcm.com"],
    "Claim Processing -CBOS":  ["nkumar06@r1rcm.com", "kwilson10@r1rcm.com"],
    "Credit Balance-CBOS":     ["rkumar205@r1rcm.com", "rlakshminarayan@r1rcm.com"],
}

rpa = {
    "status":          "idle",   # idle | running | otp_needed | done | error
    "log":             [],
    "otp_value":       None,
    "auth_number":     None,
    "download_path":   None,
    "report_run_time": None,
    "analysis":        None,     # last completed analysis (any LOB)
    "analyses":        {},       # lob_name -> analysis dict (accumulates across runs)
    "stop_requested":  False,
    "send_email":      True,
    "current_lob":     "CBOS AR",
}
rpa_lock = threading.Lock()

STATUS_COLOR = {
    "idle":       "#95a5a6",
    "running":    "#2980b9",
    "otp_needed": "#e67e22",
    "done":       "#27ae60",
    "error":      "#e74c3c",
}


def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    with rpa_lock:
        rpa["log"].append(line)
        rpa["log"] = rpa["log"][-60:]
    print(msg, flush=True)


def notify_auth_number(number):
    """Show a Windows toast notification with the MS Authenticator number."""
    try:
        script = f"""
$xml = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent(
    [Windows.UI.Notifications.ToastTemplateType]::ToastText02)
$xml.GetElementsByTagName('text')[0].AppendChild(
    $xml.CreateTextNode('QA Suite RPA - MFA Required')) | Out-Null
$xml.GetElementsByTagName('text')[1].AppendChild(
    $xml.CreateTextNode('Tap number {number} in Microsoft Authenticator')) | Out-Null
$toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier(
    'QA Suite RPA').Show($toast)
"""
        subprocess.Popen(
            ["powershell", "-WindowStyle", "Hidden", "-Command", script],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        log(f"Desktop notification sent: tap {number} in Authenticator.")
    except Exception as e:
        log(f"Toast notification failed: {e}")


# ── Report analysis ─────────────────────────────────────────────────────
def analyze_report(filepath, run_time, facility_targets=None):
    """Parse the audit Excel and return quality dashboard data.

    facility_targets: list of facility names to filter to, or None to use all.
    """
    import pandas as pd
    try:
        df = pd.read_excel(filepath, header=0, skiprows=[1])
        # Normalize facility names: remove stray spaces after underscores (e.g. "CBOS_ Arkansas" → "CBOS_Arkansas")
        df["Facility Group Name"] = df["Facility Group Name"].astype(str).str.replace(r'_\s+', '_', regex=True).str.strip()
        # Collect distinct Process values from the FULL file (before facility filter)
        _proc_col = next((c for c in df.columns if c.strip().lower() == "process"), None)
        processes = sorted(df[_proc_col].dropna().astype(str).str.strip().unique().tolist()) if _proc_col else []
        if facility_targets is not None:
            df = df[df["Facility Group Name"].isin(facility_targets)].copy()
        else:
            df = df.copy()
        if df.empty:
            return None

        # Normalise quality score to float 0-100
        def parse_score(v):
            if pd.isna(v):
                return None
            s = str(v).replace("%", "").strip()
            try:
                return float(s)
            except Exception:
                return None

        df["_score"] = df["Total Quality Score"].apply(parse_score)
        df["_correction_blank"] = df["Correction Made"].isna()

        # Count auto-accepted: rows where Correction Made contains "auto accept" (case-insensitive)
        df["_auto_accepted"] = df["Correction Made"].astype(str).str.strip().str.lower().str.contains("auto accept", na=False)

        fac_list = facility_targets if facility_targets else sorted(df["Facility Group Name"].dropna().unique().tolist())
        summary = {}
        for fac in fac_list:
            sub = df[df["Facility Group Name"] == fac]
            total        = len(sub)
            at_100       = (sub["_score"] == 100).sum()
            errors       = (sub["_score"] < 100).sum() if total else 0
            pending      = sub[(sub["_score"] < 100) & sub["_correction_blank"]]
            auto_accepted = int(sub["_auto_accepted"].sum())
            summary[fac] = {
                "total":         total,
                "at_100":        int(at_100),
                "errors":        int(errors),
                "pending":       int(len(pending)),
                "pct":           f"{sub['_score'].mean():.1f}%" if total else "N/A",
                "auto_accepted": auto_accepted,
            }

        # Build pending-corrections rows
        pending_df = df[(df["_score"] < 100) & df["_correction_blank"]]

        def _safe(v):
            """Return empty string for NaN/None, otherwise strip string."""
            try:
                if pd.isna(v):
                    return ""
            except Exception:
                pass
            return str(v).strip() if v is not None else ""

        def _safe_date(v):
            """Return formatted date string or empty."""
            try:
                if pd.isna(v):
                    return ""
                return pd.to_datetime(v).strftime("%m/%d/%Y")
            except Exception:
                return str(v).strip() if v else ""

        rows = []
        for _, r in pending_df.iterrows():
            aud = r["Audited Date"]
            try:
                hrs = round((run_time - pd.to_datetime(aud)).total_seconds() / 3600, 1)
            except Exception:
                hrs = "?"
            rows.append({
                "facility":     _safe(r.get("Facility Group Name", "")),
                "analyst":      _safe(r.get("Analyst Name", "")),
                "claim":        _safe(r.get("Claim Number", r.get("Claim #", r.get("Account Number", "")))),
                "score":        r["Total Quality Score"],
                "hours":        hrs,
                "correction":   _safe(r.get("Correction Made", "")),
                "rebuttal":     _safe(r.get("Rebuttal", "")),
                "aud_comment":  _safe(r.get("Auditor Comment on Rebuttal", "")),
                "last_reb_dt":  _safe_date(r.get("Last Rebuttal Date", "")),
                "last_res_dt":  _safe_date(r.get("Last Rebuttal Response date", "")),
            })

        return {"summary": summary, "pending": rows, "run_time": run_time.strftime("%Y-%m-%d %H:%M:%S"), "processes": processes}
    except Exception as e:
        log(f"Analysis error: {e}")
        return None


# ── OTP reader via local Outlook ─────────────────────────────────────────
def get_otp_from_outlook(wait_sec=45):
    try:
        import pythoncom, win32com.client
        pythoncom.CoInitialize()
        log("Checking Outlook inbox for OTP...")
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = ns.GetDefaultFolder(6)
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
                        break
                    combined = f"{msg.Subject} {msg.Body}"
                    m = re.search(r'\b(\d{6})\b', combined)
                    if m:
                        log(f"OTP auto-read: {m.group(1)}")
                        return m.group(1)
                except Exception:
                    continue
            log("OTP not yet in inbox, retrying...")
            time.sleep(5)
        return None
    except Exception as e:
        log(f"Outlook read failed: {e}")
        return None
    finally:
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass


# ── Email dashboard ─────────────────────────────────────────────────────
FACILITY_EMAILS = {
    "CBOS_Arkansas Valley":              "AReddy01@r1rcm.com",
    "CBOS_Gila Regional Medical Center": "nkannat@r1rcm.com",
    "CBOS_Hardtner Medical Center":      "dkaursardar@r1rcm.com",
    "CBOS_Lincoln Health":                "dkaursardar@r1rcm.com",
    "CBOS_Memorial Health":              "AReddy01@r1rcm.com",
    "CBOS_Modoc Medical Center":         "dkaursardar@r1rcm.com",
    "CBOS_Seneca Healthcare District":   "kwilson10@r1rcm.com",
    "CBOS_Van Diest Medical Center":     "nkumar06@r1rcm.com",
}

FACILITY_COLORS = {
    "CBOS_Seneca Healthcare District":   "#1a6fa8",
    "CBOS_Van Diest Medical Center":     "#117a65",
    "CBOS_Memorial Health":              "#6c3483",
    "CBOS_Gila Regional Medical Center": "#784212",
    "CBOS_Arkansas Valley":              "#c0392b",
    "CBOS_Hardtner Medical Center":      "#1abc9c",
    "CBOS_Lincoln":                       "#d35400",
    "CBOS_Modoc Medical Center":         "#2c3e50",
}


def build_email_html(analysis):
    """Build Outlook-compatible inline-CSS HTML for the dashboard email."""
    run_time = analysis.get("run_time", "")

    # ── Summary table ────────────────────────────────────────────────────
    summary_rows = ""
    for fac, s in analysis["summary"].items():
        bg    = FACILITY_COLORS.get(fac, "#2c3e50")
        if s["errors"] == 0:
            score_color = "#27ae60"
        elif s["pending"] == 0:
            score_color = "#e67e22"
        else:
            score_color = "#e74c3c"
        summary_rows += f"""
        <tr>
          <td style="padding:8px 12px;border:1px solid #ddd;font-weight:bold;color:{bg}">{fac.replace('CBOS_','')}</td>
          <td style="padding:8px 12px;border:1px solid #ddd;text-align:center">{s['total']}</td>
          <td style="padding:8px 12px;border:1px solid #ddd;text-align:center;font-weight:bold;color:{score_color}">{s['pct']}</td>
          <td style="padding:8px 12px;border:1px solid #ddd;text-align:center;color:{'#e74c3c' if s['errors'] else '#27ae60'}">{s['errors']}</td>
          <td style="padding:8px 12px;border:1px solid #ddd;text-align:center;color:{'#e74c3c' if s['pending'] else '#27ae60'}">{s['pending']}</td>
        </tr>"""

    # ── Pending corrections table ─────────────────────────────────────────
    pending = analysis.get("pending", [])
    if pending:
        pending_rows = ""
        for r in pending:
            hrs_color  = "#c0392b" if isinstance(r["hours"], (int, float)) and r["hours"] >= 48 else "#333"
            cmt        = r.get("aud_comment", "")
            cmt_style  = "font-size:10px;" if len(cmt) > 80 else ""
            reb        = r.get("rebuttal", "")
            reb_style  = "font-size:10px;" if len(reb) > 80 else ""
            pending_rows += f"""
            <tr>
              <td style="padding:7px 10px;border:1px solid #f5c6c6">{r['facility'].replace('CBOS_','')}</td>
              <td style="padding:7px 10px;border:1px solid #f5c6c6">{r['analyst']}</td>
              <td style="padding:7px 10px;border:1px solid #f5c6c6">{r['claim']}</td>
              <td style="padding:7px 10px;border:1px solid #f5c6c6;text-align:center">{r['score']}</td>
              <td style="padding:7px 10px;border:1px solid #f5c6c6;text-align:center;color:{hrs_color};font-weight:{'bold' if hrs_color!='#333' else 'normal'}">{r['hours']}</td>
              <td style="padding:7px 10px;border:1px solid #f5c6c6">{r.get('correction','')}</td>
              <td style="padding:7px 10px;border:1px solid #f5c6c6;{reb_style}">{reb}</td>
              <td style="padding:7px 10px;border:1px solid #f5c6c6;max-width:220px;word-wrap:break-word;{cmt_style}">{cmt}</td>
              <td style="padding:7px 10px;border:1px solid #f5c6c6;text-align:center">{r.get('last_reb_dt','')}</td>
              <td style="padding:7px 10px;border:1px solid #f5c6c6;text-align:center">{r.get('last_res_dt','')}</td>
            </tr>"""
        pending_section = f"""
        <h3 style="color:#c0392b;font-family:Arial,sans-serif;margin:24px 0 8px">
          Pending Corrections &nbsp;<span style="font-size:13px;font-weight:normal">(score &lt; 100% &amp; no correction made)</span>
        </h3>
        <table style="border-collapse:collapse;width:100%;font-family:Arial,sans-serif;font-size:13px">
          <thead>
            <tr style="background:#c0392b;color:white">
              <th style="padding:8px 10px;text-align:left">Facility</th>
              <th style="padding:8px 10px;text-align:left">Analyst</th>
              <th style="padding:8px 10px;text-align:left">Claim #</th>
              <th style="padding:8px 10px;text-align:center">Score</th>
              <th style="padding:8px 10px;text-align:center">Hrs Since Audit</th>
              <th style="padding:8px 10px;text-align:left">Correction Made</th>
              <th style="padding:8px 10px;text-align:left">Rebuttal</th>
              <th style="padding:8px 10px;text-align:left">Auditor Comment on Rebuttal</th>
              <th style="padding:8px 10px;text-align:center">Last Rebuttal Date</th>
              <th style="padding:8px 10px;text-align:center">Last Rebuttal Resp Date</th>
            </tr>
          </thead>
          <tbody>{pending_rows}</tbody>
        </table>"""
    else:
        pending_section = """<p style="font-family:Arial,sans-serif;color:#27ae60;font-weight:bold">
          All errors have corrections recorded.</p>"""

    return f"""
    <html><body style="margin:0;padding:0;background:#f0f2f5">
    <div style="max-width:820px;margin:20px auto;font-family:Arial,sans-serif">

      <div style="background:#2c3e50;color:white;padding:16px 22px;border-radius:8px 8px 0 0">
        <h2 style="margin:0;font-size:18px">QA Audit Quality Dashboard</h2>
        <div style="font-size:12px;opacity:.8;margin-top:4px">Report run time: {run_time}</div>
      </div>

      <div style="background:white;padding:20px 22px;border:1px solid #ddd;border-top:none">

        <h3 style="color:#2c3e50;margin:0 0 12px;font-size:14px">Quality Score Summary</h3>
        <table style="border-collapse:collapse;width:100%;font-size:13px">
          <thead>
            <tr style="background:#2c3e50;color:white">
              <th style="padding:8px 12px;text-align:left">Facility Group</th>
              <th style="padding:8px 12px;text-align:center">Total Audits</th>
              <th style="padding:8px 12px;text-align:center">Quality Score</th>
              <th style="padding:8px 12px;text-align:center">Errors</th>
              <th style="padding:8px 12px;text-align:center">No Correction</th>
            </tr>
          </thead>
          <tbody>{summary_rows}</tbody>
        </table>

        {pending_section}

        <p style="font-size:11px;color:#aaa;margin-top:24px;border-top:1px solid #eee;padding-top:10px">
          Generated automatically by QA Suite RPA &nbsp;|&nbsp; {run_time}
        </p>
      </div>
    </div>
    </body></html>"""


def build_no_errors_email_html(run_time, label):
    """Brief 'no outstanding errors' notification email."""
    return f"""
    <html><body style="margin:0;padding:0;background:#f0f2f5">
    <div style="max-width:600px;margin:20px auto;font-family:Arial,sans-serif">
      <div style="background:#2c3e50;color:white;padding:16px 22px;border-radius:8px 8px 0 0">
        <h2 style="margin:0;font-size:18px">QA Audit Quality Dashboard</h2>
        <div style="font-size:12px;opacity:.8;margin-top:4px">Report run time: {run_time}</div>
      </div>
      <div style="background:white;padding:24px 22px;border:1px solid #ddd;border-top:none;text-align:center">
        <div style="font-size:48px;margin-bottom:12px">&#10003;</div>
        <h3 style="color:#27ae60;margin:0 0 10px;font-size:18px">{label}</h3>
        <p style="color:#555;font-size:14px;margin:0">
          No outstanding rework items found. All audited records have corrections recorded or are within quality threshold.
        </p>
        <p style="font-size:11px;color:#aaa;margin-top:24px;border-top:1px solid #eee;padding-top:10px">
          Generated automatically by QA Suite RPA &nbsp;|&nbsp; {run_time}
        </p>
      </div>
    </div>
    </body></html>"""


def send_dashboard_email(analysis, lob_name="CBOS AR"):
    """Send email via Outlook.

    CBOS AR: per-facility emails using FACILITY_EMAILS dict.
    Other LOBs: one email per LOB to LOB_EMAIL_TO recipients.
    Always sends — brief 'no outstanding errors' email when pending == 0.
    """
    try:
        import pythoncom, win32com.client
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")

        if lob_name == "CBOS AR":
            # ── Per-facility emails (CBOS AR legacy behaviour) ────────────
            for fac, s in analysis["summary"].items():
                if s["total"] == 0:
                    log(f"Email skipped for {fac.replace('CBOS_', '')} — no audits in report.")
                    continue
                recipient = FACILITY_EMAILS.get(fac)
                if not recipient:
                    log(f"No email configured for {fac} — skipping.")
                    continue
                mail = outlook.CreateItem(0)
                mail.To      = recipient
                mail.Subject = f"QA Audit — {fac.replace('CBOS_', '')} — {analysis['run_time']}"
                if s["pending"] == 0:
                    mail.HTMLBody = build_no_errors_email_html(
                        analysis["run_time"],
                        f"No Outstanding Errors — {fac.replace('CBOS_', '')}"
                    )
                    mail.Send()
                    log(f"No-errors email sent to {recipient} for {fac.replace('CBOS_', '')}")
                else:
                    fac_analysis = {
                        "summary": {fac: s},
                        "pending": [r for r in analysis["pending"] if r["facility"] == fac],
                        "run_time": analysis["run_time"],
                    }
                    mail.HTMLBody = build_email_html(fac_analysis)
                    mail.Send()
                    log(f"Email sent to {recipient} for {fac.replace('CBOS_', '')}")
        else:
            # ── Single LOB-level email ─────────────────────────────────────
            recipients = LOB_EMAIL_TO.get(lob_name, [])
            if not recipients:
                log(f"No email recipients configured for LOB '{lob_name}' — skipping.")
                return
            pending_count = sum(s["pending"] for s in analysis["summary"].values())
            mail = outlook.CreateItem(0)
            mail.To      = "; ".join(recipients)
            mail.Subject = f"QA Audit — {lob_name} — {analysis['run_time']}"
            if pending_count == 0:
                mail.HTMLBody = build_no_errors_email_html(
                    analysis["run_time"],
                    f"No Outstanding Errors — {lob_name}"
                )
                mail.Send()
                log(f"No-errors email sent to {', '.join(recipients)} for {lob_name}")
            else:
                mail.HTMLBody = build_email_html(analysis)
                mail.Send()
                log(f"Email sent to {', '.join(recipients)} for {lob_name}")
    except Exception as e:
        log(f"Email send failed: {e}")
    finally:
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass


# ── MS O365 login ────────────────────────────────────────────────────────
def handle_ms_login(page):
    log("MS O365 login started...")

    # Email
    try:
        page.wait_for_selector("input[type='email']", timeout=15_000)
        page.fill("input[type='email']", EMAIL)
        for s in ["#idSIButton9", "input[type='submit']", "button[type='submit']"]:
            try:
                page.click(s, timeout=3_000)
                break
            except Exception:
                continue
        log("Email entered.")
    except PWTimeout:
        log("Email field not found.")

    # Password
    try:
        page.wait_for_selector("input[type='password']", timeout=15_000)
        page.fill("input[type='password']", PASSWORD)
        for s in ["#idSIButton9", "input[type='submit']", "button[type='submit']"]:
            try:
                page.click(s, timeout=3_000)
                break
            except Exception:
                continue
        log("Password entered.")
    except PWTimeout:
        log("Password field not found.")

    # ── Wait for whatever MFA screen appears (up to 40 s) ────────────────
    log("Checking for MFA or direct login...")
    MFA_SELECTORS = [
        ("#idRichContext_DisplaySign", "auth_number"),
        (".displaySign",               "auth_number"),
        ("[class*='DisplaySign']",     "auth_number"),
        ("[id*='DisplaySign']",        "auth_number"),
        ("input[name='otc']",          "otp_input"),
        ("input[autocomplete='one-time-code']", "otp_input"),
        ("[data-value='EmailOtpContextData']",  "email_mfa"),
        ("input[value='EmailOtpContextData']",  "email_mfa"),
        ("#idSIButton9",               "stay_signed_in"),
        ("text=Reports",               "app_loaded"),
    ]
    detected_type = None
    detected_sel  = None
    deadline = time.time() + 40
    while time.time() < deadline:
        with rpa_lock:
            if rpa.get("stop_requested"):
                return
        for sel, kind in MFA_SELECTORS:
            try:
                el = page.locator(sel)
                if el.count() > 0 and el.first.is_visible():
                    detected_type = kind
                    detected_sel  = sel
                    break
            except Exception:
                continue
        if detected_type:
            break
        time.sleep(1)

    if detected_type in ("app_loaded", None):
        log("No MFA required — already signed in or skipped.")
    else:
        log(f"MFA detected: {detected_type}")

    # ── Handle Authenticator number matching ──────────────────────────────
    if detected_type == "auth_number":
        auth_number = page.locator(detected_sel).first.inner_text().strip()
        log(f"AUTHENTICATOR NUMBER: {auth_number}")
        with rpa_lock:
            rpa["auth_number"] = auth_number
            rpa["status"]      = "otp_needed"
        log("Tap that number in your Authenticator app (2 min window)...")
        notify_auth_number(auth_number)

        approved = False
        deadline = time.time() + 120
        while time.time() < deadline:
            try:
                stay = page.locator("#idSIButton9")
                if stay.count() > 0 and stay.is_visible():
                    stay.click()
                    log("Stay signed in: Yes")
                    approved = True
                    break
                if page.locator("text=Reports").count() > 0:
                    approved = True
                    break
            except Exception:
                pass
            time.sleep(1)

        if approved:
            try:
                page.wait_for_selector("text=Reports", timeout=30_000)
            except PWTimeout:
                pass
            log("Login complete!")
            with rpa_lock:
                rpa["status"] = "running"
        else:
            log("Authenticator approval timed out.")
            with rpa_lock:
                rpa["status"] = "error"
        return

    # ── Handle email MFA option ───────────────────────────────────────────
    if detected_type == "email_mfa":
        page.locator(detected_sel).first.click()
        page.wait_for_timeout(800)
        for btn_s in ["#idSubmit_ProofUp_Redirect", "#idSIButton9",
                       "input[type='submit']", "button[type='submit']"]:
            try:
                btn = page.locator(btn_s)
                if btn.count() > 0 and btn.first.is_visible():
                    btn.first.click()
                    break
            except Exception:
                continue
        log("Email MFA selected — waiting for OTP box...")
        # Re-detect — now should show OTP input
        detected_type = None
        deadline = time.time() + 30
        while time.time() < deadline:
            for sel, kind in [("input[name='otc']", "otp_input"),
                               ("input[autocomplete='one-time-code']", "otp_input")]:
                try:
                    if page.locator(sel).count() > 0:
                        detected_type = "otp_input"
                        detected_sel  = sel
                        break
                except Exception:
                    continue
            if detected_type:
                break
            time.sleep(1)

    # ── Handle OTP entry ──────────────────────────────────────────────────
    if detected_type == "otp_input":
        log("OTP box found — reading from Outlook...")
        otp = get_otp_from_outlook(wait_sec=45)
        if not otp:
            log("Auto-read failed. Enter OTP in the UI.")
            with rpa_lock:
                rpa["status"] = "otp_needed"
            deadline = time.time() + 120
            while time.time() < deadline:
                with rpa_lock:
                    otp = rpa.get("otp_value")
                if otp:
                    break
                time.sleep(2)
        if otp:
            page.locator(detected_sel).first.fill(str(otp))
            for btn_s in ["#idSubmit_SAOTCC_Continue", "#idSIButton9",
                           "input[type='submit']", "button[type='submit']"]:
                try:
                    btn = page.locator(btn_s)
                    if btn.count() > 0 and btn.first.is_visible():
                        btn.first.click()
                        break
                except Exception:
                    continue
            log("OTP submitted.")
            with rpa_lock:
                rpa["status"] = "running"
        else:
            log("OTP timeout.")

    # ── Stay signed in? ───────────────────────────────────────────────────
    if detected_type == "stay_signed_in":
        page.locator("#idSIButton9").click()
        log("Stay signed in: Yes")
        return

    # ── Already on app ────────────────────────────────────────────────────
    if detected_type == "app_loaded":
        log("Already authenticated.")
        return

    # Final "Stay signed in?" after OTP
    try:
        page.wait_for_selector("#idSIButton9", timeout=10_000)
        page.click("#idSIButton9")
        log("Stay signed in: Yes")
    except PWTimeout:
        pass


# ── Calendar picker ──────────────────────────────────────────────────────
def click_calendar_day(page, day_number):
    tiles = page.locator("button.react-calendar__tile")
    for i in range(tiles.count()):
        tile = tiles.nth(i)
        cls  = tile.get_attribute("class") or ""
        if "neighboringMonth" in cls or "otherMonth" in cls:
            continue
        if tile.inner_text().strip() == str(day_number):
            tile.click()
            page.wait_for_timeout(400)
            return True
    log(f"WARNING: calendar tile {day_number} not found")
    return False


def set_audit_dates(page, today, month_start):
    """Set the Audited Date range. Tries two strategies in order."""

    # ── Strategy 1: React-aware JS setter on the hidden type="date" inputs ──
    # The picker hides two native date inputs: name="daterange_from" / "daterange_to"
    # Setting them via the React native-setter trick is the most reliable approach.
    try:
        found = page.locator("input[name='daterange_from']").count()
        if found:
            log("Setting dates via JS React setter on hidden inputs.")
            page.evaluate(f"""() => {{
                const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
                function reactSet(name, value) {{
                    const el = document.querySelector("input[name='" + name + "']");
                    if (!el) return;
                    setter.call(el, value);
                    el.dispatchEvent(new Event('input',  {{bubbles: true}}));
                    el.dispatchEvent(new Event('change', {{bubbles: true}}));
                }}
                reactSet('daterange_from', '{month_start.strftime("%Y-%m-%d")}');
                reactSet('daterange_to',   '{today.strftime("%Y-%m-%d")}');
            }}""")
            page.wait_for_timeout(600)
            if page.locator("button:has-text('Run Report'):not([disabled])").count() > 0:
                log("Dates accepted — Run Report button is enabled.")
                return
            log("JS setter did not enable button — trying calendar.")
    except Exception as e:
        log(f"JS setter failed: {e} — trying calendar.")

    # ── Strategy 2: open calendar and click tiles ───────────────────────────
    if page.locator(".react-calendar").count() > 0:
        page.keyboard.press("Escape")
        page.wait_for_timeout(500)

    opened = False
    for sel in [
        "button.react-daterange-picker__calendar-button",
        ".react-daterange-picker__inputGroup",
        ".react-daterange-picker__wrapper",
    ]:
        try:
            el = page.locator(sel).first
            if el.count() > 0 and el.is_visible():
                el.click()
                page.wait_for_timeout(600)
                if page.locator(".react-calendar").count() > 0:
                    opened = True
                    break
        except Exception:
            continue

    if opened:
        log(f"Selecting start: day {month_start.day}")
        click_calendar_day(page, month_start.day)
        log(f"Selecting end: day {today.day}")
        click_calendar_day(page, today.day)
        page.wait_for_timeout(600)
    else:
        log("WARNING: Could not set dates — calendar did not open.")


def purge_old_downloads():
    """Delete all files in DOWNLOAD_DIR before saving a new one."""
    for f in os.listdir(DOWNLOAD_DIR):
        fp = os.path.join(DOWNLOAD_DIR, f)
        if os.path.isfile(fp):
            try:
                os.remove(fp)
                log(f"Deleted old file: {f}")
            except Exception as e:
                log(f"Could not delete {f}: {e}")


def latest_download():
    """Return path to the most recently modified file in DOWNLOAD_DIR, or None."""
    files = [os.path.join(DOWNLOAD_DIR, f) for f in os.listdir(DOWNLOAD_DIR)
             if os.path.isfile(os.path.join(DOWNLOAD_DIR, f))]
    return max(files, key=os.path.getmtime) if files else None


# ── LOB selector ─────────────────────────────────────────────────────────
def select_lob(page, lob_name):
    """
    Select the global LOB from the top-right header dropdown using NATIVE mouse clicks.

    CRITICAL: React ignores programmatic element.click() calls (JS evaluate).
    Only real browser events (fired via page.mouse.click) trigger React's
    synthetic event handlers. All clicks here use page.mouse.click(x, y).

    Dropdown structure (confirmed from screenshots):
      Trigger: ⊞ <span>CurrentLOB</span> ▾   (top-right header, y < 80px)
      Options: plain div/span items appearing below y=100px when open

    Returns True if LOB was confirmed active, False if selection failed.
    """
    page.wait_for_timeout(800)

    # ── Read current LOB from header (JS read-only — safe) ───────────────────
    def current_header_lob():
        try:
            return page.evaluate("""
                (lobKeys) => {
                    for (const el of document.querySelectorAll('span')) {
                        const r = el.getBoundingClientRect();
                        if (lobKeys.includes(el.textContent.trim()) && r.top < 120 && r.width > 0)
                            return el.textContent.trim();
                    }
                    return null;
                }
            """, LOB_LIST)
        except Exception:
            return None

    cur = current_header_lob()
    log(f"[LOB-SEL] header={cur!r}  target={lob_name!r}")
    if cur == lob_name:
        log(f"LOB '{lob_name}' already selected.")
        return True

    # ── Get bounding boxes of the header span AND its ancestors (JS read-only) ─
    ancestor_boxes = page.evaluate("""
        (lobKeys) => {
            // Find the LOB name span in the header band
            let sp = null;
            for (const el of document.querySelectorAll('span')) {
                const r = el.getBoundingClientRect();
                if (lobKeys.includes(el.textContent.trim()) && r.top < 120 && r.width > 0) {
                    sp = el; break;
                }
            }
            if (!sp) return [];
            const boxes = [];
            let el = sp;
            for (let i = 0; i <= 8 && el; i++) {
                const r = el.getBoundingClientRect();
                if (r.width > 0 && r.height > 0) {
                    boxes.push({
                        level: i,
                        tag:   el.tagName,
                        cls:   (el.className || '').trim().split(/\s+/)[0],
                        x:     r.left + r.width  / 2,
                        y:     r.top  + r.height / 2,
                    });
                }
                el = el.parentElement;
            }
            return boxes;
        }
    """, LOB_LIST)

    if not ancestor_boxes:
        log("[LOB-SEL] Header LOB span not found.")
        try:
            page.screenshot(path=os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug_lob_select.png"))
        except Exception:
            pass
        return False

    log(f"[LOB-SEL] Ancestors: " + " | ".join(
        f"L{b['level']} {b['tag']}.{b['cls']} ({b['x']:.0f},{b['y']:.0f})" for b in ancestor_boxes
    ))

    # ── Baseline: count LOB-text elements already visible below the header ────
    # CRITICAL: The page body (report list, breadcrumbs, etc.) may already
    # contain LOB names below y=100px. We must compare AGAINST this baseline —
    # the dropdown is only "open" when the count INCREASES beyond baseline.
    _count_js = """
        (lobKeys) => {
            let n = 0;
            for (const el of document.querySelectorAll('*')) {
                const r = el.getBoundingClientRect();
                if (lobKeys.includes(el.textContent.trim()) && r.top > 100 && r.height > 0 && r.width > 0)
                    n++;
            }
            return n;
        }
    """
    baseline_count = page.evaluate(_count_js, LOB_LIST)
    log(f"[LOB-SEL] baseline LOB elements below header: {baseline_count}")

    def dropdown_is_open():
        try:
            count = page.evaluate(_count_js, LOB_LIST)
            return count > baseline_count   # new items appeared → dropdown opened
        except Exception:
            return False

    # ── Step 1: native mouse click on each ancestor until dropdown opens ──────
    opened = False
    for box in ancestor_boxes:
        try:
            page.mouse.click(box['x'], box['y'])
            page.wait_for_timeout(600)
            if dropdown_is_open():
                log(f"[LOB-SEL] Dropdown opened — level {box['level']} {box['tag']}.{box['cls']}")
                opened = True
                break
            log(f"[LOB-SEL] Level {box['level']} click — dropdown still closed.")
        except Exception as e:
            log(f"[LOB-SEL] Level {box['level']} click error: {e}")

    if not opened:
        log("[LOB-SEL] All ancestor clicks failed — saving debug snapshot.")
        try:
            page.screenshot(path=os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug_lob_select.png"))
            with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug_lob_select.html"), "w", encoding="utf-8") as fh:
                fh.write(page.content())
            log("[LOB-SEL] Saved debug_lob_select.png / .html")
        except Exception:
            pass
        log(f"WARNING: Could not open LOB dropdown for '{lob_name}'.")
        return False

    # ── Step 2: find target option — smallest bounding box wins (leaf element) ─
    # Using smallest area avoids clicking a container div that wraps the option.
    opt_box = page.evaluate("""
        (tgt) => {
            const matches = [];
            for (const el of document.querySelectorAll('*')) {
                const r = el.getBoundingClientRect();
                if (el.textContent.trim() === tgt && r.top > 100 && r.height > 0 && r.width > 0)
                    matches.push({ x: r.left + r.width/2, y: r.top + r.height/2,
                                   tag: el.tagName, area: r.width * r.height });
            }
            if (!matches.length) return null;
            // smallest area = most likely the actual clickable leaf option
            matches.sort((a, b) => a.area - b.area);
            return matches[0];
        }
    """, lob_name)

    if not opt_box:
        log(f"[LOB-SEL] Option '{lob_name}' not visible in open dropdown.")
        page.keyboard.press("Escape")
        return False

    # ── Step 3: native mouse click on the target option ───────────────────────
    log(f"[LOB-SEL] Clicking option '{lob_name}' ({opt_box['tag']}, area={opt_box['area']:.0f}) at ({opt_box['x']:.0f}, {opt_box['y']:.0f})")
    page.mouse.click(opt_box['x'], opt_box['y'])
    page.wait_for_timeout(2000)   # allow page context to switch

    new_lob = current_header_lob()
    if new_lob == lob_name:
        log(f"LOB '{lob_name}' confirmed active in header.")
        return True
    else:
        log(f"[LOB-SEL] WARNING: Header shows '{new_lob}' after clicking — expected '{lob_name}'.")
        return False


# ── RPA background thread ────────────────────────────────────────────────
def rpa_thread(today, month_start, lob_names=None):
    if lob_names is None:
        lob_names = ["CBOS AR"]
    with rpa_lock:
        rpa.update({"status": "running", "log": [], "otp_value": None,
                    "auth_number": None, "download_path": None,
                    "stop_requested": False, "current_lob": lob_names[0]})

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=["--disable-blink-features=AutomationControlled"],
            )
            context = browser.new_context(
                viewport={"width": 1400, "height": 900},
                accept_downloads=True,
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                ),
            )
            page = context.new_page()

            log(f"Opening {SITE_URL}...")
            page.goto(SITE_URL, wait_until="domcontentloaded")
            page.wait_for_timeout(2000)

            # Login if needed
            try:
                page.wait_for_selector("input[type='email'], input[type='password']",
                                        timeout=8_000)
                handle_ms_login(page)
            except PWTimeout:
                log("Login page not detected — already authenticated.")

            with rpa_lock:
                rpa["status"] = "running"

            # Wait for app dashboard
            log("Waiting for app dashboard...")
            try:
                page.wait_for_selector("text=Reports", timeout=180_000)
                log("App loaded.")
            except PWTimeout:
                log("ERROR: Dashboard not reached. Check credentials.")
                with rpa_lock:
                    rpa["status"] = "error"
                browser.close()
                return

            def run_one_lob(lob_name):
                # Step 1: Select the LOB in the global context switcher (top-of-page dropdown).
                # This must happen BEFORE navigating to the report page so the report runs
                # under the correct LOB's data context.
                log(f"Selecting LOB: {lob_name}")
                if not select_lob(page, lob_name):
                    log(f"ERROR: LOB switch to '{lob_name}' failed — skipping this LOB to avoid wrong-data download.")
                    return
                page.wait_for_timeout(1500)   # allow page to stabilise after LOB switch

                # Step 2: Navigate to Reports → Audit Report within the now-selected LOB context
                log("Reports -> Audit Report")
                page.click("text=Reports")
                page.wait_for_load_state("networkidle", timeout=30_000)
                page.wait_for_selector("text=Audit Report", timeout=15_000)
                page.click("text=Audit Report")
                page.wait_for_load_state("networkidle", timeout=30_000)
                page.wait_for_timeout(800)

                # Step 3: Open the report request form
                log("Clicking Request Report...")
                page.wait_for_selector("text=Request Report", timeout=15_000)
                page.click("text=Request Report")
                page.wait_for_selector("text=Audited Date", timeout=30_000)
                page.wait_for_timeout(800)

                # Step 4: Set month-to-date date range
                log(f"Setting dates: {month_start.strftime('%m/%d/%Y')} to {today.strftime('%m/%d/%Y')}")
                set_audit_dates(page, today, month_start)
                log("Dates set.")

                # Run Report
                # The button may stay [disabled] even after dates are set (React validation quirk).
                # Wait briefly for it to become enabled; if not, remove the attribute via JS and click.
                log("Waiting for Run Report button to be enabled...")
                enabled_btn = page.locator("button:has-text('Run Report'):not([disabled])")
                try:
                    enabled_btn.wait_for(state="visible", timeout=5_000)
                    log("Button enabled — clicking normally.")
                except PWTimeout:
                    log("Button still disabled — removing attribute via JS and clicking.")
                    page.evaluate("""
                        const btn = document.querySelector('button[form="editFilterOptions"]')
                                  || [...document.querySelectorAll('button')]
                                       .find(b => b.textContent.trim() === 'Run Report');
                        if (btn) btn.removeAttribute('disabled');
                    """)
                    page.wait_for_timeout(300)

                log("Clicking Run Report...")
                run_btn = page.locator("button:has-text('Run Report')").first

                def dismiss_modal(pg):
                    """Close any confirmation/info modal that appears after Run Report."""
                    for sel in ["button:has-text('Cancel')", "button:has-text('Close')",
                                "button:has-text('OK')", "[aria-label='Close']", "button.modal-close"]:
                        try:
                            el = pg.locator(sel).first
                            if el.count() > 0 and el.is_visible():
                                el.click()
                                log(f"Modal dismissed ({sel}).")
                                pg.wait_for_timeout(1000)
                                return
                        except Exception:
                            continue
                    # Try Escape key
                    try:
                        pg.keyboard.press("Escape")
                        pg.wait_for_timeout(500)
                    except Exception:
                        pass

                def get_download_cell_coords(pg):
                    """Scroll the last cell of the first data row into view, return its viewport centre."""
                    return pg.evaluate("""() => {
                        const selectors = ['table tbody tr', 'table tr'];
                        for (const sel of selectors) {
                            const rows = document.querySelectorAll(sel);
                            for (const row of rows) {
                                const cells = row.querySelectorAll('td');
                                if (cells.length === 0) continue;
                                const lastCell = cells[cells.length - 1];
                                lastCell.scrollIntoView({block: 'center', inline: 'nearest'});
                                const rect = lastCell.getBoundingClientRect();
                                if (rect.width > 0 && rect.height > 0) {
                                    return {x: rect.x + rect.width / 2, y: rect.y + rect.height / 2};
                                }
                            }
                        }
                        return null;
                    }""")

                def finish_download(dl, run_time):
                    """Save download, run analysis, send email. Returns True on success."""
                    abbrev     = LOB_ABBREV.get(lob_name, lob_name.replace(" ", "_").replace("-", ""))
                    # Fixed filename per LOB — prevents duplicate files when old file is locked in Excel
                    fixed_name = f"report_{abbrev}.xlsx"
                    save_path  = os.path.join(DOWNLOAD_DIR, fixed_name)

                    # Delete all stale timestamped files for this LOB (e.g. *_AR.xlsx)
                    for _f in os.listdir(DOWNLOAD_DIR):
                        _fp = os.path.join(DOWNLOAD_DIR, _f)
                        if _f != fixed_name and _f.endswith(f"_{abbrev}.xlsx") and os.path.isfile(_fp):
                            try:
                                os.remove(_fp)
                                log(f"Deleted old file: {_f}")
                            except Exception as _e:
                                log(f"Could not delete {_f}: {_e}")

                    # Try to remove the fixed file so the new download can be saved cleanly
                    if os.path.isfile(save_path):
                        try:
                            os.remove(save_path)
                        except Exception:
                            pass  # locked — Playwright will overwrite or we fall back below

                    # Save
                    try:
                        dl.save_as(save_path)
                        log(f"Saved: {fixed_name}")
                    except Exception as _save_err:
                        # Fallback: timestamped name when fixed file is still locked
                        suggested = dl.suggested_filename
                        if suggested:
                            base, ext = os.path.splitext(suggested)
                            fname = f"{base}_{abbrev}{ext}"
                        else:
                            fname = f"audit_{today.strftime('%Y%m%d')}_{abbrev}.xlsx"
                        save_path = os.path.join(DOWNLOAD_DIR, fname)
                        dl.save_as(save_path)
                        log(f"Saved (fallback, fixed file locked): {fname}")
                    with rpa_lock:
                        rpa["download_path"]   = save_path
                        rpa["report_run_time"] = run_time
                    fac_targets = LOB_FACILITY_TARGETS.get(lob_name)  # None = all facilities
                    analysis = analyze_report(save_path, run_time, fac_targets)
                    with rpa_lock:
                        rpa["analysis"] = analysis
                        if analysis:
                            rpa["analyses"][lob_name] = analysis
                    if analysis:
                        pending_count = sum(v["pending"] for v in analysis["summary"].values())
                        log(f"Analysis complete [{lob_name}] — {pending_count} pending corrections found.")
                    with rpa_lock:
                        do_email = rpa.get("send_email", False)
                    if analysis and do_email:
                        log(f"Auto-sending email for {lob_name}...")
                        send_dashboard_email(analysis, lob_name)
                    return True

                # Snapshot row count BEFORE clicking Run Report so we can detect the new row
                def count_table_rows(pg):
                    return pg.evaluate("""() => {
                        for (const sel of ['table tbody tr', 'table tr']) {
                            const rows = [...document.querySelectorAll(sel)].filter(
                                r => r.querySelectorAll('td').length >= 2
                            );
                            if (rows.length) return rows.length;
                        }
                        return 0;
                    }""") or 0

                rows_before = count_table_rows(page)
                log(f"Rows before Run Report: {rows_before}")

                try:
                    with page.expect_download(timeout=20_000) as dl_info:
                        run_btn.click()
                    finish_download(dl_info.value, datetime.now())
                except Exception:
                    # Server-side report generation — dismiss confirmation modal, then click download arrow
                    log("Report queued (server-side). Dismissing modal...")
                    dismiss_modal(page)
                    page.wait_for_timeout(2000)

                    log("Waiting for NEW report row to appear in table...")
                    download_saved = False
                    deadline = time.time() + 300   # 5 min total — server generation can be slow

                    def get_first_row_status(pg):
                        """Return the Status cell text of the first data row, or '' if not found."""
                        return pg.evaluate("""() => {
                            const selectors = ['table tbody tr', 'table tr'];
                            for (const sel of selectors) {
                                const rows = document.querySelectorAll(sel);
                                for (const row of rows) {
                                    const cells = row.querySelectorAll('td');
                                    if (cells.length < 2) continue;
                                    for (const cell of cells) {
                                        const t = cell.innerText.trim();
                                        if (t === 'Completed' || t === 'In-Progress' || t === 'Failed') return t;
                                    }
                                    return '';
                                }
                            }
                            return '';
                        }""") or ''

                    # Phase 0: wait until a NEW row appears (row count increases)
                    new_row_deadline = time.time() + 60  # give 60s for new row to appear
                    while time.time() < new_row_deadline:
                        with rpa_lock:
                            if rpa["stop_requested"]:
                                log("Stop requested — aborting.")
                                break
                        try:
                            page.reload(wait_until="networkidle")
                            page.wait_for_timeout(1500)
                        except Exception:
                            pass
                        rows_now = count_table_rows(page)
                        if rows_now > rows_before:
                            log(f"New report row appeared ({rows_before} -> {rows_now}). Waiting for Completed status...")
                            break
                        log(f"New row not yet visible (still {rows_now} rows) — retrying in 8s...")
                        page.wait_for_timeout(8000)
                    else:
                        log("WARNING: New report row never appeared — may download wrong file.")

                    # Phase 1: wait until first-row status == Completed
                    while time.time() < deadline:
                        with rpa_lock:
                            if rpa["stop_requested"]:
                                log("Stop requested — aborting.")
                                break
                        status_text = get_first_row_status(page)
                        if status_text == "Completed":
                            log("Report status: Completed — attempting download...")
                            break
                        elif status_text == "Failed":
                            log("Report generation failed on server.")
                            break
                        else:
                            log(f"Report status: {status_text or 'waiting...'} — checking again in 8s...")
                            page.wait_for_timeout(8000)
                            page.reload(wait_until="networkidle")
                            page.wait_for_timeout(2000)

                    # Phase 2: click download once Completed
                    while time.time() < deadline and not download_saved:
                        with rpa_lock:
                            if rpa["stop_requested"]:
                                log("Stop requested — aborting download.")
                                break
                        coords = get_download_cell_coords(page)
                        if not coords:
                            log("Table row not found — waiting 5s...")
                            page.wait_for_timeout(5000)
                            continue

                        log(f"Clicking download icon at ({coords['x']:.0f}, {coords['y']:.0f})...")
                        try:
                            # Strategy 1: SVG locator with Playwright (handles scroll + visibility)
                            svg_loc = page.locator(
                                "table tbody tr:first-child td:last-child svg"
                            ).first
                            with page.expect_download(timeout=60_000) as dl_info:
                                if svg_loc.count() > 0:
                                    svg_loc.scroll_into_view_if_needed()
                                    page.wait_for_timeout(300)
                                    svg_loc.click(force=True)
                                else:
                                    # Strategy 2: coordinate click (scroll already done in get_download_cell_coords)
                                    page.wait_for_timeout(300)
                                    page.mouse.click(coords["x"], coords["y"])
                            finish_download(dl_info.value, datetime.now())
                            download_saved = True
                        except Exception as e:
                            log(f"Click failed: {e}")
                            # Strategy 3: JS dispatchEvent as last resort
                            try:
                                log("Trying JS-dispatch click fallback...")
                                with page.expect_download(timeout=60_000) as dl_info:
                                    page.evaluate("""() => {
                                        const selectors = ['table tbody tr', 'table tr'];
                                        for (const sel of selectors) {
                                            const rows = document.querySelectorAll(sel);
                                            for (const row of rows) {
                                                const cells = row.querySelectorAll('td');
                                                if (cells.length === 0) continue;
                                                const lastCell = cells[cells.length - 1];
                                                const target = lastCell.querySelector('svg') || lastCell;
                                                target.dispatchEvent(
                                                    new MouseEvent('click', {bubbles: true, cancelable: true})
                                                );
                                                return;
                                            }
                                        }
                                    }""")
                                finish_download(dl_info.value, datetime.now())
                                download_saved = True
                            except Exception as e2:
                                log(f"JS-dispatch also failed: {e2} — retrying in 5s...")
                                page.wait_for_timeout(5000)

                    if not download_saved:
                        # Save debug snapshot
                        try:
                            debug_png  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug_download.png")
                            debug_html = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug_download.html")
                            page.screenshot(path=debug_png, full_page=True)
                            with open(debug_html, "w", encoding="utf-8") as f:
                                f.write(page.content())
                            # Log live DOM buttons via JS
                            btn_info = page.evaluate("""() => {
                                const els = document.querySelectorAll('button, a');
                                return [...els].filter(e => {
                                    const r = e.getBoundingClientRect();
                                    return r.width > 0 && r.height > 0;
                                }).slice(0, 20).map(e => ({
                                    tag: e.tagName, cls: e.className.slice(0,60),
                                    aria: e.getAttribute('aria-label')||'',
                                    title: e.getAttribute('title')||'',
                                    text: e.innerText.trim().slice(0,30)
                                }));
                            }""")
                            lines = [f"  [{b['tag']}] text='{b['text']}' cls='{b['cls']}' aria='{b['aria']}' title='{b['title']}'" for b in (btn_info or [])]
                            log("Visible buttons on page:\n" + "\n".join(lines))
                            log("Debug snapshot saved: debug_download.png + debug_download.html")
                        except Exception as de:
                            log(f"Debug capture failed: {de}")
                        log("Auto-download not captured. Check downloads folder manually.")

            # ── Loop through all LOBs ─────────────────────────────────────
            for i, lob_name in enumerate(lob_names):
                with rpa_lock:
                    if rpa["stop_requested"]:
                        log("Stop requested — aborting.")
                        break
                    rpa["current_lob"] = lob_name
                log(f"[{i+1}/{len(lob_names)}] Starting LOB: {lob_name}")
                run_one_lob(lob_name)

            with rpa_lock:
                rpa["status"] = "done"
            log("All LOBs complete!" if len(lob_names) > 1 else "Done!")
            browser.close()

    except Exception as e:
        log(f"RPA error: {e}")
        with rpa_lock:
            rpa["status"] = "error"


# ── Dash layout ──────────────────────────────────────────────────────────
app = dash.Dash(__name__, title="QA Suite Report Runner",
                suppress_callback_exceptions=True)
app.layout = html.Div(style={
    "fontFamily": "Arial, sans-serif", "backgroundColor": "#f0f2f5",
    "minHeight": "100vh", "padding": "30px 0",
}, children=[
html.Div(style={"maxWidth": "820px", "margin": "0 auto", "padding": "0 20px"}, children=[

    html.H2("QA Suite Audit Report", style={
        "color": "#2c3e50", "margin": "0 0 20px",
        "borderBottom": "3px solid #2c3e50", "paddingBottom": "10px",
    }),

    # Date range
    html.Div(style={
        "background": "white", "borderRadius": "8px", "padding": "16px 20px",
        "marginBottom": "14px", "boxShadow": "0 1px 4px rgba(0,0,0,.08)",
    }, children=[
        html.Div("Audit Date Range", style={"fontWeight": "bold", "color": "#555",
                                             "fontSize": "12px", "marginBottom": "6px"}),
        html.Div(id="date-display", style={"fontSize": "22px", "color": "#2c3e50",
                                            "fontWeight": "bold", "letterSpacing": "1px"}),
    ]),

    # Run button + LOB selector + OTP
    html.Div(style={
        "background": "white", "borderRadius": "8px", "padding": "18px 20px",
        "marginBottom": "14px", "boxShadow": "0 1px 4px rgba(0,0,0,.08)",
    }, children=[
        # LOB selector row
        html.Div(style={"display": "flex", "alignItems": "center", "gap": "10px",
                        "marginBottom": "12px", "flexWrap": "wrap"}, children=[
            html.Span("Line of Business:", style={"fontWeight": "bold", "color": "#555",
                                                   "fontSize": "13px", "whiteSpace": "nowrap"}),
            dcc.Dropdown(
                id="lob-selector",
                options=[{"label": lob, "value": lob} for lob in LOB_LIST],
                value="CBOS AR",
                clearable=False,
                style={"width": "260px", "fontSize": "13px"},
            ),
        ]),
        html.Div(style={"display": "flex", "gap": "10px", "alignItems": "center", "flexWrap": "wrap"}, children=[
            html.Button("Run Report", id="run-btn", n_clicks=0, style={
                "padding": "11px 32px", "backgroundColor": "#2c3e50", "color": "white",
                "border": "none", "borderRadius": "6px", "cursor": "pointer",
                "fontSize": "15px", "fontWeight": "bold", "letterSpacing": ".5px",
            }),
            html.Button("Run All LOBs", id="run-all-btn", n_clicks=0, style={
                "padding": "11px 28px", "backgroundColor": "#117a65", "color": "white",
                "border": "none", "borderRadius": "6px", "cursor": "pointer",
                "fontSize": "14px", "fontWeight": "bold", "letterSpacing": ".5px",
            }),
            html.Button("Reset", id="reset-btn", n_clicks=0, style={
                "padding": "11px 20px", "backgroundColor": "#e74c3c", "color": "white",
                "border": "none", "borderRadius": "6px", "cursor": "pointer",
                "fontSize": "13px", "fontWeight": "bold",
            }),
            dcc.Checklist(
                id="email-toggle",
                options=[{"label": " Send email notifications", "value": "yes"}],
                value=["yes"],
                style={"marginLeft": "10px", "fontSize": "13px", "color": "#2c3e50",
                       "display": "flex", "alignItems": "center"},
            ),
        ]),
        # Auth panel (hidden by default, shows when MFA needed)
        html.Div(id="otp-panel", style={"display": "none", "marginTop": "16px"}, children=[
            # Authenticator number matching box
            html.Div(id="auth-number-box", style={"display": "none"}, children=[
                html.Div("Open Microsoft Authenticator and tap this number:",
                         style={"color": "#e67e22", "fontWeight": "bold", "marginBottom": "10px"}),
                html.Div(id="auth-number-display", style={
                    "fontSize": "56px", "fontWeight": "900", "color": "white",
                    "backgroundColor": "#e67e22", "borderRadius": "12px",
                    "padding": "10px 30px", "display": "inline-block",
                    "letterSpacing": "8px", "marginBottom": "8px",
                }),
                html.Div("Waiting for approval...",
                         style={"color": "#7f8c8d", "fontSize": "13px", "fontStyle": "italic"}),
            ]),
            # OTP code entry (email OTP fallback)
            html.Div(id="otp-entry-box", style={"display": "none"}, children=[
                html.Div("OTP sent to your email — enter it below:",
                         style={"color": "#e67e22", "fontWeight": "bold", "marginBottom": "8px"}),
                html.Div(style={"display": "flex", "gap": "10px", "alignItems": "center"}, children=[
                    dcc.Input(id="otp-input", type="text", maxLength=6,
                              placeholder="6-digit code",
                              style={"padding": "9px 12px", "fontSize": "20px", "width": "150px",
                                     "borderRadius": "5px", "border": "2px solid #e67e22",
                                     "letterSpacing": "6px", "textAlign": "center"}),
                    html.Button("Submit", id="otp-submit", n_clicks=0, style={
                        "padding": "9px 20px", "backgroundColor": "#e67e22", "color": "white",
                        "border": "none", "borderRadius": "5px", "cursor": "pointer",
                        "fontWeight": "bold",
                    }),
                ]),
                html.Div(id="otp-msg", style={"marginTop": "6px", "fontSize": "12px", "color": "#888"}),
            ]),
        ]),
    ]),

    # Status
    html.Div(style={
        "background": "white", "borderRadius": "8px", "padding": "16px 20px",
        "marginBottom": "14px", "boxShadow": "0 1px 4px rgba(0,0,0,.08)",
    }, children=[
        html.Div(style={"display": "flex", "alignItems": "center", "gap": "10px",
                        "marginBottom": "10px"}, children=[
            html.Span("Status:", style={"fontWeight": "bold", "color": "#555", "fontSize": "13px"}),
            html.Span(id="status-badge", children="Idle", style={
                "padding": "3px 14px", "borderRadius": "12px", "color": "white",
                "fontSize": "12px", "fontWeight": "bold", "backgroundColor": "#95a5a6",
            }),
        ]),
        html.Pre(id="log-display", style={
            "fontFamily": "Consolas, monospace", "fontSize": "11.5px",
            "backgroundColor": "#1e2a35", "color": "#a8d8ea",
            "padding": "12px 14px", "borderRadius": "5px",
            "minHeight": "90px", "maxHeight": "220px",
            "overflowY": "auto", "margin": "0", "whiteSpace": "pre-wrap",
        }),
    ]),

    # ── Quality Dashboard ────────────────────────────────────────────────
    html.Div(id="dashboard-section", style={"marginTop": "18px", "display": "none"}, children=[

        html.H3("Quality Dashboard", style={
            "color": "#2c3e50", "margin": "0 0 12px",
            "borderBottom": "2px solid #2c3e50", "paddingBottom": "6px",
            "fontSize": "15px", "letterSpacing": ".5px",
        }),

        # LOB slicer — pre-created buttons (hidden until data exists for that LOB)
        html.Div(id="lob-slicer-container", style={"marginBottom": "14px"}, children=[
            html.Div(id="lob-slicer-tabs", style={"display": "flex", "gap": "6px", "flexWrap": "wrap"},
                     children=[
                         html.Button(
                             lob,
                             id=f"lob-tab-{lob.replace(' ', '_').replace('-', '_')}",
                             n_clicks=0,
                             style={"display": "none"},
                         )
                         for lob in LOB_LIST
                     ]),
        ]),
        dcc.Store(id="dashboard-lob", storage_type="memory", data=None),

        html.Div(id="quality-cards", style={"marginBottom": "16px"}),

        html.Div(style={"display": "flex", "alignItems": "center", "justifyContent": "space-between",
                        "marginBottom": "8px"}, children=[
            html.Div("Pending Corrections (Score < 100% & No Correction Made)", style={
                "fontWeight": "bold", "color": "#e74c3c", "fontSize": "13px",
            }),
            html.Button("⬇ Download", id="download-pending-btn", n_clicks=0, style={
                "backgroundColor": "#2980b9", "color": "white", "border": "none",
                "padding": "5px 14px", "borderRadius": "4px", "cursor": "pointer",
                "fontSize": "12px", "fontWeight": "bold",
            }),
        ]),
        dcc.Download(id="download-pending"),
        html.Div(id="dashboard-run-time", style={"fontSize": "11px", "color": "#888", "marginBottom": "8px"}),

        html.Div(id="process-validation-banner"),

        dash_table.DataTable(
            id="pending-table",
            columns=[
                {"name": "Facility",                    "id": "facility"},
                {"name": "Analyst",                     "id": "analyst"},
                {"name": "Claim Number",                "id": "claim"},
                {"name": "Quality Score",               "id": "score"},
                {"name": "Hrs Since Audit",             "id": "hours"},
                {"name": "Correction Made",             "id": "correction"},
                {"name": "Rebuttal",                    "id": "rebuttal"},
                {"name": "Auditor Comment on Rebuttal", "id": "aud_comment"},
                {"name": "Last Rebuttal Date",          "id": "last_reb_dt"},
                {"name": "Last Rebuttal Resp Date",     "id": "last_res_dt"},
            ],
            data=[],
            style_table={"overflowX": "auto"},
            style_cell={"textAlign": "left", "padding": "6px 10px",
                        "fontFamily": "Arial", "fontSize": "12px",
                        "whiteSpace": "normal", "height": "auto"},
            style_header={"backgroundColor": "#c0392b", "color": "white",
                          "fontWeight": "bold", "fontSize": "12px",
                          "whiteSpace": "normal"},
            style_cell_conditional=[
                {"if": {"column_id": "aud_comment"},
                 "maxWidth": "260px", "fontSize": "10px", "whiteSpace": "normal"},
                {"if": {"column_id": "rebuttal"},
                 "maxWidth": "180px", "fontSize": "10px", "whiteSpace": "normal"},
                {"if": {"column_id": "correction"},
                 "maxWidth": "160px", "fontSize": "11px", "whiteSpace": "normal"},
                {"if": {"column_id": "facility"},  "minWidth": "130px"},
                {"if": {"column_id": "claim"},     "minWidth": "110px"},
                {"if": {"column_id": "last_reb_dt"},  "minWidth": "110px", "textAlign": "center"},
                {"if": {"column_id": "last_res_dt"},  "minWidth": "110px", "textAlign": "center"},
                {"if": {"column_id": "hours"},     "textAlign": "center", "minWidth": "80px"},
                {"if": {"column_id": "score"},     "textAlign": "center", "minWidth": "80px"},
            ],
            style_data_conditional=[
                {"if": {"row_index": "odd"}, "backgroundColor": "#fff5f5"},
                {"if": {"filter_query": "{hours} >= 48"},
                 "color": "#c0392b", "fontWeight": "bold"},
            ],
        ),
        html.Div(id="dashboard-msg", style={"marginTop": "8px", "fontSize": "12px",
                                            "color": "#888", "minHeight": "16px"}),
    ]),

    # Memory store — clears on browser refresh, controls dashboard visibility
    dcc.Store(id="dash-active", storage_type="memory", data=False),
    dcc.Store(id="email-pref", storage_type="memory", data=True),
    dcc.Interval(id="ticker", interval=1500, n_intervals=0),
])])


# ── Callbacks ────────────────────────────────────────────────────────────

@app.callback(
    Output("download-pending", "data"),
    Input("download-pending-btn", "n_clicks"),
    prevent_initial_call=True,
)
def download_pending_claims(n_clicks):
    with rpa_lock:
        analysis = rpa.get("analysis")
    if not analysis or not analysis.get("pending"):
        return no_update
    import pandas as pd
    cols = {
        "facility":    "Facility",
        "analyst":     "Analyst Name",
        "claim":       "Claim Number",
        "score":       "Quality Score",
        "hours":       "Hrs Since Audit",
        "correction":  "Correction Made",
        "rebuttal":    "Rebuttal",
        "aud_comment": "Auditor Comment on Rebuttal",
        "last_reb_dt": "Last Rebuttal Date",
        "last_res_dt": "Last Rebuttal Response Date",
    }
    df = pd.DataFrame(analysis["pending"])[list(cols.keys())].rename(columns=cols)
    return dcc.send_data_frame(df.to_excel, "pending_claims.xlsx", index=False, sheet_name="Pending")


@app.callback(Output("date-display", "children"), Input("ticker", "n_intervals"))
def update_date(_):
    today = date.today()
    return f"{today.replace(day=1).strftime('%m/%d/%Y')}   to   {today.strftime('%m/%d/%Y')}"


@app.callback(
    Output("email-pref", "data"),
    Input("email-toggle", "value"),
    prevent_initial_call=False,
)
def sync_email_pref(value):
    enabled = bool(value)
    with rpa_lock:
        rpa["send_email"] = enabled
    return enabled


@app.callback(
    Output("run-btn", "disabled"),
    Output("run-btn", "style"),
    Output("run-all-btn", "disabled"),
    Output("run-all-btn", "style"),
    Output("status-badge", "children"),
    Output("status-badge", "style"),
    Output("log-display", "children"),
    Output("otp-panel", "style"),
    Output("auth-number-box", "style"),
    Output("auth-number-display", "children"),
    Output("otp-entry-box", "style"),
    Output("dash-active", "data"),
    Input("ticker", "n_intervals"),
    Input("run-btn", "n_clicks"),
    Input("run-all-btn", "n_clicks"),
    State("dash-active", "data"),
    State("email-pref", "data"),
    State("lob-selector", "value"),
    prevent_initial_call=False,
)
def tick_and_run(_, run_clicks, run_all_clicks, dash_currently_active, email_pref, selected_lob):
    if ctx.triggered_id == "run-btn" and run_clicks:
        with rpa_lock:
            if rpa["status"] not in ("running", "otp_needed"):
                rpa["send_email"] = bool(email_pref) if email_pref is not None else True
                today       = date.today()
                month_start = today.replace(day=1)
                lob = selected_lob or "CBOS AR"
                t = threading.Thread(target=rpa_thread, args=(today, month_start, [lob]), daemon=True)
                t.start()
    elif ctx.triggered_id == "run-all-btn" and run_all_clicks:
        with rpa_lock:
            if rpa["status"] not in ("running", "otp_needed"):
                rpa["send_email"] = bool(email_pref) if email_pref is not None else True
                today       = date.today()
                month_start = today.replace(day=1)
                t = threading.Thread(target=rpa_thread, args=(today, month_start, RUN_ALL_ORDER), daemon=True)
                t.start()

    with rpa_lock:
        status      = rpa["status"]
        log_lines   = list(rpa["log"])
        auth_number = rpa.get("auth_number")

    busy = status in ("running", "otp_needed")
    color = STATUS_COLOR.get(status, "#95a5a6")
    btn_style = {
        "padding": "11px 32px", "color": "white", "border": "none",
        "borderRadius": "6px", "cursor": "default" if busy else "pointer",
        "fontSize": "15px", "fontWeight": "bold", "letterSpacing": ".5px",
        "backgroundColor": "#95a5a6" if busy else "#2c3e50",
        "opacity": "0.7" if busy else "1",
    }
    run_all_btn_style = {
        "padding": "11px 28px", "color": "white", "border": "none",
        "borderRadius": "6px", "cursor": "default" if busy else "pointer",
        "fontSize": "14px", "fontWeight": "bold", "letterSpacing": ".5px",
        "backgroundColor": "#95a5a6" if busy else "#117a65",
        "opacity": "0.7" if busy else "1",
    }
    badge_style = {
        "padding": "3px 14px", "borderRadius": "12px", "color": "white",
        "fontSize": "12px", "fontWeight": "bold", "backgroundColor": color,
    }
    log_text = "\n".join(log_lines) if log_lines else "Press Run Report to start."

    # Outer panel visibility
    otp_panel_style = {"display": "block", "marginTop": "16px"} \
                      if status == "otp_needed" else {"display": "none"}

    # Authenticator number box vs OTP text entry
    if status == "otp_needed" and auth_number:
        auth_box_style  = {"display": "block"}
        otp_entry_style = {"display": "none"}
    elif status == "otp_needed" and not auth_number:
        auth_box_style  = {"display": "none"}
        otp_entry_style = {"display": "block"}
    else:
        auth_box_style  = {"display": "none"}
        otp_entry_style = {"display": "none"}

    label = "Running..." if status == "running" else \
            "Approve in Authenticator..." if (status == "otp_needed" and auth_number) else \
            "Waiting for OTP..." if status == "otp_needed" else "Run Report"

    # Activate dashboard only when a new download completes in this session
    dash_active = True if status == "done" else (dash_currently_active or False)

    return (busy, btn_style,
            busy, run_all_btn_style,
            status.replace("_", " ").title(), badge_style,
            log_text,
            otp_panel_style, auth_box_style, auth_number or "", otp_entry_style,
            dash_active)


@app.callback(
    Output("otp-msg", "children"),
    Input("otp-submit", "n_clicks"),
    State("otp-input", "value"),
    prevent_initial_call=True,
)
def submit_otp(_, val):
    if val and len(str(val).strip()) == 6:
        with rpa_lock:
            rpa["otp_value"] = str(val).strip()
            rpa["status"]    = "running"
        return "OTP submitted."
    return "Enter a valid 6-digit OTP."


@app.callback(
    Output("dash-active", "data", allow_duplicate=True),
    Input("reset-btn", "n_clicks"),
    prevent_initial_call=True,
)
def reset_app(_):
    with rpa_lock:
        rpa["stop_requested"] = True
        rpa["status"]         = "idle"
        rpa["log"]            = ["[Reset] App reset to idle."]
        rpa["otp_value"]      = None
        rpa["auth_number"]    = None
        rpa["analysis"]       = None
        rpa["analyses"]       = {}
        rpa["current_lob"]    = "CBOS AR"
    return False   # hides dashboard


CARD_COLORS = {
    "CBOS_Seneca Healthcare District":   "#1a6fa8",
    "CBOS_Van Diest Medical Center":     "#117a65",
    "CBOS_Memorial Health":              "#6c3483",
    "CBOS_Gila Regional Medical Center": "#784212",
    "CBOS_Arkansas Valley":              "#c0392b",
    "CBOS_Hardtner Medical Center":      "#1abc9c",
    "CBOS_Lincoln":                       "#d35400",
    "CBOS_Modoc Medical Center":         "#2c3e50",
}

LOB_TAB_COLORS = {
    "CBOS AR":                "#2c3e50",
    "CashPosting-CBOS":       "#1a6fa8",
    "Claim Processing -CBOS": "#117a65",
    "Credit Balance-CBOS":    "#6c3483",
}


def _build_cards(analysis):
    th_style = {
        "padding": "8px 14px", "textAlign": "center",
        "background": "#2c3e50", "color": "white",
        "fontSize": "12px", "fontWeight": "bold",
        "border": "1px solid #dde3e8", "whiteSpace": "nowrap",
    }
    th_left = {**th_style, "textAlign": "left"}
    rows = []
    for fac, s in analysis["summary"].items():
        color       = CARD_COLORS.get(fac, "#2c3e50")
        score_color = "#27ae60" if s["errors"] == 0 else ("#e67e22" if s["pending"] == 0 else "#e74c3c")
        err_color   = "#e74c3c" if s["errors"]  else "#27ae60"
        pend_color  = "#e74c3c" if s["pending"] else "#27ae60"
        td = {"padding": "7px 14px", "border": "1px solid #dde3e8", "fontSize": "13px"}
        rows.append(html.Tr([
            html.Td(fac.replace("CBOS_", ""), style={**td, "fontWeight": "bold",
                                                      "color": "white", "background": color,
                                                      "whiteSpace": "nowrap"}),
            html.Td(str(s["total"]),       style={**td, "textAlign": "center"}),
            html.Td(str(s["auto_accepted"]), style={**td, "textAlign": "center", "color": "#27ae60", "fontWeight": "bold"}),
            html.Td(s["pct"],              style={**td, "textAlign": "center", "color": score_color, "fontWeight": "bold"}),
            html.Td(str(s["errors"]),      style={**td, "textAlign": "center", "color": err_color,   "fontWeight": "bold"}),
            html.Td(str(s["pending"]),     style={**td, "textAlign": "center", "color": pend_color,  "fontWeight": "bold"}),
        ]))
    table = html.Table([
        html.Thead(html.Tr([
            html.Th("Facility",       style=th_left),
            html.Th("Total Audits",   style=th_style),
            html.Th("Auto Accepted",  style=th_style),
            html.Th("Quality Score",  style=th_style),
            html.Th("Errors",         style=th_style),
            html.Th("No Correction",  style=th_style),
        ])),
        html.Tbody(rows),
    ], style={
        "borderCollapse": "collapse", "width": "100%",
        "boxShadow": "0 1px 4px rgba(0,0,0,.08)", "borderRadius": "8px",
        "overflow": "hidden", "marginBottom": "16px",
    })
    return [table]


# ── LOB slicer tab click → update store ──────────────────────────────────
@app.callback(
    Output("dashboard-lob", "data"),
    [Input(f"lob-tab-{lob.replace(' ', '_').replace('-', '_')}", "n_clicks") for lob in LOB_LIST],
    State("dashboard-lob", "data"),
    prevent_initial_call=True,
)
def lob_tab_clicked(*args):
    # args = n_clicks for each LOB tab + current store value
    store_val = args[-1]
    triggered = ctx.triggered_id
    if not triggered:
        return store_val
    # Map button id back to LOB name
    for lob in LOB_LIST:
        btn_id = f"lob-tab-{lob.replace(' ', '_').replace('-', '_')}"
        if triggered == btn_id:
            return lob
    return store_val


def _lob_btn_style(lob, is_active, is_visible):
    if not is_visible:
        return {"display": "none"}
    bg = LOB_TAB_COLORS.get(lob, "#2c3e50")
    return {
        "padding": "6px 16px", "borderRadius": "20px", "fontSize": "12px",
        "fontWeight": "bold", "cursor": "pointer", "border": "none",
        "backgroundColor": bg if is_active else "#dde3e8",
        "color": "white" if is_active else "#555",
        "boxShadow": "0 2px 4px rgba(0,0,0,.15)" if is_active else "none",
        "display": "inline-block",
    }


@app.callback(
    Output("dashboard-section",   "style"),
    *[Output(f"lob-tab-{lob.replace(' ', '_').replace('-', '_')}", "style") for lob in LOB_LIST],
    Output("quality-cards",       "children"),
    Output("pending-table",             "data"),
    Output("dashboard-run-time",        "children"),
    Output("process-validation-banner", "children"),
    Output("dashboard-msg",             "children"),
    Input("dash-active",          "data"),
    Input("ticker",               "n_intervals"),
    Input("dashboard-lob",        "data"),
)
def update_dashboard(dash_active, _, selected_lob):

    hidden    = {"marginTop": "18px", "display": "none"}
    shown     = {"marginTop": "18px", "display": "block"}
    hide_btn  = {"display": "none"}
    n_lobs    = len(LOB_LIST)

    if not dash_active:
        return (hidden, *([hide_btn] * n_lobs), [], [], "", html.Div(), "")

    with rpa_lock:
        analyses  = dict(rpa.get("analyses", {}))
        last_lob  = rpa.get("current_lob", "CBOS AR")

    if not analyses:
        return (hidden, *([hide_btn] * n_lobs), [], [], "", html.Div(), "")

    # Default to the last-run LOB if nothing explicitly chosen or choice no longer available
    display_lob = selected_lob if (selected_lob and selected_lob in analyses) else last_lob
    if display_lob not in analyses:
        display_lob = next(lob for lob in LOB_LIST if lob in analyses)

    # ── Build per-button styles ───────────────────────────────────────────
    btn_styles = [_lob_btn_style(lob, lob == display_lob, lob in analyses) for lob in LOB_LIST]

    analysis     = analyses[display_lob]
    cards        = _build_cards(analysis)
    pending_rows = analysis.get("pending", [])
    run_time_str = f"Showing: {display_lob}  |  Last updated: {analysis.get('run_time', '')}"
    msg = (f"{len(pending_rows)} claim(s) with score < 100% and no correction made."
           if pending_rows else "All errors have corrections recorded.")

    # ── Process validation banner ─────────────────────────────────────────
    processes = analysis.get("processes", [])
    badge_style = {
        "display": "inline-block", "padding": "3px 10px",
        "borderRadius": "12px", "background": "#eaf4fb",
        "border": "1px solid #aed6f1", "fontSize": "12px",
        "color": "#1a5276", "marginRight": "6px", "marginBottom": "4px",
    }
    email_badge_style = {
        "display": "inline-block", "padding": "3px 10px",
        "borderRadius": "12px", "background": "#fef9e7",
        "border": "1px solid #f9ca24", "fontSize": "12px",
        "color": "#7d6608", "marginRight": "6px", "marginBottom": "4px",
    }
    label_style = {
        "fontWeight": "bold", "fontSize": "13px",
        "color": "#555", "marginRight": "8px",
    }

    # Build email recipient badges for this LOB
    email_children = []
    if display_lob == "CBOS AR":
        # Per-facility recipients
        fac_lines = []
        for fac, email in FACILITY_EMAILS.items():
            short = fac.replace("CBOS_", "")
            fac_lines.append(html.Span(f"{short} → {email}", style=email_badge_style))
        email_children = [html.Span("Email recipients (per facility): ", style=label_style)] + fac_lines
    else:
        recipients = LOB_EMAIL_TO.get(display_lob, [])
        if recipients:
            email_children = (
                [html.Span("Email recipients: ", style=label_style)]
                + [html.Span(r, style=email_badge_style) for r in recipients]
            )
        else:
            email_children = [html.Span("Email recipients: ", style=label_style),
                              html.Span("(none configured)", style={**email_badge_style, "color": "#999"})]

    banner_rows = []
    if processes:
        banner_rows.append(html.Div([
            html.Span("Processes in this report: ", style=label_style),
            *[html.Span(p, style=badge_style) for p in processes],
        ], style={"marginBottom": "8px"}))
    if email_children:
        banner_rows.append(html.Div(email_children))

    process_banner = html.Div(banner_rows, style={
        "background": "white", "borderRadius": "8px",
        "padding": "10px 16px", "marginBottom": "10px",
        "boxShadow": "0 1px 4px rgba(0,0,0,.08)",
    }) if banner_rows else html.Div()

    return (shown, *btn_styles, cards, pending_rows, run_time_str, process_banner, msg)


def scheduled_run():
    """Triggered by scheduler — starts RPA if not already running."""
    with rpa_lock:
        if rpa["status"] in ("running", "otp_needed"):
            log("Scheduled run skipped — already running.")
            return
        rpa["send_email"] = True  # scheduled runs always send email
    today       = date.today()
    month_start = today.replace(day=1)
    log(f"Scheduled run started at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} IST")
    t = threading.Thread(target=rpa_thread, args=(today, month_start), daemon=True)
    t.start()


def setup_scheduler():
    from apscheduler.schedulers.background import BackgroundScheduler
    from apscheduler.triggers.cron import CronTrigger
    import pytz

    ist = pytz.timezone("Asia/Kolkata")
    scheduler = BackgroundScheduler(timezone=pytz.utc)

    # Daily 7:00 PM IST
    scheduler.add_job(
        scheduled_run,
        CronTrigger(hour=19, minute=0, timezone=ist),
        id="run_7pm_ist",
        name="Daily 7 PM IST",
        replace_existing=True,
    )
    print("Scheduled: Daily 7:00 PM IST")

    # Daily 12:00 AM IST (midnight)
    scheduler.add_job(
        scheduled_run,
        CronTrigger(hour=0, minute=0, timezone=ist),
        id="run_12am_ist",
        name="Daily 12 AM IST",
        replace_existing=True,
    )
    print("Scheduled: Daily 12:00 AM IST")

    # Daily 2:00 AM IST
    scheduler.add_job(
        scheduled_run,
        CronTrigger(hour=2, minute=0, timezone=ist),
        id="run_2am_ist",
        name="Daily 2 AM IST",
        replace_existing=True,
    )
    print("Scheduled: Daily 2:00 AM IST")

    scheduler.start()
    return scheduler


if __name__ == "__main__":
    print(f"Starting QA Suite Report Runner...")
    print(f"Open http://localhost:{PORT} in your browser")
    scheduler = setup_scheduler()
    app.run(debug=False, port=PORT)
