"""
Debug script — captures modal HTML so we can identify the date picker structure.
Run once, log in manually, then check debug_output.txt and modal_screenshot.png
"""

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

URL = "https://qualityauditsuite.r1rcm.com/"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, slow_mo=600)
    page = browser.new_page(viewport={"width": 1400, "height": 900})
    page.goto(URL, wait_until="domcontentloaded")

    print("Log in manually, then wait...")
    page.wait_for_selector("text=Reports", timeout=300_000)
    print("Login detected.")

    page.click("text=Reports")
    page.wait_for_load_state("networkidle")
    page.wait_for_selector("text=Audit Report")
    page.click("text=Audit Report")
    page.wait_for_load_state("networkidle")
    page.wait_for_selector("text=Request Report")
    page.click("text=Request Report")
    page.wait_for_selector("text=Audited Date", timeout=30_000)
    page.wait_for_timeout(1500)  # let modal fully render

    # Screenshot
    page.screenshot(path="modal_screenshot.png")
    print("Screenshot saved: modal_screenshot.png")

    # Dump full page HTML (just the body)
    html = page.evaluate("document.body.innerHTML")
    with open("debug_output.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("HTML saved: debug_output.html")

    # List all inputs
    inputs = page.locator("input")
    count = inputs.count()
    print(f"\nFound {count} input elements:")
    for i in range(count):
        inp = inputs.nth(i)
        try:
            typ   = inp.get_attribute("type") or "(none)"
            name  = inp.get_attribute("name") or ""
            ph    = inp.get_attribute("placeholder") or ""
            cls   = inp.get_attribute("class") or ""
            val   = inp.input_value()
            vis   = inp.is_visible()
            print(f"  [{i}] type={typ} name={name!r} placeholder={ph!r} value={val!r} visible={vis}")
            print(f"       class={cls[:80]}")
        except Exception as e:
            print(f"  [{i}] ERROR: {e}")

    # HTML around "Audited Date"
    audited_html = page.evaluate("""
        () => {
            const el = [...document.querySelectorAll('*')]
                .find(e => e.textContent.trim() === 'Audited Date *' || e.textContent.trim() === 'Audited Date');
            if (!el) return 'NOT FOUND';
            let node = el;
            for (let i = 0; i < 4; i++) node = node.parentElement;
            return node ? node.outerHTML : 'no parent';
        }
    """)
    with open("audited_date_html.txt", "w", encoding="utf-8") as f:
        f.write(audited_html)
    print("\nAudited Date container HTML saved: audited_date_html.txt")

    input("\nDone. Press Enter to close browser...")
    browser.close()
