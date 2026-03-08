"""Debug: save Fiverr search page HTML to disk so we can find the gig count element."""
import time, re, subprocess
from playwright.sync_api import sync_playwright

def notify(msg):
    try:
        subprocess.run(["osascript","-e",f'display notification "{msg}" with title "Fiverr Debug"'],timeout=3)
    except: pass

def bring_chrome():
    try:
        subprocess.run(["osascript","-e",'tell application "Google Chrome" to activate'],timeout=3)
    except: pass

def is_captcha(html, url=""):
    if any(u in url for u in ["/search/gigs", "/categories/"]):
        return False
    hl = html.lower()
    return ("it needs a human touch" in hl or "complete the task and we" in hl) and ('id="px-captcha"' in html or 'class="px-loader-wrapper"' in html)

with sync_playwright() as p:
    try:
        browser = p.chromium.launch(channel="chrome", headless=False,
            args=["--start-maximized","--disable-blink-features=AutomationControlled"])
    except:
        browser = p.chromium.launch(headless=False,
            args=["--start-maximized","--disable-blink-features=AutomationControlled","--no-sandbox"])

    ctx = browser.new_context(
        user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        viewport={"width":1400,"height":900}, locale="en-US")
    ctx.add_init_script("Object.defineProperty(navigator,'webdriver',{get:()=>undefined});window.chrome={runtime:{}};")
    page = ctx.new_page()

    notify("Opening Fiverr for debug...")
    page.goto("https://www.fiverr.com", wait_until="load", timeout=40000)
    time.sleep(4)
    bring_chrome()

    html = ""
    try: html = page.content()
    except: pass

    if is_captcha(html, page.url):
        print("CAPTCHA detected — solve it in Chrome, script will continue...")
        notify("CAPTCHA! Solve in Chrome")
        for _ in range(72):
            time.sleep(5)
            try: cur_url = page.url
            except: cur_url = ""
            try: html = page.content()
            except: html = ""
            if html and not is_captcha(html, cur_url):
                print(f"CAPTCHA cleared! URL: {cur_url}")
                break

    # Now do the search
    url = "https://www.fiverr.com/search/gigs?query=WordPress+website+development&source=top-bar"
    print(f"\nGoing to search: {url}")
    try:
        page.goto(url, wait_until="load", timeout=40000)
    except Exception as e:
        print(f"nav: {e}")

    # Wait longer for dynamic content to load
    print("Waiting 8s for JS to render...")
    time.sleep(8)
    bring_chrome()

    # Try waiting for specific elements
    for wait_sel in ["[class*='count']","[class*='result']","h1","h2","h3","main"]:
        try:
            page.wait_for_selector(wait_sel, timeout=5000)
            print(f"Element '{wait_sel}' appeared")
            break
        except: pass

    time.sleep(2)
    try:
        html = page.content()
    except Exception as e:
        print(f"content error: {e}")
        html = ""

    print(f"\nPage URL: {page.url}")
    print(f"HTML length: {len(html)}")

    # Save full HTML
    with open("/tmp/fiverr_search_debug.html","w",encoding="utf-8") as f:
        f.write(html)
    print("Saved full HTML to /tmp/fiverr_search_debug.html")

    # Try to find count patterns
    print("\n--- Searching for count patterns ---")
    patterns = [
        r'([\d,]+)\s+[Ss]ervices',
        r'([\d,]+)\s+[Rr]esults',
        r'([\d,]+)\s+[Gg]igs',
        r'"total_count"\s*:\s*(\d+)',
        r'"count"\s*:\s*(\d+)',
        r'totalCount["\s]*:\s*(\d+)',
        r'gigs_count["\s]*:\s*(\d+)',
        r'"total"\s*:\s*(\d+)',
        r'(\d[\d,]+)\s+freelancers',
    ]
    for pat in patterns:
        m = re.search(pat, html)
        if m:
            print(f"  FOUND: '{pat}' → {m.group(0)}")

    # Try JavaScript
    print("\n--- JavaScript extraction ---")
    try:
        result = page.evaluate("""() => {
            // Try window state
            try { if(window.__INITIAL_STATE__) return JSON.stringify(window.__INITIAL_STATE__).substring(0,500); } catch(e){}
            // Try React fiber
            try {
                const el = document.querySelector('[class*="result"], [class*="count"], h1, h2');
                if(el) return el.textContent;
            } catch(e){}
            return 'no data found';
        }""")
        print(f"  JS result: {str(result)[:300]}")
    except Exception as e:
        print(f"  JS error: {e}")

    # Print all visible text that contains digits near 'service/result/gig'
    print("\n--- All text elements with potential counts ---")
    try:
        texts = page.evaluate("""() => {
            const out = [];
            document.querySelectorAll('*').forEach(el => {
                if(el.children.length === 0) {
                    const t = el.textContent.trim();
                    if(/\\d/.test(t) && t.length < 60 && /service|result|gig|found|available/i.test(t))
                        out.push(t);
                }
            });
            return out.slice(0,30);
        }""")
        for t in (texts or []):
            print(f"  {repr(t)}")
    except Exception as e:
        print(f"  error: {e}")

    browser.close()
print("\nDebug complete. Check /tmp/fiverr_search_debug.html")
