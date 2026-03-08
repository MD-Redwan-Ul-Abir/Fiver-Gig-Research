"""Quick test: scrape only 3 sub-niches to verify everything works.
This uses macOS notifications and auto-polling — no input() needed.
"""
import time, random, re, subprocess, os, json

_BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
_CONFIG_FILE = os.path.join(_BASE_DIR, "hub_config.json")

def _load_threshold():
    try:
        with open(_CONFIG_FILE) as f:
            return int(json.load(f).get("threshold", 1800))
    except Exception:
        return 1800

THRESHOLD = _load_threshold()

TEST_QUERIES = [
    "Corporate Company Website",
    "AI SaaS Platform Development",
    "WordPress Blog Content Website",
]

def notify(title, message):
    """Send a macOS notification."""
    try:
        script = f'display notification "{message}" with title "{title}" sound name "Glass"'
        subprocess.run(["osascript", "-e", script], timeout=5)
    except Exception:
        pass

def bring_chrome_front():
    """Bring Chrome window to the foreground."""
    try:
        subprocess.run(["osascript", "-e",
            'tell application "Google Chrome" to activate'], timeout=5)
    except Exception:
        pass

def is_captcha_page(html, url=""):
    """Return True only when Fiverr is actively blocking with a challenge page."""
    # If we're on a real Fiverr page URL → not a captcha
    if any(u in url for u in ["/search/gigs", "/categories/", "/gigs/search"]):
        return False
    html_l = html.lower()
    # Must have the challenge TITLE
    is_challenge = (
        "it needs a human touch" in html_l or
        "complete the task and we" in html_l
    )
    # AND the active challenge widget element
    has_widget = (
        'id="px-captcha"' in html or
        'class="px-loader-wrapper"' in html
    )
    return is_challenge and has_widget

def safe_content(page, retries=3):
    for i in range(retries):
        try:
            time.sleep(1.0)
            return page.content()
        except Exception:
            if i < retries - 1:
                time.sleep(2)
    return ""

def wait_for_captcha_clear(page, label="", timeout_minutes=6):
    """Poll until CAPTCHA disappears. Notify user every minute."""
    notify("Fiverr CAPTCHA", "Please solve the CAPTCHA in Chrome window to continue scraping")
    bring_chrome_front()

    for i in range(timeout_minutes * 12):  # check every 5s
        time.sleep(5)
        current_url = ""
        try:
            current_url = page.url
        except Exception:
            pass
        html = safe_content(page)
        if html and not is_captcha_page(html, current_url):
            print(f"  ✓ CAPTCHA solved! URL: {current_url[:60]}")
            notify("Fiverr Scraper", "CAPTCHA solved — searching Fiverr now...")
            time.sleep(2)
            return True
        if i % 12 == 0 and i > 0:
            mins = i * 5 // 60
            print(f"  Still waiting ({mins} min)... URL: {current_url[:50]}")
            notify("Fiverr CAPTCHA", f"Still waiting — please solve CAPTCHA in Chrome ({mins} min elapsed)")
            bring_chrome_front()
    print("  CAPTCHA wait timed out.")
    return False

def extract_gig_count(html, page):
    # Strategy 1: Visible "84,000+ results" (note the + sign Fiverr uses)
    for pat in [r'([\d,]+)\+?\s+[Rr]esults', r'([\d,]+)\+?\s+[Ss]ervices', r'([\d,]+)\+?\s+[Gg]igs']:
        m = re.search(pat, html)
        if m:
            raw = m.group(1).replace(",","")
            if raw.isdigit() and int(raw) > 0: return int(raw)
    # Strategy 2: JSON "total" key in embedded page JSON (Fiverr uses "total":84073)
    for pat in [r'"total"\s*:\s*(\d+)', r'"total_count"\s*:\s*(\d+)']:
        for m in re.finditer(pat, html):
            val = int(m.group(1))
            if 10 <= val <= 500000: return val
    # Strategy 3: JavaScript DOM walk
    try:
        c = page.evaluate("""() => {
            try {
                const s = window.__INITIAL_STATE__;
                if (s && s.search && s.search.total_count) return s.search.total_count;
            } catch(e) {}
            const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
            let node;
            while ((node = walker.nextNode())) {
                const t = node.textContent.trim();
                const m = t.match(/^(\\d[\\d,]*)\\+?\\s+(results|services|gigs)$/i);
                if (m) return parseInt(m[1].replace(/,/g,''));
            }
            return null;
        }""")
        if c: return int(c)
    except: pass
    return None

def extract_stats(page, html):
    reviews, queues = [], []
    try:
        data = page.evaluate("""() => {
            const out = [];
            for (const el of document.querySelectorAll('*')) {
                if (el.children.length > 0) continue;
                const t = el.textContent.trim();
                const r = t.match(/^\\((\\d[\\d,]*)\\)$/);
                if (r) out.push({type:'r', val: parseInt(r[1].replace(/,/g,''))});
                const q = t.match(/^(\\d+)\\s+[Oo]rders? in [Qq]ueue$/);
                if (q) out.push({type:'q', val: parseInt(q[1])});
            }
            return out;
        }""")
        for item in (data or []):
            if item["type"]=="r" and 0 < item["val"] < 100000: reviews.append(item["val"])
            elif item["type"]=="q" and 0 < item["val"] < 10000: queues.append(item["val"])
    except: pass
    if not reviews:
        for m in re.finditer(r'\((\d[\d,]*)\)', html):
            v = int(m.group(1).replace(",",""))
            if 1 <= v <= 50000: reviews.append(v)
        reviews = list(set(reviews))[:20]
    if not queues:
        for m in re.finditer(r'(\d+)\s+[Oo]rders? in [Qq]ueue', html):
            queues.append(int(m.group(1)))
    avg_r = round(sum(reviews)/len(reviews),1) if reviews else None
    avg_q = round(sum(queues)/len(queues),1)   if queues  else None
    return avg_r, avg_q

# ── Run ──────────────────────────────────────────────────────────────────────
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

print("="*55)
print("  Fiverr Scraper - Test (3 sub-niches)")
print("="*55)
notify("Fiverr Scraper", "Opening Chrome for Fiverr scraping test...")

with sync_playwright() as p:
    try:
        browser = p.chromium.launch(
            channel="chrome", headless=False,
            args=["--start-maximized","--disable-blink-features=AutomationControlled",
                  "--no-first-run","--no-default-browser-check","--disable-extensions"]
        )
        print("Using system Chrome")
    except Exception:
        browser = p.chromium.launch(
            headless=False,
            args=["--start-maximized","--disable-blink-features=AutomationControlled","--no-sandbox"]
        )
        print("Using Playwright Chromium")

    ctx = browser.new_context(
        user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        viewport={"width":1400,"height":900}, locale="en-US",
        timezone_id="America/New_York",
    )
    ctx.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
        Object.defineProperty(navigator, 'plugins',   { get: () => [1,2,3,4,5] });
        Object.defineProperty(navigator, 'languages', { get: () => ['en-US','en'] });
        window.chrome = { runtime: {} };
    """)
    page = ctx.new_page()

    print("\nOpening Fiverr homepage...")
    try:
        page.goto("https://www.fiverr.com", wait_until="load", timeout=40000)
    except Exception as e:
        print(f"  Note: {e}")
    time.sleep(4)
    bring_chrome_front()

    html = safe_content(page)
    cur_url = ""
    try: cur_url = page.url
    except: pass
    if is_captcha_page(html, cur_url):
        print("\n⚠️  CAPTCHA on homepage! A Chrome window has opened on your screen.")
        print("  Please solve the CAPTCHA in that Chrome window.")
        print("  (This script will automatically continue once you solve it)\n")
        wait_for_captcha_clear(page, "homepage")
    else:
        print(f"  OK — Fiverr loaded (URL: {cur_url[:60]})")
        notify("Fiverr Scraper", "Fiverr loaded — starting searches now")
        time.sleep(2)

    print("\nStarting 3 test searches...\n")

    results = []
    for i, q in enumerate(TEST_QUERIES, 1):
        query = q.replace(" ","+")
        url = f"https://www.fiverr.com/search/gigs?query={query}&source=top-bar"
        print(f"[{i}/3] Searching: {q}")
        try:
            page.goto(url, wait_until="load", timeout=40000)
        except Exception as e:
            print(f"  Nav error: {e}")
        time.sleep(random.uniform(4, 6))
        bring_chrome_front()

        html = safe_content(page)
        print(f"  URL: {page.url[:70]}")

        if is_captcha_page(html, page.url if page else ""):
            print("  ⚠️  CAPTCHA mid-scrape! Solve in Chrome window...")
            wait_for_captcha_clear(page, q)
            try:
                page.goto(url, wait_until="load", timeout=40000)
                time.sleep(4)
                html = safe_content(page)
            except Exception: pass

        count = extract_gig_count(html, page)
        avg_r, avg_q = extract_stats(page, html)
        verdict = "✅ Positive" if (count and count <= THRESHOLD) else "❌ Negative" if count else "❓ N/A"

        print(f"  Gig count   : {count:,}" if count else "  Gig count   : NOT FOUND")
        print(f"  Avg reviews : {avg_r}")
        print(f"  Avg queue   : {avg_q}")
        print(f"  Result      : {verdict}\n")
        results.append((q, count, avg_r, avg_q, verdict))
        time.sleep(random.uniform(4, 6))

    browser.close()

print("="*55)
print("  TEST COMPLETE")
print("="*55)
for q, cnt, r, qu, v in results:
    print(f"  {q[:35]:35s} | {str(cnt):>8} gigs | {v}")
notify("Fiverr Scraper", f"Test complete! Searched {len(results)} sub-niches.")
