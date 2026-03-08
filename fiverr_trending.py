#!/usr/bin/env python3
"""
Fiverr Trending Services Scanner
────────────────────────────────────────────────────────────────────
Scans Fiverr category pages and homepage for currently popular/
trending services and gig categories.

Writes results to trending_results.json in the project folder.
"""

import os, sys, time, json
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_FILE = os.path.join(BASE_DIR, "trending_results.json")

PAGES = [
    ("https://www.fiverr.com",                                                         "Homepage"),
    ("https://www.fiverr.com/categories/programming-tech",                             "Programming & Tech"),
    ("https://www.fiverr.com/categories/digital-marketing",                            "Digital Marketing"),
    ("https://www.fiverr.com/categories/programming-tech/ai-coding",                   "AI Coding"),
    ("https://www.fiverr.com/categories/programming-tech/web-programming",             "Web Dev"),
    ("https://www.fiverr.com/categories/programming-tech/chatbot-development",         "Chatbot Dev"),
    ("https://www.fiverr.com/categories/digital-marketing/search-engine-optimization", "SEO"),
    ("https://www.fiverr.com/categories/digital-marketing/social-media-marketing",     "Social Media"),
]


def extract_page(page, url, section):
    items = []
    try:
        page.goto(url, wait_until="load", timeout=36000)
        time.sleep(3)

        data = page.evaluate("""() => {
            const seen = new Set();
            const items = [];

            function add(text, type) {
                const t = (text || '').trim();
                const key = t.toLowerCase().slice(0, 60);
                if (t && t.length > 3 && t.length < 90 && !seen.has(key)) {
                    seen.add(key);
                    items.push({ text: t, type });
                }
            }

            // Popular/trending chips on homepage
            ['[class*="popular"] a', '[class*="trending"] a',
             '[class*="suggestion"] a', '[class*="search-tag"] a'].forEach(sel => {
                document.querySelectorAll(sel).forEach(el => add(el.innerText, 'popular'));
            });

            // Subcategory cards
            ['[class*="sub-category"] h3', '[class*="subcategory"] h3',
             '[class*="category-card"] h3', '[class*="category-card"] p',
             '[class*="service-card"] h3', 'a[class*="category-link"]'].forEach(sel => {
                document.querySelectorAll(sel).forEach(el => add(el.innerText, 'category'));
            });

            // Top gig titles (first 10)
            const gigSels = ['[class*="gig-card"] h3', '[class*="GigCard"] h3',
                             '[class*="gig-title"]', 'article h3'];
            for (const sel of gigSels) {
                const els = document.querySelectorAll(sel);
                if (els.length > 0) {
                    [...els].slice(0, 10).forEach(el => add(el.innerText, 'gig'));
                    break;
                }
            }

            // "Best Selling" / "Editor's Choice" labels
            document.querySelectorAll('[class*="badge"], [class*="label"], [class*="tag"]').forEach(el => {
                const t = (el.innerText || '').toLowerCase();
                if (['best seller', 'trending', "editor's choice", 'top rated'].some(k => t.includes(k))) {
                    const card = el.closest('article, [class*="card"]');
                    if (card) {
                        const h = card.querySelector('h3, h2, [class*="title"]');
                        if (h) add(h.innerText, 'hot');
                    }
                }
            });

            return items.slice(0, 35);
        }""")

        for d in data:
            d["section"] = section
            d["url"]     = url
            items.append(d)
            print(f"  [{section:22s}] {d['type'].upper():10s}  {d['text'][:58]}")

    except Exception as e:
        print(f"  [{section}] Error: {e}")

    return items


def run_trending():
    print(f"\n{'='*62}")
    print(f"  Fiverr Trending Services Scanner")
    print(f"{'='*62}\n")

    all_items = []

    with sync_playwright() as pw:
        try:
            browser = pw.chromium.launch(
                channel="chrome", headless=False,
                args=["--start-maximized",
                      "--disable-blink-features=AutomationControlled",
                      "--no-first-run", "--no-default-browser-check"])
        except Exception:
            browser = pw.chromium.launch(
                headless=False,
                args=["--start-maximized",
                      "--disable-blink-features=AutomationControlled",
                      "--no-sandbox"])

        ctx = browser.new_context(
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            viewport={"width": 1400, "height": 900},
            locale="en-US",
        )
        ctx.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
            Object.defineProperty(navigator, 'plugins',   { get: () => [1,2,3,4,5] });
            window.chrome = { runtime: {} };
        """)
        page = ctx.new_page()

        for url, section in PAGES:
            print(f"\n  Scanning: {section} …")
            items = extract_page(page, url, section)
            all_items.extend(items)
            time.sleep(random.uniform(2.5, 4.0) if len(PAGES) > 1 else 0)

        browser.close()

    # Deduplicate
    seen  = set()
    unique = []
    for d in all_items:
        k = d["text"].lower()[:50]
        if k not in seen:
            seen.add(k)
            unique.append(d)

    # Group by type for summary
    by_type = {}
    for d in unique:
        by_type.setdefault(d["type"], []).append(d["text"])

    print(f"\n{'='*62}")
    print(f"  Scan complete — {len(unique)} unique trending items found")
    for t, items in by_type.items():
        print(f"    {t.upper():12s}  {len(items)} items")
    print(f"  Results saved to: trending_results.json")
    print(f"  View in the Trends page of Fiverr Research Hub.")
    print(f"{'='*62}\n")

    result = {
        "timestamp": __import__("datetime").datetime.now().isoformat(),
        "total":     len(unique),
        "by_type":   {t: len(v) for t, v in by_type.items()},
        "items":     unique,
    }
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)


import random
if __name__ == "__main__":
    run_trending()
