#!/usr/bin/env python3
"""
Fiverr Niche Discovery
────────────────────────────────────────────────────────────────────
Searches Fiverr for a seed keyword and discovers related sub-niches
from autocomplete, search suggestions, and category pages.

Outputs results to stdout AND writes discovery_results.json.
Usage:  python3 fiverr_niche_discovery.py "your keyword"
"""

import os, sys, re, time, json, random, argparse
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_FILE = os.path.join(BASE_DIR, "discovery_results.json")

CATEGORIES = [
    ("https://www.fiverr.com/categories/programming-tech",    "Programming & Tech"),
    ("https://www.fiverr.com/categories/digital-marketing",  "Digital Marketing"),
]


def safe_goto(page, url, retries=2):
    for i in range(retries):
        try:
            page.goto(url, wait_until="load", timeout=38000)
            return True
        except Exception as e:
            if i < retries - 1:
                time.sleep(3)
            else:
                print(f"    Warning: could not load {url[:60]} — {e}")
    return False


def discover_niches(seed_keyword, max_results=30):
    print(f"\n{'='*62}")
    print(f"  Fiverr Niche Discovery")
    print(f"  Seed keyword: \"{seed_keyword}\"")
    print(f"{'='*62}\n")

    discovered = []

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

        # ── 1. Homepage autocomplete ──────────────────────────────────────
        print("  Opening Fiverr…")
        safe_goto(page, "https://www.fiverr.com")
        time.sleep(3)

        print(f"\n  Step 1 — Autocomplete suggestions for \"{seed_keyword}\"")
        try:
            box = page.locator(
                '[data-testid="search-bar-input"], '
                'input[placeholder*="Search"], input[name="query"]'
            ).first
            if box.count() > 0:
                box.click()
                time.sleep(0.5)
                box.fill(seed_keyword)
                time.sleep(2.5)

                suggestions = page.evaluate("""() => {
                    const items = [];
                    const sels = [
                        '[data-testid*="suggestion"]',
                        '[class*="autocomplete"] li',
                        '[class*="suggestion"] li',
                        '[role="option"]',
                        '[class*="search-suggest"] li',
                        'ul[class*="suggest"] li',
                    ];
                    for (const sel of sels) {
                        const els = document.querySelectorAll(sel);
                        if (els.length > 0) {
                            els.forEach(el => {
                                const t = el.innerText?.trim();
                                if (t && t.length > 2 && t.length < 80) items.push(t);
                            });
                            break;
                        }
                    }
                    return [...new Set(items)].slice(0, 15);
                }""")

                if suggestions:
                    print(f"  Found {len(suggestions)} autocomplete suggestions:")
                    for s in suggestions:
                        print(f"    → {s}")
                        discovered.append({
                            "keyword": s,
                            "type": "autocomplete",
                            "source": "Fiverr autocomplete",
                        })
                else:
                    print("  (No autocomplete suggestions captured)")
        except Exception as e:
            print(f"  Autocomplete error: {e}")

        # ── 2. Search results page — related searches ─────────────────────
        print(f"\n  Step 2 — Search results related tags")
        query = seed_keyword.replace(" ", "+").replace("&", "%26")
        url   = (f"https://www.fiverr.com/search/gigs?query={query}"
                 f"&source=top-bar&search_in=everywhere")
        safe_goto(page, url)
        time.sleep(3)

        try:
            related = page.evaluate("""() => {
                const items = [];
                const sels = [
                    '[data-testid*="related"]',
                    '[class*="related-search"] a',
                    '[class*="people-also"] a',
                    '[class*="suggestion-chip"]',
                    '[class*="search-tag"] a',
                    'a[class*="chip"]',
                    '[class*="tag"] a',
                ];
                for (const sel of sels) {
                    document.querySelectorAll(sel).forEach(el => {
                        const t = (el.innerText || el.textContent || '').trim();
                        if (t && t.length > 2 && t.length < 80) items.push(t);
                    });
                }
                // Also grab subcategory filter chips
                document.querySelectorAll('[class*="filter"] label, [class*="subcategory"] a').forEach(el => {
                    const t = el.innerText?.trim();
                    if (t && t.length > 2 && t.length < 60) items.push(t);
                });
                return [...new Set(items)].slice(0, 25);
            }""")

            if related:
                print(f"  Found {len(related)} related search terms:")
                for r in related:
                    print(f"    → {r}")
                    if not any(d["keyword"].lower() == r.lower() for d in discovered):
                        discovered.append({
                            "keyword": r,
                            "type": "related_search",
                            "source": "Fiverr related searches",
                        })
            else:
                print("  (No related searches captured on this page)")
        except Exception as e:
            print(f"  Related search error: {e}")

        # Try to read result count
        try:
            count_text = page.evaluate("""() => {
                const candidates = [
                    ...document.querySelectorAll('h1, h2, [class*="count"], [data-testid*="count"]')
                ];
                for (const el of candidates) {
                    const t = el.innerText || '';
                    if (/\\d+/.test(t) && /(result|service|gig)/i.test(t)) return t.trim();
                }
                return '';
            }""")
            if count_text:
                print(f"\n  Fiverr results info: {count_text[:100]}")
        except Exception:
            pass

        # ── 3. Category pages ─────────────────────────────────────────────
        print(f"\n  Step 3 — Category pages sub-niche scan")
        kw_words = seed_keyword.lower().split()

        for cat_url, cat_name in CATEGORIES:
            try:
                if not safe_goto(page, cat_url):
                    continue
                time.sleep(2.5)

                sub_cats = page.evaluate("""() => {
                    const items = [];
                    const sels = [
                        '[class*="sub-category"] a',
                        '[class*="subcategory"] a',
                        '[class*="category-card"] h3',
                        '[class*="service-card"] h3',
                        '[class*="leaf"] a',
                        'a[class*="category-link"]',
                    ];
                    for (const sel of sels) {
                        document.querySelectorAll(sel).forEach(el => {
                            const t = (el.innerText || el.textContent || '').trim();
                            if (t && t.length > 2 && t.length < 80) items.push(t);
                        });
                    }
                    return [...new Set(items)].slice(0, 40);
                }""")

                relevant = [
                    s for s in sub_cats
                    if any(w in s.lower() for w in kw_words)
                    and not any(d["keyword"].lower() == s.lower() for d in discovered)
                ]

                if relevant:
                    print(f"  [{cat_name}] {len(relevant)} relevant sub-categories:")
                    for s in relevant[:12]:
                        print(f"    → {s}")
                        discovered.append({
                            "keyword": s,
                            "type": "category",
                            "source": f"Fiverr — {cat_name}",
                        })
                else:
                    print(f"  [{cat_name}] No closely related sub-categories found")
            except Exception as e:
                print(f"  [{cat_name}] Error: {e}")

        browser.close()

    # Deduplicate by keyword (case-insensitive)
    seen  = set()
    uniq  = []
    for d in discovered:
        k = d["keyword"].strip().lower()
        if k not in seen and len(k) > 2:
            seen.add(k)
            uniq.append(d)
    uniq = uniq[:max_results]

    # Save results
    result = {
        "seed_keyword": seed_keyword,
        "timestamp":    __import__("datetime").datetime.now().isoformat(),
        "count":        len(uniq),
        "discoveries":  uniq,
    }
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    print(f"\n{'='*62}")
    print(f"  Discovery complete!")
    print(f"  Total unique sub-niches found: {len(uniq)}")
    print(f"  Saved to: discovery_results.json")
    print(f"  Open Niche Manager → Import Results to add them.")
    print(f"{'='*62}\n")

    return uniq


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Discover Fiverr sub-niches from a seed keyword")
    parser.add_argument("keyword", nargs="?", default="", help="Seed keyword")
    parser.add_argument("--max", type=int, default=30, help="Max results")
    args = parser.parse_args()

    keyword = args.keyword.strip()
    if not keyword:
        print("Enter a seed keyword: ", end="", flush=True)
        keyword = input().strip()

    if keyword:
        discover_niches(keyword, args.max)
    else:
        print("No keyword provided. Exiting.")
        sys.exit(1)
