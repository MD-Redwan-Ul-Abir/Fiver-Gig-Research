#!/usr/bin/env python3
"""
Fiverr AI Gig Creator
──────────────────────────────────────────────────────────────────────
Reads the latest (or a specified) submission_top_gigs CSV produced by
the Submission Analyzer, then uses the configured AI (ChatGPT / Gemini)
to research the top 5 gigs for each submission and generate an optimized
new gig — title, description, suggested price, and key differentiators.

Results are saved (appended) to:
  Excel and Images/my_gigs.csv

Usage:
  python fiverr_gig_creator.py               → auto-detect latest CSV
  python fiverr_gig_creator.py "filename.csv" → use specific CSV

my_gigs.csv columns:
  Generated_At, Submission, Row, AI_Provider,
  Gig_Title, Gig_Description, Suggested_Price,
  Key_Differentiators, Source_CSV
"""

import os, sys, csv, json
from datetime import datetime

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE_DIR        = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR      = os.path.join(BASE_DIR, "Excel and Images")
MY_GIGS_CSV     = os.path.join(OUTPUT_DIR, "my_gigs.csv")
SUBMISSION_CSV  = os.path.join(OUTPUT_DIR, "submission_top_gigs.csv")

MY_GIGS_FIELDS = [
    "Generated_At", "Submission", "Row", "AI_Provider",
    "Gig_Title", "Gig_Description", "Suggested_Price",
    "Key_Differentiators", "Source_CSV",
]

# ── API manager ────────────────────────────────────────────────────────────────
sys.path.insert(0, BASE_DIR)
try:
    from api_manager import call_ai, extract_json, load_api_config
    _API_OK = True
except ImportError as e:
    _API_OK = False
    _API_ERR = str(e)


# ── Prompt builder ─────────────────────────────────────────────────────────────
_PROMPT_TEMPLATE = """\
You are an expert Fiverr gig strategist with deep marketplace knowledge.

I am creating a brand-new gig in the "{submission}" niche on Fiverr.
Below are the top {count} best-selling gigs in this niche, ranked by performance score
(score = reviews + orders_in_queue × 20). Study them carefully.

{gig_blocks}
══════════════════════════════════════════════════════
YOUR TASK
══════════════════════════════════════════════════════
Based on these top performers, craft the BEST possible new Fiverr gig that:

1. TITLE — Highly searchable, keyword-rich, compelling. STRICT 80-character maximum.
2. DESCRIPTION — Professional, persuasive, 250-500 words. Speak directly to the buyer's
   pain points. Highlight deliverables, process, and guarantee/credibility signals.
3. SUGGESTED_PRICE — A competitive starting (Basic package) price in USD based on the
   market research above.
4. KEY_DIFFERENTIATORS — 3-5 short bullet points that make this gig stand out from
   the competitors analyzed above.

Return ONLY a valid JSON object — no markdown fences, no explanation, no extra text:
{{
  "gig_title": "Your optimized title here (max 80 chars)",
  "gig_description": "Your full professional description here",
  "suggested_price": "$XX",
  "key_differentiators": ["differentiator 1", "differentiator 2", "differentiator 3"]
}}"""

_GIG_BLOCK_TEMPLATE = """\
── GIG #{rank}  (Reviews: {reviews} | Orders in Queue: {orders} | Price: {price}) ──
Title      : {title}
Description: {description}
"""


def build_prompt(submission, gigs):
    """
    gigs = list of dicts with keys: rank, title, description, pricing, reviews, orders
    """
    blocks = []
    for g in gigs:
        desc = (g.get("Gig_Description") or "").strip()
        if len(desc) > 600:
            desc = desc[:597] + "…"
        blocks.append(_GIG_BLOCK_TEMPLATE.format(
            rank        = g.get("Rank", "?"),
            reviews     = g.get("Reviews", 0),
            orders      = g.get("Orders", 0),
            price       = g.get("Pricing", "N/A"),
            title       = (g.get("Gig_Title") or "N/A").strip(),
            description = desc or "N/A",
        ))
    return _PROMPT_TEMPLATE.format(
        submission = submission,
        count      = len(gigs),
        gig_blocks = "\n".join(blocks),
    )


# ── CSV helpers ────────────────────────────────────────────────────────────────
def load_submission_csv(csv_path):
    """
    Read submission_top_gigs.csv (the single shared file) and return:
      { submission_name: [row_dict, ...] }  — only the MOST RECENT batch per submission.

    Each row has a Scraped_At timestamp written by the Submission Analyzer.
    When the same submission was analyzed multiple times (across different runs),
    we keep only the rows whose Scraped_At equals the latest value for that submission,
    so the Gig Creator always works with fresh data.
    Rows within each group are ordered by Rank ascending.
    """
    all_rows = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            sub = row.get("Submission", "").strip()
            if sub:
                all_rows.append(row)

    if not all_rows:
        return {}

    # Find the latest Scraped_At per submission
    latest_ts: dict[str, str] = {}
    for row in all_rows:
        sub = row["Submission"]
        ts  = row.get("Scraped_At", "")
        if ts > latest_ts.get(sub, ""):
            latest_ts[sub] = ts

    # Keep only rows from the most recent batch for each submission
    groups: dict[str, list] = {}
    for row in all_rows:
        sub = row["Submission"]
        if row.get("Scraped_At", "") == latest_ts[sub]:
            groups.setdefault(sub, []).append(row)

    # Sort each group by Rank
    for sub in groups:
        groups[sub].sort(key=lambda r: int(r.get("Rank") or 99))

    return groups


def append_to_my_gigs(entry):
    """Append one entry dict to my_gigs.csv, creating the file + header if needed."""
    file_exists = os.path.exists(MY_GIGS_CSV)
    with open(MY_GIGS_CSV, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=MY_GIGS_FIELDS)
        if not file_exists:
            writer.writeheader()
        writer.writerow(entry)


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    print(f"\n{'='*64}")
    print(f"  Fiverr AI Gig Creator")
    print(f"{'='*64}\n")

    # ── Verify AI is configured ───────────────────────────────────────────────
    if not _API_OK:
        print(f"  ERROR: Could not import api_manager: {_API_ERR}")
        sys.exit(1)

    api_cfg  = load_api_config()
    provider = api_cfg.get("provider", "").strip()
    api_key  = api_cfg.get("api_key",  "").strip()
    if not api_key or not provider:
        print("  ERROR: No AI API key configured.")
        print("  → Open the Hub → Settings → AI Configuration and save your key.")
        sys.exit(1)

    provider_label = {"gemini": "Google Gemini", "openai": "OpenAI ChatGPT"}.get(
        provider, provider)
    print(f"  AI Provider : {provider_label}")

    # ── Locate source CSV ─────────────────────────────────────────────────────
    if not os.path.exists(SUBMISSION_CSV):
        print(f"  ERROR: {SUBMISSION_CSV} not found.")
        print("  → Run 'Submission Analyzer' first to generate gig data.")
        sys.exit(1)

    print(f"  Source CSV  : {SUBMISSION_CSV}")
    print(f"  Output CSV  : {MY_GIGS_CSV}\n")

    # ── Load submission data (most recent batch per submission) ───────────────
    groups = load_submission_csv(SUBMISSION_CSV)
    if not groups:
        print("  ERROR: Source CSV is empty or has no valid rows.")
        sys.exit(1)

    total      = len(groups)
    done_count = 0
    fail_count = 0

    print(f"  Submissions to process: {total}")
    for sub in groups:
        print(f"    • {sub[:60]}  ({len(groups[sub])} gig(s))")
    print()

    # ── Process each submission ───────────────────────────────────────────────
    for idx, (submission, gigs) in enumerate(groups.items(), 1):
        print(f"  [{idx:>2}/{total}]  {submission[:60]}")
        print(f"    Gigs available: {len(gigs)}")

        row_num = gigs[0].get("Row", "") if gigs else ""

        # Build prompt
        prompt = build_prompt(submission, gigs)
        prompt_preview = prompt[:200].replace("\n", " ")
        print(f"    Prompt  : {prompt_preview}…")

        # Call AI
        print(f"    Calling {provider_label}…")
        try:
            raw_response = call_ai(prompt, max_tokens=2048)
        except Exception as e:
            print(f"    ERROR from AI: {e}")
            fail_count += 1
            continue

        # Parse JSON
        try:
            data = extract_json(raw_response)
        except Exception as e:
            print(f"    ERROR parsing AI response: {e}")
            print(f"    Raw response (first 400 chars):\n    {raw_response[:400]}")
            fail_count += 1
            continue

        gig_title       = str(data.get("gig_title",         "")).strip()
        gig_description = str(data.get("gig_description",   "")).strip()
        suggested_price = str(data.get("suggested_price",   "N/A")).strip()
        differentiators = data.get("key_differentiators", [])
        if isinstance(differentiators, list):
            differentiators_str = " | ".join(str(d).strip() for d in differentiators)
        else:
            differentiators_str = str(differentiators).strip()

        # Enforce title length limit (Fiverr max is 80 chars)
        if len(gig_title) > 80:
            gig_title = gig_title[:77] + "…"

        print(f"    ✓  Title     : {gig_title}")
        print(f"    ✓  Price     : {suggested_price}")
        print(f"    ✓  Desc      : {gig_description[:100]}…" if len(gig_description) > 100
              else f"    ✓  Desc      : {gig_description}")
        print(f"    ✓  USPs      : {differentiators_str[:120]}")

        # Append to my_gigs.csv
        entry = {
            "Generated_At":       datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Submission":         submission,
            "Row":                row_num,
            "AI_Provider":        provider_label,
            "Gig_Title":          gig_title,
            "Gig_Description":    gig_description,
            "Suggested_Price":    suggested_price,
            "Key_Differentiators": differentiators_str,
            "Source_CSV":         os.path.basename(SUBMISSION_CSV),
        }
        append_to_my_gigs(entry)
        done_count += 1
        print(f"    Saved to my_gigs.csv ✓\n")

    # ── Summary ───────────────────────────────────────────────────────────────
    print(f"{'='*64}")
    print(f"  Done!")
    print(f"  Submissions processed : {done_count}")
    if fail_count:
        print(f"  Failures             : {fail_count}")
    print(f"  Output               : {MY_GIGS_CSV}")
    print(f"{'='*64}\n")


if __name__ == "__main__":
    main()
