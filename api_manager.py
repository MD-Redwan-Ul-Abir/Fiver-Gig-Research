#!/usr/bin/env python3
"""
API Manager — Fiverr Research Hub
──────────────────────────────────────────────────────────────────
Handles AI API key storage, retrieval, and calls for both
Google Gemini and OpenAI ChatGPT.

Stores config in api_config.json in the same folder as this file.
"""

import os, json, re, ssl
import urllib.request, urllib.error


def _get_ssl_context():
    """Return an SSL context that works on macOS Python.org installs."""
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        pass
    try:
        return ssl.create_default_context()
    except Exception:
        pass
    return ssl._create_unverified_context()

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "api_config.json")

PROVIDERS = {
    "gemini": "Google Gemini",
    "openai": "OpenAI ChatGPT",
}


# ── Config helpers ─────────────────────────────────────────────────────────────

def load_api_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE) as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_api_config(provider: str, api_key: str):
    cfg = load_api_config()
    cfg["provider"] = provider.lower().strip()
    cfg["api_key"]  = api_key.strip()
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)


def has_api_key() -> bool:
    cfg = load_api_config()
    return bool(cfg.get("api_key", "").strip() and cfg.get("provider", "").strip())


def get_provider_label() -> str:
    cfg = load_api_config()
    return PROVIDERS.get(cfg.get("provider", ""), "Not configured")


def get_masked_key() -> str:
    cfg = load_api_config()
    k = cfg.get("api_key", "")
    if len(k) > 8:
        return k[:4] + "••••••••" + k[-4:]
    return "••••••••" if k else "—"


# ── AI call dispatcher ─────────────────────────────────────────────────────────

def call_ai(prompt: str, max_tokens: int = 8192) -> str:
    cfg      = load_api_config()
    provider = cfg.get("provider", "").lower()
    api_key  = cfg.get("api_key", "").strip()
    if not api_key:
        raise ValueError("No API key configured. Go to Settings → AI Configuration.")
    if provider == "gemini":
        return _call_gemini(api_key, prompt, max_tokens)
    elif provider == "openai":
        return _call_openai(api_key, prompt, max_tokens)
    else:
        raise ValueError(f"Unknown provider '{provider}'. Set it in Settings.")


def _call_gemini(api_key: str, prompt: str, max_tokens: int = 8192) -> str:
    url  = ("https://generativelanguage.googleapis.com/v1beta/models/"
            f"gemini-2.0-flash:generateContent?key={api_key}")
    body = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "maxOutputTokens": max_tokens,
            "temperature": 0.7,
        },
    }).encode()
    req = urllib.request.Request(
        url, data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=90, context=_get_ssl_context()) as resp:
            data = json.loads(resp.read())
        return data["candidates"][0]["content"]["parts"][0]["text"]
    except urllib.error.HTTPError as e:
        body_text = e.read().decode(errors="replace")[:400]
        raise ValueError(f"Gemini API error {e.code}: {body_text}")


def _call_openai(api_key: str, prompt: str, max_tokens: int = 8192) -> str:
    url  = "https://api.openai.com/v1/chat/completions"
    body = json.dumps({
        "model":       "gpt-4o-mini",
        "messages":    [{"role": "user", "content": prompt}],
        "max_tokens":  max_tokens,
        "temperature": 0.7,
        "response_format": {"type": "json_object"},
    }).encode()
    req = urllib.request.Request(
        url, data=body,
        headers={
            "Content-Type":  "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=90, context=_get_ssl_context()) as resp:
            data = json.loads(resp.read())
        return data["choices"][0]["message"]["content"]
    except urllib.error.HTTPError as e:
        body_text = e.read().decode(errors="replace")[:400]
        raise ValueError(f"OpenAI API error {e.code}: {body_text}")


def test_api_key(provider: str, api_key: str):
    """Returns (True, success_msg) or (False, error_msg)."""
    try:
        ping = '{"status":"ok"}'
        if provider == "gemini":
            _call_gemini(api_key, f"Reply with exactly this JSON: {ping}", max_tokens=30)
        else:
            _call_openai(api_key, f"Reply with exactly this JSON: {ping}", max_tokens=30)
        return True, "API key is valid and working!"
    except Exception as e:
        return False, str(e)


# ── JSON extractor ─────────────────────────────────────────────────────────────

def extract_json(text: str):
    """Extract first valid JSON object from an AI text response."""
    text = text.strip()
    # Strip markdown fences
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"\s*```\s*$",       "", text, flags=re.MULTILINE)
    text = text.strip()
    # Direct parse
    try:
        return json.loads(text)
    except Exception:
        pass
    # Find first { … }
    start = text.find("{")
    if start != -1:
        depth = 0
        for i, c in enumerate(text[start:], start):
            if   c == "{": depth += 1
            elif c == "}":
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(text[start:i + 1])
                    except Exception:
                        break
    raise ValueError("Could not extract JSON from AI response:\n" + text[:600])


# ── Prompt constants ───────────────────────────────────────────────────────────

# Legacy single-call prompt (kept for backwards compat)
NICHE_PROMPT = """\
You are a senior Fiverr marketplace analyst. List every niche and sub-niche
available on Fiverr under Programming & Tech and Digital Marketing.
Return ONLY a valid JSON: {"categories":[{"section":"...","niche":"...","sub_niches":[{"name":"...","keyword":"..."}]}]}
"""

# ── Comprehensive hardcoded Fiverr niche list ──────────────────────────────────
# Based on Fiverr's actual category tree (2024-2025).
# Used as the seed for Phase 1 so every real Fiverr category is covered.
FIVERR_NICHES = [
    # ── Programming & Tech ─────────────────────────────────────────────────────
    ("Programming & Tech", "Website Development"),
    ("Programming & Tech", "WordPress"),
    ("Programming & Tech", "E-Commerce Development"),
    ("Programming & Tech", "Mobile App Development"),
    ("Programming & Tech", "Game Development"),
    ("Programming & Tech", "Web Programming & Tech"),
    ("Programming & Tech", "Desktop Applications"),
    ("Programming & Tech", "APIs & Integrations"),
    ("Programming & Tech", "Databases"),
    ("Programming & Tech", "Cloud Computing & DevOps"),
    ("Programming & Tech", "Cybersecurity & Data Protection"),
    ("Programming & Tech", "Data Science & ML"),
    ("Programming & Tech", "AI Applications & Automation"),
    ("Programming & Tech", "No-Code & Low-Code Development"),
    ("Programming & Tech", "Blockchain & Cryptocurrency"),
    ("Programming & Tech", "QA & Testing"),
    ("Programming & Tech", "Web Scraping"),
    ("Programming & Tech", "IT & Networking Support"),
    ("Programming & Tech", "Convert Files"),
    ("Programming & Tech", "User Testing"),
    # ── Digital Marketing ──────────────────────────────────────────────────────
    ("Digital Marketing", "Search Engine Optimization (SEO)"),
    ("Digital Marketing", "Local SEO"),
    ("Digital Marketing", "E-Commerce SEO"),
    ("Digital Marketing", "Social Media Marketing"),
    ("Digital Marketing", "Social Media Management"),
    ("Digital Marketing", "Content Marketing"),
    ("Digital Marketing", "Email Marketing"),
    ("Digital Marketing", "Paid Advertising (PPC & Ads)"),
    ("Digital Marketing", "Video Marketing"),
    ("Digital Marketing", "Influencer Marketing"),
    ("Digital Marketing", "Affiliate Marketing"),
    ("Digital Marketing", "Community Management"),
    ("Digital Marketing", "Marketing Strategy & Consulting"),
    ("Digital Marketing", "Web Analytics & Tracking"),
    ("Digital Marketing", "E-Commerce Marketing"),
    ("Digital Marketing", "Mobile Marketing"),
    ("Digital Marketing", "Podcast Marketing"),
    ("Digital Marketing", "Brand Marketing & Strategy"),
    ("Digital Marketing", "Public Relations"),
    ("Digital Marketing", "Marketing Automation"),
]

# ── Per-niche expansion prompt (used once per niche in Phase 2) ────────────────
SUBNICHE_PROMPT_TEMPLATE = """\
You are a Fiverr marketplace expert with deep knowledge of every service sold on Fiverr.

Your task: list EVERY SINGLE sub-niche that real freelancers offer under the
"{niche}" category in Fiverr's "{section}" section.

Be EXHAUSTIVE. Think about:
- All the specific services buyers search for in this niche
- Both mainstream AND niche/specialised services
- Emerging 2024-2025 trends (AI-powered, automation, no-code, etc.)
- Different technology stacks, platforms, industries, and use-cases

Return ONLY a valid JSON object — no markdown, no explanation:

{{
  "sub_niches": [
    {{"name": "Specific Service Name", "keyword": "exact fiverr search phrase"}},
    {{"name": "Another Service",       "keyword": "another search phrase"}}
  ]
}}

Rules:
- Aim for 25-50 sub-niches (more is better — include everything)
- Each "name" must describe a specific, buyable Fiverr service
- Each "keyword" must be the exact phrase a buyer would type into Fiverr search
- NO duplicates, NO vague entries like "Other" or "Miscellaneous"
- Return ONLY the JSON
"""
