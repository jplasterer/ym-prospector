"""
Starkweather YM Prospector — Core Logic
Searches for YourMembership associations, enriches with ProPublica data,
scores against ICP, and saves to Excel.
"""

import os
import re
import json
import time
import requests
from datetime import datetime
from pathlib import Path

import anthropic
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ---------------------------------------------------------------------------
# Load .env directly — bypasses dotenv parsing quirks with special characters
# ---------------------------------------------------------------------------

def _load_env_file():
    """Read the .env file and inject values into os.environ."""
    env_path = Path(__file__).parent.parent / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        key = key.strip()
        val = val.strip().strip('"').strip("'")
        if key:
            os.environ[key] = val  # always override — our .env takes priority

_load_env_file()

BRAVE_API_KEY = os.environ.get("BRAVE_API_KEY", "")


def _get_anthropic_client() -> anthropic.Anthropic:
    """Create Anthropic client on demand so the key is always fresh."""
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    return anthropic.Anthropic(api_key=api_key)

# ---------------------------------------------------------------------------
# Sector definitions
# ---------------------------------------------------------------------------

SECTORS = {
    "Healthcare": [
        "nursing association", "physician assistant association",
        "pharmacy association", "dental association", "veterinary association",
    ],
    "Legal & Compliance": [
        "bar association", "paralegal association",
        "compliance officers association", "legal professionals association",
    ],
    "Engineering & Construction": [
        "civil engineers association", "structural engineers association",
        "contractors association", "construction association",
    ],
    "Finance & Accounting": [
        "CPA association", "financial planners association",
        "accountants association", "credit union association",
    ],
    "Education & Credentialing": [
        "educators association", "training professionals association",
        "HR professionals association", "credentialing association",
    ],
    "Technology & IT": [
        "IT professionals association", "cybersecurity association",
        "data professionals association", "technology association",
    ],
    "Real Estate & Property": [
        "realtors association", "appraisers association",
        "property managers association", "real estate association",
    ],
    "Nonprofit & Association Management": [
        "nonprofit executives association", "association managers",
        "nonprofit professionals association",
    ],
    "Manufacturing & Industrial Trade": [
        "manufacturers association", "industrial association",
        "trade association manufacturing",
    ],
    "Environmental & Energy": [
        "environmental professionals association", "energy professionals association",
        "sustainability association",
    ],
    "Media & Communications": [
        "PR professionals association", "marketing association",
        "communications professionals association",
    ],
    "Canada": [
        "Canadian professional association", "Canada trade association",
        "Canadian association members",
    ],
}

# ---------------------------------------------------------------------------
# ICP system prompt — shared across all Claude calls
# ---------------------------------------------------------------------------

ICP_SYSTEM_PROMPT = """You are a research analyst for Starkweather Association Services, identifying ideal prospects for our YourMembership (YM) AMS managed services.

IDEAL CUSTOMER PROFILE (ICP):
Must-Have Criteria (ALL required):
- Using YourMembership (YM) AMS platform — confirmed or strongly indicated
- Professional association or trade association
- United States or Canada based
- 3–8 full-time staff (sweet spot: 4–5)
- Annual revenue $2.5M+ (sweet spot: ~$3M)

Complexity Signals (score higher when present):
- LMS/Education platform: Crowd Wisdom, Freestone, TopClass, Absorb, Docebo, LearnUpon, ePath Learning
- Online community: Higher Logic, Personify Community, Mighty Networks
- Email marketing: Informz, Feathr, Constant Contact
- Events platform: Cvent, EventMobi
- Chapter management: running regional chapters or affiliates
- Certification/credentialing program: CE credits, CPE, CME

DISQUALIFIERS (exclude these):
- Volunteer-run organizations with no paid staff
- Annual revenue below $1.5M
- Fewer than 3 or more than 15 full-time employees
- Homeowners associations (HOAs)
- Country clubs, yacht clubs, golf clubs, athletic clubs
- Alumni associations (school or university-based)
- Political organizations or PACs
- Churches, dioceses, religious organizations
- Foundations (separate entities from parent associations)
- Association Management Companies (AMCs) — competitors/peers
- Local volunteer chapters of national associations without own budget/staff

KNOWN LOW-REVENUE PATTERNS (flag but do not surface for ICP outreach):
- State-level nursing associations (e.g. Virginia Nurses Association, Florida Nurses Association,
  Arizona Nurses Association) consistently report revenue of $800K–$1.65M on 990 filings — well
  below ICP minimum. Treat all state nursing associations as Exclude unless 990 confirms $2.5M+.
- Similarly, state-level chapter associations in other professions (social work, physical therapy,
  occupational therapy) tend to run lean. Apply extra skepticism before placing in Tier 2.
- These orgs ARE legitimate YM users with real operational pain — flag them in the
  "project_opportunity_notes" field as candidates for smaller-value project or advisory work,
  even though they do not qualify for full managed services.

YM SIGNALS to look for — in order of confidence:

CONFIRMED (set ym_confirmed = "Yes"):
- "yourmembership.com" in member portal URLs (login.yourmembership.com, members.yourmembership.com)
- "Powered by YourMembership" text in page footer OR on the member portal login page
- YM_SIGNAL tags prepended to page content — these were detected in raw HTML source

STRONG (set ym_confirmed = "Likely"):
- Login page URL contains "/Login.aspx" — this path is a hallmark of YM-hosted portals
- "Access member community" or "Member login" button/link that leads to a yourmembership.com URL
- YM job board embeds (jobs.yourmembership.com)
- login_aspx_content field is populated — we successfully fetched /Login.aspx on their domain
- "yourmembership" string appears in page source even if not in a URL

MODERATE (set ym_confirmed = "Unconfirmed" but note in pain_signal_notes):
- Association lists a separate "member portal" or "community" link pointing off-site that
  could be a YM-hosted portal (investigate further)
- Any reference to YourMembership in search snippets without direct URL confirmation

DATA SOURCE RELIABILITY (use in this order):
1. ProPublica 990 — most reliable for revenue and staff
2. LinkedIn company page — most reliable for current headcount
3. Organization's own website
4. ZoomInfo/RocketReach — LAST RESORT, often wildly inaccurate for small associations

TIER DEFINITIONS:
- Strong Fit: Confirmed YM + staff 3–8 + revenue $2.5M+ + at least one complexity platform → Ready for outreach
- Medium Fit: Confirmed YM + meets size/revenue BUT missing one data point to confirm → Needs enrichment
- Unlikely Fit: YM likely but unconfirmed, OR borderline on size/revenue → Watch list
- Not a Fit: Outside ICP criteria — log reason clearly

CRITICAL SCORING RULE — LinkedIn contact does NOT affect tier:
A missing decision maker name, title, or LinkedIn URL must NEVER cause a prospect to be
downgraded or held at a lower tier. Tiers are determined solely by:
  (1) YM platform confirmation
  (2) Staff size (3–8 FTE)
  (3) Annual revenue ($2.5M+)
  (4) Complexity signals (additional platforms)
If all four criteria are met, assign Strong Fit regardless of whether a contact was found.
"""

# ---------------------------------------------------------------------------
# Brave Search
# ---------------------------------------------------------------------------

def brave_search(query: str, count: int = 10) -> list[dict]:
    """Call Brave Search API and return list of {title, url, description}."""
    headers = {
        "Accept": "application/json",
        "Accept-Encoding": "gzip",
        "X-Subscription-Token": BRAVE_API_KEY,
    }
    params = {"q": query, "count": count, "country": "us", "search_lang": "en"}
    try:
        resp = requests.get(
            "https://api.search.brave.com/res/v1/web/search",
            headers=headers,
            params=params,
            timeout=12,
        )
        resp.raise_for_status()
        data = resp.json()
        results = []
        for item in data.get("web", {}).get("results", []):
            results.append({
                "title": item.get("title", ""),
                "url": item.get("url", ""),
                "description": item.get("description", ""),
            })
        return results
    except Exception as e:
        return []


# ---------------------------------------------------------------------------
# Page fetching
# ---------------------------------------------------------------------------

def fetch_page(url: str, timeout: int = 8) -> str:
    """Fetch a webpage and return cleaned text (truncated to 3,000 chars).

    YM detection improvements:
    - Scans raw HTML source for YourMembership strings BEFORE stripping tags
      (catches references in scripts, meta tags, hidden elements, etc.)
    - Preserves footer text — 'Powered by YourMembership' is often in the footer
    - Prepends any detected YM signals as YM_SIGNAL: tags so Claude sees them first
    """
    try:
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            )
        }
        resp = requests.get(url, headers=headers, timeout=timeout, allow_redirects=True)
        resp.raise_for_status()

        # --- Scan raw HTML for YM signals before any tag-stripping ---
        raw_lower = resp.text.lower()
        ym_signals = []
        if "powered by yourmembership" in raw_lower:
            ym_signals.append("YM_SIGNAL: 'Powered by YourMembership' found in page source (CONFIRMED YM user)")
        if "yourmembership.com" in raw_lower:
            ym_signals.append("YM_SIGNAL: 'yourmembership.com' URL found in page source")
        if "login.aspx" in raw_lower and "yourmembership" in raw_lower:
            ym_signals.append("YM_SIGNAL: '/Login.aspx' path found alongside YourMembership reference (strong YM indicator)")
        elif "login.aspx" in raw_lower:
            ym_signals.append("YM_SIGNAL: '/Login.aspx' path found — possible YM login page")
        if "yourmembership" in raw_lower and "yourmembership.com" not in raw_lower:
            ym_signals.append("YM_SIGNAL: 'YourMembership' string found in page source (not just a URL)")

        # --- Parse HTML, keeping footer (often contains 'Powered by YM') ---
        soup = BeautifulSoup(resp.text, "html.parser")

        # Extract footer text separately so it isn't lost
        footer_text = ""
        for footer_tag in soup.find_all(["footer", "div"], class_=lambda c: c and "footer" in c.lower() if c else False):
            footer_text = footer_tag.get_text(separator=" ", strip=True)[:500]
            break

        for tag in soup(["script", "style", "nav", "header"]):  # footer intentionally kept
            tag.decompose()

        body_text = soup.get_text(separator=" ", strip=True)

        # Assemble: YM signals first, then body, footer appended if not already present
        prefix = "\n".join(ym_signals) + "\n" if ym_signals else ""
        full_text = prefix + body_text
        if footer_text and footer_text not in full_text:
            full_text = full_text[:2500] + f"\n[FOOTER]: {footer_text}"

        return full_text[:3000]
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Claude API helper
# ---------------------------------------------------------------------------

def claude_call(prompt: str, model: str = "claude-sonnet-4-6", max_tokens: int = 4096) -> str:
    """Make a single Claude API call and return the text response."""
    client = _get_anthropic_client()
    message = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        system=ICP_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


def extract_json(text: str) -> list:
    """Extract a JSON array from a text response."""
    # Try direct parse first
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    # Find JSON array in text
    match = re.search(r"\[.*\]", text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass
    return []


# ---------------------------------------------------------------------------
# Search query generation
# ---------------------------------------------------------------------------

def generate_search_queries(sector: str, keywords: list[str]) -> list[str]:
    """Generate YM-fingerprint search queries for a sector.

    Query patterns derived from known YM site indicators:
    - login.yourmembership.com / members.yourmembership.com subdomains
    - 'Powered by YourMembership' footer text
    - site:yourmembership.com (YM-hosted portals indexed by Google)
    - 'member login' / 'access member community' button text leading to YM
    - /Login.aspx URL path — characteristic of YM member portal login pages
    """
    queries = []
    for kw in keywords[:3]:
        queries.append(f'"login.yourmembership.com" {kw}')
        queries.append(f'"members.yourmembership.com" {kw}')
    queries.append(f'"powered by yourmembership" {keywords[0]}')
    queries.append(f'site:yourmembership.com {keywords[0]}')
    queries.append(f'"yourmembership" "member login" {keywords[0]} association')
    # /Login.aspx is a characteristic URL path on YM member portals
    queries.append(f'inurl:Login.aspx "yourmembership" {keywords[0]} association')
    # "access member community" is a common YM portal CTA button label
    queries.append(f'"access member community" {keywords[0]} association')
    # Catch associations where YM portal lives on a separate subdomain
    queries.append(f'"yourmembership" "sign in" OR "member portal" {keywords[0]} association')
    return queries[:10]


# ---------------------------------------------------------------------------
# Step 1: Identify candidate organizations from search results
# ---------------------------------------------------------------------------

def identify_candidates(search_results: list[dict], sector: str, status_fn=None) -> list[dict]:
    """Use Claude to identify candidate YM associations from raw search results."""
    if not search_results:
        return []

    results_text = "\n".join([
        f"{i+1}. {r['title']}\n   URL: {r['url']}\n   Snippet: {r['description']}"
        for i, r in enumerate(search_results[:50])
    ])

    prompt = f"""I've run web searches for YourMembership (YM) associations in the {sector} sector.

SEARCH RESULTS:
{results_text}

Task: Identify the most promising candidate organizations that appear to use YourMembership AMS and match our ICP.

Rules:
- Only include professional associations or trade associations
- Exclude HOAs, clubs, alumni orgs, political orgs, AMCs, foundations, churches
- Look for ANY of these YM signals in the URLs or snippets:
  * "yourmembership.com" appearing in a URL (login.yourmembership.com, members.yourmembership.com, etc.)
  * "Powered by YourMembership" text in a snippet
  * "/Login.aspx" in a URL — this is a characteristic YM member portal login path
  * "access member community" or "member login" button text in a snippet that links to YM
  * YM job board embeds (jobs.yourmembership.com)
  * Any mention of "YourMembership" in the snippet text
- Each candidate should have their OWN main website (not yourmembership.com itself)

Return a JSON array (no other text) with 15–20 objects, each containing:
{{
  "name": "Full organization name",
  "website": "Their main website URL (not yourmembership.com)",
  "ym_signal": "Specific YM signal found (e.g. 'login.yourmembership.com in URL', '/Login.aspx path', 'Powered by YourMembership in snippet')",
  "org_type": "Type of association",
  "geography": "US or Canada if determinable, else Unknown",
  "initial_notes": "Any relevant observations from the snippet"
}}

JSON array:"""

    if status_fn:
        status_fn("Asking Claude to identify candidate organizations from search results...")

    response = claude_call(prompt)
    candidates = extract_json(response)

    # Filter out entries without a real website
    candidates = [
        c for c in candidates
        if c.get("website") and "yourmembership.com" not in c.get("website", "")
    ]
    return candidates


# ---------------------------------------------------------------------------
# Step 2: Research each candidate
# ---------------------------------------------------------------------------

def fetch_org_subpages(base_url: str) -> str:
    """
    Try common staff/leadership/contact sub-pages and return combined text.
    Stops after finding 2 pages with substantive content to keep runtime short.
    """
    if not base_url:
        return ""
    base = base_url.rstrip("/")
    sub_paths = [
        "/staff", "/team", "/leadership", "/about", "/about-us",
        "/contact", "/contact-us", "/board", "/our-team", "/meet-the-team",
    ]
    found = []
    for path in sub_paths:
        content = fetch_page(f"{base}{path}", timeout=6)
        if content and len(content) > 200:
            found.append(f"[{path}]:\n{content[:1000]}")
            if len(found) >= 2:
                break
    return "\n\n".join(found)


def research_candidate(candidate: dict, status_fn=None) -> dict:
    """
    Enrich a candidate org with four data sources:
    1. Main website + sub-pages (staff, leadership, contact info)
    2. ProPublica 990 page (real revenue + staff numbers)
    3. LinkedIn company snippet (staff size range)
    4. LinkedIn people search (decision-maker names + titles)
    """
    name = candidate.get("name", "Unknown")
    website = candidate.get("website", "")

    # --- 1. Main website ---
    candidate["page_content"] = fetch_page(website) if website else ""

    # --- 1b. Sub-pages: /staff /team /leadership /contact /about ---
    candidate["subpage_content"] = fetch_org_subpages(website)

    # --- 2. ProPublica: search, then fetch the actual 990 page ---
    propublica_results = brave_search(
        f'site:projects.propublica.org/nonprofits "{name}"', count=3
    )
    if propublica_results:
        candidate["propublica_url"] = propublica_results[0]["url"]
        candidate["propublica_snippet"] = propublica_results[0]["description"]
        # Fetch the real ProPublica page to extract actual revenue/staff figures
        pp_content = fetch_page(propublica_results[0]["url"])
        candidate["propublica_content"] = pp_content
    else:
        candidate["propublica_url"] = ""
        candidate["propublica_snippet"] = ""
        candidate["propublica_content"] = ""

    # --- 3. LinkedIn company page (staff size range) ---
    linkedin_co = brave_search(
        f'site:linkedin.com/company "{name}"', count=2
    )
    candidate["linkedin_company_snippet"] = (
        linkedin_co[0]["description"] if linkedin_co else ""
    )

    # --- 4. LinkedIn people search (decision-maker names + titles) ---
    linkedin_people = brave_search(
        f'"{name}" (\"executive director\" OR \"chief executive\" OR \"CEO\" OR '
        f'\"director of membership\" OR \"executive vice president\") site:linkedin.com/in',
        count=4,
    )
    candidate["linkedin_people"] = [
        {"title": r["title"], "url": r["url"], "snippet": r["description"]}
        for r in linkedin_people
    ]

    # --- 5. Probe /Login.aspx — characteristic YM member portal path ---
    # Many YM associations host their member portal on their own domain at /Login.aspx.
    # Fetching this page confirms YM usage even if the main site has no "Powered by" text.
    if website:
        login_url = website.rstrip("/") + "/Login.aspx"
        login_content = fetch_page(login_url, timeout=6)
        # Only keep if it returned real content (not a 404 redirect to homepage)
        if login_content and len(login_content) > 100:
            candidate["login_aspx_content"] = login_content[:600]
        else:
            candidate["login_aspx_content"] = ""
    else:
        candidate["login_aspx_content"] = ""

    # Also try a broader web search for email/phone on the org's own domain
    if website:
        domain = website.replace("https://", "").replace("http://", "").split("/")[0]
        contact_results = brave_search(
            f'site:{domain} (email OR "contact us" OR phone OR "executive director" OR CEO)',
            count=3,
        )
        candidate["contact_search_snippets"] = [
            r["description"] for r in contact_results
        ]
    else:
        candidate["contact_search_snippets"] = []

    time.sleep(0.5)
    return candidate


# ---------------------------------------------------------------------------
# Step 3: Score and tier all researched candidates
# ---------------------------------------------------------------------------

def score_prospects(candidates: list[dict], sector: str, status_fn=None) -> list[dict]:
    """Use Claude to score, tier, and produce the final structured prospect list."""
    if not candidates:
        return []

    # Trim content fields for token efficiency while preserving all sources
    trimmed = []
    for c in candidates:
        entry = {
            k: v for k, v in c.items()
            if k not in ("page_content", "subpage_content", "propublica_content")
        }
        entry["page_excerpt"] = c.get("page_content", "")[:600]
        entry["subpage_excerpt"] = c.get("subpage_content", "")[:800]
        entry["propublica_excerpt"] = c.get("propublica_content", "")[:800]
        # login_aspx_content is already trimmed at source; include as-is
        entry["login_aspx_content"] = c.get("login_aspx_content", "")
        trimmed.append(entry)

    candidates_text = json.dumps(trimmed, indent=2)

    prompt = f"""Here are {len(trimmed)} candidate organizations researched for the {sector} sector.

Each candidate includes:
- page_excerpt: text from their main website homepage. NOTE: This may begin with YM_SIGNAL: lines
  detected from raw HTML source — these are high-confidence YM indicators even if not visible
  in the rendered page text.
- subpage_excerpt: text from /staff, /team, /leadership, /about, or /contact pages — use this for decision-maker names, titles, emails, and phone numbers
- login_aspx_content: content fetched from the org's /Login.aspx path. If populated, this STRONGLY
  confirms YM usage — /Login.aspx is a hallmark URL path for YM member portal login pages.
- propublica_excerpt: actual content from ProPublica 990 page — use this for REVENUE and STAFF COUNT (most reliable source)
- propublica_snippet: search snippet from ProPublica result
- linkedin_company_snippet: LinkedIn company page snippet — often contains employee range (e.g. "11-50 employees")
- linkedin_people: list of LinkedIn people search results — use title/snippet to identify the decision-maker name and role
- contact_search_snippets: snippets from the org's own domain about contacts/leadership

PRIORITY for data extraction:
- Revenue → propublica_excerpt first, then propublica_snippet
- Staff count → propublica_excerpt first, then linkedin_company_snippet
- Decision maker name/title → linkedin_people snippets first, then subpage_excerpt
- Email/phone → subpage_excerpt first, then contact_search_snippets

CANDIDATES:
{candidates_text}

Score and tier each organization against our ICP.

Return a JSON array (no other text) where each object contains ALL of these fields:
{{
  "organization_name": "Full organization name",
  "website_url": "Their website URL",
  "city": "City or empty string",
  "state_province": "State/Province or empty string",
  "country": "United States or Canada or empty string",
  "staff_size": "Number or range (e.g. '4-6') or empty string",
  "annual_revenue": "Dollar amount (e.g. '$3.2M') from 990 if available, else empty string",
  "ym_confirmed": "Yes / Likely / Unconfirmed",
  "lms_platform": "Platform name if found, else empty string",
  "community_platform": "Platform name if found, else empty string",
  "other_platforms": "Comma-separated list of other platforms, else empty string",
  "decision_maker_name": "Name if findable, else empty string",
  "decision_maker_title": "Title if findable (ED, CEO, Director of Membership, etc.), else empty string",
  "decision_maker_email": "Email if findable, else empty string",
  "decision_maker_linkedin_url": "Full linkedin.com/in/... profile URL from the linkedin_people results — use the url field from the best matching result, else empty string",
  "pain_signal_notes": "Key pain signals visible from public info (bandwidth, renewal risk, events, etc.)",
  "association_scope": "Local / State / National / International — classify based on membership geography. Local = single city/metro. State = statewide chapter. National = serves entire US or US+Canada. International = global membership.",
  "tier": "Strong Fit or Medium Fit or Unlikely Fit or Not a Fit",
  "exclude_reason": "Brief reason if Not a Fit, else empty string",
  "project_opportunity_notes": "If excluded due to revenue/size but still a legitimate YM association with operational needs, note the potential for smaller-value project or advisory work (e.g. 'Small state nursing assoc at $1.6M revenue — outside ICP but could benefit from a YM audit or CE workflow project'). Leave empty if not applicable."
}}

CRITICAL: A missing decision maker name, title, email, or LinkedIn URL must NEVER cause a
prospect to be placed in a lower tier. Tier is determined only by YM confirmation, staff size,
revenue, and complexity platforms. Assign Strong Fit when those four criteria are met even if
no contact was found.

Be conservative: only assign Strong Fit when YM is confirmed AND size/revenue signals are strong.
Assign Not a Fit (with reason) for anything outside the ICP.
For excluded orgs that are legitimate associations with real YM usage, always populate project_opportunity_notes.

JSON array:"""

    if status_fn:
        status_fn("Scoring and tiering all prospects with Claude...")

    response = claude_call(prompt, max_tokens=6000)
    return extract_json(response)


# ---------------------------------------------------------------------------
# Step 4: Save to Excel
# ---------------------------------------------------------------------------

HEADERS = [
    "Organization Name", "Website URL", "City", "State/Province", "Country",
    "Association Scope", "Staff Size", "Annual Revenue", "YM Confirmed?",
    "LMS Platform", "Community Platform", "Other Platforms",
    "Key Decision Maker Name", "Key Decision Maker Title", "Key Decision Maker Email",
    "Decision Maker LinkedIn URL",
    "Pain Signal Notes", "Priority Tier", "Exclude Reason",
    "Project Opportunity Notes",
]

FIELD_MAP = [
    "organization_name", "website_url", "city", "state_province", "country",
    "association_scope", "staff_size", "annual_revenue", "ym_confirmed",
    "lms_platform", "community_platform", "other_platforms",
    "decision_maker_name", "decision_maker_title", "decision_maker_email",
    "decision_maker_linkedin_url",
    "pain_signal_notes", "tier", "exclude_reason",
    "project_opportunity_notes",
]

TIER_COLORS = {
    "Strong Fit":   "C6EFCE",   # Green
    "Medium Fit":   "FFEB9C",   # Yellow
    "Unlikely Fit": "FFCC99",   # Orange
    "Not a Fit":    "FFC7CE",   # Red
}

SHEET_CONFIGS = [
    ("Strong Fit – Ready for Outreach", "Strong Fit"),
    ("Medium Fit – Needs Enrichment",   "Medium Fit"),
    ("Unlikely Fit – Watch List",       "Unlikely Fit"),
    ("Not a Fit",                       "Not a Fit"),
]


def _sector_slug(sector: str) -> str:
    """Convert a sector name to a safe filename slug. E.g. 'Legal & Compliance' → 'legal-compliance'"""
    slug = sector.lower()
    slug = re.sub(r"[&/\\]", "-", slug)
    slug = re.sub(r"[^a-z0-9\-]", "", slug)
    slug = re.sub(r"-+", "-", slug).strip("-")
    return slug


def save_to_excel(prospects: list[dict], output_dir: str, sector: str = "") -> str:
    """Save the scored prospect list to a 4-sheet Excel file."""
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    date_str = datetime.now().strftime("%Y-%m-%d")
    sector_part = f"_{_sector_slug(sector)}" if sector else ""
    output_path = str(Path(output_dir) / f"ym_prospects{sector_part}_{date_str}.xlsx")

    wb = Workbook()
    first_sheet = True

    for sheet_name, tier_filter in SHEET_CONFIGS:
        ws = wb.active if first_sheet else wb.create_sheet(sheet_name)
        if first_sheet:
            ws.title = sheet_name
            first_sheet = False

        # Header row
        ws.append(HEADERS)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E79")
            cell.alignment = Alignment(horizontal="center", wrap_text=True)

        # Filter prospects for this sheet
        sheet_prospects = [
            p for p in prospects
            if p.get("tier", "").strip() == tier_filter
        ]

        fill_color = TIER_COLORS.get(tier_filter, "FFFFFF")
        for prospect in sheet_prospects:
            row = [str(prospect.get(f, "") or "") for f in FIELD_MAP]
            ws.append(row)
            for cell in ws[ws.max_row]:
                cell.fill = PatternFill("solid", fgColor=fill_color)
                cell.alignment = Alignment(wrap_text=True)

        # Column widths
        col_widths = [30, 35, 15, 15, 15, 14, 12, 15, 15, 20, 20, 25, 25, 25, 30, 45, 40, 12, 30, 45]
        for i, width in enumerate(col_widths, 1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = width

    wb.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# Main orchestrator
# ---------------------------------------------------------------------------

def enrich_tier2(prospects: list[dict], status_fn=None) -> list[dict]:
    """
    Re-research Tier 2 prospects to fill missing data points and attempt
    to push them to Tier 1. Focuses on:
    - ProPublica page fetch for real revenue + staff numbers
    - LinkedIn people search via Brave for decision-maker names/titles
    - Org sub-pages for contact info
    Returns updated prospect list with revised tiers.
    """
    tier2 = [p for p in prospects if p.get("tier") == "Medium Fit"]
    other = [p for p in prospects if p.get("tier") != "Medium Fit"]

    if not tier2:
        return prospects

    enriched = []
    for i, prospect in enumerate(tier2):
        name = prospect.get("organization_name", "Unknown")
        website = prospect.get("website_url", "")

        if status_fn:
            status_fn(f"Enriching {i+1}/{len(tier2)}: {name}...")

        # --- ProPublica: fetch real 990 data ---
        if not prospect.get("annual_revenue") or not prospect.get("staff_size"):
            pp_results = brave_search(
                f'site:projects.propublica.org/nonprofits "{name}"', count=3
            )
            if pp_results:
                pp_url = pp_results[0]["url"]
                pp_content = fetch_page(pp_url)
                prospect["propublica_content_enriched"] = pp_content[:1500]
            else:
                # Try broader search with revenue keywords
                pp_results2 = brave_search(
                    f'"{name}" annual revenue 990 nonprofit', count=3
                )
                prospect["propublica_content_enriched"] = " | ".join(
                    r["description"] for r in pp_results2
                )

        # --- LinkedIn: people search for decision-maker ---
        if not prospect.get("decision_maker_name"):
            li_people = brave_search(
                f'"{name}" ("executive director" OR "CEO" OR "chief executive" '
                f'OR "director of membership") site:linkedin.com/in',
                count=4,
            )
            prospect["linkedin_people_enriched"] = [
                {"title": r["title"], "url": r["url"], "snippet": r["description"]}
                for r in li_people
            ]

            # Also try broader LinkedIn company people search
            li_co = brave_search(
                f'"{name}" executive director OR CEO linkedin', count=3
            )
            prospect["linkedin_broad_enriched"] = [
                r["description"] for r in li_co
            ]

        # --- Sub-pages: contact/staff info ---
        if website and not prospect.get("decision_maker_email"):
            prospect["subpage_enriched"] = fetch_org_subpages(website)

        enriched.append(prospect)
        time.sleep(0.4)

    # Re-score enriched Tier 2 prospects with Claude
    if status_fn:
        status_fn("Re-scoring enriched prospects with Claude...")

    enriched_text = json.dumps(enriched, indent=2)

    prompt = f"""These are Medium Fit prospects that have been re-researched to find missing data.
New data fields added:
- propublica_content_enriched: actual ProPublica 990 page content with revenue/staff figures
- linkedin_people_enriched: LinkedIn people search results for decision-maker names/titles/URLs
- linkedin_broad_enriched: broader LinkedIn search snippets
- subpage_enriched: staff/contact sub-page content

PROSPECTS:
{enriched_text}

For each prospect:
1. Extract revenue and staff from propublica_content_enriched (look for dollar amounts and employee counts)
2. Extract decision-maker name, title, and linkedin.com/in/... URL from linkedin_people_enriched
3. Extract email/phone from subpage_enriched
4. Re-score against ICP and assign updated tier (can promote to Strong Fit if data now supports it)

Tier values: Strong Fit / Medium Fit / Unlikely Fit / Not a Fit

CRITICAL: Missing decision maker contact info must NEVER prevent promotion to Strong Fit.
Tier is determined only by YM confirmation, staff size, revenue, and complexity platforms.

Return a JSON array with ALL standard fields (including decision_maker_linkedin_url) plus updated tier assignments.
Be specific with revenue (e.g. "$2.8M") and staff (e.g. "5") when found in the data.
JSON:"""

    response = claude_call(prompt)
    rescored = extract_json(response)

    if rescored:
        return other + rescored
    return other + enriched


def run_session(
    sector: str,
    num_targets: int = 12,
    status_fn=None,
    output_dir: str = None,
) -> tuple[list[dict], str]:
    """
    Run a full prospecting session for the given sector.
    Returns (prospects_list, output_excel_path).
    """
    if output_dir is None:
        output_dir = str(Path(__file__).parent.parent / "output" / "nightly")

    keywords = SECTORS.get(sector, ["professional association"])

    # --- Step 1: Run searches ---
    if status_fn:
        status_fn(f"Generating search queries for: {sector}")
    queries = generate_search_queries(sector, keywords)

    all_results = []
    for i, query in enumerate(queries):
        if status_fn:
            status_fn(f"Search {i + 1}/{len(queries)}: {query[:70]}...")
        results = brave_search(query, count=10)
        all_results.extend(results)
        time.sleep(0.5)

    # Deduplicate by URL
    seen_urls: set[str] = set()
    unique_results = []
    for r in all_results:
        if r["url"] not in seen_urls:
            seen_urls.add(r["url"])
            unique_results.append(r)

    if status_fn:
        status_fn(f"Collected {len(unique_results)} unique results. Identifying candidates...")

    # --- Step 2: Identify candidates ---
    candidates = identify_candidates(unique_results, sector, status_fn)

    if not candidates:
        if status_fn:
            status_fn("No candidates identified. Try a different sector or check API keys.")
        return [], ""

    if status_fn:
        status_fn(f"Found {len(candidates)} candidates. Researching each one...")

    # --- Step 3: Research candidates (cap at num_targets + 5 buffer) ---
    research_cap = min(len(candidates), num_targets + 5)
    researched = []
    for i, candidate in enumerate(candidates[:research_cap]):
        name = candidate.get("name", "Unknown")
        if status_fn:
            status_fn(f"Researching {i + 1}/{research_cap}: {name}...")
        researched.append(research_candidate(candidate, status_fn))

    # --- Step 4: Score and tier ---
    prospects = score_prospects(researched, sector, status_fn)

    if not prospects:
        if status_fn:
            status_fn("Scoring step returned no results. Check Anthropic API key.")
        return [], ""

    # --- Step 5: Save to Excel ---
    if status_fn:
        status_fn("Saving results to Excel...")
    output_path = save_to_excel(prospects, output_dir, sector=sector)

    if status_fn:
        t_strong   = sum(1 for p in prospects if p.get("tier") == "Strong Fit")
        t_medium   = sum(1 for p in prospects if p.get("tier") == "Medium Fit")
        t_unlikely = sum(1 for p in prospects if p.get("tier") == "Unlikely Fit")
        status_fn(
            f"Done! {t_strong} Strong Fit | {t_medium} Medium Fit | "
            f"{t_unlikely} Unlikely Fit | Saved to {Path(output_path).name}"
        )

    return prospects, output_path
