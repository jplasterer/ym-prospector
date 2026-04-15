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
from dotenv import load_dotenv

# Explicitly resolve .env relative to this file so it's found regardless of cwd
_ENV_PATH = Path(__file__).parent.parent / ".env"
load_dotenv(dotenv_path=_ENV_PATH)

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
BRAVE_API_KEY = os.getenv("BRAVE_API_KEY")

anthropic_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

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

YM SIGNALS to look for:
- "yourmembership.com" in member portal URLs (login.yourmembership.com, members.yourmembership.com)
- "Powered by YourMembership" in footers
- YM job board embeds
- YM member portal redirects

DATA SOURCE RELIABILITY (use in this order):
1. ProPublica 990 — most reliable for revenue and staff
2. LinkedIn company page — most reliable for current headcount
3. Organization's own website
4. ZoomInfo/RocketReach — LAST RESORT, often wildly inaccurate for small associations

TIER DEFINITIONS:
- Tier 1: Confirmed YM + staff 3–8 + revenue $2.5M+ + at least one complexity platform → Ready for outreach
- Tier 2: Confirmed YM + meets size/revenue BUT missing one data point to confirm → Needs enrichment
- Tier 3: YM likely but unconfirmed, OR borderline on size/revenue → Watch list
- Exclude: Outside ICP criteria — log reason clearly
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
    """Fetch a webpage and return cleaned text (truncated to 3,000 chars)."""
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
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer", "header"]):
            tag.decompose()
        text = soup.get_text(separator=" ", strip=True)
        return text[:3000]
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Claude API helper
# ---------------------------------------------------------------------------

def claude_call(prompt: str, model: str = "claude-sonnet-4-6", max_tokens: int = 4096) -> str:
    """Make a single Claude API call and return the text response."""
    message = anthropic_client.messages.create(
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
    """Generate YM-fingerprint search queries for a sector."""
    queries = []
    for kw in keywords[:3]:
        queries.append(f'"login.yourmembership.com" {kw}')
        queries.append(f'"members.yourmembership.com" {kw}')
    queries.append(f'"powered by yourmembership" {keywords[0]}')
    queries.append(f'site:yourmembership.com {keywords[0]}')
    queries.append(f'"yourmembership" "member login" {keywords[0]} association')
    return queries[:8]


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
- Look for YM signals in the URLs or snippets (yourmembership.com domains, "Powered by YM", etc.)
- Each candidate should have their OWN main website (not yourmembership.com itself)

Return a JSON array (no other text) with 15–20 objects, each containing:
{{
  "name": "Full organization name",
  "website": "Their main website URL (not yourmembership.com)",
  "ym_signal": "What YM signal was found (e.g. 'login.yourmembership.com in URL')",
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
        trimmed.append(entry)

    candidates_text = json.dumps(trimmed, indent=2)

    prompt = f"""Here are {len(trimmed)} candidate organizations researched for the {sector} sector.

Each candidate includes:
- page_excerpt: text from their main website homepage
- subpage_excerpt: text from /staff, /team, /leadership, /about, or /contact pages — use this for decision-maker names, titles, emails, and phone numbers
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
  "pain_signal_notes": "Key pain signals visible from public info (bandwidth, renewal risk, events, etc.)",
  "association_scope": "Local / State / National / International — classify based on membership geography. Local = single city/metro. State = statewide chapter. National = serves entire US or US+Canada. International = global membership.",
  "tier": "Tier 1 or Tier 2 or Tier 3 or Exclude",
  "exclude_reason": "Brief reason if Exclude, else empty string",
  "project_opportunity_notes": "If excluded due to revenue/size but still a legitimate YM association with operational needs, note the potential for smaller-value project or advisory work (e.g. 'Small state nursing assoc at $1.6M revenue — outside ICP but could benefit from a YM audit or CE workflow project'). Leave empty if not applicable."
}}

Be conservative: only assign Tier 1 when YM is confirmed AND size/revenue signals are strong.
Assign Exclude (with reason) for anything outside the ICP.
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
    "Pain Signal Notes", "Priority Tier", "Exclude Reason",
    "Project Opportunity Notes",
]

FIELD_MAP = [
    "organization_name", "website_url", "city", "state_province", "country",
    "association_scope", "staff_size", "annual_revenue", "ym_confirmed",
    "lms_platform", "community_platform", "other_platforms",
    "decision_maker_name", "decision_maker_title", "decision_maker_email",
    "pain_signal_notes", "tier", "exclude_reason",
    "project_opportunity_notes",
]

TIER_COLORS = {
    "Tier 1": "C6EFCE",   # Green
    "Tier 2": "FFEB9C",   # Yellow
    "Tier 3": "FFCC99",   # Orange
    "Exclude": "FFC7CE",  # Red
}

SHEET_CONFIGS = [
    ("Tier 1 – Ready for Outreach", "Tier 1"),
    ("Tier 2 – Needs Enrichment",   "Tier 2"),
    ("Tier 3 – Watch List",         "Tier 3"),
    ("Excluded",                    "Exclude"),
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
        if tier_filter == "Exclude":
            sheet_prospects = [
                p for p in prospects
                if p.get("tier", "").strip().lower() in ("exclude", "excluded")
            ]
        else:
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
        col_widths = [30, 35, 15, 15, 15, 14, 12, 15, 15, 20, 20, 25, 25, 25, 30, 40, 12, 30, 45]
        for i, width in enumerate(col_widths, 1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = width

    wb.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# Main orchestrator
# ---------------------------------------------------------------------------

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
        tier1 = sum(1 for p in prospects if p.get("tier") == "Tier 1")
        tier2 = sum(1 for p in prospects if p.get("tier") == "Tier 2")
        tier3 = sum(1 for p in prospects if p.get("tier") == "Tier 3")
        status_fn(
            f"Done! {tier1} Tier 1 | {tier2} Tier 2 | {tier3} Tier 3 | "
            f"Saved to {Path(output_path).name}"
        )

    return prospects, output_path
