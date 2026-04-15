# Starkweather Prospector — Build Architecture

## Vision

A nightly automated tool that discovers YourMembership associations matching
Starkweather's ICP, enriches them with decision-maker data, automates LinkedIn
outreach to set appointments, and pushes qualified contacts into Go High Level.

Comparable reference: BuiltWith.com — but purpose-built for Starkweather's
specific target segment and integrated end-to-end with outreach and CRM.

---

## Enrichment Decision Log

This section records qualification decisions made during research so future runs
apply consistent logic. Update this section as new patterns emerge.

### Confirmed Disqualification Patterns (from April 2026 seed list)

| Pattern | Example | Rule |
|---|---|---|
| Staff > 15 | NFRC (~35 staff), NY-BEST (51-200) | Move to Tier 3 |
| Revenue < $1.5M on 990 | ARM $1.04M, AICUO $1.49M | Move to Tier 3 |
| Local volunteer chapter of national org | SHRM SW FL, Acadiana SHRM | Move to Tier 3 unless own 990 filed |
| Athletic / alumni club | Iowa State Cyclone Letter Winners | Move to Tier 3 |
| Foundation entity (not membership org) | Foundation of College of Diplomates ABPD | Move to Tier 3 |
| Association Management Company | Agentis Management LLC | Move to Tier 3 (competitor/peer) |
| Inactive / ceased organization | Academy of Integrative Pain Management | Move to Tier 3 |
| Non-association domain | hannahstreet.com, richmondatwork.com, laborear.com | Move to Tier 3 |

### Confirmed Tier 2 Patterns Needing 990 Lookup

| Org | What We Know | What to Look Up |
|---|---|---|
| NASW Texas (naswtx.org) | 4 staff confirmed | Revenue — ProPublica EIN 74-1860732 |
| Bar Assoc Montgomery County (barmont.org) | 4 staff confirmed, 501(c)(6) | Revenue — ProPublica EIN 52-0807787 |
| HRACC (hracc.org) | 7 staff confirmed | Revenue — ignore ZoomInfo $11M, use 990 |
| BISA (bisanet.org) | Revenue $3.1M confirmed | Exact staff count — ZoomInfo says 11-50 |
| Louisiana Pharmacists (louisianapharmacists.com) | Professional assoc, 1892 | Both — ProPublica EIN 72-0246230 |
| American Sports Builders (sportsbuilders.org) | 600+ members, trade assoc | Revenue — ProPublica EIN 13-2561230 |

### Data Source Reliability Ranking

Use sources in this order — stop when you have confirmed data:

1. **ProPublica Nonprofit Explorer 990** — most reliable for revenue and staff
2. **LinkedIn company page** — most reliable for current headcount
3. **Candid / GuideStar** — alternative to ProPublica for 990 data
4. **Cause IQ** — good alternative for nonprofit financial data
5. **Organization's own website** — annual reports, staff pages
6. **ZoomInfo / RocketReach** — LAST RESORT, often wildly inaccurate for small associations

---

## Phase 1 (MVP): Nightly Discovery + Scored Excel Output

**Status**: Build now in Claude Code

**What it does**:
- Runs nightly web searches using YM-specific search patterns (see CLAUDE.md)
- Visits each candidate organization's website to confirm YM usage and stack
- Looks up revenue/staff via ProPublica 990 / Candid
- Scores each prospect against ICP (Tier 1/2/3)
- Outputs a clean Excel file: `ym_prospects_YYYY-MM-DD.xlsx`

**Tools needed**:
- Claude Code with web search enabled
- xlsx skill (openpyxl)
- No paid APIs required

**Trigger**:
- Manual: run via `claude code` in the `~/ym-prospector/` folder using the
  prospecting prompt from CLAUDE.md
- Scheduled: set up as a nightly Claude scheduled task once MVP is validated

**Output location**: `~/ym-prospector/output/`

**Success criteria**:
- Generates 10–15 new qualified prospects per run
- Tier 1 prospects have confirmed YM + size + revenue signals
- Excel output matches required field structure

---

## Phase 2: Decision-Maker Enrichment

**Adds to Phase 1 output**: Full contact details for key decision-makers

**What it does**:
- For each Tier 1 prospect, searches LinkedIn for the Executive Director,
  Director of Membership, IT Director, or COO
- Finds verified email address where possible (LinkedIn, website, email
  pattern inference from known contacts)
- Adds direct phone if available from website

**Tools to add**:
- Apollo.io Plugin (Claude Code plugin) — best option for verified emails
  and phone numbers at scale
- OR: Vibe Prospecting MCP — similar capability, different data source
- LinkedIn Sales Navigator (if available) — direct LinkedIn data

**Trigger**: Runs as second step in the same nightly session, enriching
Phase 1 Tier 1 results before saving the final Excel

---

## Phase 3: LinkedIn Outreach via Claude Chrome

**Adds**: Automated LinkedIn connection requests and follow-up messages
to decision-makers at Tier 1 prospects

**What it does**:
- Opens LinkedIn in Claude Chrome
- Sends personalized connection requests to decision-makers using
  Starkweather's messaging templates (see outreach templates below)
- After connection accepted: sends follow-up message with call booking link
- Logs outreach status back to the prospect Excel file

**Tools to add**:
- Claude in Chrome (already available in Cowork)
- Go High Level calendar link for booking discovery calls
- Outreach message templates (stored in `~/ym-prospector/templates/`)

**Outreach message templates** (to be developed):
- Connection request: Pain-point-led, specific to their org type
- Follow-up 1 (after connect): Value proposition + ask for 15 min
- Follow-up 2 (7 days later): Case study or proof point + re-ask
- Templates should reference: YM expertise, bandwidth relief, revenue protection

**Important**: Claude Chrome outreach requires explicit approval per session —
it will show you the message before sending each batch.

---

## Phase 4: Go High Level CRM Integration

**Adds**: Automatic push of qualified prospects into GHL, deduplication
against existing contacts, and calendar scheduling

**What it does**:
- Before surfacing any prospect, checks GHL CRM for existing contact/account
- Creates new GHL contact record for net-new Tier 1 prospects
- Assigns to appropriate pipeline stage ("New YM Prospect")
- Attaches research notes and pain signals to the GHL contact
- Adds GHL calendar booking link to outreach messages

**Tools to add**:
- Go High Level MCP (check MCP Registry in Claude Code — community GHL MCP
  exists; may need API key setup)
- GHL API key from your GHL account settings

**GHL pipeline stages to set up**:
1. New YM Prospect (auto-created by tool)
2. LinkedIn Connected (manual update after Phase 3)
3. Discovery Scheduled
4. Discovery Completed
5. Proposal Sent
6. Closed Won / Closed Lost

---

## Folder Structure

```
~/ym-prospector/
├── CLAUDE.md               ← Project briefing (ICP, methodology, output spec)
├── ARCHITECTURE.md         ← This file
├── templates/
│   ├── linkedin_connect.md     ← Connection request templates by org type
│   ├── linkedin_followup_1.md  ← First follow-up after connect
│   └── linkedin_followup_2.md  ← Second follow-up / case study
├── exclusions/
│   └── existing_clients.txt    ← Domain list of current Starkweather clients
├── output/
│   ├── seed_list/
│   │   └── starkweather_seed_prospects.xlsx   ← Processed BuiltWith data
│   └── nightly/
│       └── ym_prospects_YYYY-MM-DD.xlsx       ← Nightly discovery runs
└── scripts/
    └── process_builtwith.py    ← One-time seed list processor (already built)
```

---

## Nightly Run Prompt (Phase 1 MVP)

Save this as `~/ym-prospector/run_nightly.md` and paste into Claude Code
each night (or use as a scheduled task prompt once validated):

```
You are running the Starkweather YM Prospector nightly discovery session.
Read CLAUDE.md for full ICP criteria, search methodology, and output format.

TODAY'S SECTOR FOCUS: [paste from the weekly rotation in CLAUDE.md]

Steps:
1. Run 5–8 web searches using the YM search patterns in CLAUDE.md for today's sector
2. For each candidate org found (target 15–20 candidates):
   a. Visit their website to confirm YM usage and check for LMS/community platforms
   b. Look up revenue and staff on ProPublica (propublica.org/nonprofits) or Candid
   c. Find the primary decision-maker on LinkedIn
   d. Score against ICP and assign Tier 1/2/3
3. Build the output table with all required fields
4. Save as output/nightly/ym_prospects_[TODAY'S DATE].xlsx with 4-sheet structure
5. Print a brief summary: how many found, tier breakdown, any notable prospects

Target: 10–15 qualified prospects this session (Tier 1 + 2 combined).
Do not include any organization that already appears in previous output files.
```

---

## Success Metrics

| Metric | Target |
|---|---|
| Nightly run yield (Tier 1+2) | 10–15 prospects |
| Sessions to reach 150-account target | ~15 runs |
| Tier 1 conversion to discovery call | 10–15% |
| Discovery to opportunity | 50% |
| Opportunity to close | 25% |
| Deals needed to hit $1M goal | 21 at $40K+ |

---

## Known Limitations of Web-Search Approach vs. BuiltWith

| Factor | Web Search (current) | BuiltWith (future option) |
|---|---|---|
| Cost | Free | ~$500+/mo subscription |
| Volume per run | 10–15 manual discoveries | Hundreds of auto-exports |
| Data freshness | Current | Updated regularly |
| Contact data | Requires manual LinkedIn lookup | Often pre-populated |
| YM confirmation | Inferred from signals | Confirmed by tech tracking |
| Revenue/staff | Requires 990 lookup | Pre-populated estimates |
| Canada coverage | Requires separate search rotation | Separate export |

**Recommendation**: Web search is sufficient to build the initial 150–250 account
list. Revisit BuiltWith subscription when manual run volume exceeds 2–3 hours
per session or when Phase 2 enrichment costs are evaluated.
