# Starkweather Prospector — Project Briefing

## Who We Are

Starkweather Association Services (https://starkweather.us) is a managed service and
solutions provider for associations running on the YourMembership (YM) AMS platform.
We have the deepest YourMembership expertise of any firm in the market. We own and
execute revenue-impacting YM workflows for lean associations that cannot afford to have
critical operations stall.

**Commercial goal**: Sign 21 new YM association clients at $40,000+ annual value
to exceed $1M in annual contracted revenue. This requires building a named-account
list of 150–250 qualified YM associations, generating approximately 168 strong
discovery conversations.

## What This Tool Does

This project runs nightly web searches to discover US and Canada based trade and
professional associations that are currently using the YourMembership AMS platform,
then scores them against our Ideal Customer Profile. Results are output as a
scored Excel file. Prospects not already in our Go High Level CRM are surfaced
for outreach.

## Ideal Customer Profile (ICP)

### Must-Have Criteria (all required)

| Criterion | Target |
|---|---|
| Platform | Confirmed or strongly indicated YourMembership AMS user |
| Org type | Professional association or trade association |
| Geography | United States or Canada |
| Staff size | 3–8 full-time employees (sweet spot: 4–5) |
| Annual revenue | $2.5M+ (sweet spot: ~$3M) |

### Complexity Signals (score higher when present)

An ideal prospect runs YM alongside one or more additional platforms, signaling
operational complexity and integration pain:

- **LMS / Education platform**: Crowd Wisdom, Freestone, TopClass, Absorb,
  Docebo, LearnUpon, ePath Learning, or any dedicated CE/certification platform
- **Online community**: Higher Logic, Personify Community, Mighty Networks,
  or similar member community platform
- **Email marketing**: Informz, Feathr, Constant Contact, or similar
- **Events platform**: Cvent, EventMobi, or similar standalone event tool
- **Chapter management**: Running regional chapters or affiliates
- **Certification / credentialing program**: CE credits, CPE, CME, etc.

The ideal prospect has YM PLUS at least one of these. The more platforms they
juggle, the stronger the pain signal.

### Pain Points to Listen For During Research

These are the pains that make prospects most likely to buy:

1. **Bandwidth shortage** — small staff, too much YM work, things falling behind
2. **Renewal risk** — dues renewal workflows unreliable, membership revenue at risk
3. **Events failure risk** — event setup, registration, or reporting breaking down
4. **Education/certification delivery** — CE program stalling or manual
5. **Reporting and integrations** — data stuck in silos, manual exports
6. **Slow support** — frustrated with YM's own support tier, looking for experts
7. **Reactive firefighting** — constantly fixing problems after they happen

### Disqualifiers (exclude these)

- Volunteer-run organizations with no paid staff
- Annual revenue below $1.5M (cannot sustain a $40K+ relationship)
- Fewer than 3 full-time employees
- Staff larger than 15 (strong internal capacity, less pain)
- Organizations looking only for ad hoc/hourly break-fix help
- Strong internal YM admin with no visible operational pain
- Homeowners associations (HOAs)
- Country clubs, yacht clubs, golf clubs
- Alumni associations (school or university-based)
- Athletic/letter-winner clubs and sports fan organizations
- Political organizations or PACs
- Churches, dioceses, religious organizations
- Foundations (separate entities from their parent associations)
- Association Management Companies (AMCs) — they ARE competitors/peers
- Local volunteer chapters of national associations (e.g. local SHRM, NAHRO chapters)
  unless they clearly have paid staff and operate independently with their own budget
- Organizations that are already Starkweather clients

## Web Search Discovery Methodology

### Why This Works Without a BuiltWith Subscription

YourMembership hosts member portals on its own infrastructure. Many YM associations
have member login pages, job boards, and community portals at URLs containing
`yourmembership.com`. This creates searchable fingerprints.

### Primary Search Patterns

Use these search query templates to find new YM associations:

```
"login.yourmembership.com" [industry keyword]
"members.yourmembership.com" [industry keyword]
site:yourmembership.com [industry keyword] association
"yourmembership" "association" "member login" [industry]
"powered by yourmembership" [industry]
inurl:yourmembership.com [association type]
```

### Industry/Sector Search Rotation

Rotate through these sectors each nightly run to build coverage systematically:

Week 1: Healthcare (nursing, physician, dental, pharmacy, veterinary)
Week 2: Legal and compliance (bar associations, paralegal, compliance officers)
Week 3: Engineering and construction (civil, structural, mechanical, contractors)
Week 4: Finance and accounting (CPA, financial planners, credit unions)
Week 5: Education and credentialing (educators, training professionals, HR)
Week 6: Technology and IT (IT professionals, cybersecurity, data)
Week 7: Real estate and property (Realtors, appraisers, property managers)
Week 8: Nonprofit and association management (AMC industry, nonprofit execs)
Week 9: Manufacturing and industry (manufacturers, industrial, trade)
Week 10: Environmental and energy (environmental professionals, energy)
Week 11: Media and communications (PR, marketing, communications)
Week 12: Canada-specific (run all above sectors with Canada geography filter)

### Verification Steps Per Prospect

For each candidate org found via web search:

1. **Confirm YM usage**: Visit their member portal or website footer for YM signals
   - Look for: "yourmembership.com" in login URLs, "Powered by YourMembership",
     YM job board embeds, YM member portal redirects
2. **Check for LMS or community**: Visit Education/Learning/Community sections
   - Note platform name if visible in footer, URL, or "Powered by" text
3. **Verify staff size**: Check LinkedIn company page for headcount
   - Do NOT trust ZoomInfo or RocketReach revenue/headcount estimates for small
     associations — these are consistently unreliable. Always prefer LinkedIn
     headcount and 990 revenue.
4. **Verify revenue**: Search ProPublica Nonprofit Explorer (propublica.org/nonprofits)
   or Candid (candid.org) for IRS 990 data (revenue, staff count)
   - Search by organization name or EIN
   - Use most recent 990 filing available
   - Note: 501(c)(6) trade associations ARE on ProPublica but not on Charity Navigator
   - Note: Regional chapter/affiliate 990s are filed separately from national orgs
5. **Find decision-maker**: LinkedIn search for ED, CEO, Director of Membership,
   or IT Director at the organization
6. **Score complexity**: Count how many platforms beyond YM they appear to run

### Common Enrichment Traps (Learned from Seed List Research)

These mistakes were identified during enrichment of the April 2026 seed list.
Avoid them in future runs:

| Trap | What to Do Instead |
|---|---|
| Trusting ZoomInfo/RocketReach revenue for small assocs | Always verify via ProPublica 990 — third-party estimates can be off by 10x |
| Assuming a local chapter has its own budget | Check if the chapter files its own 990 — many don't and are volunteer-run |
| Treating foundations as associations | "[Org] Foundation" is a different entity — verify it has its own membership |
| Including AMCs as prospects | Association Management Companies are peers/competitors, not prospects |
| Counting volunteer activity as staff | LinkedIn "employees" sometimes includes volunteers — cross-check with 990 |
| Assuming inactive domains are dead | Some YM portals on old domains are still active — check Last Found date |

## Output Requirements

Save each nightly run as: `ym_prospects_YYYY-MM-DD.xlsx`

### Required Fields Per Prospect

| Field | Source |
|---|---|
| Organization Name | Website / LinkedIn |
| Website URL | Discovery search |
| City, State/Province, Country | Website / LinkedIn |
| Staff Size | LinkedIn / 990 |
| Annual Revenue | ProPublica 990 / Candid |
| YM Confirmed? | Yes / Likely / Unconfirmed |
| LMS Platform | Website inspection |
| Community Platform | Website inspection |
| Other Platforms | Website inspection |
| Key Decision Maker Name | LinkedIn |
| Key Decision Maker Title | LinkedIn |
| Key Decision Maker Email | Website / LinkedIn |
| Pain Signal Notes | Research notes |
| Priority Tier | See below |

### Tier Definitions

- **Tier 1 — Qualified**: Confirmed YM + staff 3–8 + revenue $2.5M+ + at least one complexity platform
- **Tier 2 — Qualified (enrich)**: Confirmed YM + meets size/revenue BUT missing one data point to confirm
- **Tier 3 — Watch**: YM likely but unconfirmed, or borderline on size/revenue
- **Exclude**: Outside ICP criteria — log reason, do not surface for outreach

### Sheet Structure

Sheet 1: `Tier 1 – Ready for Outreach`
Sheet 2: `Tier 2 – Needs Enrichment`
Sheet 3: `Tier 3 – Watch List`
Sheet 4: `Excluded` (with reason column)

## Positioning Context for Research Notes

When you find a prospect, note which of these pain signals are visible from
public information. These become the hook for outreach:

- **Renewal risk**: Does their membership renewal page look manual or broken?
- **Events**: Do they run an annual conference with registration issues?
- **Education**: Do they have CE/certification programs that look complex?
- **Tech stack**: Are they clearly running 3+ platforms that don't talk to each other?
- **Hiring signals**: Are they posting for YM admin, database manager, or similar roles?
  (Hiring = YM pain they can't solve internally)
- **New ED**: Did they recently hire a new Executive Director?
  (New leader = likely reviewing vendors and processes)

## Target Volume

- **Per nightly run**: 10–15 newly researched prospects
- **Total named account list target**: 150–250 qualified organizations
- **Estimated sessions to reach target**: 15–20 nightly runs covering all sectors
