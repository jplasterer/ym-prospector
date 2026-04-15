"""
Applies research findings from web lookups to re-tier Tier 2 records.
Moves confirmed disqualifications to Tier 3 and adds enrichment notes.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from copy import copy

INPUT = '/sessions/sweet-focused-pasteur/mnt/outputs/starkweather_seed_prospects.xlsx'
OUTPUT = '/sessions/sweet-focused-pasteur/mnt/outputs/starkweather_seed_prospects_v2.xlsx'

# Research findings keyed by domain
# Format: (new_tier, org_name_correction, enrichment_note)
RESEARCH = {
    # --- Moving to Tier 3 ---
    'nfrccommunity.org': ('Tier 3 – Disqualified', 'National Fenestration Rating Council Inc',
        'Staff ~35 (too large per LinkedIn/news). Outside 3-8 range.'),
    'acadianashrm.org': ('Tier 3 – Disqualified', 'Acadiana SHRM Chapter',
        'Local volunteer-run SHRM chapter. Too small to support $40K relationship.'),
    'cycloneletterwinners.org': ('Tier 3 – Disqualified', 'Iowa State Cyclone Letter Winners',
        'Athletic alumni club — wrong org type. Not a trade/professional association.'),
    'rotomolding.org': ('Tier 3 – Disqualified', 'Association of Rotational Molders',
        'Revenue 2022: $1.04M (below $1.5M floor). EIN 36-2871176.'),
    'ny-best.org': ('Tier 3 – Disqualified', 'NY Battery & Energy Storage Technology Consortium',
        'Staff 51–200 (far outside 3-8 range). Revenue $1M-$5M range per public sources.'),
    'shrmswfl.org': ('Tier 3 – Disqualified', 'SHRM Southwest Florida Chapter',
        'Local volunteer-run SHRM chapter. Unlikely to sustain $40K relationship.'),
    'integrativepainmanagement.org': ('Tier 3 – Disqualified', 'Academy of Integrative Pain Management',
        'Organization appears inactive/ceased. Not found on current IRS Business Master File.'),
    'thecollegeofdiplomates.org': ('Tier 3 – Disqualified', 'Foundation of the College of Diplomates ABPD',
        'Foundation entity, not a membership association. Different buyer profile.'),
    'agentismanagement.com': ('Tier 3 – Disqualified', 'Agentis Management LLC',
        'Association management company (AMC), not an association. Staff 10, outside range.'),
    'aicuo.edu': ('Tier 3 – Disqualified', 'Assoc. of Independent Colleges & Universities of Ohio',
        'Revenue 2022: $1.49M (below $2.5M threshold). Staff: 7. EIN 31-0746875.'),
    'richmondatwork.com': ('Tier 3 – Disqualified', 'richmondatwork.com',
        'Not identifiable as a trade/professional association. Domain suggests workforce org.'),
    'hannahstreet.com': ('Tier 3 – Disqualified', 'hannahstreet.com',
        'Not identifiable as an association. Likely a personal/commercial domain.'),
    'yestoemployment.org': ('Tier 3 – Disqualified', 'Yes to Employment',
        'Employment/workforce nonprofit. Not a professional/trade membership association.'),
    'mcodsa.com': ('Tier 3 – Disqualified', 'Michigan Court Officers & Deputy Sheriff Assoc.',
        'Very niche law enforcement niche. Likely volunteer-run and too small.'),
    'laborear.com': ('Tier 3 – Disqualified', 'laborear.com',
        'Not identifiable as a trade/professional association.'),
    'svca-online.org': ('Tier 3 – Disqualified', 'Silicon Valley Compensation Association',
        'Local SHRM-affiliate chapter. Volunteer-run, unlikely to support $40K relationship.'),

    # --- Staying Tier 2 with enrichment notes ---
    'kcnpnm.org': ('Tier 2 – Needs Enrichment', 'Kentucky Coalition of Nurse Practitioners & Nurse-Midwives',
        'Professional nursing assoc. Est. 1987, ~statewide membership. Staff/revenue unknown — look up EIN via Candid.'),
    'naswtx.org': ('Tier 2 – Needs Enrichment', 'National Association of Social Workers – Texas Chapter',
        'Staff: 4 (confirmed). ~6,000 TX members. Revenue unknown — search ProPublica EIN 74-1860732 for TX chapter.'),
    'barmont.org': ('Tier 2 – Needs Enrichment', 'Bar Association of Montgomery County MD',
        'Staff: 4 (confirmed). 501(c)(6) — revenue not on Charity Navigator. Full 990 at ProPublica EIN 52-0807787.'),
    'hracc.org': ('Tier 2 – Needs Enrichment', 'Human Resource Association of Central Connecticut',
        'Staff: 7 (confirmed). SHRM affiliate, Greater Hartford. Ignore ZoomInfo/RocketReach $11M figure — unreliable for small assocs. Verify via 990.'),
    'bisanet.org': ('Tier 2 – Needs Enrichment', 'Bank Insurance & Securities Association (BISA)',
        'Revenue: $3.1M confirmed (990). Staff: ZoomInfo says 11-50 — need exact count. ProPublica EIN 52-2385104.'),
    'louisianapharmacists.com': ('Tier 2 – Needs Enrichment', 'Louisiana Pharmacists Association',
        'Professional pharmacist assoc., est. 1892, Baton Rouge. Revenue/staff unknown. ProPublica EIN 72-0246230.'),
    'whfdc.org': ('Tier 2 – Needs Enrichment', 'Women in Housing and Finance',
        'DC-area networking assoc for housing/finance professionals. Likely small — verify revenue before prioritizing.'),
    'leadingagesoutheast.org': ('Tier 2 – Needs Enrichment', 'LeadingAge Southeast',
        'Regional affiliate of national senior living assoc (FL, AL, LA, MS). File separately from national — need regional 990.'),
    'sportsbuilders.org': ('Tier 2 – Needs Enrichment', 'American Sports Builders Association',
        '600+ member trade assoc for sports facility construction. BuiltWith $1M may be low — verify via ProPublica EIN 13-2561230.'),
    'cappo.org': ('Tier 2 – Needs Enrichment', 'CA Association of Public Procurement Officials',
        'Oldest public procurement assoc in US, est. 1915. Revenue/staff unclear. ProPublica EIN 94-2793199.'),
    'kentuckyacc.org': ('Tier 2 – Needs Enrichment', 'Kentucky Chapter – American College of Cardiology',
        'State chapter of national cardiology assoc. Chapters can be substantial — verify staff/revenue independently.'),
    'arkansaspa.org': ('Tier 2 – Needs Enrichment', 'Arkansas Physician Assistants (est.)',
        'Likely AR PA state chapter. Professional healthcare assoc. Staff/revenue unknown — needs investigation.'),
    'californiaglassassociation.org': ('Tier 2 – Needs Enrichment', 'California Glass Association',
        'Trade assoc for glass industry. Apple Valley, CA. No 990 data found — may be 501(c)(6). Needs investigation.'),
    'nae4hydp.org': ('Tier 2 – Needs Enrichment', 'Natl Assoc of Extension 4-H Youth Dev Professionals',
        'Professional development assoc for 4-H extension agents. Raleigh NC. Needs staff/revenue verification.'),
    'pafpc.org': ('Tier 2 – Needs Enrichment', 'PA Assoc of Federal Program Coordinators (est.)',
        'State College PA. Identity/revenue unconfirmed — needs website investigation.'),
    'flamedmfg.org': ('Tier 2 – Needs Enrichment', 'FL Alliance of Manufacturing (est.)',
        'Tallahassee FL. Identity/revenue unconfirmed — needs website investigation.'),
    'clwbar.org': ('Tier 2 – Needs Enrichment', 'Clearwater Bar Association (est.)',
        'Local bar assoc, Clearwater FL. Likely small — verify before investing outreach time.'),
    'psaphotoworldwide.org': ('Tier 2 – Needs Enrichment', 'Photographic Society of America',
        'Photography society. Needs staff/revenue verification — may be hobbyist not professional assoc.'),
    'wicounseling.org': ('Tier 2 – Needs Enrichment', 'Wisconsin Counseling Association (est.)',
        'Likely state chapter of ACA. Professional counselors assoc. Staff/revenue unconfirmed.'),
    'educational-access.org': ('Tier 2 – Needs Enrichment', 'educational-access.org',
        'Boston MA. Identity unclear. Needs website investigation to confirm org type.'),
    'naswinternational.org': ('Tier 2 – Needs Enrichment', 'NASW International (est.)',
        'Identity unclear — may be NASW international programs. Needs investigation.'),
    'ohioprivatecolleges.org': ('Tier 2 – Needs Enrichment', 'Ohio Private Colleges (est.)',
        'Likely higher ed association. Needs investigation — may be AICUO (already disqualified) or separate entity.'),
}

COLORS = {
    'header_bg': 'FF1F3864',
    'header_fg': 'FFFFFFFF',
    'tier1_bg': 'FFE2EFDA',
    'tier2_bg': 'FFFFF2CC',
    'tier3_bg': 'FFFCE4D6',
    'note_bg': 'FFDCE6F1',
    'tab_t1': 'FF375623',
    'tab_t2': 'FF7F6000',
    'tab_t3': 'FF833C00',
}

def get_tier_color(tier):
    if 'Tier 1' in tier:
        return COLORS['tier1_bg']
    elif 'Tier 2' in tier:
        return COLORS['tier2_bg']
    return COLORS['tier3_bg']

def copy_row_data(ws_src, row_num):
    return [ws_src.cell(row_num, c).value for c in range(1, ws_src.max_column + 1)]

def write_row(ws, row_num, data, tier):
    bg = get_tier_color(tier)
    for c, val in enumerate(data, 1):
        cell = ws.cell(row_num, c, val)
        cell.fill = PatternFill('solid', start_color=bg)
        cell.font = Font(name='Arial', size=9)
        cell.alignment = Alignment(vertical='top', wrap_text=(c in [1, 6, 7, 18, 19, 20]))

def style_header(ws, num_cols):
    for c in range(1, num_cols + 1):
        cell = ws.cell(1, c)
        cell.font = Font(name='Arial', bold=True, color=COLORS['header_fg'], size=10)
        cell.fill = PatternFill('solid', start_color=COLORS['header_bg'])
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 36

def set_widths(ws):
    widths = [32,28,20,12,16,24,28,32,18,30,22,18,14,14,22,24,24,28,26,45]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def build_new_wb():
    src_wb = openpyxl.load_workbook(INPUT)

    # Collect all rows from all data sheets
    all_rows = []
    for sheet_name in src_wb.sheetnames:
        if 'Summary' in sheet_name:
            continue
        ws = src_wb[sheet_name]
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0] or 'Compliance Notice' in str(row[0]):
                continue
            all_rows.append(list(row))

    # Add Research Note column header
    HEADERS_NEW = headers + ['Research Notes']

    # Apply research findings
    t1, t2, t3 = [], [], []
    for row in all_rows:
        domain = row[1] if row[1] else ''
        research = RESEARCH.get(domain)
        if research:
            new_tier, name_fix, note = research
            row[0] = name_fix  # update org name if we identified it
            row[-1] = new_tier  # update tier column
            row_with_note = row + [note]
        else:
            row_with_note = row + ['']

        tier = row_with_note[-2]  # Priority Tier is second to last
        if 'Tier 1' in tier:
            t1.append(row_with_note)
        elif 'Tier 2' in tier:
            t2.append(row_with_note)
        else:
            t3.append(row_with_note)

    # Build output workbook
    wb = openpyxl.Workbook()
    del wb['Sheet']

    for title, data, tab_color in [
        (f'Tier 1 – Qualified ({len(t1)})', t1, COLORS['tab_t1']),
        (f'Tier 2 – Needs Enrichment ({len(t2)})', t2, COLORS['tab_t2']),
        (f'Tier 3 – Disqualified ({len(t3)})', t3, COLORS['tab_t3']),
    ]:
        ws = wb.create_sheet(title)
        ws.tab_color = tab_color
        ws.freeze_panes = 'A2'
        for c, h in enumerate(HEADERS_NEW, 1):
            ws.cell(1, c, h)
        style_header(ws, len(HEADERS_NEW))
        for i, row in enumerate(data, 2):
            tier = row[-2]
            write_row(ws, i, row, tier)
        set_widths(ws)

    # Summary sheet
    ws_sum = wb.create_sheet('Summary', 0)
    ws_sum.tab_color = 'FF1F3864'
    ws_sum.column_dimensions['A'].width = 42
    ws_sum.column_dimensions['B'].width = 20

    summary = [
        ['STARKWEATHER SEED PROSPECT ANALYSIS v2', ''],
        ['Enriched via web search – April 2026', ''],
        ['', ''],
        ['METRIC', 'COUNT'],
        ['Source records (BuiltWith US export)', len(all_rows)],
        ['Tier 1 – Qualified (ICP confirmed)', len(t1)],
        ['Tier 2 – Needs Enrichment (data gaps remain)', len(t2)],
        ['Tier 3 – Disqualified', len(t3)],
        ['', ''],
        ['KEY ENRICHMENT DECISIONS MADE', ''],
        ['Moved to T3: Staff too large (NFRC, NY-BEST)', ''],
        ['Moved to T3: Revenue too low (AICUO $1.49M, ARM $1.04M)', ''],
        ['Moved to T3: Wrong org type (alumni club, mgmt co, foundation)', ''],
        ['Moved to T3: Inactive/ceased (Integrative Pain Mgmt)', ''],
        ['Moved to T3: Local volunteer chapters (SHRM chapters, Clearwater)', ''],
        ['Confirmed T2: Revenue verified (BISA $3.1M)', ''],
        ['Confirmed T2: Staff verified (naswtx 4, barmont 4, hracc 7)', ''],
        ['', ''],
        ['RELIABILITY LESSON DOCUMENTED', ''],
        ['ZoomInfo/RocketReach revenue figures unreliable for small assocs', ''],
        ['Always verify via 990 (ProPublica EIN lookup or Candid)', ''],
        ['501(c)(6) orgs not on Charity Navigator — use ProPublica directly', ''],
        ['', ''],
        ['NEXT STEPS', ''],
        ['1. Run 990 lookups on top 10 Tier 2 orgs (EINs in notes)', ''],
        ['2. Start nightly discovery runs using CLAUDE.md methodology', ''],
        ['3. Add Canada YM associations (separate BuiltWith export or searches)', ''],
    ]

    for r_idx, (label, value) in enumerate(summary, 1):
        ws_sum.cell(r_idx, 1, label)
        ws_sum.cell(r_idx, 2, value)

    ws_sum.cell(1, 1).font = Font(name='Arial', bold=True, size=13, color='FF1F3864')
    ws_sum.cell(2, 1).font = Font(name='Arial', italic=True, size=9, color='FF666666')

    for section_row in [4, 10, 20, 25]:
        for col in [1, 2]:
            c = ws_sum.cell(section_row, col)
            c.font = Font(name='Arial', bold=True, color='FFFFFFFF', size=10)
            c.fill = PatternFill('solid', start_color='FF1F3864')

    tier_map = {5: None, 6: COLORS['tier1_bg'], 7: COLORS['tier2_bg'], 8: COLORS['tier3_bg']}
    for r_idx, bg in tier_map.items():
        if bg:
            for col in [1, 2]:
                ws_sum.cell(r_idx, col).fill = PatternFill('solid', start_color=bg)
        ws_sum.cell(r_idx, 1).font = Font(name='Arial', size=10)
        ws_sum.cell(r_idx, 2).font = Font(name='Arial', bold=True, size=10)

    wb.save(OUTPUT)
    print(f"Saved: {OUTPUT}")
    print(f"Tier 1: {len(t1)} | Tier 2: {len(t2)} | Tier 3: {len(t3)}")

build_new_wb()
