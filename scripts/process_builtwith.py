import csv
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

INPUT_CSV = '/tmp/builtwith_data/YourMembership_websites_in_the_United_States_-_2026-04-11.csv'
OUTPUT_XLSX = '/sessions/sweet-focused-pasteur/mnt/outputs/starkweather_seed_prospects.xlsx'

ASSOC_KEYWORDS = [
    'association', 'society', 'council', 'federation', 'institute',
    'network', 'guild', 'league', 'alliance', 'consortium', 'chapter',
    'academy', 'college of', 'board of', 'bureau', 'chamber', 'conference',
    'congress', 'foundation', 'professionals', 'practitioners'
]

EXCLUDE_KEYWORDS = [
    'homeowner', 'hoa', 'country club', 'yacht', 'golf', 'alumni',
    'school', 'university', 'church', 'temple', 'diocese', 'parish',
    'political', 'pac', 'campaign'
]

COMPLEXITY_PLATFORMS = [
    'informz', 'higher logic', 'topclass', 'freestone', 'crowd wisdom',
    'absorb', 'docebo', 'learnupon', 'epath', 'nimble ams', 'fonteva',
    'salesforce', 'hubspot', 'marketo', 'mailchimp', 'constant contact',
    'clover', 'authorize.net', 'stripe', 'paypal', 'imis', 'memberclicks'
]

def parse_revenue(rev_str):
    if not rev_str or rev_str.strip() == '':
        return None
    clean = re.sub(r'[,$]', '', rev_str.strip())
    try:
        return float(clean)
    except:
        return None

def parse_employees(emp_str):
    if not emp_str or emp_str.strip() == '':
        return None
    clean = emp_str.strip().replace(',', '')
    try:
        return int(float(clean))
    except:
        return None

def get_primary_email(emails_str):
    if not emails_str:
        return ''
    emails = [e.strip() for e in emails_str.split(';') if e.strip()]
    # Prefer non-generic emails
    for e in emails:
        local = e.split('@')[0].lower() if '@' in e else ''
        if not any(g in local for g in ['info', 'contact', 'admin', 'office', 'general', 'hello']):
            return e
    return emails[0] if emails else ''

def get_primary_phone(phones_str):
    if not phones_str:
        return ''
    phones = [p.replace('ph:', '').strip() for p in phones_str.split(';') if p.strip()]
    return phones[0] if phones else ''

def parse_people(people_str):
    if not people_str:
        return '', ''
    contacts = []
    for entry in people_str.split(';'):
        entry = entry.strip()
        if ' - ' in entry:
            parts = entry.split(' - ', 1)
            name = parts[0].strip()
            title = parts[1].strip()
            contacts.append((name, title))
    # Prefer exec director, CEO, executive director
    priority_titles = ['executive director', 'ceo', 'president', 'coo', 'chief executive',
                       'director of membership', 'membership director', 'it director',
                       'director of technology', 'director of education']
    for name, title in contacts:
        for pt in priority_titles:
            if pt in title.lower():
                return name, title
    return (contacts[0][0], contacts[0][1]) if contacts else ('', '')

def score_org(row):
    company = (row.get('Company', '') or '').lower()
    domain = (row.get('Root Domain', '') or '').lower()
    name_check = company if company else domain

    # Org type check
    is_assoc = any(k in name_check for k in ASSOC_KEYWORDS)
    is_excluded = any(k in name_check for k in EXCLUDE_KEYWORDS)
    vertical = (row.get('Vertical', '') or '').lower()
    if 'business' in vertical or 'professional' in vertical:
        is_assoc = True

    employees = parse_employees(row.get('Employees', ''))
    revenue = parse_revenue(row.get('Sales Revenue', ''))

    # Stack signals
    all_tech = ' '.join([
        row.get('CMS Platform', '') or '',
        row.get('CRM Platform', '') or '',
        row.get('Marketing Automation Platform', '') or '',
        row.get('Payment Platforms', '') or '',
        row.get('eCommerce Platform', '') or ''
    ]).lower()
    complexity_found = [p for p in COMPLEXITY_PLATFORMS if p in all_tech]

    # Staff fit
    if employees is None:
        staff_fit = 'Unknown – enrich'
        staff_score = 0
    elif 3 <= employees <= 8:
        staff_fit = f'Qualified ({employees})'
        staff_score = 2
    elif 9 <= employees <= 15:
        staff_fit = f'Marginal ({employees})'
        staff_score = 1
    else:
        staff_fit = f'Outside range ({employees})'
        staff_score = -1

    # Revenue fit
    if revenue is None:
        rev_fit = 'Unknown – enrich'
        rev_score = 0
    elif revenue >= 2_500_000:
        rev_fit = f'Qualified (${revenue:,.0f})'
        rev_score = 2
    elif revenue >= 1_000_000:
        rev_fit = f'Marginal (${revenue:,.0f})'
        rev_score = 1
    else:
        rev_fit = f'Too small (${revenue:,.0f})'
        rev_score = -1

    # Org type fit
    if is_excluded:
        org_fit = 'Excluded (wrong category)'
        org_score = -2
    elif is_assoc:
        org_fit = 'Likely association'
        org_score = 1
    else:
        org_fit = 'Unclear – review name'
        org_score = 0

    total = staff_score + rev_score + org_score

    # Tier assignment
    if is_excluded or staff_score == -1 or rev_score == -1:
        tier = 'Tier 3 – Disqualified'
    elif staff_score == 2 and rev_score == 2:
        tier = 'Tier 1 – Qualified'
    elif staff_score >= 1 and rev_score >= 1:
        tier = 'Tier 1 – Qualified'
    else:
        tier = 'Tier 2 – Needs Enrichment'

    return {
        'staff_fit': staff_fit,
        'rev_fit': rev_fit,
        'org_fit': org_fit,
        'complexity': '; '.join(complexity_found) if complexity_found else '',
        'tier': tier,
        'total_score': total
    }

def load_rows():
    rows = []
    with open(INPUT_CSV, encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows

def safe(row, key):
    val = row.get(key)
    return val.strip() if val else ''

def build_output_row(row):
    company = safe(row, 'Company') or safe(row, 'Root Domain')
    domain = safe(row, 'Root Domain')
    city = safe(row, 'City')
    state = safe(row, 'State')
    location = f"{city}, {state}" if city and state else city or state or ''
    name, title = parse_people(safe(row, 'People'))
    email = get_primary_email(safe(row, 'Emails'))
    phone = get_primary_phone(safe(row, 'Telephones'))
    linkedin = safe(row, 'LinkedIn')
    mkt_auto = safe(row, 'Marketing Automation Platform')
    crm = safe(row, 'CRM Platform')
    first_detected = safe(row, 'First Detected')
    last_found = safe(row, 'Last Found')
    s = score_org(row)
    return [
        company, domain, location,
        safe(row, 'Employees'),
        safe(row, 'Sales Revenue'),
        name, title, email, phone, linkedin,
        mkt_auto, crm,
        first_detected, last_found,
        s['org_fit'], s['staff_fit'], s['rev_fit'],
        s['complexity'], s['tier']
    ]

HEADERS = [
    'Organization Name', 'Website', 'Location',
    'Employees (BuiltWith)', 'Revenue (BuiltWith)',
    'Key Contact Name', 'Key Contact Title', 'Email', 'Phone', 'LinkedIn',
    'Marketing Automation', 'CRM Platform',
    'First on YM', 'Last Confirmed YM',
    'Org Type Fit', 'Staff Fit', 'Revenue Fit',
    'Complexity Signals', 'Priority Tier'
]

# Colors
COLORS = {
    'header_bg': 'FF1F3864',   # dark navy
    'header_fg': 'FFFFFFFF',
    'tier1_bg': 'FFE2EFDA',    # light green
    'tier2_bg': 'FFFFF2CC',    # light yellow
    'tier3_bg': 'FFFCE4D6',    # light red
    'alt_row': 'FFF2F2F2',
    'tab_t1': 'FF375623',
    'tab_t2': 'FF7F6000',
    'tab_t3': 'FF833C00',
}

def style_header(ws, num_cols):
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(name='Arial', bold=True, color=COLORS['header_fg'], size=10)
        cell.fill = PatternFill('solid', start_color=COLORS['header_bg'])
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 36

def style_data_row(ws, row_num, tier, num_cols):
    if 'Tier 1' in tier:
        bg = COLORS['tier1_bg']
    elif 'Tier 2' in tier:
        bg = COLORS['tier2_bg']
    else:
        bg = COLORS['tier3_bg']
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.fill = PatternFill('solid', start_color=bg)
        cell.font = Font(name='Arial', size=9)
        cell.alignment = Alignment(vertical='top', wrap_text=(col in [1, 6, 7, 18]))

def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def add_sheet(wb, title, data_rows, tab_color):
    ws = wb.create_sheet(title)
    ws.tab_color = tab_color
    ws.freeze_panes = 'A2'
    ws.append(HEADERS)
    style_header(ws, len(HEADERS))
    for i, row in enumerate(data_rows, 2):
        ws.append(row)
        tier = row[-1] if row else ''
        style_data_row(ws, i, tier, len(HEADERS))
    set_col_widths(ws, [
        32, 28, 20,   # name, website, location
        12, 16,        # employees, revenue
        24, 28,        # contact name, title
        32, 18, 30,    # email, phone, linkedin
        22, 18,        # mkt auto, crm
        14, 14,        # first/last on YM
        22, 24, 24,    # org/staff/rev fit
        28, 26         # complexity, tier
    ])
    return ws

def main():
    rows = load_rows()
    t1, t2, t3 = [], [], []
    for row in rows:
        out = build_output_row(row)
        tier = out[-1]
        if 'Tier 1' in tier:
            t1.append(out)
        elif 'Tier 2' in tier:
            t2.append(out)
        else:
            t3.append(out)

    wb = openpyxl.Workbook()
    # Remove default sheet
    del wb['Sheet']

    add_sheet(wb, f'Tier 1 – Qualified ({len(t1)})', t1, COLORS['tab_t1'])
    add_sheet(wb, f'Tier 2 – Needs Enrichment ({len(t2)})', t2, COLORS['tab_t2'])
    add_sheet(wb, f'Tier 3 – Disqualified ({len(t3)})', t3, COLORS['tab_t3'])

    # Summary sheet
    ws_sum = wb.create_sheet('Summary', 0)
    ws_sum.tab_color = 'FF1F3864'
    ws_sum.column_dimensions['A'].width = 35
    ws_sum.column_dimensions['B'].width = 18

    summary_data = [
        ['STARKWEATHER SEED PROSPECT ANALYSIS', ''],
        ['Source: BuiltWith YM US Export – April 11 2026', ''],
        ['', ''],
        ['METRIC', 'COUNT'],
        ['Total records in source file', len(rows)],
        ['Tier 1 – Qualified (meets known ICP)', len(t1)],
        ['Tier 2 – Needs Enrichment (data gaps)', len(t2)],
        ['Tier 3 – Disqualified (outside ICP)', len(t3)],
        ['', ''],
        ['ICP CRITERIA APPLIED', ''],
        ['Staff size target', '3–8 employees'],
        ['Revenue target', '$2.5M+ annual'],
        ['Org types targeted', 'Prof. & trade associations'],
        ['Platform confirmed', 'YourMembership AMS'],
        ['', ''],
        ['NEXT STEPS', ''],
        ['1. Enrich Tier 2 via ProPublica 990 lookups', ''],
        ['2. Verify org type for "Unclear" Tier 2 rows', ''],
        ['3. Find decision-maker contacts for Tier 1', ''],
        ['4. Add Canada YM associations to this list', ''],
        ['5. Build nightly web-search discovery tool', ''],
    ]

    for r_idx, (label, value) in enumerate(summary_data, 1):
        ws_sum.cell(r_idx, 1, label)
        ws_sum.cell(r_idx, 2, value)

    # Style summary
    ws_sum.cell(1, 1).font = Font(name='Arial', bold=True, size=13, color='FF1F3864')
    ws_sum.cell(2, 1).font = Font(name='Arial', italic=True, size=9, color='FF666666')
    for col in [1, 2]:
        ws_sum.cell(4, col).font = Font(name='Arial', bold=True, color='FFFFFFFF', size=10)
        ws_sum.cell(4, col).fill = PatternFill('solid', start_color='FF1F3864')
        ws_sum.cell(10, col).font = Font(name='Arial', bold=True, color='FFFFFFFF', size=10)
        ws_sum.cell(10, col).fill = PatternFill('solid', start_color='FF1F3864')
        ws_sum.cell(16, col).font = Font(name='Arial', bold=True, color='FFFFFFFF', size=10)
        ws_sum.cell(16, col).fill = PatternFill('solid', start_color='FF1F3864')

    for r_idx in [5, 6, 7, 8]:
        tier_colors = {5: None, 6: COLORS['tier1_bg'], 7: COLORS['tier2_bg'], 8: COLORS['tier3_bg']}
        if tier_colors[r_idx]:
            for col in [1, 2]:
                ws_sum.cell(r_idx, col).fill = PatternFill('solid', start_color=tier_colors[r_idx])
        ws_sum.cell(r_idx, 1).font = Font(name='Arial', size=10)
        ws_sum.cell(r_idx, 2).font = Font(name='Arial', bold=True, size=10)

    wb.save(OUTPUT_XLSX)
    print(f"Saved: {OUTPUT_XLSX}")
    print(f"Tier 1: {len(t1)} | Tier 2: {len(t2)} | Tier 3: {len(t3)} | Total: {len(rows)}")

main()
