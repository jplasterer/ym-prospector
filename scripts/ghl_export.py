"""
GHL Export Utility
Converts prospector Excel output to a Go High Level contact import CSV.
"""

import io
import re
import csv
import pandas as pd
from pathlib import Path


# GHL CSV column headers (must match GHL import field names exactly)
GHL_COLUMNS = [
    "First Name",
    "Last Name",
    "Email",
    "Phone",
    "Company Name",
    "Website",
    "City",
    "State",
    "Country",
    "Tags",
    "Source",
    "Notes",
    # Custom fields — GHL will create these if they don't exist
    "YM Confirmed",
    "Annual Revenue",
    "Staff Size",
    "Association Scope",
    "LMS Platform",
    "Community Platform",
    "Pain Signals",
    "Priority Tier",
    "Prospector Run Date",
]


def split_name(full_name: str) -> tuple[str, str]:
    """Split 'First Last' into (first, last). Handles empty/single-word names."""
    if not full_name or str(full_name).strip() in ("", "nan", "None"):
        return "", ""
    parts = str(full_name).strip().split(" ", 1)
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], parts[1]


def build_tags(row: pd.Series, run_date: str) -> str:
    """Build a comma-separated tag string for GHL from prospect data."""
    tags = ["YM Prospect"]

    tier = str(row.get("Priority Tier", "")).strip()
    if tier:
        tags.append(tier)

    scope = str(row.get("Association Scope", "")).strip()
    if scope and scope not in ("nan", "None", ""):
        tags.append(f"Scope: {scope}")

    ym = str(row.get("YM Confirmed?", "")).strip()
    if ym == "Yes":
        tags.append("YM Confirmed")
    elif ym == "Likely":
        tags.append("YM Likely")

    if run_date:
        tags.append(f"Prospector: {run_date}")

    return ", ".join(tags)


def build_notes(row: pd.Series) -> str:
    """Combine pain signal notes and other context into a GHL notes field."""
    parts = []

    pain = str(row.get("Pain Signal Notes", "")).strip()
    if pain and pain not in ("nan", "None", ""):
        parts.append(f"Pain Signals: {pain}")

    lms = str(row.get("LMS Platform", "")).strip()
    community = str(row.get("Community Platform", "")).strip()
    other = str(row.get("Other Platforms", "")).strip()

    platforms = [p for p in [lms, community, other] if p and p not in ("nan", "None", "")]
    if platforms:
        parts.append(f"Platforms: {', '.join(platforms)}")

    proj = str(row.get("Project Opportunity Notes", "")).strip()
    if proj and proj not in ("nan", "None", ""):
        parts.append(f"Project Opportunity: {proj}")

    return " | ".join(parts)


def excel_to_ghl_csv(
    excel_path: str,
    tiers: list[str] = None,
    run_date: str = "",
) -> tuple[str, int]:
    """
    Read a prospector Excel file and return a GHL-ready CSV string.

    Args:
        excel_path: Path to the ym_prospects_*.xlsx file
        tiers: List of tiers to include (default: Tier 1 + Tier 2)
        run_date: Date string for tagging (e.g. '2026-04-13')

    Returns:
        (csv_string, row_count)
    """
    if tiers is None:
        tiers = ["Strong Fit", "Medium Fit"]

    if not run_date:
        # Infer from filename
        name = Path(excel_path).stem  # e.g. ym_prospects_2026-04-13
        run_date = name.replace("ym_prospects_", "")

    # Read all sheets and combine
    xl = pd.read_excel(excel_path, sheet_name=None)
    frames = []
    for sheet_name, df in xl.items():
        if df.empty:
            continue
        # Normalize tier column — sheet name tells us the tier
        frames.append(df)

    if not frames:
        return "", 0

    all_prospects = pd.concat(frames, ignore_index=True)

    # Filter to requested tiers
    if "Priority Tier" in all_prospects.columns:
        filtered = all_prospects[all_prospects["Priority Tier"].isin(tiers)]
    else:
        filtered = all_prospects

    if filtered.empty:
        return "", 0

    # Build GHL rows
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=GHL_COLUMNS, extrasaction="ignore")
    writer.writeheader()

    for _, row in filtered.iterrows():
        first, last = split_name(row.get("Key Decision Maker Name", ""))

        email = str(row.get("Key Decision Maker Email", "")).strip()
        if email in ("nan", "None"):
            email = ""

        website = str(row.get("Website URL", "")).strip()
        if website in ("nan", "None"):
            website = ""

        city = str(row.get("City", "")).strip()
        if city in ("nan", "None"):
            city = ""

        state = str(row.get("State/Province", "")).strip()
        if state in ("nan", "None"):
            state = ""

        country = str(row.get("Country", "")).strip()
        if country in ("nan", "None", ""):
            country = "United States"

        revenue = str(row.get("Annual Revenue", "")).strip()
        if revenue in ("nan", "None"):
            revenue = ""

        staff = str(row.get("Staff Size", "")).strip()
        if staff in ("nan", "None"):
            staff = ""

        scope = str(row.get("Association Scope", "")).strip()
        if scope in ("nan", "None"):
            scope = ""

        lms = str(row.get("LMS Platform", "")).strip()
        if lms in ("nan", "None"):
            lms = ""

        community = str(row.get("Community Platform", "")).strip()
        if community in ("nan", "None"):
            community = ""

        writer.writerow({
            "First Name": first,
            "Last Name": last,
            "Email": email,
            "Phone": "",
            "Company Name": str(row.get("Organization Name", "")).strip(),
            "Website": website,
            "City": city,
            "State": state,
            "Country": country,
            "Tags": build_tags(row, run_date),
            "Source": "YM Prospector",
            "Notes": build_notes(row),
            "YM Confirmed": str(row.get("YM Confirmed?", "")).strip(),
            "Annual Revenue": revenue,
            "Staff Size": staff,
            "Association Scope": scope,
            "LMS Platform": lms,
            "Community Platform": community,
            "Pain Signals": str(row.get("Pain Signal Notes", "")).strip(),
            "Priority Tier": str(row.get("Priority Tier", "")).strip(),
            "Prospector Run Date": run_date,
        })

    csv_string = output.getvalue()
    row_count = len(filtered)
    return csv_string, row_count
