"""
Starkweather YM Prospector — Streamlit UI
Run with: streamlit run app.py
"""

import sys
from pathlib import Path

# Ensure project root is on the path
sys.path.insert(0, str(Path(__file__).parent))

import yaml
import streamlit as st
import pandas as pd
import streamlit_authenticator as stauth
from pathlib import Path
from dotenv import load_dotenv

from scripts.prospector import run_session, enrich_tier2, SECTORS
from scripts.ghl_export import excel_to_ghl_csv

load_dotenv()

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Starkweather YM Prospector",
    page_icon=str(Path(__file__).parent / "assets" / "logo.png"),
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    /* Starkweather brand colors */
    .stButton > button[kind="primary"] {
        background-color: #1b5897;
        border-color: #1b5897;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #154778;
        border-color: #154778;
    }
    /* Metric cards */
    div[data-testid="metric-container"] {
        background-color: #f0f4f9;
        border-left: 4px solid #1b5897;
        border-radius: 6px;
        padding: 0.5rem;
    }
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #f0f4f9;
        border-radius: 6px 6px 0 0;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1b5897 !important;
        color: white !important;
    }
    /* Sidebar background */
    section[data-testid="stSidebar"] {
        background-color: #f0f4f9;
        border-right: 2px solid #1b5897;
    }
</style>
""", unsafe_allow_html=True)

OUTPUT_DIR = str(Path(__file__).parent / "output" / "nightly")

# ---------------------------------------------------------------------------
# Version
# ---------------------------------------------------------------------------

def _get_version() -> str:
    """Read version from VERSION file."""
    version_path = Path(__file__).parent / "VERSION"
    try:
        return version_path.read_text(encoding="utf-8").strip()
    except Exception:
        return "unknown"

APP_VERSION = _get_version()

# ---------------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------------

_auth_config_path = Path(__file__).parent / "auth_config.yaml"
with open(_auth_config_path) as f:
    _auth_config = yaml.safe_load(f)

authenticator = stauth.Authenticate(
    _auth_config["credentials"],
    _auth_config["cookie"]["name"],
    _auth_config["cookie"]["key"],
    _auth_config["cookie"]["expiry_days"],
)

authenticator.login()

if st.session_state.get("authentication_status") is False:
    st.error("Incorrect username or password.")
    st.stop()
elif st.session_state.get("authentication_status") is None:
    st.info("Please log in to access the Starkweather YM Prospector.")
    st.stop()

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------

if "results" not in st.session_state:
    st.session_state.results = None
if "output_path" not in st.session_state:
    st.session_state.output_path = None
if "last_sector" not in st.session_state:
    st.session_state.last_sector = None

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_run_history() -> list[Path]:
    path = Path(OUTPUT_DIR)
    if not path.exists():
        return []
    return sorted(path.glob("ym_prospects_*.xlsx"), reverse=True)


def tier_color(tier: str) -> str:
    return {
        "Strong Fit":   "#C6EFCE",
        "Medium Fit":   "#FFEB9C",
        "Unlikely Fit": "#FFCC99",
        "Not a Fit":    "#FFC7CE",
    }.get(tier, "#FFFFFF")


def linkedin_message_url(profile_url: str) -> str:
    """Build a LinkedIn message compose URL from a profile URL."""
    if not profile_url or str(profile_url).strip() in ("", "nan", "None"):
        return ""
    import re
    match = re.search(r"linkedin\.com/in/([^/?#\s]+)", str(profile_url))
    if match:
        vanity = match.group(1)
        return f"https://www.linkedin.com/messaging/compose?recipient={vanity}"
    return str(profile_url)  # fallback to profile URL


def style_tier_row(row):
    color = tier_color(row.get("tier", ""))
    return [f"background-color: {color}"] * len(row)


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

with st.sidebar:
    logo_path = Path(__file__).parent / "assets" / "logo.png"
    if logo_path.exists():
        st.image(str(logo_path), width=160)
    st.markdown("### YM Prospector")
    st.caption(f"Logged in as **{st.session_state.get('name', '')}**")
    st.caption(f"Version {APP_VERSION}")
    authenticator.logout("Log Out", "sidebar")
    st.divider()

    # Chrome Mode toggle
    st.markdown("#### 🔗 Chrome + LinkedIn Mode")
    chrome_mode = st.toggle(
        "Claude in Chrome active",
        value=st.session_state.get("chrome_mode", False),
        help="Enable this if you have the Claude Chrome extension running and are logged into LinkedIn",
    )
    st.session_state["chrome_mode"] = chrome_mode

    if chrome_mode:
        st.success("Chrome Mode ON")
        st.markdown("""
        **Setup checklist:**
        - ✅ Claude Chrome extension is active
        - 🔲 LinkedIn is open and you are logged in
        - 🔲 Your LinkedIn profile is visible

        **Open LinkedIn now:**
        """)
        st.link_button("Open LinkedIn", "https://www.linkedin.com/feed/")
        linkedin_ready = st.checkbox("LinkedIn is open and I'm logged in")
        st.session_state["linkedin_ready"] = linkedin_ready
        if linkedin_ready:
            st.success("Ready to auto-enrich!")
    else:
        st.warning(
            "Chrome Mode is OFF. Enable it to unlock automatic "
            "LinkedIn enrichment for Tier 2 prospects."
        )

    st.divider()

    sector = st.selectbox(
        "Sector Focus",
        options=list(SECTORS.keys()),
        help="Industry sector to search for YM associations",
    )

    num_targets = st.slider(
        "Target Prospects",
        min_value=8,
        max_value=20,
        value=12,
        help="How many candidates to research per session",
    )

    st.divider()

    run_button = st.button(
        "Run Prospecting Session",
        type="primary",
        use_container_width=True,
    )

    st.divider()
    st.markdown(
        """
        **Fit Guide**
        - 🟢 Strong Fit — Ready for outreach
        - 🟡 Medium Fit — Needs enrichment
        - 🟠 Unlikely Fit — Watch list
        - 🔴 Not a Fit — Outside ICP
        """
    )

# ---------------------------------------------------------------------------
# Main content
# ---------------------------------------------------------------------------

st.markdown("""
<div style="background-color:#1b5897;color:white;padding:1rem 1.5rem;border-radius:8px;margin-bottom:1rem;">
    <h2 style="color:white;margin:0;font-family:'Open Sans',sans-serif;">YM Prospector</h2>
    <p style="color:rgba(255,255,255,0.85);margin:0.2rem 0 0 0;font-size:0.9rem;">
        YourMembership Association Discovery — Starkweather Association Services
    </p>
</div>
""", unsafe_allow_html=True)

tab_run, tab_results, tab_history, tab_ghl = st.tabs(
    ["Run Session", "Results", "Run History", "Export to GHL"]
)

# ---------------------------------------------------------------------------
# Tab 1: Run Session
# ---------------------------------------------------------------------------

with tab_run:
    if run_button:
        st.session_state.results = None
        st.session_state.output_path = None
        st.session_state.last_sector = sector

        log_lines = []

        with st.status(
            f"Running prospecting session: **{sector}**", expanded=True
        ) as status:

            def update(msg: str):
                st.write(msg)
                log_lines.append(msg)

            results, output_path = run_session(
                sector=sector,
                num_targets=num_targets,
                status_fn=update,
                output_dir=OUTPUT_DIR,
            )

            st.session_state.results = results
            st.session_state.output_path = output_path

            if results:
                t1 = sum(1 for p in results if p.get("tier") == "Strong Fit")
                t2 = sum(1 for p in results if p.get("tier") == "Medium Fit")
                t3 = sum(1 for p in results if p.get("tier") == "Unlikely Fit")
                ex = len(results) - t1 - t2 - t3
                status.update(
                    label=f"Complete — {t1} Strong Fit | {t2} Medium Fit | {t3} Unlikely Fit | {ex} Not a Fit",
                    state="complete",
                )
                st.success("Session complete! View results in the **Results** tab.")
            else:
                status.update(label="Session ended with no results.", state="error")
                st.error("No prospects returned. Check your API keys in the .env file.")

    else:
        st.info(
            "Select a sector in the sidebar and click **Run Prospecting Session** to begin."
        )
        st.markdown(
            """
            **What happens during a session:**
            1. Runs 6–8 Brave web searches using YourMembership URL fingerprints
            2. Claude identifies candidate associations from the results
            3. Each candidate's website is fetched and ProPublica/LinkedIn are searched
            4. Claude scores every candidate against the ICP and assigns Tier 1/2/3
            5. Results are saved to a color-coded Excel file in `output/nightly/`

            **Typical run time:** 3–6 minutes for 12 prospects
            """
        )

# ---------------------------------------------------------------------------
# Tab 2: Results
# ---------------------------------------------------------------------------

with tab_results:
    results = st.session_state.results
    output_path = st.session_state.output_path

    if results:
        # Summary metrics
        t1 = sum(1 for p in results if p.get("tier") == "Strong Fit")
        t2 = sum(1 for p in results if p.get("tier") == "Medium Fit")
        t3 = sum(1 for p in results if p.get("tier") == "Unlikely Fit")
        ex = len(results) - t1 - t2 - t3

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("🟢 Strong Fit", t1)
        col2.metric("🟡 Medium Fit", t2)
        col3.metric("🟠 Unlikely Fit", t3)
        col4.metric("🔴 Not a Fit", ex)

        st.divider()

        # Auto-enrichment for Tier 2 prospects
        tier2_count = sum(1 for p in results if p.get("tier") == "Medium Fit")
        chrome_mode = st.session_state.get("chrome_mode", False)
        linkedin_ready = st.session_state.get("linkedin_ready", False)

        if tier2_count > 0:
            if not chrome_mode:
                st.warning(
                    f"⚠️ **{tier2_count} Medium Fit prospect(s) need enrichment.** "
                    "Enable Chrome + LinkedIn Mode in the sidebar to auto-enrich them."
                )
            elif not linkedin_ready:
                st.warning(
                    f"⚠️ **{tier2_count} Medium Fit prospect(s) ready to enrich.** "
                    "Check 'LinkedIn is open and I'm logged in' in the sidebar to proceed."
                )
            else:
                st.info(
                    f"✅ Chrome + LinkedIn active. **{tier2_count} Medium Fit prospect(s)** "
                    "can be auto-enriched with ProPublica revenue and LinkedIn leadership data."
                )
                if st.button(
                    f"🔍 Auto-Enrich {tier2_count} Medium Fit Prospect(s)",
                    type="primary",
                ):
                    with st.status("Enriching Tier 2 prospects...", expanded=True) as enrich_status:
                        def enrich_update(msg):
                            st.write(msg)

                        updated = enrich_tier2(
                            st.session_state.results,
                            status_fn=enrich_update,
                        )
                        st.session_state.results = updated

                        # Re-save Excel with enriched data
                        from scripts.prospector import save_to_excel
                        new_path = save_to_excel(
                            updated,
                            OUTPUT_DIR,
                            sector=st.session_state.get("last_sector", ""),
                        )
                        st.session_state.output_path = new_path

                        new_t1 = sum(1 for p in updated if p.get("tier") == "Strong Fit")
                        promoted = new_t1 - t1
                        enrich_status.update(
                            label=f"Enrichment complete! "
                                  f"{promoted} prospect(s) promoted to Tier 1.",
                            state="complete",
                        )
                    st.rerun()

        st.divider()

        # Download button
        if output_path and Path(output_path).exists():
            with open(output_path, "rb") as f:
                st.download_button(
                    label="Download Excel File",
                    data=f,
                    file_name=Path(output_path).name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        # Results table
        df = pd.DataFrame(results)

        # Filters
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            fit_filter = st.multiselect(
                "Filter by Fit",
                options=["Strong Fit", "Medium Fit", "Unlikely Fit", "Not a Fit"],
                default=["Strong Fit", "Medium Fit", "Unlikely Fit", "Not a Fit"],
            )
        with col_f2:
            scope_filter = st.multiselect(
                "Filter by Scope",
                options=["Local", "State", "National", "International"],
                default=["Local", "State", "National", "International"],
                help="National/International associations tend to have higher revenue",
            )

        filtered = df.copy()
        if "tier" in df.columns and fit_filter:
            filtered = filtered[filtered["tier"].isin(fit_filter)]
        if "association_scope" in df.columns and scope_filter:
            filtered = filtered[filtered["association_scope"].isin(scope_filter)]

        # Build display columns — always show LinkedIn URL + message; show
        # exclude reason and project notes when "Not a Fit" is in the filter
        showing_not_a_fit = "Not a Fit" in fit_filter
        display_cols = [
            "organization_name", "website_url", "association_scope",
            "ym_confirmed", "staff_size", "annual_revenue",
            "tier",
            "decision_maker_name", "decision_maker_title",
            "decision_maker_linkedin_url",
            "pain_signal_notes",
        ]
        if showing_not_a_fit:
            display_cols += ["exclude_reason", "project_opportunity_notes"]

        available_cols = [c for c in display_cols if c in filtered.columns]

        rename_map = {
            "organization_name": "Organization",
            "website_url": "Website",
            "association_scope": "Scope",
            "ym_confirmed": "YM?",
            "staff_size": "Staff",
            "annual_revenue": "Revenue",
            "tier": "Fit",
            "decision_maker_name": "Contact Name",
            "decision_maker_title": "Contact Title",
            "decision_maker_linkedin_url": "LinkedIn",
            "pain_signal_notes": "Pain Signals",
            "exclude_reason": "Exclude Reason",
            "project_opportunity_notes": "Project Opportunity",
        }

        filtered_df = filtered[available_cols].rename(columns=rename_map)

        # Add a Message column derived from the LinkedIn URL
        if "LinkedIn" in filtered_df.columns:
            filtered_df["Message"] = filtered_df["LinkedIn"].apply(linkedin_message_url)

        st.dataframe(
            filtered_df,
            use_container_width=True,
            height=500,
            column_config={
                "Website": st.column_config.LinkColumn(
                    "Website", display_text="Visit Site"
                ),
                "LinkedIn": st.column_config.LinkColumn(
                    "LinkedIn", display_text="View Profile"
                ),
                "Message": st.column_config.LinkColumn(
                    "Message", display_text="💬 Message"
                ),
            },
        )

    else:
        st.info("No results yet. Run a prospecting session first.")

# ---------------------------------------------------------------------------
# Tab 3: Run History
# ---------------------------------------------------------------------------

with tab_history:
    history = get_run_history()

    if history:
        st.subheader(f"Past Runs ({len(history)} sessions)")

        for file in history:
            col1, col2 = st.columns([4, 1])
            with col1:
                size_kb = file.stat().st_size / 1024
                st.markdown(f"**{file.name}**  \n`{size_kb:.1f} KB`")
            with col2:
                with open(file, "rb") as f:
                    st.download_button(
                        label="Download",
                        data=f,
                        file_name=file.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=str(file),
                    )
            st.divider()
    else:
        st.info(
            "No past runs found. Completed sessions will appear here automatically."
        )

# ---------------------------------------------------------------------------
# Tab 4: Export to GHL
# ---------------------------------------------------------------------------

with tab_ghl:
    st.subheader("Export Contacts to Go High Level")
    st.markdown(
        "Generate a CSV file ready to import directly into GHL. "
        "Each prospect becomes a contact with tags, source, notes, and custom fields pre-populated."
    )

    history = get_run_history()

    if not history:
        st.info("No prospecting runs found. Run a session first.")
    else:
        st.divider()

        # File selector
        file_options = {f.name: f for f in history}
        selected_file_name = st.selectbox(
            "Select Prospecting Run",
            options=list(file_options.keys()),
        )
        selected_file = file_options[selected_file_name]

        # Tier selector
        tiers_to_export = st.multiselect(
            "Include Fits",
            options=["Strong Fit", "Medium Fit", "Unlikely Fit"],
            default=["Strong Fit", "Medium Fit"],
            help="Strong Fit and Medium Fit are recommended for outreach. Unlikely Fit can be added for awareness.",
        )

        st.divider()

        if st.button("Generate GHL Import CSV", type="primary"):
            try:
                csv_string, row_count = excel_to_ghl_csv(
                    excel_path=str(selected_file),
                    tiers=tiers_to_export,
                )

                if row_count == 0:
                    st.warning("No contacts found for the selected tiers. Try including more tiers.")
                else:
                    run_date = selected_file_name.replace("ym_prospects_", "").replace(".xlsx", "")
                    csv_filename = f"ghl_import_{run_date}.csv"

                    st.success(f"Ready to import — {row_count} contacts")

                    st.download_button(
                        label=f"Download GHL CSV ({row_count} contacts)",
                        data=csv_string,
                        file_name=csv_filename,
                        mime="text/csv",
                    )

                    st.markdown("**How to import into GHL:**")
                    st.markdown(
                        """
                        1. Open Go High Level → Contacts
                        2. Click **Import** → **Import Contacts**
                        3. Upload the CSV file downloaded above
                        4. Map columns if prompted (fields should auto-map)
                        5. Review and confirm the import
                        6. Contacts will appear tagged as **YM Prospect** with their tier
                        """
                    )

            except Exception as e:
                st.error(f"Export failed: {e}")
