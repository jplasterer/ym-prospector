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

from scripts.prospector import run_session, SECTORS
from scripts.ghl_export import excel_to_ghl_csv

load_dotenv()

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Starkweather YM Prospector",
    layout="wide",
    initial_sidebar_state="expanded",
)

OUTPUT_DIR = str(Path(__file__).parent / "output" / "nightly")

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
        "Tier 1": "#C6EFCE",
        "Tier 2": "#FFEB9C",
        "Tier 3": "#FFCC99",
        "Exclude": "#FFC7CE",
    }.get(tier, "#FFFFFF")


def style_tier_row(row):
    color = tier_color(row.get("tier", ""))
    return [f"background-color: {color}"] * len(row)


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

with st.sidebar:
    st.markdown("## YM Prospector")
    st.caption("Starkweather Association Services")
    st.caption(f"Logged in as **{st.session_state.get('name', '')}**")
    authenticator.logout("Log Out", "sidebar")
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
        **Tier Guide**
        - 🟢 Tier 1 — Ready for outreach
        - 🟡 Tier 2 — Needs enrichment
        - 🟠 Tier 3 — Watch list
        - 🔴 Exclude — Outside ICP
        """
    )

# ---------------------------------------------------------------------------
# Main content
# ---------------------------------------------------------------------------

st.title("Starkweather YM Prospector")
st.caption("YourMembership Association Discovery — Powered by Brave Search + Claude")

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
                t1 = sum(1 for p in results if p.get("tier") == "Tier 1")
                t2 = sum(1 for p in results if p.get("tier") == "Tier 2")
                t3 = sum(1 for p in results if p.get("tier") == "Tier 3")
                ex = len(results) - t1 - t2 - t3
                status.update(
                    label=f"Complete — {t1} Tier 1 | {t2} Tier 2 | {t3} Tier 3 | {ex} Excluded",
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
        t1 = sum(1 for p in results if p.get("tier") == "Tier 1")
        t2 = sum(1 for p in results if p.get("tier") == "Tier 2")
        t3 = sum(1 for p in results if p.get("tier") == "Tier 3")
        ex = len(results) - t1 - t2 - t3

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Tier 1 – Ready", t1)
        col2.metric("Tier 2 – Enrich", t2)
        col3.metric("Tier 3 – Watch", t3)
        col4.metric("Excluded", ex)

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
        display_cols = [
            "organization_name", "website_url", "association_scope",
            "ym_confirmed", "staff_size", "annual_revenue",
            "tier", "pain_signal_notes",
        ]
        available_cols = [c for c in display_cols if c in df.columns]

        rename_map = {
            "organization_name": "Organization",
            "website_url": "Website",
            "association_scope": "Scope",
            "ym_confirmed": "YM?",
            "staff_size": "Staff",
            "annual_revenue": "Revenue",
            "tier": "Tier",
            "pain_signal_notes": "Pain Signals",
        }

        display_df = df[available_cols].rename(columns=rename_map)

        # Filters
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            tier_filter = st.multiselect(
                "Filter by Tier",
                options=["Tier 1", "Tier 2", "Tier 3", "Exclude"],
                default=["Tier 1", "Tier 2", "Tier 3"],
            )
        with col_f2:
            scope_filter = st.multiselect(
                "Filter by Scope",
                options=["Local", "State", "National", "International"],
                default=["Local", "State", "National", "International"],
                help="National/International associations tend to have higher revenue",
            )

        filtered = df.copy()
        if "tier" in df.columns and tier_filter:
            filtered = filtered[filtered["tier"].isin(tier_filter)]
        if "association_scope" in df.columns and scope_filter:
            filtered = filtered[filtered["association_scope"].isin(scope_filter)]
        filtered_df = filtered[available_cols].rename(columns=rename_map)

        st.dataframe(
            filtered_df,
            use_container_width=True,
            height=500,
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
            "Include Tiers",
            options=["Tier 1", "Tier 2", "Tier 3"],
            default=["Tier 1", "Tier 2"],
            help="Tier 1 and 2 are recommended for outreach. Tier 3 can be added for awareness.",
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
