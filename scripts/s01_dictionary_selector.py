"""
AGEVAL Step 1 - Variable Dictionary Selector
Standalone run: python -m streamlit run scripts/s01_dictionary_selector.py
Also imported as a page by the unified scripts/app.py
"""

import streamlit as st
import pandas as pd
import os

DICT_PATH = os.path.join(os.path.dirname(__file__), "../dictionary/variables_master.xlsx")
PERSONALIZED_PATH = os.path.join(os.path.dirname(__file__), "../dictionary/variables_personalized.csv")

PIPELINE_COLS = {
    "questionnaire_include":    "📋 Add to Questionnaire",
    "surv_calculation_include": "🔢 ODK Calculate",
    "quality_include":          "✅ Quality Check",
    "descriptive_include":      "📊 Descriptive Stats",
    "model_role":               "📐 Model Role",
}


@st.cache_data
def load_dict():
    return pd.read_excel(DICT_PATH, sheet_name="variables_master")


def render(standalone: bool = False):
    """Render the dictionary selector. Set standalone=True to also set page config
    and title (only needed when this file is run directly, not from the unified app)."""

    if standalone:
        st.set_page_config(page_title="AGEVAL - Variable Selector", layout="wide")

    st.title("AGEVAL - Variable Dictionary Selector")
    st.caption("Select which variables to include at each pipeline stage. Save your filtered dictionary to proceed.")

    if not os.path.exists(DICT_PATH):
        st.error(f"Master dictionary not found at: `{os.path.abspath(DICT_PATH)}`")
        st.info("Place your `variables_master.xlsx` file in the `dictionary/` folder.")
        return

    df = load_dict()
    TOPICS = sorted(df["topic"].dropna().unique().tolist())

    # ── Sidebar filters ─────────────────────────────────────────────────────
    st.sidebar.header("Filter variables")
    selected_topics = st.sidebar.multiselect("Topic", TOPICS, default=TOPICS)
    search = st.sidebar.text_input("Search variable name or label")

    filtered = df[df["topic"].isin(selected_topics)].copy()
    if search:
        mask = (
            filtered["variable_name"].str.contains(search, case=False, na=False) |
            filtered["label_english"].str.contains(search, case=False, na=False) |
            filtered["label_spanish"].str.contains(search, case=False, na=False)
        )
        filtered = filtered[mask]

    st.sidebar.markdown("---")
    st.sidebar.markdown(f"**{len(filtered)}** variables shown")

    # ── Pipeline stage summary ──────────────────────────────────────────────
    cols = st.columns(5)
    stage_keys = list(PIPELINE_COLS.keys())

    for i, (key, label) in enumerate(PIPELINE_COLS.items()):
        if key == "model_role":
            count = int((df[key] > 0).sum())
        else:
            count = int(df[key].sum())
        cols[i].metric(label, count)

    st.markdown("---")

    # ── Editable table per topic ────────────────────────────────────────────
    st.subheader("Edit pipeline flags")
    st.info("Add to Questionnaire: 1=include, 0=exclude. For Model Role: 0=excluded, 1=dependent, 2=independent, 3=control.")

    edited_frames = []

    for topic in selected_topics:
        topic_df = filtered[filtered["topic"] == topic].copy()
        if topic_df.empty:
            continue

        with st.expander(f"**{topic.upper()}** — {len(topic_df)} variables", expanded=True):
            display_cols = ["variable_name", "label_spanish", "surv_type"] + stage_keys
            editable = st.data_editor(
                topic_df[display_cols].reset_index(drop=True),
                column_config={
                    "variable_name":          st.column_config.TextColumn("Variable", disabled=True, width="medium"),
                    "label_spanish":          st.column_config.TextColumn("Label (ES)", disabled=True, width="large"),
                    "surv_type":              st.column_config.TextColumn("Type", disabled=True, width="small"),
                    "questionnaire_include":  st.column_config.CheckboxColumn("Enabled", default=False),
                    "surv_calculate_include": st.column_config.NumberColumn("ODK Calc", min_value=0, max_value=1, step=1),
                    "quality_include":        st.column_config.NumberColumn("Quality", min_value=0, max_value=1, step=1),
                    "descriptive_include":    st.column_config.NumberColumn("Desc. Stats", min_value=0, max_value=1, step=1),
                    "model_role":             st.column_config.NumberColumn("Model Role (0-3)", min_value=0, max_value=3, step=1),
                },
                width="stretch",
                key=f"editor_{topic}",
                hide_index=True,
            )
            topic_df = topic_df.reset_index(drop=True)
            topic_df[display_cols] = editable
            edited_frames.append(topic_df)

    # ── Save / download ──────────────────────────────────────────────────────
    st.markdown("---")

    if edited_frames:
        personalized = pd.concat(edited_frames)
        download_df = personalized[personalized["questionnaire_include"] == 1].copy()

        c1, c2 = st.columns(2)
        with c1:
            if st.button("💾 Save to project (dictionary/variables_personalized.csv)", type="primary"):
                os.makedirs(os.path.dirname(PERSONALIZED_PATH), exist_ok=True)
                download_df.to_csv(PERSONALIZED_PATH, index=False, encoding="utf-8-sig")
                st.session_state["dict_saved"] = True
                st.success(f"Saved {len(download_df)} variables to `{os.path.abspath(PERSONALIZED_PATH)}`. "
                           "Steps 2-4 will now pick this up automatically.")
        with c2:
            csv = download_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                label="⬇️ Download a copy (CSV)",
                data=csv,
                file_name="variables_personalized.csv",
                mime="text/csv",
            )
    else:
        st.warning("No variables selected — adjust filters first.")

    # ── Variable detail inspector ───────────────────────────────────────────
    st.markdown("---")
    st.subheader("Variable inspector")
    if len(filtered) > 0:
        selected_var = st.selectbox("Select a variable to inspect", filtered["variable_name"].tolist())
        if selected_var:
            row = df[df["variable_name"] == selected_var].iloc[0]
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(f"**Label (EN):** {row['label_english']}")
                st.markdown(f"**Label (ES):** {row['label_spanish']}")
                st.markdown(f"**Type:** `{row['surv_type']}`")
                st.markdown(f"**Topic:** {row['topic']}")
                if pd.notna(row.get("surv_choices")) and row["surv_choices"]:
                    st.markdown(f"**Choices:** {row['surv_choices']}")
                if pd.notna(row.get("surv_constraint")) and row["surv_constraint"]:
                    st.markdown(f"**Constraint:** `{row['surv_constraint']}`")
            with c2:
                st.markdown(f"**ODK Calculate:** `{row.get('surv_calculation', 'n/a')}`")
                st.markdown(f"**Output variable:** `{row.get('surv_calculation_output', 'n/a')}`")
                st.markdown(f"**Quality range:** {row.get('quality_min', '—')} – {row.get('quality_max', '—')}")
                st.markdown(f"**Outlier SD threshold:** {row.get('quality_outlier_sd', '—')}")
                if pd.notna(row.get("model_notes")) and row["model_notes"]:
                    st.markdown(f"**Model notes:** {row['model_notes']}")


if __name__ == "__main__":
    render(standalone=True)
