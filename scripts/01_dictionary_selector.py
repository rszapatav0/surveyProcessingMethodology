"""
AGEVAL Step 1 — Variable Dictionary Selector
Run: python -m streamlit run scripts/01_dictionary_selector.py
"""

import streamlit as st
import pandas as pd
import os

DICT_PATH = os.path.join(os.path.dirname(__file__), "../dictionary/variables_master.csv")

st.set_page_config(page_title="AGEVAL — Variable Selector", layout="wide")

st.title("AGEVAL — Variable Dictionary Selector")
st.caption("Select which variables to include at each pipeline stage. Save your filtered dictionary to proceed.")

# ── Load ──────────────────────────────────────────────────────────────────────
@st.cache_data
def load_dict():
    return pd.read_csv(DICT_PATH)

df = load_dict()

PIPELINE_COLS = {
    "questionnaire_include": "📋 Questionnaire",
    "odk_calculate_include": "🔢 ODK Calculate",
    "quality_include":       "✅ Quality Check",
    "descriptive_include":   "📊 Descriptive Stats",
    "model_role":            "📐 Model Role",
}

TOPICS = sorted(df["topic"].dropna().unique().tolist())

# ── Sidebar filters ───────────────────────────────────────────────────────────
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

# ── Pipeline stage summary ────────────────────────────────────────────────────
cols = st.columns(5)
stage_labels = list(PIPELINE_COLS.values())
stage_keys   = list(PIPELINE_COLS.keys())

for i, (key, label) in enumerate(PIPELINE_COLS.items()):
    if key == "model_role":
        count = int((df[key] > 0).sum())
    else:
        count = int(df[key].sum())
    cols[i].metric(label, count)

st.markdown("---")

# ── Editable table per topic ──────────────────────────────────────────────────
st.subheader("Edit pipeline flags")
st.info("Set 1 = include, 0 = exclude. For Model Role: 0=excluded, 1=dependent, 2=independent, 3=control.")

edited_frames = []

for topic in selected_topics:
    topic_df = filtered[filtered["topic"] == topic].copy()
    if topic_df.empty:
        continue

    with st.expander(f"**{topic.upper()}** — {len(topic_df)} variables", expanded=True):
        display_cols = ["variable_name", "label_spanish", "question_type"] + stage_keys
        editable = st.data_editor(
            topic_df[display_cols].reset_index(drop=True),
            column_config={
                "variable_name":          st.column_config.TextColumn("Variable", disabled=True, width="medium"),
                "label_spanish":          st.column_config.TextColumn("Label (ES)", disabled=True, width="large"),
                "question_type":          st.column_config.TextColumn("Type", disabled=True, width="small"),
                "questionnaire_include":  st.column_config.NumberColumn("Questionnaire", min_value=0, max_value=1, step=1),
                "odk_calculate_include":  st.column_config.NumberColumn("ODK Calc", min_value=0, max_value=1, step=1),
                "quality_include":        st.column_config.NumberColumn("Quality", min_value=0, max_value=1, step=1),
                "descriptive_include":    st.column_config.NumberColumn("Desc. Stats", min_value=0, max_value=1, step=1),
                "model_role":             st.column_config.NumberColumn("Model Role (0-3)", min_value=0, max_value=3, step=1),
            },
            width="stretch",
            key=f"editor_{topic}",
            hide_index=True,
        )
        # Merge edits back into topic_df
        topic_df[display_cols] = editable
        edited_frames.append(topic_df)

# ── Merge and save ────────────────────────────────────────────────────────────
st.markdown("---")

# Save button with customized download path
if edited_frames:
    personalized = pd.concat(edited_frames)
    csv = personalized.to_csv(index=False)
    st.download_button(
        label="💾 Download dictionary (CSV)",
        data=csv,
        file_name="variables_personalized.csv",
        mime="text/csv",
        type="primary",
    )
else:
    st.warning("No variables selected — adjust filters first.")
    

# ── Variable detail inspector ─────────────────────────────────────────────────
st.markdown("---")
st.subheader("Variable inspector")
selected_var = st.selectbox("Select a variable to inspect", filtered["variable_name"].tolist())
if selected_var:
    row = df[df["variable_name"] == selected_var].iloc[0]
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**Label (EN):** {row['label_english']}")
        st.markdown(f"**Label (ES):** {row['label_spanish']}")
        st.markdown(f"**Type:** `{row['question_type']}`")
        st.markdown(f"**Topic:** {row['topic']}")
        if pd.notna(row.get("choices")) and row["choices"]:
            st.markdown(f"**Choices:** {row['choices']}")
        if pd.notna(row.get("constraint")) and row["constraint"]:
            st.markdown(f"**Constraint:** `{row['constraint']}`")
    with c2:
        st.markdown(f"**ODK Calculate:** `{row.get('odk_calculate_expr', 'n/a')}`")
        st.markdown(f"**Output variable:** `{row.get('odk_calculate_output', 'n/a')}`")
        st.markdown(f"**Quality range:** {row.get('quality_min', '—')} – {row.get('quality_max', '—')}")
        st.markdown(f"**Outlier SD threshold:** {row.get('quality_outlier_sd', '—')}")
        if pd.notna(row.get("model_notes")) and row["model_notes"]:
            st.markdown(f"**Model notes:** {row['model_notes']}")
