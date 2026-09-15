"""
AGEVAL Harmonization - Step 1 - Dictionary Merge
Standalone run: python -m streamlit run harmonization/01_code/s01_dictionary_merge.py
Also imported as a page by the unified harmonization/01_code/app.py

Combines the master dictionary with the personalized baseline and endline
dictionaries (each produced by the individual baseline/endline pipeline's
Step 1 - Dictionary Selector) into a single harmonization dictionary:

    * The first 8 columns come from the master dictionary:
      topic, subtopic, variable_name, var_type, label_spanish, label_english,
      surv_type, surv_choices. `surv_type`/`surv_choices` are carried through
      unedited — they identify select_one/select_multiple variables and their
      response options for later steps (e.g. comparison plots), since
      `var_type` alone isn't always reliable for that (a select_multiple can be
      tagged `var_type=numerical` in the master dictionary).
    * `baseline` / `endline` are auto-detected: 1 if the variable is present
      in that round's personalized dictionary (i.e. it was included in that
      round's survey), 0 otherwise. These can be unchecked/overridden here.
    * `plots`, `outcome_var`, `control_var`, `logarithm_var`, `fe_var`,
      `cluster_var` are pre-loaded from the master dictionary's own values
      and can be edited here.
    * Only variables present in at least one round (baseline == 1 or
      endline == 1) are kept.
"""

import streamlit as st
import pandas as pd
import os
import yaml

# Read config
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "s00_harmonizationConfig.yaml")
with open(CONFIG_PATH, "r") as f:
    config = yaml.safe_load(f)

# Paths
BASE_PATH = os.path.normpath(os.path.join(os.path.dirname(CONFIG_PATH), config["paths"]["base"]))
MASTER_PATH = os.path.join(BASE_PATH, config["paths"]["dictionary_master"])
DICT_BASELINE_PATH = os.path.join(BASE_PATH, config["paths"]["dictionary_baseline"])
DICT_ENDLINE_PATH = os.path.join(BASE_PATH, config["paths"]["dictionary_endline"])
DICT_HARMONIZATION_PATH = os.path.join(BASE_PATH, config["paths"]["dictionary_harmonization"])

# First columns taken directly from the master dictionary, unedited
BASE_COLS = ["topic", "subtopic", "variable_name", "var_type", "label_spanish", "label_english",
             "surv_type", "surv_choices"]

# Auto-detected round-membership columns
ROUND_COLS = ["baseline", "endline"]

# The harmonization-analysis flags a variable can be toggled into.
# Pre-loaded from the master dictionary's own values; editable here.
# (outcome_var/control_var/logarithm_var/fe_var/cluster_var will be used by a
# future regression-modeling step on the combined baseline+endline data.)
HARMONIZATION_FLAG_COLS = {
    "plots": "Plots",
    "outcome_var": "Outcome Variable",
    "control_var": "Control Variable",
    "logarithm_var": "Logarithm",
    "fe_var": "Fixed Effects Variable",
    "cluster_var": "Cluster Variable",
}


@st.cache_data
def load_master():
    return pd.read_excel(MASTER_PATH, sheet_name="variables_master")


@st.cache_data
def load_sections():
    """Load the 'sections' sheet, which defines the canonical display order
    and labels for topics and subtopics."""
    sections = pd.read_excel(MASTER_PATH, sheet_name="sections")

    topics = (
        sections[sections["level"] == "topic"]
        .sort_values("order")
        .set_index("key")
        .to_dict(orient="index")
    )
    subtopics = (
        sections[sections["level"] == "subtopic"]
        .sort_values("order")
        .set_index("key")
        .to_dict(orient="index")
    )
    return topics, subtopics


def _ordered_keys(present_keys, section_info):
    """Return present_keys ordered per section_info's 'order' field.
    Any keys not found in section_info are appended alphabetically at the end."""
    present_keys = list(dict.fromkeys(present_keys))  # de-dup, keep first-seen order as fallback
    known = [k for k in present_keys if k in section_info]
    unknown = sorted(k for k in present_keys if k not in section_info)
    known.sort(key=lambda k: section_info[k]["order"])
    return known + unknown


def _section_label(key, section_info):
    """Nice display label for a topic/subtopic key, falling back to the raw key."""
    info = section_info.get(key)
    if info is None:
        return key
    label = info.get("label_spanish") or info.get("label_english") or key
    # strip markdown heading markers sometimes present in the source labels (e.g. "### Parcelas")
    return str(label).lstrip("#").strip()


def load_round_variables(path):
    """Return the set of variable_name values present in a round's
    personalized dictionary CSV (i.e. the variables included in that
    round's survey). Returns an empty set if the file doesn't exist yet."""
    if not path or not os.path.exists(path):
        return set()
    round_df = pd.read_csv(path)
    if "variable_name" not in round_df.columns:
        return set()
    return set(round_df["variable_name"].dropna().astype(str))


def build_harmonization_dict(master_df, baseline_vars, endline_vars):
    """Build the harmonization dictionary: base columns from master,
    auto-detected baseline/endline membership, and the analysis flags
    pre-loaded from master. Keeps only rows present in at least one round."""
    merged = master_df[BASE_COLS].copy()
    varnames = master_df["variable_name"].astype(str)

    merged["baseline"] = varnames.isin(baseline_vars).astype(int)
    merged["endline"] = varnames.isin(endline_vars).astype(int)

    for col in HARMONIZATION_FLAG_COLS:
        if col in master_df.columns:
            merged[col] = master_df[col].fillna(0).astype(bool).astype(int)
        else:
            merged[col] = 0

    merged = merged[(merged["baseline"] == 1) | (merged["endline"] == 1)]
    return merged.reset_index(drop=True)


def render(standalone: bool = False):
    """Render the dictionary merge step. Set standalone=True to also set
    page config and title (only needed when this file is run directly,
    not from the unified harmonization app.py)."""

    if standalone:
        st.set_page_config(page_title="AGEVAL Harmonization - Dictionary Merge", layout="wide")

    st.title("AGEVAL Harmonization - Step 1 — Dictionary Merge")
    st.caption(
        "Combines the master dictionary with the personalized baseline and endline dictionaries. "
        "`baseline`/`endline` are auto-detected from each round's personalized dictionary; "
        "`plots`, `outcome_var`, `control_var`, `logarithm_var`, `fe_var` and `cluster_var` are "
        "pre-loaded from the master dictionary. "
        "Uncheck anything you want to exclude, then save the combined harmonization dictionary."
    )

    if not os.path.exists(MASTER_PATH):
        st.error(f"Master dictionary not found at: `{os.path.abspath(MASTER_PATH)}`")
        st.info("Place your `variables_master.xlsx` file in the `dictionary/` folder.")
        return

    baseline_missing = not os.path.exists(DICT_BASELINE_PATH)
    endline_missing = not os.path.exists(DICT_ENDLINE_PATH)
    if baseline_missing:
        st.warning(
            f"Personalized baseline dictionary not found at `{os.path.abspath(DICT_BASELINE_PATH)}`. "
            "Baseline will be treated as not applied to any variable until it is generated."
        )
    if endline_missing:
        st.warning(
            f"Personalized endline dictionary not found at `{os.path.abspath(DICT_ENDLINE_PATH)}`. "
            "Endline will be treated as not applied to any variable until it is generated."
        )

    master_df = load_master()
    topic_info, subtopic_info = load_sections()

    baseline_vars = load_round_variables(DICT_BASELINE_PATH)
    endline_vars = load_round_variables(DICT_ENDLINE_PATH)

    df = build_harmonization_dict(master_df, baseline_vars, endline_vars)

    if df.empty:
        st.warning(
            "No variables found in either round's personalized dictionary yet. "
            "Complete Step 1 (Dictionary Selector) for baseline and/or endline first."
        )
        return

    TOPICS = _ordered_keys(df["topic"].dropna().unique().tolist(), topic_info)

    # ── Summary metrics ─────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total variables", len(df))
    c2.metric("In baseline", int(df["baseline"].sum()))
    c3.metric("In endline", int(df["endline"].sum()))
    c4.metric("In both rounds", int(((df["baseline"] == 1) & (df["endline"] == 1)).sum()))

    # ── Sidebar filters ─────────────────────────────────────────────────────
    st.sidebar.header("Filter variables")
    selected_topics = st.sidebar.multiselect(
        "Topic", TOPICS, default=TOPICS,
        format_func=lambda k: _section_label(k, topic_info),
    )
    # st.multiselect preserves click order, not option order — re-sort so the
    # rest of the page always follows the sections-sheet order regardless.
    selected_topics = _ordered_keys(selected_topics, topic_info)

    round_filter = st.sidebar.multiselect(
        "Present in round",
        ["Baseline", "Endline", "Both"],
        default=["Baseline", "Endline", "Both"],
    )
    search = st.sidebar.text_input("Search variable name or label")

    filtered = df[df["topic"].isin(selected_topics)].copy()

    round_mask = pd.Series(False, index=filtered.index)
    if "Baseline" in round_filter:
        round_mask |= (filtered["baseline"] == 1) & (filtered["endline"] == 0)
    if "Endline" in round_filter:
        round_mask |= (filtered["endline"] == 1) & (filtered["baseline"] == 0)
    if "Both" in round_filter:
        round_mask |= (filtered["baseline"] == 1) & (filtered["endline"] == 1)
    filtered = filtered[round_mask]

    if search:
        mask = (
            filtered["variable_name"].str.contains(search, case=False, na=False) |
            filtered["label_spanish"].str.contains(search, case=False, na=False) |
            filtered["label_english"].str.contains(search, case=False, na=False)
        )
        filtered = filtered[mask]

    st.sidebar.markdown("---")
    st.sidebar.markdown(f"**{len(filtered)}** variables shown")

    flag_keys = list(HARMONIZATION_FLAG_COLS.keys())
    editable_keys = ROUND_COLS + flag_keys

    st.markdown("---")

    # ── Editable table per topic ────────────────────────────────────────────
    st.subheader("Edit harmonization flags")
    st.info(
        "`baseline`/`endline` were auto-detected from each round's personalized dictionary — "
        "uncheck to exclude a round for that variable. Check `plots`, `outcome_var`, `control_var`, "
        "`logarithm_var`, `fe_var` and `cluster_var` to include the variable in that part of the "
        "harmonized analysis."
    )

    edited_frames = []
    display_cols = ["variable_name", "label_spanish", "var_type"] + editable_keys
    column_config = {
        "variable_name":   st.column_config.TextColumn("Variable", disabled=True, width="medium"),
        "label_spanish":   st.column_config.TextColumn("Label (ES)", disabled=True, width="large"),
        "var_type":        st.column_config.TextColumn("Tipo de variable", disabled=True, width="small"),
        "baseline":        st.column_config.CheckboxColumn("Baseline", default=False),
        "endline":         st.column_config.CheckboxColumn("Endline", default=False),
        "plots":           st.column_config.CheckboxColumn("Plots", default=False),
        "outcome_var":     st.column_config.CheckboxColumn("Outcome Var", default=False),
        "control_var":     st.column_config.CheckboxColumn("Control Var", default=False),
        "logarithm_var":   st.column_config.CheckboxColumn("Logarithm", default=False),
        "fe_var":          st.column_config.CheckboxColumn("FE Var", default=False),
        "cluster_var":     st.column_config.CheckboxColumn("Cluster Var", default=False),
    }

    for topic in selected_topics:
        topic_df = filtered[filtered["topic"] == topic].copy()
        if topic_df.empty:
            continue

        topic_label = _section_label(topic, topic_info)

        with st.expander(f"**{topic_label.upper()}** — {len(topic_df)} variables", expanded=True):
            # Split the topic into its subtopics, ordered per the sections sheet.
            present_subtopics = _ordered_keys(
                topic_df["subtopic"].dropna().unique().tolist(), subtopic_info
            )

            if len(present_subtopics) <= 1:
                # Single (or no) subtopic — no need for a nested header.
                subtopic_groups = [(None, topic_df)]
            else:
                subtopic_groups = [
                    (sub, topic_df[topic_df["subtopic"] == sub].copy())
                    for sub in present_subtopics
                ]

            for sub, sub_df in subtopic_groups:
                if sub_df.empty:
                    continue
                if sub is not None:
                    sub_label = _section_label(sub, subtopic_info)
                    st.markdown(f"##### {sub_label} — {len(sub_df)} variables")

                editor_input = sub_df[display_cols].reset_index(drop=True)
                # Coerce the flag columns to real booleans so they always
                # render as checkboxes regardless of how they're stored.
                editor_input[editable_keys] = editor_input[editable_keys].fillna(0).astype(bool)

                editable = st.data_editor(
                    editor_input,
                    column_config=column_config,
                    width="stretch",
                    key=f"harmonize_editor_{topic}_{sub or 'all'}",
                    hide_index=True,
                )
                sub_df = sub_df.reset_index(drop=True)
                sub_df[display_cols] = editable
                edited_frames.append(sub_df)

    # ── Save / download ──────────────────────────────────────────────────────
    st.markdown("---")

    if edited_frames:
        harmonized = pd.concat(edited_frames)
        # Store flags back as 0/1 ints for downstream harmonization steps.
        harmonized[editable_keys] = harmonized[editable_keys].astype(bool).astype(int)
        # Re-apply the "at least one round" rule in case both were unchecked.
        download_df = harmonized[(harmonized["baseline"] == 1) | (harmonized["endline"] == 1)].copy()
        excluded = len(harmonized) - len(download_df)

        c1, c2 = st.columns(2)
        with c1:
            if st.button("💾 Save combined dictionary (harmonization/02_dictionary/variables_personalized.csv)", type="primary"):
                os.makedirs(os.path.dirname(DICT_HARMONIZATION_PATH), exist_ok=True)
                download_df.to_csv(DICT_HARMONIZATION_PATH, index=False, encoding="utf-8-sig")
                st.session_state["harmonization_dict_saved"] = True
                st.success(
                    f"Saved {len(download_df)} variables to `{os.path.abspath(DICT_HARMONIZATION_PATH)}`. "
                    + (f"({excluded} excluded — unchecked in both rounds.) " if excluded else "")
                    + "Later harmonization steps will pick this up automatically."
                )
        with c2:
            csv = download_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                label="⬇️ Download a copy (CSV)",
                data=csv,
                file_name="variables_harmonization.csv",
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
            master_row = master_df[master_df["variable_name"] == selected_var].iloc[0]
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(f"**Label (EN):** {master_row['label_english']}")
                st.markdown(f"**Label (ES):** {master_row['label_spanish']}")
                st.markdown(f"**Tipo de variable:** `{master_row['var_type']}`")
                st.markdown(f"**Type:** `{master_row.get('surv_type', 'n/a')}`")
                st.markdown(f"**Topic:** {master_row['topic']}")
            with c2:
                st.markdown(f"**In baseline:** {'✅' if row['baseline'] == 1 else '❌'}")
                st.markdown(f"**In endline:** {'✅' if row['endline'] == 1 else '❌'}")
                st.markdown(f"**Quality range:** {master_row.get('quality_min', '—')} – {master_row.get('quality_max', '—')}")
                st.markdown(f"**Outlier SD threshold:** {master_row.get('quality_outlier_sd', '—')}")


if __name__ == "__main__":
    render(standalone=True)