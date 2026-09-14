"""
AGEVAL Harmonization - Step 3 - Treatment Assignment
Standalone run: python -m streamlit run harmonization/01_code/s03_treatment_assignment.py
Also imported as a page by the unified harmonization/01_code/app.py

Uploads (or picks from the project folders) the long baseline+endline
database (output of Step 2) and a treatment-assignment Excel file, and:

    * Shows which columns the two files share, so you can pick the
      merge key(s) — a single respondent ID for individual-level
      assignment, or one or more columns (e.g. a community/group ID)
      for group-level assignment where many long-database rows share
      one treatment record.
    * Lets you pick which treatment column is the numeric code (for
      regressions) and which is the label, from the treatment file.
    * Left-joins treatment onto the long database on the chosen key(s),
      so every existing observation is kept unchanged (no rows added,
      dropped, or duplicated) and the treatment columns land on both
      baseline and endline rows for the same respondent/group.
    * Saves the resulting database.

No merge/quality checks beyond validating the treatment file's merge key
is unique happen here — this step only assigns treatment.
"""

import streamlit as st
import pandas as pd
import os
import glob
import yaml

# Read config
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "s00_harmonizationConfig.yaml")
with open(CONFIG_PATH, "r") as f:
    config = yaml.safe_load(f)

harmonize_cfg = config.get("harmonize", {})

# Paths
BASE_PATH = os.path.normpath(os.path.join(os.path.dirname(CONFIG_PATH), config["paths"]["base"]))
DATA_HARMONIZATION_DIR = os.path.join(BASE_PATH, config["paths"]["data_harmonization"])
TREATMENT_DEFAULT_PATH = os.path.join(BASE_PATH, config["paths"]["treatment_assignment"])
TREATMENT_DIR = os.path.dirname(TREATMENT_DEFAULT_PATH)

# Harmonization settings (all overridable in s00_harmonizationConfig.yaml)
ID_VAR = harmonize_cfg.get("id_variable", "respondent_id")
PERIOD_VAR = harmonize_cfg.get("period_variable", "period")
BASELINE_VALUE = harmonize_cfg.get("baseline_period_value", 0)
ENDLINE_VALUE = harmonize_cfg.get("endline_period_value", 1)
LONG_DATA_FILENAME = harmonize_cfg.get("long_data_filename", "data_long.csv")
DEFAULT_TREATMENT_VAR = harmonize_cfg.get("treatment_var", "treatment")
DEFAULT_TREATMENT_LABEL = harmonize_cfg.get("treatment_label", "treatment_label")
DEFAULT_OUT_FILENAME = harmonize_cfg.get("long_treatment_filename", "data_long_treatment.csv")


# ── I/O helpers ──────────────────────────────────────────────────────────────
def list_data_files(folder, patterns=("*.csv", "*.xlsx", "*.xls")):
    if not folder or not os.path.isdir(folder):
        return []
    files = []
    for pattern in patterns:
        files.extend(glob.glob(os.path.join(folder, pattern)))
    return sorted(files)


def read_table(source, sheet_name=0):
    """Read a CSV/XLSX table from either a filesystem path (str) or a
    Streamlit UploadedFile-like object. Returns None if source is None."""
    if source is None:
        return None
    name = getattr(source, "name", source)
    ext = os.path.splitext(str(name))[1].lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(source, sheet_name=sheet_name)
    else:
        df = pd.read_csv(source)
    df.columns = df.columns.str.strip()  # normalize whitespace in headers
    return df


def pick_or_upload(folder, label, key_prefix, patterns=("*.csv", "*.xlsx", "*.xls"),
                    upload_types=("csv", "xlsx", "xls"), default_path=None):
    """Let the user either pick a file already sitting in `folder` (the
    project's configured path) or upload one from their computer. An
    upload always takes precedence over the dropdown pick. If neither is
    chosen and `default_path` exists on disk, fall back to it silently.
    Returns (source, display_name) or (None, None)."""
    st.markdown(f"**{label}**")
    files = list_data_files(folder, patterns)
    folder_display = os.path.relpath(folder, BASE_PATH) if folder else "n/a"
    options = ["— choose a file from the project —"] + [os.path.relpath(f, BASE_PATH) for f in files]
    choice = st.selectbox(f"From `{folder_display}/`", options, key=f"{key_prefix}_select")

    uploaded = st.file_uploader(
        "…or upload a file", type=list(upload_types), key=f"{key_prefix}_upload"
    )

    if uploaded is not None:
        return uploaded, uploaded.name
    if choice != "— choose a file from the project —":
        p = os.path.join(BASE_PATH, choice)
        return p, os.path.basename(p)
    if default_path and os.path.exists(default_path):
        return default_path, os.path.basename(default_path)
    return None, None


# ── Core logic ───────────────────────────────────────────────────────────────
def find_duplicate_keys(df, merge_keys):
    """Return the rows of df whose merge_keys combination is duplicated."""
    dup_mask = df.duplicated(subset=merge_keys, keep=False)
    return df.loc[dup_mask].sort_values(merge_keys)


def assign_treatment(long_df, treatment_df, merge_keys, treatment_var, treatment_label,
                      extra_cols=None):
    """Left-join treatment onto the long database on merge_keys. Validates
    the treatment side is unique on merge_keys (m:1) so no existing
    observation is duplicated or dropped. Returns (merged_df, diagnostics dict)
    or raises ValueError with a human-readable message on a validation failure."""
    extra_cols = extra_cols or []
    keep_cols = list(dict.fromkeys(list(merge_keys) + [treatment_var, treatment_label] + list(extra_cols)))
    treat_subset = treatment_df[keep_cols].copy()

    n_before = len(long_df)
    try:
        merged = long_df.merge(
            treat_subset, on=list(merge_keys), how="left",
            validate="m:1", indicator=True,
        )
    except pd.errors.MergeError as e:
        raise ValueError(
            "The treatment file has more than one row for at least one merge-key "
            f"combination, so the merge can't guarantee a clean many-to-one match: {e}"
        )

    if len(merged) != n_before:
        raise ValueError(
            f"Merge changed the number of rows ({n_before} -> {len(merged)}). "
            "This shouldn't happen with a validated many-to-one merge — check your merge keys."
        )

    matched = int((merged["_merge"] == "both").sum())
    unmatched_df = merged.loc[merged["_merge"] == "left_only", list(merge_keys)].drop_duplicates()
    merged = merged.drop(columns=["_merge"])

    diagnostics = {
        "n_rows": len(merged),
        "n_matched": matched,
        "n_unmatched": n_before - matched,
        "unmatched_keys": unmatched_df,
    }
    return merged, diagnostics


# ── Streamlit page ───────────────────────────────────────────────────────────
def render(standalone: bool = False):
    """Render the treatment-assignment step. Set standalone=True to also
    set page config and title (only needed when this file is run
    directly, not from the unified harmonization app.py)."""

    if standalone:
        st.set_page_config(page_title="AGEVAL Harmonization - Treatment Assignment", layout="wide")

    st.title("AGEVAL Harmonization - Step 3 — Assign Treatment")
    st.caption(
        "Merges a treatment-assignment file into the long baseline+endline database. "
        "Existing observations are left unchanged — this is a left join keyed on the "
        "column(s) you pick, validated so no row is duplicated or dropped. "
        "No quality checks happen here."
    )

    # ── Long database input ───────────────────────────────────────────────────
    st.subheader("1. Long database (from Step 2)")
    long_source, long_name = pick_or_upload(
        DATA_HARMONIZATION_DIR, "Long baseline + endline database", "long_db",
        default_path=os.path.join(DATA_HARMONIZATION_DIR, LONG_DATA_FILENAME),
    )
    if long_source is None:
        st.info("Select or upload the long database to continue.")
        return
    long_df = read_table(long_source)
    st.success(f"Loaded `{long_name}` — {long_df.shape[0]} rows, {long_df.shape[1]} columns.")

    st.markdown("---")

    # ── Treatment file input ──────────────────────────────────────────────────
    st.subheader("2. Treatment assignment file")
    treat_source, treat_name = pick_or_upload(
        TREATMENT_DIR, "Treatment assignment (Excel)", "treatment",
        patterns=("*.xlsx", "*.xls", "*.csv"), upload_types=("xlsx", "xls", "csv"),
        default_path=TREATMENT_DEFAULT_PATH,
    )
    if treat_source is None:
        st.info("Select or upload the treatment assignment file to continue.")
        return
    treatment_df = read_table(treat_source)
    st.success(f"Loaded `{treat_name}` — {treatment_df.shape[0]} rows, {treatment_df.shape[1]} columns.")

    with st.expander("Preview treatment file", expanded=False):
        st.dataframe(treatment_df.head(20), width="stretch")

    st.markdown("---")

    # ── Common columns / merge keys ───────────────────────────────────────────
    st.subheader("3. Merge key(s)")
    common_cols = sorted(set(long_df.columns) & set(treatment_df.columns))
    if not common_cols:
        st.error(
            "The long database and the treatment file don't share any column names, "
            "so there's no key to merge on. Rename the matching ID column(s) so they match "
            "exactly (e.g. both called `respondent_id`) and re-upload."
        )
        return

    st.caption(f"Columns found in both files: {', '.join(common_cols)}")

    level = st.radio(
        "Treatment assignment level",
        [
            "Respondent / individual level — one treatment row per respondent",
            "Community / group level — one treatment row shared by many respondents",
        ],
        help="This only changes the suggested default key below — the merge itself "
             "works the same way for both levels.",
    )
    suggested_default = [ID_VAR] if (level.startswith("Respondent") and ID_VAR in common_cols) else []

    merge_keys = st.multiselect(
        "Column(s) that uniquely identify a treatment record (select one or more)",
        common_cols, default=suggested_default,
    )
    if not merge_keys:
        st.info("Select at least one merge key to continue.")
        return

    # Validate uniqueness on the treatment side up front.
    dup_rows = find_duplicate_keys(treatment_df, merge_keys)
    if not dup_rows.empty:
        n_keys_dup = dup_rows[merge_keys].drop_duplicates().shape[0]
        st.error(
            f"The treatment file has {n_keys_dup} merge-key combination(s) that appear more "
            f"than once ({len(dup_rows)} rows total). Each key combination must map to exactly "
            "one treatment record — fix the treatment file before continuing."
        )
        with st.expander("Show duplicated rows"):
            st.dataframe(dup_rows, width="stretch")
        return

    # How many long-database rows share each treatment key — sanity check for the level chosen.
    fanout = long_df.groupby(merge_keys, dropna=False).size()
    f1, f2, f3 = st.columns(3)
    f1.metric("Distinct keys in long database", fanout.shape[0])
    f2.metric("Rows per key (min–max)", f"{int(fanout.min())}–{int(fanout.max())}")
    f3.metric("Rows per key (median)", f"{fanout.median():.1f}")

    st.markdown("---")

    # ── Treatment columns ──────────────────────────────────────────────────────
    st.subheader("4. Treatment columns")
    treat_candidate_cols = [c for c in treatment_df.columns if c not in merge_keys]
    if not treat_candidate_cols:
        st.error("The treatment file has no columns left to use as treatment variables besides the merge key(s).")
        return

    c1, c2 = st.columns(2)
    with c1:
        default_idx = treat_candidate_cols.index(DEFAULT_TREATMENT_VAR) if DEFAULT_TREATMENT_VAR in treat_candidate_cols else 0
        treatment_var = st.selectbox("Treatment (numeric code, for regressions)", treat_candidate_cols, index=default_idx)
    with c2:
        label_candidates = [c for c in treat_candidate_cols if c != treatment_var]
        default_idx_label = label_candidates.index(DEFAULT_TREATMENT_LABEL) if DEFAULT_TREATMENT_LABEL in label_candidates else 0
        treatment_label = st.selectbox("Treatment (label)", label_candidates, index=default_idx_label if label_candidates else 0)

    extra_options = [c for c in treat_candidate_cols if c not in (treatment_var, treatment_label)]
    extra_cols = st.multiselect(
        "Additional columns to bring over from the treatment file (optional)",
        extra_options, default=[],
    )

    if treatment_var == treatment_label:
        st.error("The numeric treatment column and the label column must be different columns.")
        return

    st.markdown("---")

    # ── Merge ────────────────────────────────────────────────────────────────
    st.subheader("5. Assign treatment")
    if st.button("🎯 Assign treatment to the long database", type="primary"):
        try:
            merged_df, diagnostics = assign_treatment(
                long_df, treatment_df, merge_keys, treatment_var, treatment_label, extra_cols,
            )
        except ValueError as e:
            st.error(str(e))
            merged_df, diagnostics = None, None

        if merged_df is not None:
            st.session_state["harmonization_treatment_df"] = merged_df
            st.session_state["harmonization_treatment_diagnostics"] = diagnostics
            st.session_state["harmonization_treatment_cols"] = (treatment_var, treatment_label)

    merged_df = st.session_state.get("harmonization_treatment_df")
    diagnostics = st.session_state.get("harmonization_treatment_diagnostics")
    chosen_cols = st.session_state.get("harmonization_treatment_cols")

    if merged_df is not None and diagnostics is not None and chosen_cols == (treatment_var, treatment_label):
        st.success(
            f"Treatment assigned — {diagnostics['n_rows']} rows preserved "
            f"({diagnostics['n_matched']} matched, {diagnostics['n_unmatched']} unmatched)."
        )

        if diagnostics["n_unmatched"] > 0:
            with st.expander(f"⚠️ {diagnostics['n_unmatched']} observation(s) with no matching treatment record", expanded=False):
                st.caption("These rows keep all their original values; the treatment columns are left missing.")
                st.dataframe(diagnostics["unmatched_keys"], hide_index=True, width="stretch")

        if PERIOD_VAR in merged_df.columns:
            with st.expander("Treatment coverage by period (baseline vs endline)", expanded=True):
                crosstab = pd.crosstab(merged_df[PERIOD_VAR], merged_df[treatment_label], dropna=False)
                st.caption(
                    "Confirms the treatment label is available for both rounds "
                    f"(`{PERIOD_VAR}` = {BASELINE_VALUE} is baseline, {ENDLINE_VALUE} is endline)."
                )
                st.dataframe(crosstab, width="stretch")

        with st.expander("Preview merged database", expanded=True):
            st.dataframe(merged_df.head(50), width="stretch")

        st.markdown("---")
        st.subheader("6. Save")
        out_filename = st.text_input("Output filename", value=DEFAULT_OUT_FILENAME)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("💾 Save to project (data_harmonization folder)", type="primary"):
                os.makedirs(DATA_HARMONIZATION_DIR, exist_ok=True)
                out_path = os.path.join(DATA_HARMONIZATION_DIR, out_filename)
                merged_df.to_csv(out_path, index=False, encoding="utf-8-sig")
                st.success(f"Saved database with treatment assigned to `{os.path.abspath(out_path)}`.")
        with c2:
            csv = merged_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                label="⬇️ Download a copy (CSV)",
                data=csv,
                file_name=out_filename or DEFAULT_OUT_FILENAME,
                mime="text/csv",
            )
    else:
        st.info("Click **Assign treatment** above to run the merge.")


if __name__ == "__main__":
    render(standalone=True)
