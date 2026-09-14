"""
AGEVAL Harmonization - Step 2 - Combine Baseline + Endline into a Long Database
Standalone run: python -m streamlit run harmonization/01_code/s02_combine_long.py
Also imported as a page by the unified harmonization/01_code/app.py

Uploads (or picks from the project folders) the baseline database, the
endline database, and the harmonization dictionary (output of Step 1), and:

    * Filters each round's data to the variables the dictionary flags for
      that round (`baseline` == 1 / `endline` == 1), reporting any variable
      the dictionary expects but the data doesn't have, and any column the
      data has that the dictionary doesn't flag for that round (dropped).
    * Coerces every variable to a consistent dtype across both rounds,
      based on the dictionary's `var_type` (numerical, dummy, date,
      categorical/text/constraint), reporting any values that failed to
      parse. A variable the dictionary's `surv_type` marks as
      select_one/select_multiple is always kept as text, even if `var_type`
      says otherwise — this avoids destroying multi-select choice codes
      that are sometimes mistagged `var_type=numerical` in the master
      dictionary.
    * Appends baseline and endline vertically into one long database and
      adds a `period` column (0 = baseline, 1 = endline).
    * Keeps the participant/survey ID column's values unchanged (only
      normalized to a consistent string representation so it lines up
      across rounds).

No merge/quality checks are performed here — this step only builds the
long database. That's a later step.
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
DICT_HARMONIZATION_PATH = os.path.join(BASE_PATH, config["paths"]["dictionary_harmonization"])
DATA_BASELINE_DIR = os.path.join(BASE_PATH, config["paths"]["data_baseline"])
DATA_ENDLINE_DIR = os.path.join(BASE_PATH, config["paths"]["data_endline"])
DATA_HARMONIZATION_DIR = os.path.join(BASE_PATH, config["paths"]["data_harmonization"])

# Harmonization settings (all overridable in s00_harmonizationConfig.yaml)
DEFAULT_ID_VAR = harmonize_cfg.get("id_variable", "respondent_id")
PERIOD_VAR = harmonize_cfg.get("period_variable", "period")
BASELINE_VALUE = harmonize_cfg.get("baseline_period_value", 0)
ENDLINE_VALUE = harmonize_cfg.get("endline_period_value", 1)
DEFAULT_OUT_FILENAME = harmonize_cfg.get("long_data_filename", "data_long.csv")

# Dictionary var_type -> how to coerce that column so both rounds share a dtype
NUMERIC_TYPES = {"numerical"}
DUMMY_TYPES = {"dummy"}
DATE_TYPES = {"date"}
# everything else (categorical, text, constraint, unknown) -> pandas "string"


# ── I/O helpers ──────────────────────────────────────────────────────────────
def list_data_files(folder, patterns=("*.csv", "*.xlsx", "*.xls")):
    if not folder or not os.path.isdir(folder):
        return []
    files = []
    for pattern in patterns:
        files.extend(glob.glob(os.path.join(folder, pattern)))
    return sorted(files)


def read_table(source):
    """Read a CSV/XLSX table from either a filesystem path (str) or a
    Streamlit UploadedFile-like object. Returns None if source is None."""
    if source is None:
        return None
    name = getattr(source, "name", source)
    ext = os.path.splitext(str(name))[1].lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(source)
    else:
        df = pd.read_csv(source)
    df.columns = df.columns.str.strip()  # normalize whitespace in headers
    return df


def pick_or_upload(folder, label, key_prefix, patterns=("*.csv", "*.xlsx", "*.xls"),
                    upload_types=("csv", "xlsx", "xls")):
    """Let the user either pick a file already sitting in `folder` (the
    project's configured path) or upload one from their computer. An
    upload always takes precedence over the dropdown pick.
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
    return None, None


# ── Core logic ───────────────────────────────────────────────────────────────
def filter_by_round(df, dict_df, round_col, id_var):
    """Keep only the ID column plus the variables the dictionary flags for
    this round. Reports variables the dictionary expects but the data is
    missing, and columns present in the data but not flagged (dropped)."""
    expected = [v for v in dict_df.loc[dict_df[round_col] == 1, "variable_name"].tolist() if v != id_var]
    cols = list(df.columns)
    present = [v for v in expected if v in cols]
    missing = [v for v in expected if v not in cols]
    extra = [c for c in cols if c not in expected and c != id_var]

    keep_cols = ([id_var] if id_var in cols else []) + present
    filtered = df[keep_cols].copy()
    return filtered, present, missing, extra


def coerce_series(series, var_type):
    """Coerce a column to a consistent dtype based on its dictionary
    var_type. Returns (coerced_series, n_parse_failures)."""
    non_null_before = series.notna().sum()

    if var_type in NUMERIC_TYPES:
        out = pd.to_numeric(series, errors="coerce")
    elif var_type in DUMMY_TYPES:
        out = pd.to_numeric(series, errors="coerce").round()
        out = out.astype("Int64")
    elif var_type in DATE_TYPES:
        out = pd.to_datetime(series, errors="coerce", format="mixed")
    else:  # categorical, text, constraint, or unrecognized -> plain string
        out = series.astype("string").str.strip()

    non_null_after = out.notna().sum()
    n_failures = max(0, int(non_null_before) - int(non_null_after))
    return out, n_failures


def effective_var_types(dict_df):
    """Map variable_name -> the var_type actually used for dtype coercion.

    A select_one/select_multiple variable (per the dictionary's `surv_type`)
    is coerced as categorical/text instead, but ONLY when `var_type` says
    `numerical` or `date` — those are the mistagging cases (e.g. a
    select_multiple marked `var_type=numerical` in the master dictionary,
    which would otherwise be silently wiped out by numeric coercion).
    `var_type == "dummy"` is deliberately never overridden: a Yes/No
    question is commonly implemented as `select_one` in the ODK form, but
    `dummy` is the analyst's explicit choice to treat it as a 0/1 mean, not
    a categorical distribution — the same priority used for classifying
    variables in the comparison-plots step.
    Returns (var_types, overrides) where overrides lists (variable, old_type)
    pairs that were corrected this way, for transparency."""
    d = dict_df.drop_duplicates("variable_name").set_index("variable_name")
    var_types = d["var_type"].to_dict()
    overrides = []
    if "surv_type" in d.columns:
        surv_types = d["surv_type"].to_dict()
        for var, styp in surv_types.items():
            if styp in ("select_one", "select_multiple") and var_types.get(var) in (NUMERIC_TYPES | DATE_TYPES):
                overrides.append((var, var_types[var]))
                var_types[var] = "categorical"
    return var_types, overrides


def build_long(baseline_df, endline_df, dict_df, id_var, period_var,
                baseline_value, endline_value):
    """Coerce dtypes consistently, append baseline+endline vertically, and
    add the period indicator. Returns (long_df, coercion_report_df, type_overrides)."""
    var_types, type_overrides = effective_var_types(dict_df)
    coercion_report = []

    def process(df, label):
        df = df.copy()
        if id_var in df.columns:
            df[id_var] = df[id_var].astype("string").str.strip()
        for col in df.columns:
            if col == id_var:
                continue
            vtype = var_types.get(col, "text")
            coerced, n_fail = coerce_series(df[col], vtype)
            df[col] = coerced
            if n_fail > 0:
                coercion_report.append({
                    "round": label, "variable": col, "var_type": vtype,
                    "values_failed_to_parse": n_fail,
                })
        return df

    baseline_proc = process(baseline_df, "baseline")
    endline_proc = process(endline_df, "endline")

    baseline_proc[period_var] = baseline_value
    endline_proc[period_var] = endline_value

    long_df = pd.concat([baseline_proc, endline_proc], ignore_index=True, sort=False)

    # Column order: id, period, then dictionary order, then any leftovers.
    dict_order = [v for v in dict_df["variable_name"].tolist() if v in long_df.columns]
    ordered_cols = [c for c in [id_var, period_var] if c in long_df.columns]
    ordered_cols += [c for c in dict_order if c not in ordered_cols]
    ordered_cols += [c for c in long_df.columns if c not in ordered_cols]
    long_df = long_df[ordered_cols]

    return long_df, pd.DataFrame(coercion_report), type_overrides


# ── Streamlit page ───────────────────────────────────────────────────────────
def render(standalone: bool = False):
    """Render the combine-into-long-database step. Set standalone=True to
    also set page config and title (only needed when this file is run
    directly, not from the unified harmonization app.py)."""

    if standalone:
        st.set_page_config(page_title="AGEVAL Harmonization - Combine Long Database", layout="wide")

    st.title("AGEVAL Harmonization - Step 2 — Combine into a Long Database")
    st.caption(
        "Filters each round to the variables the harmonization dictionary flags for it, "
        "aligns names and types, appends baseline + endline vertically, and adds a `period` "
        "column (0 = baseline, 1 = endline). No merge or quality checks happen here."
    )

    # ── Dictionary input ─────────────────────────────────────────────────────
    st.subheader("1. Harmonization dictionary")
    dict_source, dict_name = pick_or_upload(
        os.path.dirname(DICT_HARMONIZATION_PATH), "Harmonization dictionary (from Step 1)",
        "dict", patterns=("*.csv", "*.xlsx"), upload_types=("csv", "xlsx"),
    )
    if dict_source is None and os.path.exists(DICT_HARMONIZATION_PATH):
        dict_source, dict_name = DICT_HARMONIZATION_PATH, os.path.basename(DICT_HARMONIZATION_PATH)

    if dict_source is None:
        st.warning("Select, upload, or generate (Step 1) a harmonization dictionary to continue.")
        return

    dict_df = read_table(dict_source)
    required_dict_cols = {"variable_name", "var_type", "baseline", "endline"}
    if not required_dict_cols.issubset(dict_df.columns):
        st.error(
            f"`{dict_name}` doesn't look like a harmonization dictionary — "
            f"missing columns: {sorted(required_dict_cols - set(dict_df.columns))}"
        )
        return
    st.success(f"Loaded dictionary `{dict_name}` — {len(dict_df)} variables.")

    st.markdown("---")

    # ── Baseline / endline data input ────────────────────────────────────────
    st.subheader("2. Baseline and endline databases")
    c1, c2 = st.columns(2)
    with c1:
        baseline_source, baseline_name = pick_or_upload(DATA_BASELINE_DIR, "Baseline database", "baseline")
    with c2:
        endline_source, endline_name = pick_or_upload(DATA_ENDLINE_DIR, "Endline database", "endline")

    if baseline_source is None or endline_source is None:
        st.info("Select or upload both the baseline and endline databases to continue.")
        return

    baseline_raw = read_table(baseline_source)
    endline_raw = read_table(endline_source)
    st.success(
        f"Loaded baseline `{baseline_name}` ({baseline_raw.shape[0]} rows, {baseline_raw.shape[1]} cols) "
        f"and endline `{endline_name}` ({endline_raw.shape[0]} rows, {endline_raw.shape[1]} cols)."
    )

    st.markdown("---")

    # ── ID variable ──────────────────────────────────────────────────────────
    st.subheader("3. Participant / survey ID")
    common_cols = [c for c in baseline_raw.columns if c in endline_raw.columns]
    id_options = common_cols if common_cols else list(baseline_raw.columns)
    default_idx = id_options.index(DEFAULT_ID_VAR) if DEFAULT_ID_VAR in id_options else 0
    id_var = st.selectbox(
        "ID column present in both databases (kept unchanged in the long database)",
        id_options, index=default_idx,
    )

    st.markdown("---")

    # ── Round-level diagnostics (what the dictionary changes per round) ──────
    st.subheader("4. Dictionary-driven changes per round")
    baseline_filtered, base_present, base_missing, base_extra = filter_by_round(
        baseline_raw, dict_df, "baseline", id_var
    )
    endline_filtered, end_present, end_missing, end_extra = filter_by_round(
        endline_raw, dict_df, "endline", id_var
    )

    d1, d2 = st.columns(2)
    with d1:
        st.markdown("**Baseline**")
        st.caption(f"{len(base_present)} variables kept (dictionary `baseline` = 1).")
        if base_extra:
            st.warning(f"{len(base_extra)} column(s) in the data are not flagged for baseline and will be dropped:")
            st.write(base_extra)
        if base_missing:
            st.error(f"{len(base_missing)} variable(s) flagged for baseline in the dictionary are missing from the data:")
            st.write(base_missing)
        if not base_extra and not base_missing:
            st.info("No changes — the baseline data already matches the dictionary exactly.")
    with d2:
        st.markdown("**Endline**")
        st.caption(f"{len(end_present)} variables kept (dictionary `endline` = 1).")
        if end_extra:
            st.warning(f"{len(end_extra)} column(s) in the data are not flagged for endline and will be dropped:")
            st.write(end_extra)
        if end_missing:
            st.error(f"{len(end_missing)} variable(s) flagged for endline in the dictionary are missing from the data:")
            st.write(end_missing)
        if not end_extra and not end_missing:
            st.info("No changes — the endline data already matches the dictionary exactly.")

    st.markdown("---")

    # ── Combine ──────────────────────────────────────────────────────────────
    st.subheader("5. Combine")
    if st.button("🔗 Combine baseline + endline into one long database", type="primary"):
        long_df, coercion_report, type_overrides = build_long(
            baseline_filtered, endline_filtered, dict_df, id_var, PERIOD_VAR,
            BASELINE_VALUE, ENDLINE_VALUE,
        )
        st.session_state["harmonization_long_df"] = long_df
        st.session_state["harmonization_coercion_report"] = coercion_report
        st.session_state["harmonization_type_overrides"] = type_overrides

    long_df = st.session_state.get("harmonization_long_df")
    coercion_report = st.session_state.get("harmonization_coercion_report")
    type_overrides = st.session_state.get("harmonization_type_overrides")

    if long_df is not None:
        st.success(f"Long database built: {len(long_df)} rows, {long_df.shape[1]} columns.")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total rows", len(long_df))
        m2.metric(f"{PERIOD_VAR} = {BASELINE_VALUE} (baseline)", int((long_df[PERIOD_VAR] == BASELINE_VALUE).sum()))
        m3.metric(f"{PERIOD_VAR} = {ENDLINE_VALUE} (endline)", int((long_df[PERIOD_VAR] == ENDLINE_VALUE).sum()))
        m4.metric("Variables", long_df.shape[1])

        if type_overrides:
            with st.expander(f"ℹ️ {len(type_overrides)} variable(s) treated as categorical based on `surv_type`", expanded=False):
                st.caption(
                    "These are flagged select_one/select_multiple in `surv_type` but tagged with a numeric-ish "
                    "`var_type` in the dictionary — kept as text instead of being forced through numeric parsing:"
                )
                st.dataframe(pd.DataFrame(type_overrides, columns=["variable", "var_type in dictionary"]),
                             hide_index=True, width="stretch")

        if coercion_report is not None and not coercion_report.empty:
            with st.expander(f"⚠️ Type-parsing warnings ({len(coercion_report)} column/round pairs)", expanded=False):
                st.caption(
                    "Values that didn't match the dictionary's `var_type` for that column were set to missing "
                    "so both rounds share a consistent type. Review these before the next (quality-check) step."
                )
                st.dataframe(coercion_report, hide_index=True, width="stretch")

        with st.expander("Preview long database", expanded=True):
            st.dataframe(long_df.head(50), width="stretch")

        st.markdown("---")
        st.subheader("6. Save")
        out_filename = st.text_input("Output filename", value=DEFAULT_OUT_FILENAME)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("💾 Save to project (data_harmonization folder)", type="primary"):
                os.makedirs(DATA_HARMONIZATION_DIR, exist_ok=True)
                out_path = os.path.join(DATA_HARMONIZATION_DIR, out_filename)
                long_df.to_csv(out_path, index=False, encoding="utf-8-sig")
                st.session_state["harmonization_long_saved"] = out_path
                st.success(f"Saved long database to `{os.path.abspath(out_path)}`.")
        with c2:
            csv = long_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                label="⬇️ Download a copy (CSV)",
                data=csv,
                file_name=out_filename or DEFAULT_OUT_FILENAME,
                mime="text/csv",
            )
    else:
        st.info("Click **Combine** above to build the long database.")


if __name__ == "__main__":
    render(standalone=True)