"""
AGEVAL — Unified Survey Design & Analysis App
Run:  python -m streamlit run scripts/app.py

Wraps steps 1-5 (dictionary selection -> ODK form -> quality checks ->
descriptive stats -> econometric models) into a single clickable app.
Reads/writes the same project folders your standalone scripts already use:
    dictionary/   config/   forms/   data_raw/   data_clean/   outputs/
"""

import streamlit as st
import pandas as pd
import numpy as np
import os
import sys
import glob
import yaml
import streamlit.components.v1 as components

# Make sibling scripts importable regardless of the working directory
SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

import s01_dictionary_selector as s01
import s02_generate_odk_form as s02
import s03_quality_check as s03
import s04_descriptive_stats as s04
import s05_run_model as s05

ROOT               = os.path.join(SCRIPTS_DIR, "..")
DICT_MASTER        = os.path.join(ROOT, "dictionary", "variables_master.xlsx")
DICT_PERSONALIZED  = os.path.join(ROOT, "dictionary", "variables_personalized.csv")
DICT_MASTER_CSV    = os.path.join(ROOT, "dictionary", "variables_master.csv")  # used by Step 5, see note in that page
CFG_PATH           = os.path.join(ROOT, "config", "config.yaml")

st.set_page_config(page_title="AGEVAL", layout="wide", page_icon="🌱")


# ── Helpers ──────────────────────────────────────────────────────────────────
def exists(path):
    return os.path.exists(path)


def list_csvs(subdir):
    folder = os.path.join(ROOT, subdir)
    if not os.path.isdir(folder):
        return []
    return sorted(glob.glob(os.path.join(folder, "*.csv")))


def pick_data_file(subdir, key_prefix, label):
    """Dropdown of CSVs found in `subdir`, plus a manual path override."""
    files = list_csvs(subdir)
    options = ["— choose a file —"] + [os.path.relpath(f, ROOT) for f in files]
    choice = st.selectbox(f"{label} (from `{subdir}/`)", options, key=f"{key_prefix}_select")
    manual = st.text_input("…or enter a path manually", value="", key=f"{key_prefix}_manual",
                            placeholder=f"e.g. {subdir}/my_file.csv")
    if manual.strip():
        return os.path.join(ROOT, manual.strip()) if not os.path.isabs(manual.strip()) else manual.strip()
    if choice != "— choose a file —":
        return os.path.join(ROOT, choice)
    return None


def status_badge(ok):
    return "🟢" if ok else "⚪"


# ── Sidebar navigation ───────────────────────────────────────────────────────
st.sidebar.title("🌱 AGEVAL")
st.sidebar.caption("Survey design & analysis pipeline")

pipeline_status = {
    "1. Dictionary Selector":  exists(DICT_PERSONALIZED),
    "2. ODK Form Generator":   len(glob.glob(os.path.join(ROOT, "forms", "*.xlsx"))) > 0,
    "3. Quality Check":        len(glob.glob(os.path.join(ROOT, "outputs", "quality", "*.html"))) > 0,
    "4. Descriptive Stats":    len(glob.glob(os.path.join(ROOT, "outputs", "stats", "*.html"))) > 0,
    "5. Econometric Model":    len(glob.glob(os.path.join(ROOT, "outputs", "models", "*.csv"))) > 0,
}

page = st.sidebar.radio(
    "Pipeline step",
    list(pipeline_status.keys()),
    format_func=lambda p: f"{status_badge(pipeline_status[p])}  {p}",
)

st.sidebar.markdown("---")
st.sidebar.caption(f"Project root:\n`{os.path.abspath(ROOT)}`")
with st.sidebar.expander("Expected folder layout"):
    st.code(
        "project/\n"
        "├── config/config.yaml\n"
        "├── dictionary/variables_master.xlsx\n"
        "├── data_raw/*.csv\n"
        "├── data_clean/*.csv\n"
        "├── forms/            (generated)\n"
        "├── outputs/          (generated)\n"
        "└── scripts/app.py    (this app)",
        language="text",
    )


# ── Page 1: Dictionary Selector ───────────────────────────────────────────────
if page == "1. Dictionary Selector":
    s01.render(standalone=False)


# ── Page 2: ODK Form Generator ────────────────────────────────────────────────
elif page == "2. ODK Form Generator":
    st.header("📋 Step 2 — Generate ODK Form")
    st.caption("Builds a KoboToolbox / ODK Central-ready XLSForm from the personalized dictionary.")

    if not exists(DICT_PERSONALIZED):
        st.warning("No personalized dictionary found yet. Go to **Step 1**, select your variables, "
                    "and click **Save to project**.")
    elif not exists(CFG_PATH):
        st.error(f"Config file not found: `{os.path.abspath(CFG_PATH)}`")
    else:
        cfg = s02.load_config()
        df = s02.load_dict()

        c1, c2, c3 = st.columns(3)
        c1.metric("Form title", cfg["project"]["name"])
        c2.metric("Form ID", cfg["project"]["form_id"])
        c3.metric("Variables included", len(df))

        with st.expander("Preview variables to include", expanded=False):
            st.dataframe(df[["variable_name", "topic", "question_type"]], width="stretch", height=300)

        if st.button("⚙️ Generate ODK XLSForm", type="primary"):
            with st.spinner("Building survey, choices and settings sheets..."):
                try:
                    out_path = s02.generate_form()
                    st.session_state["last_odk_form"] = out_path
                    st.success(f"Form generated: `{os.path.basename(out_path)}`")
                except Exception as e:
                    st.error(f"Form generation failed: {e}")

        last_form = st.session_state.get("last_odk_form")
        if last_form and exists(last_form):
            with open(last_form, "rb") as f:
                st.download_button(
                    "⬇️ Download ODK XLSForm", f, file_name=os.path.basename(last_form),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            c1, c2 = st.columns(2)
            with c1.expander("Preview 'survey' sheet"):
                st.dataframe(pd.read_excel(last_form, sheet_name="survey"), width="stretch")
            with c2.expander("Preview 'choices' sheet"):
                st.dataframe(pd.read_excel(last_form, sheet_name="choices"), width="stretch")


# ── Page 3: Quality Check ─────────────────────────────────────────────────────
elif page == "3. Quality Check":
    st.header("✅ Step 3 — Data Quality Check")
    st.caption("Applies range, missingness and outlier checks from the dictionary to a collected ODK export.")

    if not exists(DICT_PERSONALIZED):
        st.warning("No personalized dictionary found yet. Complete **Step 1** first.")
    else:
        data_path = pick_data_file("data_raw", "qc", "Collected data file")
        batch_name = st.text_input("Batch name (optional)", value="")

        if st.button("✅ Run Quality Check", type="primary", disabled=not data_path):
            with st.spinner("Checking variables against dictionary rules..."):
                try:
                    out_path = s03.run_quality_check(data_path, batch_name or None)
                    st.session_state["last_quality_report"] = out_path
                    st.success(f"Report generated: `{os.path.basename(out_path)}`")
                except Exception as e:
                    st.error(f"Quality check failed: {e}")

        last_report = st.session_state.get("last_quality_report")
        if last_report and exists(last_report):
            with open(last_report, "rb") as f:
                st.download_button("⬇️ Download HTML report", f, file_name=os.path.basename(last_report),
                                    mime="text/html")
            with open(last_report, "r", encoding="utf-8") as f:
                html = f.read()
            components.html(html, height=1400, scrolling=True)


# ── Page 4: Descriptive Stats ─────────────────────────────────────────────────
elif page == "4. Descriptive Stats":
    st.header("📊 Step 4 — Descriptive Statistics")
    st.caption("Generates charts and summary tables for every variable flagged descriptive_include=1.")

    if not exists(DICT_PERSONALIZED):
        st.warning("No personalized dictionary found yet. Complete **Step 1** first.")
    else:
        data_path = pick_data_file("data_clean", "desc", "Cleaned data file")

        if st.button("📊 Generate Descriptive Stats", type="primary", disabled=not data_path):
            with st.spinner("Rendering charts and computing summary statistics..."):
                try:
                    out_path, variables_rendered = s04.run_descriptive(data_path)
                    st.session_state["last_desc_report"] = out_path
                    st.session_state["last_desc_vars"] = variables_rendered
                    st.success(f"Report generated: `{os.path.basename(out_path)}` "
                               f"({len(variables_rendered)} variables)")
                except Exception as e:
                    st.error(f"Descriptive stats failed: {e}")

        last_report = st.session_state.get("last_desc_report")
        if last_report and exists(last_report):
            with open(last_report, "rb") as f:
                st.download_button("⬇️ Download HTML report", f, file_name=os.path.basename(last_report),
                                    mime="text/html")

            variables_rendered = st.session_state.get("last_desc_vars", [])
            current_topic = None
            for var in variables_rendered:
                if var["topic"] != current_topic:
                    current_topic = var["topic"]
                    st.markdown(f"#### {str(current_topic).upper()}")
                chart_abs = os.path.normpath(os.path.join(os.path.dirname(last_report), var["chart_path"]))
                c1, c2 = st.columns([3, 2])
                with c1:
                    if exists(chart_abs):
                        st.image(chart_abs, caption=f"{var['varname']} — {var['label']}")
                    else:
                        st.info(f"Chart image not found for {var['varname']}")
                with c2:
                    st.dataframe(
                        pd.DataFrame(list(var["stats"].items()), columns=["Statistic", "Value"]),
                        hide_index=True, width="stretch",
                    )


# ── Page 5: Econometric Model ─────────────────────────────────────────────────
elif page == "5. Econometric Model":
    st.header("📐 Step 5 — Econometric Model")
    st.caption("Runs OLS using model_role flags from the dictionary (1=dependent, 2=independent, 3=control).")
    if not exists(DICT_MASTER_CSV):
        st.warning(f"Expected dictionary file not found: `{os.path.relpath(DICT_MASTER_CSV, ROOT)}`. "
                    "Export `variables_master.xlsx` to CSV at that path, or update `s05_run_model.py`.")
    elif not s05.HAS_STATSMODELS:
        st.error("`statsmodels` is not installed. Run: `pip install statsmodels`")
    else:
        data_path = pick_data_file("data_clean", "model", "Cleaned data file")

        if st.button("📐 Run Model", type="primary", disabled=not data_path):
            with st.spinner("Fitting model..."):
                try:
                    result = s05.run_models(data_path)
                    st.session_state["last_model_result"] = result
                except Exception as e:
                    st.session_state["last_model_result"] = {"error": str(e)}

        result = st.session_state.get("last_model_result")
        if result:
            if result.get("error"):
                st.error(result["error"])
                if result.get("formula"):
                    st.caption(f"Formula attempted: `{result['formula']}`")
            else:
                st.success(f"Model fit — N = {result['n_obs']}, R² = {result['r2']}, Adj. R² = {result['r2_adj']}")
                st.markdown(f"**Formula:** `{result['formula']}`")

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Dependent", result["dep"])
                c2.metric("Independent vars", len(result["indep"]))
                c3.metric("Controls", len(result["controls"]))
                c4.metric("N obs", result["n_obs"])

                st.dataframe(result["full_table"], width="stretch")

                with st.expander("Full statsmodels summary"):
                    st.code(result["summary_txt"])

                dl1, dl2 = st.columns(2)
                with dl1:
                    csv_bytes = result["full_table"].to_csv(index=False).encode("utf-8")
                    st.download_button("⬇️ Download results (CSV)", csv_bytes,
                                        file_name="model_results.csv", mime="text/csv")
                with dl2:
                    if result.get("csv_path") and exists(result["csv_path"]):
                        st.caption(f"Also saved to `{os.path.relpath(result['csv_path'], ROOT)}`")
