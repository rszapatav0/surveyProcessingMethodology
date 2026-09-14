"""
AGEVAL — Unified Survey Design & Analysis App
Run:  python -m streamlit run baseline/01_code/app.py

Wraps steps 1-6 (dictionary selection -> ODK form -> quality checks ->
correction template -> apply corrections -> descriptive stats) into a
single clickable app. All folders are resolved from `s00_config.yaml`,
the same configuration file the standalone scripts use.
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
import s04_correction_template as s04
import s05_apply_corrections as s05
import s06_descriptive_stats as s06


# ── Config-driven paths (single source of truth: s00_config.yaml) ────────────
CFG = os.path.join(SCRIPTS_DIR, "s00_config.yaml")
with open(CFG, "r") as f:
    config = yaml.safe_load(f)
SURVEY_ROUND = config["project"]["survey_round"]

BASE        = os.path.normpath(os.path.join(os.path.dirname(CFG), config["paths"]["base"]))
MASTER      = os.path.join(BASE, config["paths"]["dictionary_master"])
DICT        = os.path.join(BASE, config["paths"]["dictionary_personalized"].format(survey_round=SURVEY_ROUND))
FORMS_OUT   = os.path.join(BASE, config["paths"]["forms_out"].format(survey_round=SURVEY_ROUND))
DATARAW     = os.path.join(BASE, config["paths"]["data_raw"].format(survey_round=SURVEY_ROUND))
DATACLEAN   = os.path.join(BASE, config["paths"]["data_clean"].format(survey_round=SURVEY_ROUND))
CORRECTIONS = os.path.join(BASE, config["paths"]["correction_files"].format(survey_round=SURVEY_ROUND))
OUTQUALITY  = os.path.join(BASE, config["paths"]["outputs_quality"].format(survey_round=SURVEY_ROUND))
OUTPLOTS    = os.path.join(BASE, config["paths"]["outputs_plots"].format(survey_round=SURVEY_ROUND))
OUTSTATS    = os.path.join(BASE, config["paths"]["outputs_stats"].format(survey_round=SURVEY_ROUND))
DICT_DIR    = os.path.dirname(DICT)

st.set_page_config(page_title="AGEVAL", layout="wide", page_icon="🌱")


# ── Helpers ──────────────────────────────────────────────────────────────────
def exists(path):
    return bool(path) and os.path.exists(path)


def list_data_files(folder, patterns=("*.csv", "*.xlsx", "*.xls")):
    if not os.path.isdir(folder):
        return []
    files = []
    for pattern in patterns:
        files.extend(glob.glob(os.path.join(folder, pattern)))
    return sorted(files)


def pick_data_file(folder, key_prefix, label):
    """Dropdown of data files found in `folder`, plus a manual path override.
    All paths are resolved relative to the project BASE directory (from config)."""
    files = list_data_files(folder)
    folder_display = os.path.relpath(folder, BASE)
    options = ["— choose a file —"] + [os.path.relpath(f, BASE) for f in files]
    choice = st.selectbox(f"{label} (from `{folder_display}/`)", options, key=f"{key_prefix}_select")
    manual = st.text_input("…or enter a path manually", value="", key=f"{key_prefix}_manual",
                            placeholder=f"e.g. {folder_display}/my_file.csv")
    if manual.strip():
        return os.path.join(BASE, manual.strip()) if not os.path.isabs(manual.strip()) else manual.strip()
    if choice != "— choose a file —":
        return os.path.join(BASE, choice)
    return None


def status_badge(ok):
    return "🟢" if ok else "⚪"


def parse_choices(choices_str):
    """'enc001:Encuestador 001 | enc002:Encuestador 002' -> DataFrame[code, label]"""
    rows = []
    if isinstance(choices_str, str) and choices_str.strip():
        for pair in choices_str.split("|"):
            pair = pair.strip()
            if not pair:
                continue
            if ":" in pair:
                code, label = pair.split(":", 1)
            else:
                code, label = pair, pair
            rows.append({"code": code.strip(), "label": label.strip()})
    return pd.DataFrame(rows, columns=["code", "label"])


def serialize_choices(df):
    """DataFrame[code, label] -> 'enc001:Encuestador 001 | enc002:Encuestador 002'"""
    df = df.dropna(subset=["code"])
    df = df[df["code"].astype(str).str.strip() != ""]
    parts = [f"{str(r.code).strip()}:{str(r.label).strip()}" for r in df.itertuples()]
    return " | ".join(parts)


# ── Sidebar navigation ───────────────────────────────────────────────────────
st.sidebar.title("🌱 AGEVAL")
st.sidebar.caption("Survey design & analysis pipeline")

pipeline_status = {
    "1. Dictionary Selector":   exists(DICT),
    "2. ODK Form Generator":    len(glob.glob(os.path.join(FORMS_OUT, "*.xlsx"))) > 0,
    "3. Quality Check":         os.path.isdir(OUTQUALITY) and len(os.listdir(OUTQUALITY)) > 0,
    "4. Correction Template":   len(glob.glob(os.path.join(CORRECTIONS, "*.xlsx"))) > 0,
    "5. Apply Corrections":     os.path.isdir(DATACLEAN) and len(os.listdir(DATACLEAN)) > 0,
    "6. Descriptive Statistics": len(glob.glob(os.path.join(OUTSTATS, "*.html"))) > 0,
}

page = st.sidebar.radio(
    "Pipeline step",
    list(pipeline_status.keys()),
    format_func=lambda p: f"{status_badge(pipeline_status[p])}  {p}",
)

st.sidebar.markdown("---")
st.sidebar.caption(f"Project base:\n`{os.path.abspath(BASE)}`")
st.sidebar.caption(f"Survey round: `{SURVEY_ROUND}`")
with st.sidebar.expander("Expected folder layout"):
    st.code(
        "dictionary/variables_master.xlsx\n"
        f"{SURVEY_ROUND}/\n"
        "├── 02_dictionary/variables_personalized.csv\n"
        "├── 03_survey/                  (generated ODK form)\n"
        "├── 04_data/\n"
        "│   ├── dataRaw/                (collected data)\n"
        "│   ├── correctionFiles/        (generated templates)\n"
        "│   └── dataClean/              (generated clean data)\n"
        "├── 05_monitoring/              (generated quality reports)\n"
        "├── 06_plots/                   (generated charts)\n"
        "└── 07_descriptiveStatistics/   (generated reports)",
        language="text",
    )


# ── Page 1: Dictionary Selector ───────────────────────────────────────────────
if page == "1. Dictionary Selector":
    s01.render(standalone=False)


# ── Page 2: ODK Form Generator ────────────────────────────────────────────────
elif page == "2. ODK Form Generator":
    st.header("📋 Step 2 — Generate ODK Form")
    st.caption("Builds a KoboToolbox / ODK Central-ready XLSForm from the personalized dictionary.")

    if not exists(DICT):
        st.warning("No personalized dictionary found yet.  Complete **Step 1** first.")
    else:
        dict_path = pick_data_file(DICT_DIR, "odk", "Personalized dictionary file")

        if not dict_path:
            st.info("Select or enter a dictionary CSV file to continue.")
        elif not exists(dict_path):
            st.error(f"File not found: `{dict_path}`")
        else:
            cfg = s02.load_config()
            df = s02.load_dict()

            c1, c2, c3 = st.columns(3)
            c1.metric("Form title", cfg["project"]["name"])
            c2.metric("Form ID", cfg["project"]["form_id"])
            c3.metric("Variables included", len(df))

            with st.expander("Preview variables to include", expanded=False):
                st.dataframe(df[["variable_name", "topic", "surv_type"]], width="stretch", height=300)

            # ── Review/edit enumerator list before generating the form ──────
            st.markdown("#### 👤 Review enumerators")
            enum_mask = df["variable_name"] == "enumerator_id"

            if not enum_mask.any():
                st.info("No `enumerator_id` variable found in the dictionary — skipping enumerator editing.")
            else:
                enum_idx = df.index[enum_mask][0]
                current_choices_str = df.loc[enum_idx, "surv_choices"]

                if "enum_choices_df" not in st.session_state:
                    st.session_state["enum_choices_df"] = parse_choices(current_choices_str)

                st.caption("Edit codes/labels, add new rows, or delete rows below. Click **Confirm enumerator list** when done.")
                edited_enum_df = st.data_editor(
                    st.session_state["enum_choices_df"],
                    num_rows="dynamic",
                    width="stretch",
                    key="enum_editor",
                    column_config={
                        "code": st.column_config.TextColumn("Code", required=True),
                        "label": st.column_config.TextColumn("Label", required=True),
                    },
                )

                if st.button("✅ Confirm enumerator list"):
                    st.session_state["enum_choices_df"] = edited_enum_df
                    new_choices_str = serialize_choices(edited_enum_df)
                    df.loc[enum_idx, "surv_choices"] = new_choices_str
                    df.to_csv(dict_path, index=False)  # persist so s02.generate_form() picks it up
                    st.session_state["enum_confirmed"] = True
                    st.success(f"Enumerator list updated ({len(edited_enum_df)} entries) and saved.")

                if st.session_state.get("enum_confirmed"):
                    st.dataframe(st.session_state["enum_choices_df"], hide_index=True, width="stretch")

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

    if not exists(DICT):
        st.warning("No personalized dictionary found yet. Complete **Step 1** first.")
    else:
        dict_path = pick_data_file(DICT_DIR, "qc_dict", "Personalized dictionary file")
        data_path = pick_data_file(DATARAW, "qc", "Collected data file")
        batch_name = st.text_input("Batch name (optional)", value="", key="qc_batch")

        # Date range filter — options are restricted to dates that actually
        # exist in the `surveyDate` column of the selected file.
        date_start = date_end = None
        if data_path:
            available_dates = s03.get_available_dates(data_path)
            if available_dates:
                col1, col2 = st.columns(2)
                with col1:
                    date_start = st.selectbox("Start date", available_dates,
                                               index=0, key="qc_date_start")
                with col2:
                    date_end = st.selectbox("End date", available_dates,
                                             index=len(available_dates) - 1, key="qc_date_end")
            else:
                st.info("No `surveyDate` values found in this file — date filter unavailable.")

        if st.button("✅ Run Quality Check", type="primary", disabled=not (data_path and dict_path)):
            with st.spinner("Checking variables against dictionary rules..."):
                try:
                    out_path = s03.run_quality_check(
                        data_path,
                        batch_name or None,
                        date_start=date_start,
                        date_end=date_end,
                        dict_path=dict_path,
                    )
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


# ── Page 4: Correction Template ───────────────────────────────────────────────
elif page == "4. Correction Template":
    st.header("📝 Step 4 — Correction Template Generator")
    st.caption("Re-runs the quality rules and packages every flagged record into an editable Excel correction workbook.")

    if not exists(DICT):
        st.warning("No personalized dictionary found yet. Complete **Step 1** first.")
    else:
        dict_path = pick_data_file(DICT_DIR, "corr_dict", "Personalized dictionary file")
        data_path = pick_data_file(DATARAW, "corr", "Collected data file")
        batch_name = st.text_input("Batch name (optional)", value="", key="corr_batch")

        date_start = date_end = None
        if data_path:
            available_dates = s03.get_available_dates(data_path)
            if available_dates:
                col1, col2 = st.columns(2)
                with col1:
                    date_start = st.selectbox("Start date", available_dates,
                                               index=0, key="corr_date_start")
                with col2:
                    date_end = st.selectbox("End date", available_dates,
                                             index=len(available_dates) - 1, key="corr_date_end")
            else:
                st.info("No `surveyDate` values found in this file — date filter unavailable.")

        if st.button("📝 Generate Correction Template", type="primary", disabled=not (data_path and dict_path)):
            with st.spinner("Evaluating quality rules and building the correction workbook..."):
                try:
                    out_path = s04.run_correction_template(
                        data_path,
                        batch_name or None,
                        date_start=date_start,
                        date_end=date_end,
                        dict_path=dict_path,
                    )
                    st.session_state["last_correction_template"] = out_path
                    st.success(f"Correction template generated: `{os.path.basename(out_path)}`")
                except Exception as e:
                    st.error(f"Correction template generation failed: {e}")

        last_template = st.session_state.get("last_correction_template")
        if last_template and exists(last_template):
            with open(last_template, "rb") as f:
                st.download_button(
                    "⬇️ Download correction template (xlsx)", f, file_name=os.path.basename(last_template),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            st.caption(
                "Open the file, fill in corrected values in the highlighted columns, "
                "then use the completed workbook in **Step 5 — Apply Corrections**."
            )


# ── Page 5: Apply Corrections ─────────────────────────────────────────────────
elif page == "5. Apply Corrections":
    st.header("🛠️ Step 5 — Apply Corrections")
    st.caption("Applies a completed correction template onto the original data and recalculates dependent variables.")

    if not exists(DICT):
        st.warning("No personalized dictionary found yet. Complete **Step 1** first.")
    else:
        dict_path = pick_data_file(DICT_DIR, "apply_dict", "Personalized dictionary file")
        data_path = pick_data_file(DATARAW, "apply_data", "Original collected data file")
        template_path = pick_data_file(CORRECTIONS, "apply_tmpl", "Completed correction template")

        c1, c2 = st.columns(2)
        with c1:
            id_col_input = st.text_input("ID column (leave blank for default: respondent_id)", value="", key="apply_id_col")
        with c2:
            suffix = st.text_input("Output filename suffix", value="_clean", key="apply_suffix")

        if st.button("🛠️ Apply Corrections", type="primary", disabled=not (data_path and template_path and dict_path)):
            with st.spinner("Applying corrections and recalculating variables..."):
                try:
                    out_path, summary = s05.apply_corrections(
                        data_path,
                        template_path,
                        id_col=id_col_input.strip() or None,
                        output_suffix=suffix.strip() or "_clean",
                        dict_path=dict_path,
                    )
                    st.session_state["last_apply_result"] = (out_path, summary)
                    st.success(f"Cleaned data saved: `{os.path.basename(out_path)}`")
                except Exception as e:
                    st.error(f"Applying corrections failed: {e}")

        result = st.session_state.get("last_apply_result")
        if result:
            out_path, summary = result

            c1, c2, c3 = st.columns(3)
            c1.metric("Records updated", summary["records_updated"])
            c2.metric("Cells updated", summary["cells_updated"])
            c3.metric("Duplicate IDs resolved", len(summary["duplicate_ids_resolved"]))

            if summary["duplicate_ids_resolved"]:
                with st.expander(f"🔁 Duplicate IDs resolved ({len(summary['duplicate_ids_resolved'])})"):
                    st.write([f"{old} → {new}" for old, new in summary["duplicate_ids_resolved"]])
            if summary["unmatched"]:
                with st.expander(f"⚠️ Unmatched rows ({len(summary['unmatched'])})"):
                    st.write(summary["unmatched"])
            if summary["skipped_variables"]:
                with st.expander(f"⚠️ Skipped variables ({len(summary['skipped_variables'])})"):
                    st.write(summary["skipped_variables"])
            if summary["skipped_calculations"]:
                with st.expander(f"⚠️ Skipped calculations ({len(summary['skipped_calculations'])})"):
                    st.write(summary["skipped_calculations"])

            if exists(out_path):
                mime = (
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    if out_path.lower().endswith((".xlsx", ".xls")) else "text/csv"
                )
                with open(out_path, "rb") as f:
                    st.download_button("⬇️ Download cleaned data", f, file_name=os.path.basename(out_path), mime=mime)


# ── Page 6: Descriptive Statistics ────────────────────────────────────────────
elif page == "6. Descriptive Statistics":
    st.header("📊 Step 6 — Descriptive Statistics")
    st.caption("Generates charts and summary tables for every variable flagged descriptive_include=1.")

    if not exists(DICT):
        st.warning("No personalized dictionary found yet. Complete **Step 1** first.")
    else:
        dict_path = pick_data_file(DICT_DIR, "desc_dict", "Personalized dictionary file")
        data_path = pick_data_file(DATACLEAN, "desc", "Cleaned data file")

        if st.button("📊 Generate Descriptive Stats", type="primary", disabled=not (data_path and dict_path)):
            with st.spinner("Rendering charts and computing summary statistics..."):
                try:
                    out_path, variables_rendered = s06.run_descriptive(data_path, dict_path=dict_path)
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