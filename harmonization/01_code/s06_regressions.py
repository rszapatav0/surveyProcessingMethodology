"""
AGEVAL Harmonization - Step 6 - RCT Regressions
Standalone run: python -m streamlit run harmonization/01_code/s06_regressions.py
Also imported as a page by the unified harmonization/01_code/app.py

Loads the harmonization dictionary (output of Step 1) and the long
baseline+endline database with treatment assigned (output of Step 3), and:

    * Automatically reshapes the long database into a respondent-level wide
      database, using the dictionary flags added in Step 1:
        - `outcome_var`  -> eligible outcome variables (baseline value kept
          as a possible lagged control, endline value is the regression
          outcome)
        - `control_var`  -> eligible baseline controls (baseline value only)
        - `logarithm_var`-> variables also offered in log form (ln(x),
          non-positive values set to missing)
        - `fe_var`       -> eligible fixed-effect variables (time-invariant:
          baseline value, falling back to endline if baseline is missing)
        - `cluster_var`  -> eligible clustering variables (same
          time-invariant rule as fixed effects)
    * Lets you build one or more regression "specifications" (treatment
      coding, baseline-outcome control, baseline controls, fixed effects,
      treatment x control interactions, and the standard-error type), each
      of which can be run against one or more outcome variables.
    * Runs the selected (specification x outcome) regressions with OLS,
      and lays the results out as standard side-by-side regression tables
      (coefficients, significance stars, SEs in parentheses, N, R-squared,
      and descriptive rows such as "<var> FE: Yes/No" and "SE: Clustered
      (<var>)").
    * Exports everything — the regression tables and the specification
      definitions — to a single HTML file and a single .tex file.
"""

import streamlit as st
import pandas as pd
import numpy as np
import os
import re
import glob
import uuid
import html as html_lib
import yaml
from datetime import datetime

import statsmodels.formula.api as smf
from statsmodels.iolib.summary2 import summary_col

# Read config
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "s00_harmonizationConfig.yaml")
with open(CONFIG_PATH, "r") as f:
    config = yaml.safe_load(f)

harmonize_cfg = config.get("harmonize", {})
project_cfg = config.get("project", {})
model_cfg = config.get("model", {})

# Paths
BASE_PATH = os.path.normpath(os.path.join(os.path.dirname(CONFIG_PATH), config["paths"]["base"]))
DICT_HARMONIZATION_PATH = os.path.join(BASE_PATH, config["paths"]["dictionary_harmonization"])
DATA_HARMONIZATION_DIR = os.path.join(BASE_PATH, config["paths"]["data_harmonization"])
OUTPUTS_REGRESSIONS_DIR = os.path.join(BASE_PATH, config["paths"]["outputs_regressions"])

# Harmonization settings (all overridable in s00_harmonizationConfig.yaml)
ID_VAR = harmonize_cfg.get("id_variable", "respondent_id")
PERIOD_VAR = harmonize_cfg.get("period_variable", "period")
BASELINE_VALUE = harmonize_cfg.get("baseline_period_value", 0)
ENDLINE_VALUE = harmonize_cfg.get("endline_period_value", 1)
TREATMENT_VAR = harmonize_cfg.get("treatment_var", "treatment")
TREATMENT_LABEL = harmonize_cfg.get("treatment_label", "treatment_label")
DEFAULT_LONG_TREATMENT_FILENAME = harmonize_cfg.get("long_treatment_filename", "data_long_treatment.csv")
LANGUAGE = project_cfg.get("language", "spanish")
PROJECT_NAME = project_cfg.get("name", "AGEVAL")
DEFAULT_OUTPUT_FORMAT = model_cfg.get("output_format", "both")  # "latex", "csv"(->html here), or "both"

DICT_FLAG_COLS = ["outcome_var", "control_var", "logarithm_var", "fe_var", "cluster_var"]


# ── I/O helpers (same pattern as previous steps) ─────────────────────────────
def list_data_files(folder, patterns=("*.csv", "*.xlsx", "*.xls")):
    if not folder or not os.path.isdir(folder):
        return []
    files = []
    for pattern in patterns:
        files.extend(glob.glob(os.path.join(folder, pattern)))
    return sorted(files)


def read_table(source, sheet_name=0):
    if source is None:
        return None
    name = getattr(source, "name", source)
    ext = os.path.splitext(str(name))[1].lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(source, sheet_name=sheet_name)
    else:
        df = pd.read_csv(source)
    df.columns = df.columns.str.strip()
    return df


def pick_or_upload(folder, label, key_prefix, patterns=("*.csv", "*.xlsx", "*.xls"),
                    upload_types=("csv", "xlsx", "xls"), default_path=None):
    st.markdown(f"**{label}**")
    files = list_data_files(folder, patterns)
    folder_display = os.path.relpath(folder, BASE_PATH) if folder else "n/a"
    options = ["— choose a file from the project —"] + [os.path.relpath(f, BASE_PATH) for f in files]
    choice = st.selectbox(f"From `{folder_display}/`", options, key=f"{key_prefix}_select")
    uploaded = st.file_uploader("…or upload a file", type=list(upload_types), key=f"{key_prefix}_upload")

    if uploaded is not None:
        return uploaded, uploaded.name
    if choice != "— choose a file from the project —":
        p = os.path.join(BASE_PATH, choice)
        return p, os.path.basename(p)
    if default_path and os.path.exists(default_path):
        return default_path, os.path.basename(default_path)
    return None, None


def get_label(row, language=LANGUAGE):
    if language == "english":
        return row.get("label_english") or row.get("label_spanish") or row["variable_name"]
    return row.get("label_spanish") or row.get("label_english") or row["variable_name"]


def sanitize_token(s):
    s = re.sub(r"[^0-9A-Za-z_]+", "_", str(s)).strip("_")
    return s or "level"


# ── Wide (respondent-level) reshape ──────────────────────────────────────────
def build_wide(long_df, dict_df, id_var, period_var, baseline_value, endline_value,
               treatment_var, treatment_label):
    """Reshape the long database into one row per respondent, using the
    dictionary's outcome_var/control_var/fe_var/cluster_var/logarithm_var
    flags. Returns (wide_df, catalog, notes)."""
    notes = []
    for col in DICT_FLAG_COLS:
        if col not in dict_df.columns:
            dict_df[col] = 0
            notes.append(f"Dictionary has no `{col}` column — treated as 0 for every variable.")

    ids = pd.Index(sorted(long_df[id_var].dropna().unique()), name=id_var)
    base = (long_df[long_df[period_var] == baseline_value]
            .drop_duplicates(subset=id_var, keep="first").set_index(id_var))
    end = (long_df[long_df[period_var] == endline_value]
           .drop_duplicates(subset=id_var, keep="first").set_index(id_var))
    wide = pd.DataFrame(index=ids)

    def col_at(frame, col):
        if col in frame.columns:
            return frame[col].reindex(ids)
        return pd.Series(np.nan, index=ids)

    def flagged(col):
        return dict_df.loc[dict_df[col].fillna(0).astype(bool), "variable_name"].astype(str).tolist()

    outcome_names = flagged("outcome_var")
    control_names = flagged("control_var")
    fe_names = flagged("fe_var")
    cluster_names = flagged("cluster_var")
    log_names = set(flagged("logarithm_var"))

    outcome_cols = {}
    for v in outcome_names:
        if v not in long_df.columns:
            notes.append(f"Outcome `{v}` is flagged but not found in the database — skipped.")
            continue
        bl_col = el_col = None
        bl = col_at(base, v)
        if bl.notna().any():
            bl_col = f"{v}__bl"
            wide[bl_col] = bl
        el = col_at(end, v)
        if el.notna().any():
            el_col = f"{v}__el"
            wide[el_col] = el
        else:
            notes.append(f"Outcome `{v}` has no endline data — it cannot be used as a regression outcome.")
        outcome_cols[v] = {"bl": bl_col, "el": el_col}

    control_cols = {}
    for v in control_names:
        if v not in long_df.columns:
            notes.append(f"Control `{v}` is flagged but not found in the database — skipped.")
            continue
        existing_bl = outcome_cols.get(v, {}).get("bl")
        if existing_bl:
            control_cols[v] = existing_bl
            continue
        bl = col_at(base, v)
        if bl.notna().any():
            col = f"{v}__bl"
            wide[col] = bl
            control_cols[v] = col
            continue
        el = col_at(end, v)
        if el.notna().any():
            col = f"{v}__bl_fromendline"
            wide[col] = el
            control_cols[v] = col
            notes.append(f"Control `{v}` has no baseline data — using endline values instead.")
        else:
            notes.append(f"Control `{v}` has no data in either round — skipped.")

    def add_time_invariant(names, label):
        cols = {}
        for v in names:
            if v not in long_df.columns:
                notes.append(f"{label} `{v}` is flagged but not found in the database — skipped.")
                continue
            bl = col_at(base, v)
            el = col_at(end, v)
            combined = bl.combine_first(el)
            if combined.notna().any():
                col = v if v not in wide.columns else f"{v}__{sanitize_token(label)}"
                wide[col] = combined
                cols[v] = col
            else:
                notes.append(f"{label} `{v}` has no data in either round — skipped.")
        return cols

    fe_cols = add_time_invariant(fe_names, "Fixed effect")
    cluster_cols = add_time_invariant(cluster_names, "Cluster")

    for tv in (treatment_var, treatment_label):
        if tv in long_df.columns:
            bl = col_at(base, tv)
            el = col_at(end, tv)
            wide[tv] = bl.combine_first(el)
        else:
            notes.append(f"Treatment column `{tv}` was not found in the database.")

    # Logarithm transforms (ln(x); non-positive values become missing)
    log_eligible = set()
    for v, cc in outcome_cols.items():
        if v in log_names:
            for key in ("bl", "el"):
                if cc.get(key):
                    log_eligible.add(cc[key])
    for v, c in control_cols.items():
        if v in log_names:
            log_eligible.add(c)

    log_cols = {}
    for c in log_eligible:
        vals = pd.to_numeric(wide[c], errors="coerce")
        n_nonpos = int((vals <= 0).sum())
        ln_col = f"ln_{c}"
        wide[ln_col] = np.where(vals > 0, np.log(vals), np.nan)
        log_cols[c] = ln_col
        if n_nonpos > 0:
            notes.append(f"`{c}`: {n_nonpos} non-positive value(s) treated as missing in `{ln_col}`.")

    catalog = {
        "outcome_cols": outcome_cols,
        "control_cols": control_cols,
        "fe_cols": fe_cols,
        "cluster_cols": cluster_cols,
        "log_cols": log_cols,
    }
    return wide.reset_index(), catalog, notes


def build_variable_catalog(dict_df, catalog):
    """Turn the wide-reshape catalog into display-friendly option lists for
    the specification builder: [{'key','display','col', ...}, ...]."""
    labels = {str(row["variable_name"]): get_label(row) for _, row in dict_df.iterrows()}

    outcome_options = []
    for v, cc in catalog["outcome_cols"].items():
        el = cc.get("el")
        if not el:
            continue
        lab = labels.get(v, v)
        outcome_options.append({"col": el, "display": f"{lab} [{v}]", "baseline_col": cc.get("bl"), "var": v})
        if el in catalog["log_cols"]:
            ln_el = catalog["log_cols"][el]
            ln_bl = catalog["log_cols"].get(cc.get("bl")) if cc.get("bl") else None
            outcome_options.append({"col": ln_el, "display": f"{lab} [{v}], log form",
                                     "baseline_col": ln_bl, "var": v})

    control_options = []
    for v, c in catalog["control_cols"].items():
        lab = labels.get(v, v)
        control_options.append({"col": c, "display": f"{lab} [{v}]"})
        if c in catalog["log_cols"]:
            control_options.append({"col": catalog["log_cols"][c], "display": f"{lab} [{v}], log form"})

    fe_options = [{"col": c, "display": f"{labels.get(v, v)} [{v}]"} for v, c in catalog["fe_cols"].items()]
    cluster_options = [{"col": c, "display": f"{labels.get(v, v)} [{v}]"} for v, c in catalog["cluster_cols"].items()]
    return outcome_options, control_options, fe_options, cluster_options


def detect_arms(wide_df, treatment_var, treatment_label):
    sub = wide_df[[treatment_var, treatment_label]].dropna(subset=[treatment_label]).drop_duplicates()
    if treatment_var in sub.columns and sub[treatment_var].notna().any():
        sub = sub.sort_values(treatment_var)
    else:
        sub = sub.sort_values(treatment_label)
    arms = sub[treatment_label].astype(str).tolist()
    default_ref = next((a for a in arms if re.search(r"control|comparison|placebo", a, re.I)), (arms[0] if arms else None))
    return arms, default_ref


def build_treatment_dummies(wide_df, arms, treatment_label):
    """One 0/1 dummy column per treatment arm, added to wide_df in place.
    Returns {arm: column_name}."""
    dummy_cols = {}
    for arm in arms:
        col = f"T_{sanitize_token(arm)}"
        while col in wide_df.columns and col not in dummy_cols.values():
            col += "_"
        wide_df[col] = (wide_df[treatment_label] == arm).astype(float)
        dummy_cols[arm] = col
    return dummy_cols


# ── Specification -> formula ─────────────────────────────────────────────────
def build_spec_inputs(spec, outcome_opt, wide_df, dummy_cols):
    """Returns (formula, data, cluster_col_or_None, caveats:list[str])."""
    caveats = []
    terms = []
    needed = [outcome_opt["col"]]

    if spec["treatment_mode"] == "categorical":
        arm_cols = [dummy_cols[a] for a in spec["treatment_arms"] if a in dummy_cols]
        terms += arm_cols
        needed += arm_cols
        mask = wide_df[TREATMENT_LABEL].isin([spec["reference_level"]] + spec["treatment_arms"])
        interact_base_cols = arm_cols
    else:
        terms.append(TREATMENT_VAR)
        needed.append(TREATMENT_VAR)
        mask = wide_df[TREATMENT_VAR].notna()
        interact_base_cols = [TREATMENT_VAR]

    if spec["include_baseline_outcome"]:
        bl = outcome_opt.get("baseline_col")
        if bl:
            terms.append(bl)
            needed.append(bl)
        else:
            caveats.append("Baseline value of this outcome is unavailable — baseline-outcome control omitted.")

    terms += spec["control_cols"]
    needed += spec["control_cols"]

    for bcol in interact_base_cols:
        for ccol in spec["interact_control_cols"]:
            terms.append(f"{bcol}:{ccol}")

    for fe in spec["fe_cols"]:
        terms.append(f"C({fe})")
        needed.append(fe)

    cluster_col = None
    if spec["se_type"] == "cluster":
        cluster_col = spec["cluster_col"]
        needed.append(cluster_col)

    needed = list(dict.fromkeys(needed))
    cols_to_keep = list(dict.fromkeys(needed + [TREATMENT_LABEL]))
    cols_to_keep = [c for c in cols_to_keep if c in wide_df.columns]
    data = wide_df.loc[mask, cols_to_keep].dropna(subset=[c for c in needed if c in wide_df.columns]).copy()

    formula = f"{outcome_opt['col']} ~ " + " + ".join(terms) if terms else f"{outcome_opt['col']} ~ 1"
    return formula, data, cluster_col, caveats


def fit_spec(spec, outcome_opt, wide_df, dummy_cols):
    formula, data, cluster_col, caveats = build_spec_inputs(spec, outcome_opt, wide_df, dummy_cols)
    if data.empty or len(data) <= data.shape[1]:
        raise ValueError(f"Not enough observations to estimate this specification (N={len(data)}).")

    model = smf.ols(formula, data=data)
    if spec["se_type"] == "robust":
        res = model.fit(cov_type="HC1")
        se_desc = "Robust (HC1)"
    elif spec["se_type"] == "cluster":
        res = model.fit(cov_type="cluster", cov_kwds={"groups": data[cluster_col]})
        se_desc = f"Clustered ({cluster_col})"
    else:
        res = model.fit()
        se_desc = "Classical"

    res.spec_name = spec["name"]
    res.spec_se_desc = se_desc
    res.spec_has_controls = "Yes" if spec["control_cols"] else "No"
    res.spec_baseline_outcome = "Yes" if (spec["include_baseline_outcome"] and outcome_opt.get("baseline_col")
                                            and outcome_opt["baseline_col"] in formula) else "No"
    if spec["treatment_mode"] == "categorical":
        res.spec_treatment_desc = f"{spec['reference_level']} (ref.) vs " + ", ".join(spec["treatment_arms"])
    else:
        res.spec_treatment_desc = "Numeric (as provided)"
    return res, formula, cluster_col, caveats


# ── Table assembly (statsmodels summary_col: coef/SE/stars, N, R2, FE rows) ─
def build_comparison_table(results, model_names, fe_cols, fe_label_map=None):
    """fe_cols: raw dictionary variable names used as C(<fe_col>) fixed
    effects in the fitted formulas. fe_label_map: optional {raw_col:
    display_label} for the "<label> FE" row names.

    Note: this deliberately does NOT use summary_col()'s own `fixed_effects`
    kwarg — that parameter was only added in newer statsmodels releases, and
    raised `TypeError: unexpected keyword argument 'fixed_effects'` on older
    installs. The "<FE> FE: Yes/No" rows and the removal of the raw C(fe)
    dummy-coefficient rows are reproduced manually below using only the
    long-stable `info_dict`/`regressor_order`/`drop_omitted` arguments."""
    fe_label_map = fe_label_map or {}
    info_dict = {
        "N": lambda x: f"{int(x.nobs)}",
        "SE type": lambda x: x.spec_se_desc,
        "Controls": lambda x: x.spec_has_controls,
        "Baseline outcome control": lambda x: x.spec_baseline_outcome,
        "Treatment": lambda x: x.spec_treatment_desc,
    }
    for fe in fe_cols:
        label = fe_label_map.get(fe, fe)
        marker = f"C({fe})"
        info_dict[f"{label} FE"] = (
            lambda x, marker=marker: "Yes" if any(marker in p for p in x.params.index) else "No"
        )

    # Regressors to show, in first-seen order across all models; every
    # C(...) fixed-effect dummy is excluded here (and therefore dropped
    # from the table via drop_omitted=True below), leaving only the
    # substantive coefficients plus the intercept.
    regressor_order = []
    for r in results:
        for p in r.params.index:
            if p not in regressor_order and not p.startswith("C(") and p != "Intercept":
                regressor_order.append(p)
    regressor_order += ["Intercept"]

    tab = summary_col(
        results, stars=True, float_format="%.3f", model_names=model_names,
        info_dict=info_dict, regressor_order=regressor_order, drop_omitted=True,
        include_r2=True,
    )
    return tab


# ── HTML / LaTeX export ──────────────────────────────────────────────────────
REPORT_CSS = """
  body { font-family: -apple-system, sans-serif; max-width: 1150px; margin: 2rem auto; color: #1a1a1a; }
  h1   { font-size: 1.5rem; font-weight: 600; border-bottom: 2px solid #2563EB; padding-bottom: .5rem; }
  h2   { font-size: 1.2rem; font-weight: 700; margin-top: 2.5rem; padding-top: 1rem; border-top: 1px solid #E2E8F0; color: #0F172A; }
  h2 .sub { font-size: .8rem; font-weight: 400; color: #64748B; }
  table.simpletable { width: 100%; border-collapse: collapse; font-size: .82rem; margin-top: .75rem; }
  table.simpletable tr:first-child th { background: #1D4ED8; color: white; padding: .4rem .6rem; text-align: left; }
  table.simpletable tr:not(:first-child) th { text-align: left; font-weight: 600; padding: .3rem .6rem; border-bottom: 1px solid #E2E8F0; background: transparent; color: inherit; }
  table.simpletable td { padding: .3rem .6rem; border-bottom: 1px solid #E2E8F0; }
  table.simpletable tr:nth-child(even) { background: #F8FAFC; }
  table.spec-table { width: 100%; border-collapse: collapse; font-size: .82rem; margin-top: .75rem; }
  table.spec-table th { background: #1D4ED8; color: white; padding: .4rem .6rem; text-align: left; }
  table.spec-table td { padding: .4rem .6rem; border-bottom: 1px solid #E2E8F0; vertical-align: top; }
  table.spec-table tr:nth-child(even) { background: #F8FAFC; }
  .note { font-size: .78rem; color: #64748B; margin-top: .4rem; }
  footer { margin-top: 3rem; font-size: .75rem; color: #94A3B8; }
"""


def esc(x):
    return html_lib.escape(str(x))


def build_html_report(project_name, generated_at, spec_defs_df, outcome_tables):
    """outcome_tables: list of (outcome_display, html_table_str)"""
    spec_html = spec_defs_df.to_html(index=False, classes="spec-table", border=0, escape=True)
    sections = []
    for outcome_display, table_html in outcome_tables:
        sections.append(f"<h2>{esc(outcome_display)}</h2>\n{table_html}")
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>{esc(project_name)} - RCT Regressions</title>
<style>{REPORT_CSS}</style>
</head>
<body>
<h1>{esc(project_name)} — RCT Regression Results</h1>
<p>Generated: {esc(generated_at)}</p>
<h2>Specification definitions</h2>
{spec_html}
<p class="note">Standard errors in parentheses. * p&lt;.1, ** p&lt;.05, *** p&lt;.01</p>
{''.join(sections)}
<footer>Generated by AGEVAL Harmonization v1.0 — {esc(project_name)}</footer>
</body>
</html>
"""


def build_latex_report(project_name, generated_at, spec_defs_df, outcome_tables_tex):
    """outcome_tables_tex: list of (outcome_display, tex_table_str)"""
    spec_tex = spec_defs_df.to_latex(index=False, escape=True)
    sections = []
    for outcome_display, tex in outcome_tables_tex:
        safe_title = outcome_display.replace("_", "\\_").replace("&", "\\&")
        # Drop the empty \caption{}/\label{} statsmodels always emits, and
        # pin the float in place instead of letting it drift to its own page.
        tex = tex.replace("\\caption{}\n", "").replace("\\label{}\n", "")
        tex = tex.replace("\\begin{table}", "\\begin{table}[!htbp]")
        sections.append(f"\\section*{{{safe_title}}}\n{tex}")
    body = "\n\n".join(sections)
    return f"""\\documentclass[11pt]{{article}}
\\usepackage[utf8]{{inputenc}}
\\usepackage[T1]{{fontenc}}
\\usepackage[margin=1in]{{geometry}}
\\usepackage{{booktabs}}
\\usepackage{{longtable}}
\\usepackage{{graphicx}}
\\begin{{document}}
\\title{{{project_name} --- RCT Regression Results}}
\\date{{{generated_at}}}
\\maketitle

\\section*{{Specification definitions}}
\\resizebox{{\\textwidth}}{{!}}{{%
{spec_tex}}}

\\bigskip
\\noindent Standard errors in parentheses. * p$<$.1, ** p$<$.05, *** p$<$.01

{body}

\\end{{document}}
"""


def spec_definitions_df(specs, col_to_display):
    rows = []
    for s in specs:
        if s["treatment_mode"] == "categorical":
            treat_desc = f"{s['reference_level']} (ref.) vs " + ", ".join(s["treatment_arms"])
        else:
            treat_desc = "Numeric treatment variable"
        rows.append({
            "Specification": s["name"],
            "Treatment": treat_desc,
            "Baseline outcome control": "Yes" if s["include_baseline_outcome"] else "No",
            "Controls": ", ".join(col_to_display.get(c, c) for c in s["control_cols"]) or "—",
            "Treatment x Control interactions": ", ".join(col_to_display.get(c, c) for c in s["interact_control_cols"]) or "—",
            "Fixed effects": ", ".join(col_to_display.get(c, c) for c in s["fe_cols"]) or "—",
            "Standard errors": ("Clustered (" + col_to_display.get(s["cluster_col"], str(s["cluster_col"])) + ")"
                                 if s["se_type"] == "cluster" else
                                 "Robust (HC1)" if s["se_type"] == "robust" else "Classical"),
            "Outcomes": ", ".join(s["outcome_displays"]),
        })
    return pd.DataFrame(rows)


# ── Streamlit page ───────────────────────────────────────────────────────────
def render(standalone: bool = False):
    if standalone:
        st.set_page_config(page_title="AGEVAL Harmonization - RCT Regressions", layout="wide")

    st.title("AGEVAL Harmonization - Step 6 — RCT Regressions")
    st.caption(
        "Reshapes the long baseline+endline database into a respondent-level wide database using the "
        "`outcome_var`/`control_var`/`logarithm_var`/`fe_var`/`cluster_var` dictionary flags, then lets you "
        "build and run OLS regression specifications and export publication-style tables."
    )

    # ── 1. Dictionary ────────────────────────────────────────────────────────
    st.subheader("1. Harmonization dictionary")
    dict_source, dict_name = pick_or_upload(
        os.path.dirname(DICT_HARMONIZATION_PATH), "Harmonization dictionary (from Step 1)",
        "reg_dict", patterns=("*.csv", "*.xlsx"), upload_types=("csv", "xlsx"),
        default_path=DICT_HARMONIZATION_PATH,
    )
    if dict_source is None:
        st.info("Select or upload the harmonization dictionary to continue.")
        return
    dict_df = read_table(dict_source)
    if "variable_name" not in dict_df.columns:
        st.error(f"`{dict_name}` doesn't look like a harmonization dictionary — missing `variable_name`.")
        return
    missing_flags = [c for c in DICT_FLAG_COLS if c not in dict_df.columns]
    if missing_flags:
        st.warning(
            f"This dictionary is missing {missing_flags} — re-run Step 1 with the updated master dictionary "
            "so these columns are generated. Treating them as 0 for now."
        )
    st.success(f"Loaded `{dict_name}` — {len(dict_df)} variables.")

    st.markdown("---")

    # ── 2. Long database with treatment ──────────────────────────────────────
    st.subheader("2. Long database (with treatment)")
    data_source, data_name = pick_or_upload(
        DATA_HARMONIZATION_DIR, "Long database with treatment assigned (from Step 3)", "reg_data",
        default_path=os.path.join(DATA_HARMONIZATION_DIR, DEFAULT_LONG_TREATMENT_FILENAME),
    )
    if data_source is None:
        st.info("Select or upload the long database to continue.")
        return
    long_df = read_table(data_source)
    st.success(f"Loaded `{data_name}` — {long_df.shape[0]} rows, {long_df.shape[1]} columns.")

    missing_cols = [c for c in (ID_VAR, PERIOD_VAR, TREATMENT_VAR, TREATMENT_LABEL) if c not in long_df.columns]
    if missing_cols:
        st.error(f"This database is missing {missing_cols} — treatment must be assigned (Step 3) first.")
        return

    st.markdown("---")

    # ── 3. Wide (respondent-level) database ──────────────────────────────────
    st.subheader("3. Respondent-level wide database")
    wide_df, catalog, notes = build_wide(
        long_df, dict_df.copy(), ID_VAR, PERIOD_VAR, BASELINE_VALUE, ENDLINE_VALUE, TREATMENT_VAR, TREATMENT_LABEL
    )
    outcome_options, control_options, fe_options, cluster_options = build_variable_catalog(dict_df, catalog)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Respondents", len(wide_df))
    c2.metric("Outcomes", len(outcome_options))
    c3.metric("Controls", len(control_options))
    c4.metric("FE variables", len(fe_options))
    c5.metric("Cluster variables", len(cluster_options))

    if notes:
        with st.expander(f"ℹ️ {len(notes)} note(s) from the reshape", expanded=False):
            for n in notes:
                st.caption(f"- {n}")

    with st.expander("Preview wide database", expanded=False):
        st.dataframe(wide_df.head(30), width="stretch")

    if not outcome_options:
        st.warning("No eligible outcome variables found — flag at least one variable `outcome_var = 1` in Step 1.")
        return

    arms, default_ref = detect_arms(wide_df, TREATMENT_VAR, TREATMENT_LABEL)
    if not arms:
        st.error("No treatment arms found — `treatment_label` is empty for every respondent.")
        return
    dummy_cols = build_treatment_dummies(wide_df, arms, TREATMENT_LABEL)
    st.caption(f"Treatment arms detected: {', '.join(arms)}")

    # Map every raw wide-database column back to a friendly display label, for
    # specification previews and the exported specification-definitions table.
    col_to_display = {}
    for opts in (outcome_options, control_options, fe_options, cluster_options):
        for o in opts:
            col_to_display[o["col"]] = o["display"]
    for arm, c in dummy_cols.items():
        col_to_display[c] = f"Treatment: {arm}"
    col_to_display[TREATMENT_VAR] = "Treatment (numeric)"

    st.markdown("---")

    # ── 4. Build specifications ──────────────────────────────────────────────
    st.subheader("4. Build a specification")
    if "reg_specs" not in st.session_state:
        st.session_state["reg_specs"] = []
    gen = st.session_state.get("reg_spec_gen", 0)

    default_name = f"Spec {len(st.session_state['reg_specs']) + 1}"
    name = st.text_input("Specification name", value=default_name, key=f"spec_name_{gen}")

    treat_mode_choice = st.radio(
        "Treatment coding", ["Categorical (dummy per arm vs. a reference)", "Numeric (treatment variable as provided)"],
        key=f"spec_treatmode_{gen}",
    )
    treatment_mode = "categorical" if treat_mode_choice.startswith("Categorical") else "numeric"

    if treatment_mode == "categorical":
        ref_idx = arms.index(default_ref) if default_ref in arms else 0
        reference_level = st.selectbox("Reference / control group", arms, index=ref_idx, key=f"spec_ref_{gen}")
        arm_choices = [a for a in arms if a != reference_level]
        treatment_arms = st.multiselect(
            "Treatment arm(s) to include (compared against the reference)",
            arm_choices, default=arm_choices, key=f"spec_arms_{gen}",
        )
    else:
        reference_level, treatment_arms = None, []
        st.caption(f"Uses `{TREATMENT_VAR}` directly as a continuous regressor.")

    include_baseline_outcome = st.checkbox(
        "Include the baseline value of the outcome as a control (ANCOVA-style)", value=True, key=f"spec_blout_{gen}",
    )

    control_display_options = [o["display"] for o in control_options]
    control_display_to_col = {o["display"]: o["col"] for o in control_options}
    controls_display = st.multiselect("Baseline controls", control_display_options, key=f"spec_controls_{gen}")

    interact_display = st.multiselect(
        "Interact treatment x control (auto-added as a control if not already selected above)",
        control_display_options, key=f"spec_interact_{gen}",
    )

    fe_display_options = [o["display"] for o in fe_options]
    fe_display_to_col = {o["display"]: o["col"] for o in fe_options}
    fe_display = st.multiselect("Fixed effects", fe_display_options, key=f"spec_fe_{gen}")

    cluster_display_options = [o["display"] for o in cluster_options]
    cluster_display_to_col = {o["display"]: o["col"] for o in cluster_options}
    se_choice = st.selectbox(
        "Standard errors",
        ["Classical (homoskedastic)", "Robust (HC1)"] + [f"Clustered: {d}" for d in cluster_display_options],
        key=f"spec_se_{gen}",
    )
    if se_choice.startswith("Classical"):
        se_type, cluster_col = "classical", None
    elif se_choice.startswith("Robust"):
        se_type, cluster_col = "robust", None
    else:
        se_type = "cluster"
        cluster_col = cluster_display_to_col[se_choice.replace("Clustered: ", "", 1)]

    outcome_display_options = [o["display"] for o in outcome_options]
    outcome_display_to_opt = {o["display"]: o for o in outcome_options}
    outcomes_for_spec = st.multiselect(
        "Run this specification for outcome(s)", outcome_display_options,
        default=outcome_display_options, key=f"spec_outcomes_{gen}",
    )

    if st.button("➕ Add specification", type="primary"):
        if not name.strip():
            st.error("Give the specification a name before adding it.")
        elif treatment_mode == "categorical" and not treatment_arms:
            st.error("Select at least one treatment arm to include.")
        elif not outcomes_for_spec:
            st.error("Select at least one outcome to run this specification on.")
        else:
            control_cols = [control_display_to_col[d] for d in controls_display]
            interact_cols = [control_display_to_col[d] for d in interact_display]
            control_cols = list(dict.fromkeys(control_cols + interact_cols))  # auto-include interacted controls
            spec = {
                "id": uuid.uuid4().hex[:8],
                "name": name.strip(),
                "treatment_mode": treatment_mode,
                "reference_level": reference_level,
                "treatment_arms": treatment_arms,
                "include_baseline_outcome": include_baseline_outcome,
                "control_cols": control_cols,
                "interact_control_cols": interact_cols,
                "fe_cols": [fe_display_to_col[d] for d in fe_display],
                "se_type": se_type,
                "cluster_col": cluster_col,
                "outcome_displays": outcomes_for_spec,
                "outcome_opts": [outcome_display_to_opt[d] for d in outcomes_for_spec],
            }
            st.session_state["reg_specs"].append(spec)
            st.session_state["reg_spec_gen"] = gen + 1
            st.success(f"Added specification “{spec['name']}”.")
            st.rerun()

    specs = st.session_state["reg_specs"]
    st.markdown(f"#### Current specifications ({len(specs)})")
    if not specs:
        st.info("No specifications yet — build one above and click **Add specification**.")
    else:
        for i, spec in enumerate(specs):
            treat_desc = (f"{spec['reference_level']} (ref.) vs " + ", ".join(spec["treatment_arms"])
                          if spec["treatment_mode"] == "categorical" else "Numeric treatment variable")
            with st.expander(f"**{spec['name']}** — {treat_desc}", expanded=False):
                st.markdown(f"- **Treatment:** {treat_desc}")
                st.markdown(f"- **Baseline outcome control:** {'Yes' if spec['include_baseline_outcome'] else 'No'}")
                st.markdown(f"- **Controls:** {', '.join(col_to_display.get(c, c) for c in spec['control_cols']) or '—'}")
                st.markdown(f"- **Treatment × control interactions:** "
                            f"{', '.join(col_to_display.get(c, c) for c in spec['interact_control_cols']) or '—'}")
                st.markdown(f"- **Fixed effects:** {', '.join(col_to_display.get(c, c) for c in spec['fe_cols']) or '—'}")
                se_label = (f"Clustered ({col_to_display.get(spec['cluster_col'], spec['cluster_col'])})"
                            if spec["se_type"] == "cluster" else
                            "Robust (HC1)" if spec["se_type"] == "robust" else "Classical")
                st.markdown(f"- **Standard errors:** {se_label}")
                st.markdown(f"- **Outcomes:** {', '.join(spec['outcome_displays'])}")
                if st.button("🗑️ Remove this specification", key=f"del_spec_{spec['id']}"):
                    st.session_state["reg_specs"] = [s for s in specs if s["id"] != spec["id"]]
                    st.rerun()

    st.markdown("---")

    # ── 5. Select & run ───────────────────────────────────────────────────────
    st.subheader("5. Select specifications to estimate")
    if not specs:
        st.info("Add at least one specification above to continue.")
        return

    combos = []
    for spec in specs:
        for disp, opt in zip(spec["outcome_displays"], spec["outcome_opts"]):
            _, data_preview, _, caveats = build_spec_inputs(spec, opt, wide_df, dummy_cols)
            combos.append({
                "Run": True, "Specification": spec["name"], "Outcome": disp,
                "N (available)": len(data_preview), "Notes": "; ".join(caveats) if caveats else "",
                "_spec_id": spec["id"], "_outcome_col": opt["col"],
            })
    combos_df = pd.DataFrame(combos)

    edited = st.data_editor(
        combos_df.drop(columns=["_spec_id", "_outcome_col"]),
        column_config={"Run": st.column_config.CheckboxColumn("Run", default=True)},
        disabled=["Specification", "Outcome", "N (available)", "Notes"],
        hide_index=True, width="stretch", key="reg_combo_editor",
    )
    run_mask = edited["Run"].tolist()

    if st.button("▶️ Run selected regressions", type="primary"):
        results = []
        errors = []
        spec_by_id = {s["id"]: s for s in specs}
        for keep, combo in zip(run_mask, combos):
            if not keep:
                continue
            spec = spec_by_id[combo["_spec_id"]]
            outcome_opt = next(o for o in spec["outcome_opts"] if o["col"] == combo["_outcome_col"])
            try:
                res, formula, cluster_col, caveats = fit_spec(spec, outcome_opt, wide_df, dummy_cols)
                results.append({
                    "spec": spec, "outcome_opt": outcome_opt, "result": res,
                    "formula": formula, "outcome_display": combo["Outcome"],
                })
            except Exception as e:
                errors.append(f"**{spec['name']}** × {combo['Outcome']}: {e}")
        st.session_state["reg_results"] = results
        st.session_state["reg_errors"] = errors

    results = st.session_state.get("reg_results", [])
    errors = st.session_state.get("reg_errors", [])

    if errors:
        with st.expander(f"⚠️ {len(errors)} regression(s) failed to estimate", expanded=True):
            for e in errors:
                st.error(e)

    if not results:
        st.info("Select the specifications to estimate above, then click **Run selected regressions**.")
        return

    st.markdown("---")

    # ── 6. Results ────────────────────────────────────────────────────────────
    st.subheader("6. Results")
    st.markdown(f"<style>{REPORT_CSS}</style>", unsafe_allow_html=True)

    by_outcome = {}
    for r in results:
        by_outcome.setdefault(r["outcome_display"], []).append(r)

    outcome_tables_html = []
    outcome_tables_tex = []
    for outcome_display, group in by_outcome.items():
        model_results = [g["result"] for g in group]
        model_names = [g["spec"]["name"] for g in group]
        fe_for_group = sorted({fe for g in group for fe in g["spec"]["fe_cols"]})
        tab = build_comparison_table(model_results, model_names, fe_for_group, col_to_display)

        st.markdown(f"**{outcome_display}**")
        st.markdown(tab.as_html(), unsafe_allow_html=True)
        with st.expander("Formulas used", expanded=False):
            for g in group:
                st.code(f"{g['spec']['name']}: {g['formula']}", language="text")

        outcome_tables_html.append((outcome_display, tab.as_html()))
        outcome_tables_tex.append((outcome_display, tab.as_latex()))
        st.markdown("---")

    # ── 7. Export ─────────────────────────────────────────────────────────────
    st.subheader("7. Export")
    spec_defs = spec_definitions_df(specs, col_to_display)
    with st.expander("Specification definitions", expanded=False):
        st.dataframe(spec_defs, hide_index=True, width="stretch")

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    html_report = build_html_report(PROJECT_NAME, generated_at, spec_defs, outcome_tables_html)
    tex_report = build_latex_report(PROJECT_NAME, generated_at, spec_defs, outcome_tables_tex)

    default_export_both = str(DEFAULT_OUTPUT_FORMAT).lower() in ("both", "latex")

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("💾 Save HTML + TEX to project (outputs_regressions folder)", type="primary"):
            os.makedirs(OUTPUTS_REGRESSIONS_DIR, exist_ok=True)
            html_path = os.path.join(OUTPUTS_REGRESSIONS_DIR, "regression_results.html")
            tex_path = os.path.join(OUTPUTS_REGRESSIONS_DIR, "regression_results.tex")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html_report)
            with open(tex_path, "w", encoding="utf-8") as f:
                f.write(tex_report)
            spec_csv_path = os.path.join(OUTPUTS_REGRESSIONS_DIR, "specification_definitions.csv")
            spec_defs.to_csv(spec_csv_path, index=False, encoding="utf-8-sig")
            st.success(f"Saved regression_results.html, regression_results.tex and specification_definitions.csv "
                       f"to `{os.path.abspath(OUTPUTS_REGRESSIONS_DIR)}`.")
    with c2:
        st.download_button("⬇️ Download HTML report", data=html_report.encode("utf-8"),
                            file_name="regression_results.html", mime="text/html")
    with c3:
        st.download_button("⬇️ Download LaTeX (.tex)", data=tex_report.encode("utf-8"),
                            file_name="regression_results.tex", mime="application/x-tex")

    with st.expander("Preview full report (HTML)", expanded=False):
        st.iframe(html_report, height=1200, width="stretch")


if __name__ == "__main__":
    render(standalone=True)