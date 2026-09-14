"""
AGEVAL Harmonization - Step 5 - Baseline/Endline Comparison Plots
Standalone run: python -m streamlit run harmonization/01_code/s05_comparison_plots.py
Also imported as a page by the unified harmonization/01_code/app.py

Loads the harmonization dictionary (output of Step 1 — must include the
`surv_type`/`surv_choices` columns) and the long database with treatment
assigned (output of Step 3), and for every variable flagged `plots == 1`
in the dictionary generates a treatment x period comparison chart:

    * Numerical / dummy: bar of the mean by treatment x period, with a 95%
      CI and N per bar, plus an "Overall" bar per period pooling all
      treatment arms.
    * Categorical & select_one: one panel per period, bars per response
      option grouped by treatment, showing counts (not stacked, not %).
    * Categorical & select_multiple: same layout, counting how many
      observations selected each option (counts need not sum to N).

`var_type` alone isn't always reliable for telling select_one/select_multiple
apart from a plain numerical variable (the master dictionary can mistag a
select_multiple as `var_type=numerical`), so classification here prioritizes
`surv_type`/`surv_choices` when they're present — this is exactly why those
two columns were added to the dictionary.

A variable available in only one round still gets a plot, with the missing
round clearly marked instead of silently showing an empty/zero bar.
"""

import streamlit as st
import pandas as pd
import numpy as np
import os
import io
import re
import glob
import base64
import zipfile
import html as html_lib
import yaml
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# Read config
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "s00_harmonizationConfig.yaml")
with open(CONFIG_PATH, "r") as f:
    config = yaml.safe_load(f)

harmonize_cfg = config.get("harmonize", {})
project_cfg = config.get("project", {})

# Paths
BASE_PATH = os.path.normpath(os.path.join(os.path.dirname(CONFIG_PATH), config["paths"]["base"]))
MASTER_PATH = os.path.join(BASE_PATH, config["paths"]["dictionary_master"])
DICT_HARMONIZATION_PATH = os.path.join(BASE_PATH, config["paths"]["dictionary_harmonization"])
DATA_HARMONIZATION_DIR = os.path.join(BASE_PATH, config["paths"]["data_harmonization"])
OUTPUTS_PLOTS_DIR = os.path.join(BASE_PATH, config["paths"]["outputs_plots"])

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

# Validated categorical palette (fixed order — never reassigned per filter)
CATEGORICAL = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"]
OVERALL_COLOR = "#94A3B8"   # neutral gray — "Overall" is a pooled aggregate, not a real arm
MUTED_TEXT = "#64748B"


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


@st.cache_data
def load_sections():
    """Topic/subtopic display order & labels, from the master dictionary's
    'sections' sheet (same convention as Step 1)."""
    if not os.path.exists(MASTER_PATH):
        return {}, {}
    sections = pd.read_excel(MASTER_PATH, sheet_name="sections")
    topics = (sections[sections["level"] == "topic"].sort_values("order")
              .set_index("key").to_dict(orient="index"))
    subtopics = (sections[sections["level"] == "subtopic"].sort_values("order")
                 .set_index("key").to_dict(orient="index"))
    return topics, subtopics


def _ordered_keys(present_keys, section_info):
    present_keys = list(dict.fromkeys(present_keys))
    known = [k for k in present_keys if k in section_info]
    unknown = sorted(k for k in present_keys if k not in section_info)
    known.sort(key=lambda k: section_info[k]["order"])
    return known + unknown


def _section_label(key, section_info):
    info = section_info.get(key)
    if info is None:
        return key
    label = info.get("label_spanish") or info.get("label_english") or key
    return str(label).lstrip("#").strip()


# ── Dictionary parsing ────────────────────────────────────────────────────────
def parse_choices(choices_str):
    """'male:Hombre|female:Mujer' -> [('male','Hombre'), ('female','Mujer')].
    Robust to missing values, extra whitespace, and a missing ':' (falls
    back to using the raw token as both code and label)."""
    pairs = []
    if not isinstance(choices_str, str) or not choices_str.strip():
        return pairs
    for chunk in choices_str.split("|"):
        chunk = chunk.strip()
        if not chunk:
            continue
        if ":" in chunk:
            code, label = chunk.split(":", 1)
        else:
            code, label = chunk, chunk
        pairs.append((code.strip(), label.strip()))
    return pairs


def get_label(row, language=LANGUAGE):
    if language == "english":
        return row.get("label_english") or row.get("label_spanish") or row["variable_name"]
    return row.get("label_spanish") or row.get("label_english") or row["variable_name"]


def classify_variable(row):
    """Return one of 'numerical', 'dummy', 'select_one', 'select_multiple',
    or None (not plottable with these rules).

    `var_type == "dummy"` always wins first: it's a deliberate analyst
    choice to treat a variable as a 0/1 mean (e.g. a Yes/No question
    implemented as `select_one` in the ODK form is still conceptually a
    dummy, not a categorical distribution). Only once `dummy` is ruled out
    do `surv_type`/`surv_choices` get priority over `var_type` — this is
    what correctly reclassifies a true select_one/select_multiple that the
    master dictionary mistagged `var_type=numerical` (e.g. a select_multiple
    variable), without misfiring on a real dummy."""
    var_type = str(row.get("var_type") or "").strip()
    if var_type == "dummy":
        return "dummy"

    surv_type = str(row.get("surv_type") or "").strip()
    surv_choices = row.get("surv_choices")
    has_choices = isinstance(surv_choices, str) and surv_choices.strip() != ""
    if has_choices and surv_type in ("select_one", "select_multiple"):
        return surv_type
    if var_type == "numerical":
        return "numerical"
    if var_type == "categorical":
        # categorical without usable surv_type/choices metadata -> best-effort select_one
        return "select_one"
    return None


def selected_codes(raw_value):
    """Parse a select_multiple raw cell (space-separated choice codes, the
    standard ODK/XLSForm export convention) into a set of codes."""
    if pd.isna(raw_value):
        return set()
    text = str(raw_value).replace(",", " ").strip()
    if not text:
        return set()
    return set(text.split())


# ── Stats ──────────────────────────────────────────────────────────────────
def mean_ci(series):
    """Mean, 95% CI bounds, and N for a numeric series (NaNs dropped)."""
    vals = pd.to_numeric(series, errors="coerce").dropna()
    n = len(vals)
    if n == 0:
        return np.nan, np.nan, np.nan, 0
    mean = vals.mean()
    if n < 2:
        return mean, mean, mean, n
    sem = vals.std(ddof=1) / np.sqrt(n)
    return mean, mean - 1.96 * sem, mean + 1.96 * sem, n


def arm_order_and_colors(df, treatment_var, treatment_label):
    """Fixed, consistent order (and color) for treatment arms across every
    chart in the app — ordered by numeric treatment code when available."""
    sub = df[[treatment_var, treatment_label]].dropna(subset=[treatment_label]).drop_duplicates()
    if treatment_var in sub.columns and sub[treatment_var].notna().any():
        sub = sub.sort_values(treatment_var)
    else:
        sub = sub.sort_values(treatment_label)
    arms = sub[treatment_label].astype(str).tolist()
    colors = {arm: CATEGORICAL[i % len(CATEGORICAL)] for i, arm in enumerate(arms)}
    return arms, colors


def round_availability(df, var, period_var, baseline_value, endline_value):
    """Which rounds actually have non-null data for this variable."""
    avail = {}
    for label, value in [("baseline", baseline_value), ("endline", endline_value)]:
        sub = df.loc[df[period_var] == value, var]
        avail[label] = bool(sub.notna().any())
    return avail


def compute_numeric_summary(df, var, period_var, treatment_label, arms, baseline_value, endline_value):
    rows = []
    for round_label, value in [("baseline", baseline_value), ("endline", endline_value)]:
        sub = df[df[period_var] == value]
        if sub[var].notna().any():
            mean, lo, hi, n = mean_ci(sub[var])
            rows.append({"round": round_label, "group": "Overall", "mean": mean, "ci_low": lo, "ci_high": hi, "n": n})
            for arm in arms:
                arm_sub = sub.loc[sub[treatment_label] == arm, var]
                mean, lo, hi, n = mean_ci(arm_sub)
                rows.append({"round": round_label, "group": arm, "mean": mean, "ci_low": lo, "ci_high": hi, "n": n})
    return pd.DataFrame(rows)


def compute_categorical_counts(df, var, choices, period_var, treatment_label, arms,
                                baseline_value, endline_value, multiple=False):
    """Returns (counts_df[round, group(arm), choice_code, choice_label, count], totals{round: n})."""
    declared_codes = [c for c, _ in choices]
    label_map = dict(choices)

    rows = []
    totals = {}
    for round_label, value in [("baseline", baseline_value), ("endline", endline_value)]:
        sub = df[df[period_var] == value]
        if not sub[var].notna().any():
            continue
        totals[round_label] = int(sub[var].notna().sum())

        if multiple:
            observed_codes = set()
            per_row_codes = sub[var].apply(selected_codes)
            for codes in per_row_codes:
                observed_codes |= codes
        else:
            observed_codes = set(sub[var].dropna().astype(str).unique())

        all_codes = list(dict.fromkeys(declared_codes + sorted(observed_codes - set(declared_codes))))

        for arm in arms:
            arm_sub = sub[sub[treatment_label] == arm]
            if multiple:
                arm_codes_series = arm_sub[var].apply(selected_codes)
                for code in all_codes:
                    count = int(arm_codes_series.apply(lambda s, c=code: c in s).sum())
                    rows.append({"round": round_label, "group": arm, "choice_code": code,
                                 "choice_label": label_map.get(code, code), "count": count})
            else:
                vc = arm_sub[var].dropna().astype(str).value_counts()
                for code in all_codes:
                    rows.append({"round": round_label, "group": arm, "choice_code": code,
                                 "choice_label": label_map.get(code, code), "count": int(vc.get(code, 0))})

    return pd.DataFrame(rows), totals


# ── Charts (validated categorical palette, one axis, direct labels) ─────────
def fig_to_base64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")


def _title(fig, ax, varname, label, subtitle):
    fig.suptitle(f"{varname} — {label}", fontsize=12, fontweight="bold", x=0.01, ha="left", y=1.04)
    ax.set_title(subtitle, fontsize=9, color=MUTED_TEXT, loc="left")


def chart_numeric(summary_df, varname, label, arms, avail, subtitle):
    groups = ["Overall"] + list(arms)
    colors = {"Overall": OVERALL_COLOR, **{a: CATEGORICAL[i % len(CATEGORICAL)] for i, a in enumerate(arms)}}
    rounds = ["baseline", "endline"]
    width = 0.8 / len(groups)
    x = np.arange(len(rounds))

    fig, ax = plt.subplots(figsize=(7.5, 3.6))
    for i, g in enumerate(groups):
        means, los, his, ns, xs = [], [], [], [], []
        for j, r in enumerate(rounds):
            if not avail[r]:
                continue
            row = summary_df[(summary_df["round"] == r) & (summary_df["group"] == g)]
            if row.empty or pd.isna(row.iloc[0]["mean"]):
                continue
            r0 = row.iloc[0]
            means.append(r0["mean"]); los.append(r0["mean"] - r0["ci_low"]); his.append(r0["ci_high"] - r0["mean"])
            ns.append(int(r0["n"])); xs.append(x[j] + i * width)
        if not means:
            continue
        bars = ax.bar(xs, means, width=width * 0.92, color=colors[g], label=g,
                       yerr=[los, his], capsize=3, error_kw={"linewidth": 1, "ecolor": "#334155"})
        for xi, m, n in zip(xs, means, ns):
            ax.text(xi, m, f"{m:.2f}\nN={n}", ha="center", va="bottom", fontsize=7.5)

    for j, r in enumerate(rounds):
        if not avail[r]:
            ax.text(x[j] + width * (len(groups) - 1) / 2, 0.02, "Not available\nin this period",
                     ha="center", va="bottom", fontsize=9, color=MUTED_TEXT, style="italic",
                     transform=ax.get_xaxis_transform())

    ax.set_xticks(x + width * (len(groups) - 1) / 2)
    ax.set_xticklabels(["Baseline", "Endline"])
    ax.spines[["top", "right"]].set_visible(False)
    ax.legend(frameon=False, fontsize=8, loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=min(len(groups), 5))
    _title(fig, ax, varname, label, subtitle)
    fig.tight_layout()
    return fig


def chart_categorical(counts_df, totals, varname, label, arms, avail, subtitle):
    rounds = ["baseline", "endline"]
    fig, axes = plt.subplots(1, 2, figsize=(11, 4.2))
    colors = {a: CATEGORICAL[i % len(CATEGORICAL)] for i, a in enumerate(arms)}

    for ax, r in zip(axes, rounds):
        if not avail[r]:
            ax.text(0.5, 0.5, "Not available\nin this period", ha="center", va="center",
                     fontsize=11, color=MUTED_TEXT, style="italic", transform=ax.transAxes)
            ax.set_xticks([]); ax.set_yticks([])
            ax.spines[["top", "right", "left", "bottom"]].set_visible(False)
            ax.set_title(r.capitalize(), fontsize=10, loc="left")
            continue

        sub = counts_df[counts_df["round"] == r]
        choice_labels = list(dict.fromkeys(sub["choice_label"].tolist()))
        n_choices = len(choice_labels)
        width = 0.8 / max(len(arms), 1)
        xpos = np.arange(n_choices)
        for i, arm in enumerate(arms):
            arm_sub = sub[sub["group"] == arm].set_index("choice_label").reindex(choice_labels)
            vals = arm_sub["count"].fillna(0).tolist()
            bars = ax.bar(xpos + i * width, vals, width=width * 0.92, color=colors[arm], label=arm)
            for xi, v in zip(xpos + i * width, vals):
                if v > 0:
                    ax.text(xi, v, f"{int(v)}", ha="center", va="bottom", fontsize=7)

        ax.set_xticks(xpos + width * (len(arms) - 1) / 2)
        ax.set_xticklabels(choice_labels, rotation=20, ha="right", fontsize=8)
        ax.spines[["top", "right"]].set_visible(False)
        ax.set_title(f"{r.capitalize()} (N={totals.get(r, 0)})", fontsize=10, loc="left")

    handles, labels_ = axes[-1].get_legend_handles_labels()
    if not handles:
        handles, labels_ = axes[0].get_legend_handles_labels()
    fig.legend(handles, labels_, frameon=False, fontsize=8, loc="upper center",
               bbox_to_anchor=(0.5, -0.02), ncol=min(len(arms), 6))
    _title(fig, axes[0], varname, label, subtitle)
    fig.tight_layout()
    return fig


# ── HTML gallery export (styled consistently with the Step 4 report) ────────
REPORT_CSS = """
  body { font-family: -apple-system, sans-serif; max-width: 1100px; margin: 2rem auto; color: #1a1a1a; }
  h1   { font-size: 1.5rem; font-weight: 600; border-bottom: 2px solid #2563EB; padding-bottom: .5rem; }
  h2   { font-size: 1.2rem; font-weight: 700; margin-top: 2.5rem; padding-top: 1rem; border-top: 1px solid #E2E8F0; color: #0F172A; }
  h2 .sub { font-size: .8rem; font-weight: 400; color: #64748B; }
  .chart { margin-top: .75rem; }
  .chart img { width: 100%; max-width: 900px; height: auto; border: 1px solid #E2E8F0; border-radius: 8px; background: white; }
  footer { margin-top: 3rem; font-size: .75rem; color: #94A3B8; }
"""


def esc(x):
    return html_lib.escape(str(x))


def build_html_gallery(project_name, batch_name, generated_at, plots):
    """plots: list of (varname, label, subtitle, base64_png)"""
    sections = []
    for varname, label, subtitle, b64 in plots:
        sections.append(f"""
        <h2>{esc(varname)} <span class="sub">— {esc(label)}</span></h2>
        <p style="color:#64748B; font-size:.85rem; margin:0;">{esc(subtitle)}</p>
        <div class="chart"><img src="data:image/png;base64,{b64}" alt="{esc(varname)} chart"></div>
        """)
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>{esc(project_name)} Comparison Plots - {esc(batch_name)}</title>
<style>{REPORT_CSS}</style>
</head>
<body>
<h1>{esc(project_name)} — Baseline/Endline Comparison Plots</h1>
<p>Database: <strong>{esc(batch_name)}</strong> &nbsp;|&nbsp; Generated: {esc(generated_at)} &nbsp;|&nbsp; Plots: <strong>{len(plots)}</strong></p>
{''.join(sections)}
<footer>Generated by AGEVAL Harmonization v1.0 — {esc(project_name)}</footer>
</body>
</html>
"""


# ── Streamlit page ───────────────────────────────────────────────────────────
def render(standalone: bool = False):
    if standalone:
        st.set_page_config(page_title="AGEVAL Harmonization - Comparison Plots", layout="wide")

    st.title("AGEVAL Harmonization - Step 5 — Baseline/Endline Comparison Plots")
    st.caption(
        "Generates a treatment x period comparison chart for every variable flagged `plots = 1` in the "
        "harmonization dictionary. Numerical/dummy variables show mean + 95% CI; categorical variables "
        "show counts per response option, grouped by treatment, one panel per period."
    )

    # ── Inputs ────────────────────────────────────────────────────────────
    st.subheader("1. Harmonization dictionary")
    dict_source, dict_name = pick_or_upload(
        os.path.dirname(DICT_HARMONIZATION_PATH), "Harmonization dictionary (from Step 1)",
        "plots_dict", patterns=("*.csv", "*.xlsx"), upload_types=("csv", "xlsx"),
        default_path=DICT_HARMONIZATION_PATH,
    )
    if dict_source is None:
        st.info("Select or upload the harmonization dictionary to continue.")
        return
    dict_df = read_table(dict_source)
    if "variable_name" not in dict_df.columns or "plots" not in dict_df.columns:
        st.error(f"`{dict_name}` doesn't look like a harmonization dictionary — missing `variable_name`/`plots` columns.")
        return
    if "surv_type" not in dict_df.columns or "surv_choices" not in dict_df.columns:
        st.warning(
            "This dictionary has no `surv_type`/`surv_choices` columns — classification will fall back to "
            "`var_type` only, which can misclassify select_multiple variables. Re-run Step 1 with the updated "
            "master dictionary columns for best results."
        )
        dict_df["surv_type"] = dict_df.get("surv_type", pd.Series(dtype="object"))
        dict_df["surv_choices"] = dict_df.get("surv_choices", pd.Series(dtype="object"))
    st.success(f"Loaded `{dict_name}` — {len(dict_df)} variables.")

    st.markdown("---")

    st.subheader("2. Long database (with treatment)")
    data_source, data_name = pick_or_upload(
        DATA_HARMONIZATION_DIR, "Long database with treatment assigned (from Step 3)", "plots_data",
        default_path=os.path.join(DATA_HARMONIZATION_DIR, DEFAULT_LONG_TREATMENT_FILENAME),
    )
    if data_source is None:
        st.info("Select or upload the long database to continue.")
        return
    df = read_table(data_source)
    st.success(f"Loaded `{data_name}` — {df.shape[0]} rows, {df.shape[1]} columns.")

    missing_cols = [c for c in (ID_VAR, PERIOD_VAR, TREATMENT_VAR, TREATMENT_LABEL) if c not in df.columns]
    if missing_cols:
        st.error(
            f"This database is missing {missing_cols} — treatment must be assigned (Step 3) before generating "
            "comparison plots."
        )
        return

    arms, _ = arm_order_and_colors(df, TREATMENT_VAR, TREATMENT_LABEL)
    if not arms:
        st.error("No treatment arms found (treatment_label is empty for every observation).")
        return
    st.caption(f"Treatment arms detected (in plotting order): {', '.join(arms)}")

    st.markdown("---")

    # ── Variable selection (topic-grouped, like Step 1) ─────────────────────
    plot_vars = dict_df[dict_df["plots"] == 1].copy()
    plot_vars = plot_vars[plot_vars["variable_name"].isin(df.columns)]
    if plot_vars.empty:
        st.warning("No variables are flagged `plots = 1` in the dictionary (or none of them exist in this database).")
        return

    topic_info, subtopic_info = load_sections()
    TOPICS = _ordered_keys(plot_vars["topic"].dropna().unique().tolist(), topic_info) if "topic" in plot_vars.columns else []

    st.sidebar.header("Filter variables")
    if TOPICS:
        selected_topics = st.sidebar.multiselect(
            "Topic", TOPICS, default=TOPICS, format_func=lambda k: _section_label(k, topic_info),
        )
        selected_topics = _ordered_keys(selected_topics, topic_info)
        plot_vars = plot_vars[plot_vars["topic"].isin(selected_topics)]
    else:
        selected_topics = []
    search = st.sidebar.text_input("Search variable name or label")
    if search:
        mask = (
            plot_vars["variable_name"].str.contains(search, case=False, na=False) |
            plot_vars.get("label_spanish", pd.Series(dtype="object")).str.contains(search, case=False, na=False) |
            plot_vars.get("label_english", pd.Series(dtype="object")).str.contains(search, case=False, na=False)
        )
        plot_vars = plot_vars[mask]
    st.sidebar.markdown("---")
    st.sidebar.markdown(f"**{len(plot_vars)}** variables to plot")

    st.markdown("---")
    st.subheader("3. Plots")

    generated_plots = []   # (varname, label, subtitle, base64_png) for export
    skipped = []

    groups = [(t, plot_vars[plot_vars["topic"] == t]) if "topic" in plot_vars.columns else (None, plot_vars)
              for t in (selected_topics or [None])]

    for topic, topic_df in groups:
        if topic_df.empty:
            continue
        topic_label = _section_label(topic, topic_info) if topic else "Variables"
        with st.expander(f"**{str(topic_label).upper()}** — {len(topic_df)} variables", expanded=True):
            for _, row in topic_df.iterrows():
                var = row["variable_name"]
                label = get_label(row)
                kind = classify_variable(row)
                avail = round_availability(df, var, PERIOD_VAR, BASELINE_VALUE, ENDLINE_VALUE)

                if kind is None:
                    skipped.append((var, row.get("var_type"), "not a plottable type (numerical/dummy/categorical)"))
                    continue
                if not avail["baseline"] and not avail["endline"]:
                    skipped.append((var, row.get("var_type"), "no data in either round"))
                    continue

                if kind in ("numerical", "dummy"):
                    subtitle = "Mean by treatment × period (95% CI, N per bar)" if kind == "numerical" else \
                               "Share = 1 by treatment × period (95% CI, N per bar)"
                    summary = compute_numeric_summary(df, var, PERIOD_VAR, TREATMENT_LABEL, arms, BASELINE_VALUE, ENDLINE_VALUE)
                    fig = chart_numeric(summary, var, label, arms, avail, subtitle)
                else:
                    choices = parse_choices(row.get("surv_choices"))
                    multiple = kind == "select_multiple"
                    subtitle = ("Count of observations selecting each option, by treatment × period (not %, may not sum to N)"
                                if multiple else "Count of responses by option and treatment × period (not %)")
                    counts, totals = compute_categorical_counts(
                        df, var, choices, PERIOD_VAR, TREATMENT_LABEL, arms, BASELINE_VALUE, ENDLINE_VALUE, multiple=multiple
                    )
                    fig = chart_categorical(counts, totals, var, label, arms, avail, subtitle)

                if not avail["baseline"] or not avail["endline"]:
                    missing_round = "endline" if not avail["endline"] else "baseline"
                    st.info(f"`{var}` is not available in **{missing_round}** — shown for the other period only.")

                st.pyplot(fig, width="content")
                b64 = fig_to_base64(fig)
                generated_plots.append((var, label, subtitle, b64))

    if skipped:
        with st.expander(f"ℹ️ {len(skipped)} flagged variable(s) skipped", expanded=False):
            st.dataframe(pd.DataFrame(skipped, columns=["variable", "var_type", "reason"]), hide_index=True, width="stretch")

    st.markdown("---")

    # ── Export ────────────────────────────────────────────────────────────
    st.subheader("4. Export")
    if not generated_plots:
        st.info("No plots were generated for the current filters.")
        return

    batch_name = os.path.splitext(data_name)[0]
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    gallery_html = build_html_gallery(PROJECT_NAME, batch_name, generated_at, generated_plots)

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("💾 Save all plots to project (outputs_plots folder)", type="primary"):
            os.makedirs(OUTPUTS_PLOTS_DIR, exist_ok=True)
            for varname, _, _, b64 in generated_plots:
                safe_name = re.sub(r"[^A-Za-z0-9_\-]", "_", varname)
                with open(os.path.join(OUTPUTS_PLOTS_DIR, f"{safe_name}.png"), "wb") as f:
                    f.write(base64.b64decode(b64))
            gallery_path = os.path.join(OUTPUTS_PLOTS_DIR, "comparison_plots.html")
            with open(gallery_path, "w", encoding="utf-8") as f:
                f.write(gallery_html)
            st.success(f"Saved {len(generated_plots)} PNGs and the gallery report to `{os.path.abspath(OUTPUTS_PLOTS_DIR)}`.")
    with c2:
        st.download_button(
            "⬇️ Download gallery (HTML)", data=gallery_html.encode("utf-8"),
            file_name=f"comparison_plots_{batch_name}.html", mime="text/html",
        )
    with c3:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for varname, _, _, b64 in generated_plots:
                safe_name = re.sub(r"[^A-Za-z0-9_\-]", "_", varname)
                zf.writestr(f"{safe_name}.png", base64.b64decode(b64))
        st.download_button(
            "⬇️ Download all plots (ZIP of PNGs)", data=zip_buf.getvalue(),
            file_name=f"comparison_plots_{batch_name}.zip", mime="application/zip",
        )

    with st.expander("Preview gallery report", expanded=False):
        st.iframe(gallery_html, height=1200, width="stretch")


if __name__ == "__main__":
    render(standalone=True)
