"""
AGEVAL Harmonization - Step 4 - Baseline/Endline Quality Check
Standalone run: python -m streamlit run harmonization/01_code/s04_quality_check.py
Also imported as a page by the unified harmonization/01_code/app.py

Loads the long database (with treatment assigned, from Step 3 — or the
plain long database from Step 2 if treatment hasn't been assigned yet)
and checks the STRUCTURE of the baseline/endline panel and the treatment
assignment:

    * Respondents matched in both rounds, baseline-only (attrition), and
      endline-only (new/unmatched), with counts and percentages.
    * Duplicate IDs within each round.
    * Treatment group sizes by period, and whether treatment is complete
      and consistent (same value in baseline and endline) for every
      respondent.
    * Attrition by treatment arm (differential attrition).

This does not check individual variables (ranges, outliers, missingness)
— that's a different kind of quality check on the raw survey data. This
module only assesses the panel/treatment structure of the harmonized
long database. Results can be exported as a standalone HTML report,
styled after the existing AGEVAL quality report.
"""

import streamlit as st
import pandas as pd
import numpy as np
import os
import io
import glob
import base64
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
DATA_HARMONIZATION_DIR = os.path.join(BASE_PATH, config["paths"]["data_harmonization"])
OUTPUTS_QUALITY_DIR = os.path.join(BASE_PATH, config["paths"]["outputs_quality"])

# Harmonization settings (all overridable in s00_harmonizationConfig.yaml)
ID_VAR = harmonize_cfg.get("id_variable", "respondent_id")
PERIOD_VAR = harmonize_cfg.get("period_variable", "period")
BASELINE_VALUE = harmonize_cfg.get("baseline_period_value", 0)
ENDLINE_VALUE = harmonize_cfg.get("endline_period_value", 1)
TREATMENT_VAR = harmonize_cfg.get("treatment_var", "treatment")
TREATMENT_LABEL = harmonize_cfg.get("treatment_label", "treatment_label")
DEFAULT_LONG_TREATMENT_FILENAME = harmonize_cfg.get("long_treatment_filename", "data_long_treatment.csv")
DEFAULT_LONG_FILENAME = harmonize_cfg.get("long_data_filename", "data_long.csv")
PROJECT_NAME = project_cfg.get("name", "AGEVAL")

# Validated categorical palette (fixed order — never reassigned per filter)
CATEGORICAL = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"]
COLOR_OK, COLOR_WARN, COLOR_ERR = "#16A34A", "#D97706", "#DC2626"


# ── I/O helpers ──────────────────────────────────────────────────────────────
def list_data_files(folder, patterns=("*.csv", "*.xlsx", "*.xls")):
    if not folder or not os.path.isdir(folder):
        return []
    files = []
    for pattern in patterns:
        files.extend(glob.glob(os.path.join(folder, pattern)))
    return sorted(files)


def read_table(source):
    if source is None:
        return None
    name = getattr(source, "name", source)
    ext = os.path.splitext(str(name))[1].lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(source)
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


# ── Core checks ──────────────────────────────────────────────────────────────
def panel_structure(df, id_var, period_var, baseline_value, endline_value):
    ids_baseline = set(df.loc[df[period_var] == baseline_value, id_var])
    ids_endline = set(df.loc[df[period_var] == endline_value, id_var])
    matched = ids_baseline & ids_endline
    baseline_only = ids_baseline - ids_endline
    endline_only = ids_endline - ids_baseline
    total_ids = ids_baseline | ids_endline

    def pct(n, d):
        return (n / d * 100) if d else 0.0

    return {
        "ids_baseline": ids_baseline, "ids_endline": ids_endline,
        "matched": matched, "baseline_only": baseline_only, "endline_only": endline_only,
        "n_total_ids": len(total_ids), "n_baseline": len(ids_baseline), "n_endline": len(ids_endline),
        "n_matched": len(matched), "n_baseline_only": len(baseline_only), "n_endline_only": len(endline_only),
        "pct_matched": pct(len(matched), len(total_ids)),
        "pct_baseline_only": pct(len(baseline_only), len(total_ids)),
        "pct_endline_only": pct(len(endline_only), len(total_ids)),
        "attrition_rate": pct(len(baseline_only), len(ids_baseline)),
        "new_entrant_rate": pct(len(endline_only), len(ids_endline)),
    }


def duplicates_by_round(df, id_var, period_var, baseline_value, endline_value):
    out = {}
    for label, value in [("baseline", baseline_value), ("endline", endline_value)]:
        sub = df[df[period_var] == value]
        counts = sub.groupby(id_var).size().rename("n_rows").reset_index()
        dup = counts[counts["n_rows"] > 1].sort_values("n_rows", ascending=False)
        out[label] = {
            "n_rows": len(sub), "n_unique_ids": sub[id_var].nunique(),
            "n_duplicate_ids": len(dup), "n_duplicate_rows": int(dup["n_rows"].sum()) if len(dup) else 0,
            "detail": dup.reset_index(drop=True),
        }
    return out


def treatment_crosstab(df, period_var, treatment_label, baseline_value, endline_value):
    labels = df[treatment_label].fillna("(missing)")
    ct = pd.crosstab(df[period_var], labels)
    ct_pct = pd.crosstab(df[period_var], labels, normalize="index") * 100
    order = [v for v in [baseline_value, endline_value] if v in ct.index]
    return ct.loc[order], ct_pct.loc[order]


def treatment_consistency(df, id_var, period_var, treatment_label, baseline_value, endline_value):
    pivot = df.pivot_table(index=id_var, columns=period_var, values=treatment_label, aggfunc="first")
    if baseline_value not in pivot.columns or endline_value not in pivot.columns:
        return pd.DataFrame(columns=[id_var, "baseline", "endline"]), 0, 0
    both = pivot.dropna(subset=[baseline_value, endline_value])
    mismatches = both[both[baseline_value] != both[endline_value]].reset_index()
    mismatches = mismatches.rename(columns={baseline_value: "baseline", endline_value: "endline"})
    return mismatches[[id_var, "baseline", "endline"]], len(both), len(mismatches)


def treatment_coverage(df, period_var, treatment_label, baseline_value, endline_value):
    rows = []
    for label, value in [("baseline", baseline_value), ("endline", endline_value)]:
        sub = df[df[period_var] == value]
        n = len(sub)
        n_missing = int(sub[treatment_label].isna().sum())
        rows.append({"round": label, "n_obs": n, "n_missing_treatment": n_missing,
                     "pct_missing": (n_missing / n * 100) if n else 0.0})
    return pd.DataFrame(rows)


def attrition_by_treatment(df, id_var, period_var, treatment_label, baseline_value, endline_value):
    base = df.loc[df[period_var] == baseline_value, [id_var, treatment_label]].drop_duplicates(subset=[id_var]).copy()
    base[treatment_label] = base[treatment_label].fillna("(missing)")
    end_ids = set(df.loc[df[period_var] == endline_value, id_var])
    base["retained"] = base[id_var].isin(end_ids)
    grp = base.groupby(treatment_label, dropna=False)["retained"].agg(n_baseline="count", n_retained="sum").reset_index()
    grp["n_attrited"] = grp["n_baseline"] - grp["n_retained"]
    grp["attrition_pct"] = np.where(grp["n_baseline"] > 0, grp["n_attrited"] / grp["n_baseline"] * 100, 0.0)
    return grp


def status_of(value, warn_at, err_at, higher_is_worse=True):
    if higher_is_worse:
        if value >= err_at:
            return "ERR"
        if value >= warn_at:
            return "WARN"
        return "OK"
    else:
        if value <= err_at:
            return "ERR"
        if value <= warn_at:
            return "WARN"
        return "OK"


# ── Charts (validated categorical palette, one axis, direct labels) ──────────
def fig_to_base64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")


def chart_panel_structure(structure):
    labels = ["Matched\n(both rounds)", "Baseline only\n(attrition)", "Endline only\n(new)"]
    values = [structure["n_matched"], structure["n_baseline_only"], structure["n_endline_only"]]
    colors = CATEGORICAL[:3]
    fig, ax = plt.subplots(figsize=(6, 3.2))
    bars = ax.bar(labels, values, color=colors, width=0.55)
    for b, v in zip(bars, values):
        ax.text(b.get_x() + b.get_width() / 2, v, f"{v}", ha="center", va="bottom", fontsize=10, fontweight="bold")
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.get_yaxis().set_visible(False)
    ax.set_title("Respondents by panel category", fontsize=11, loc="left", color="#0F172A")
    ax.tick_params(axis="x", labelsize=9)
    fig.tight_layout()
    return fig


def chart_treatment_balance(ct, baseline_value, endline_value):
    arms = list(ct.columns)
    if not arms:
        return None
    colors = (CATEGORICAL * ((len(arms) // len(CATEGORICAL)) + 1))[:len(arms)]
    periods = [p for p in [baseline_value, endline_value] if p in ct.index]
    n = len(periods)
    width = 0.8 / max(len(arms), 1)
    x = np.arange(n)
    fig, ax = plt.subplots(figsize=(6, 3.2))
    for i, arm in enumerate(arms):
        vals = [ct.loc[p, arm] if p in ct.index else 0 for p in periods]
        ax.bar(x + i * width, vals, width=width, label=str(arm), color=colors[i])
    ax.set_xticks(x + width * (len(arms) - 1) / 2)
    ax.set_xticklabels(["Baseline" if p == baseline_value else "Endline" for p in periods])
    ax.legend(frameon=False, fontsize=8, loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=min(len(arms), 4))
    ax.spines[["top", "right"]].set_visible(False)
    ax.set_title("Treatment group sizes by round", fontsize=11, loc="left", color="#0F172A")
    fig.tight_layout()
    return fig


def chart_attrition_by_treatment(attr_df, treatment_label_col):
    if attr_df.empty:
        return None
    labels = attr_df[treatment_label_col].astype(str).tolist()
    values = attr_df["attrition_pct"].tolist()
    colors = (CATEGORICAL * ((len(labels) // len(CATEGORICAL)) + 1))[:len(labels)]
    fig, ax = plt.subplots(figsize=(6, 3.2))
    bars = ax.bar(labels, values, color=colors, width=0.55)
    for b, v in zip(bars, values):
        ax.text(b.get_x() + b.get_width() / 2, v, f"{v:.1f}%", ha="center", va="bottom", fontsize=9)
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.get_yaxis().set_visible(False)
    ax.set_title("Attrition rate by treatment arm", fontsize=11, loc="left", color="#0F172A")
    ax.tick_params(axis="x", labelsize=9)
    fig.tight_layout()
    return fig


# ── HTML report (styled after the existing AGEVAL quality report) ───────────
REPORT_CSS = """
  body { font-family: -apple-system, sans-serif; max-width: 1100px; margin: 2rem auto; color: #1a1a1a; }
  h1   { font-size: 1.5rem; font-weight: 600; border-bottom: 2px solid #2563EB; padding-bottom: .5rem; }
  h2   { font-size: 1.3rem; font-weight: 700; margin-top: 2.5rem; padding-top: 1rem; border-top: 1px solid #E2E8F0; color: #0F172A; }
  h2 .sample { font-size: .8rem; font-weight: 400; color: #64748B; }
  h3   { font-size: 1.1rem; font-weight: 600; margin-top: 1.5rem; color: #1D4ED8; }
  .summary-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; margin: 1.5rem 0; }
  .card { background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 1rem; text-align: center; }
  .card .val { font-size: 2rem; font-weight: 700; }
  .card .lbl { font-size: .75rem; color: #64748B; margin-top: .25rem; }
  .ok   { color: #16A34A; }
  .warn { color: #D97706; }
  .err  { color: #DC2626; }
  table { width: 100%; border-collapse: collapse; font-size: .85rem; margin-top: .75rem; }
  th    { background: #1D4ED8; color: white; padding: .5rem .75rem; text-align: left; }
  td    { padding: .4rem .75rem; border-bottom: 1px solid #E2E8F0; }
  tr:nth-child(even) { background: #F8FAFC; }
  .badge { display: inline-block; padding: 2px 8px; border-radius: 9999px; font-size: .75rem; font-weight: 600; }
  .badge-err  { background: #FEE2E2; color: #991B1B; }
  .badge-warn { background: #FEF3C7; color: #92400E; }
  .badge-ok   { background: #D1FAE5; color: #065F46; }
  .chart-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 1.5rem; margin-top: 1rem; }
  .chart-grid img { width: 100%; height: auto; border: 1px solid #E2E8F0; border-radius: 8px; background: white; }
  footer { margin-top: 3rem; font-size: .75rem; color: #94A3B8; }
"""


def esc(x):
    return html_lib.escape(str(x))


def badge_html(status):
    cls = {"OK": "badge-ok", "WARN": "badge-warn", "ERR": "badge-err"}[status]
    return f'<span class="badge {cls}">{status}</span>'


def rows_html(rows):
    """rows: list of (check_label, status, detail_str)"""
    out = []
    for label, status, detail in rows:
        out.append(f"<tr><td>{esc(label)}</td><td>{badge_html(status)}</td><td>{esc(detail)}</td></tr>")
    return "\n".join(out)


def df_table_html(df, max_rows=200):
    if df is None or df.empty:
        return "<p style='color:#64748B; font-size:.85rem;'>No records.</p>"
    df = df.head(max_rows)
    header = "".join(f"<th>{esc(c)}</th>" for c in df.columns)
    body = []
    for _, r in df.iterrows():
        cells = "".join(f"<td>{esc(v)}</td>" for v in r.tolist())
        body.append(f"<tr>{cells}</tr>")
    note = f"<p style='color:#64748B; font-size:.75rem;'>Showing first {max_rows} rows.</p>" if len(df) == max_rows else ""
    return f"<table><tr>{header}</tr>{''.join(body)}</table>{note}"


def build_html_report(project_name, batch_name, generated_at, structure, dup, ct, ct_pct,
                       mismatches, n_both_periods, coverage_df, attrition_df,
                       treatment_label, id_var, has_treatment, badge_counts,
                       chart_panel_b64, chart_balance_b64, chart_attr_b64):
    n_warn, n_err = badge_counts["WARN"], badge_counts["ERR"]

    summary_cards = f"""
    <div class="summary-grid">
      <div class="card"><div class="val">{structure['n_total_ids']}</div><div class="lbl">Total respondents</div></div>
      <div class="card"><div class="val ok">{structure['pct_matched']:.1f}%</div><div class="lbl">Matched (both rounds)</div></div>
      <div class="card"><div class="val warn">{n_warn}</div><div class="lbl">Warnings</div></div>
      <div class="card"><div class="val err">{n_err}</div><div class="lbl">Errors</div></div>
    </div>
    """

    panel_rows = [
        ("Matched (baseline + endline)", "OK", f"{structure['n_matched']} respondents ({structure['pct_matched']:.1f}%)"),
        ("Baseline only (attrition)", status_of(structure["attrition_rate"], 10, 25),
         f"{structure['n_baseline_only']} respondents — {structure['attrition_rate']:.1f}% of baseline sample"),
        ("Endline only (new / unmatched)", status_of(structure["new_entrant_rate"], 10, 25),
         f"{structure['n_endline_only']} respondents — {structure['new_entrant_rate']:.1f}% of endline sample"),
    ]

    dup_rows = [
        (f"Duplicate {id_var} — baseline", "ERR" if dup["baseline"]["n_duplicate_ids"] > 0 else "OK",
         f"{dup['baseline']['n_duplicate_ids']} duplicated ID(s), {dup['baseline']['n_duplicate_rows']} rows"),
        (f"Duplicate {id_var} — endline", "ERR" if dup["endline"]["n_duplicate_ids"] > 0 else "OK",
         f"{dup['endline']['n_duplicate_ids']} duplicated ID(s), {dup['endline']['n_duplicate_rows']} rows"),
    ]

    charts_html = ""
    if chart_panel_b64:
        charts_html += f'<div class="chart-grid"><img src="data:image/png;base64,{chart_panel_b64}" alt="Panel structure chart"></div>'

    treatment_section = ""
    if has_treatment:
        cov_rows = []
        for _, r in coverage_df.iterrows():
            status = "ERR" if (r["round"] == "baseline" and r["n_missing_treatment"] > 0) else ("WARN" if r["n_missing_treatment"] > 0 else "OK")
            cov_rows.append((f"Treatment coverage — {r['round']}", status,
                              f"{r['n_missing_treatment']}/{r['n_obs']} missing ({r['pct_missing']:.1f}%)"))
        consistency_status = "ERR" if len(mismatches) > 0 else "OK"
        cov_rows.append(("Treatment consistency (baseline vs endline)", consistency_status,
                          f"{len(mismatches)}/{n_both_periods} matched respondents have a different treatment value between rounds"))

        balance_chart_html = (f'<img src="data:image/png;base64,{chart_balance_b64}" alt="Treatment balance chart">'
                               if chart_balance_b64 else "")
        attr_chart_html = (f'<img src="data:image/png;base64,{chart_attr_b64}" alt="Attrition by treatment arm chart">'
                            if chart_attr_b64 else "")

        treatment_section = f"""
        <h2>3. Treatment assignment</h2>
        <table>
          <tr><th>Check</th><th>Status</th><th>Detail</th></tr>
          {rows_html(cov_rows)}
        </table>

        <h3>Group sizes by round</h3>
        {df_table_html(ct.reset_index().rename(columns={ct.index.name or PERIOD_VAR: "period"}))}
        <h3>Group sizes by round (%)</h3>
        {df_table_html(ct_pct.round(1).reset_index().rename(columns={ct_pct.index.name or PERIOD_VAR: "period"}))}

        <div class="chart-grid">
          {balance_chart_html}
          {attr_chart_html}
        </div>

        <h3>Attrition by treatment arm</h3>
        {df_table_html(attrition_df.round(1))}

        <h3>Consistency mismatches</h3>
        {df_table_html(mismatches)}
        """
    else:
        treatment_section = """
        <h2>3. Treatment assignment</h2>
        <p style="color:#64748B;">No treatment columns found in this database — run Step 3 (Treatment Assignment) first.</p>
        """

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>{esc(project_name)} Harmonization Quality Report - {esc(batch_name)}</title>
<style>{REPORT_CSS}</style>
</head>
<body>
<h1>{esc(project_name)} — Harmonization Quality Report</h1>
<p>Database: <strong>{esc(batch_name)}</strong> &nbsp;|&nbsp; Generated: {esc(generated_at)} &nbsp;|&nbsp; Respondents checked: <strong>{structure['n_total_ids']}</strong></p>

{summary_cards}

<h2>1. Panel structure <span class="sample">— matched / attrition / new</span></h2>
<table>
  <tr><th>Check</th><th>Status</th><th>Detail</th></tr>
  {rows_html(panel_rows)}
</table>
{charts_html}

<h2>2. Duplicates</h2>
<table>
  <tr><th>Check</th><th>Status</th><th>Detail</th></tr>
  {rows_html(dup_rows)}
</table>
<h3>Duplicate IDs — baseline</h3>
{df_table_html(dup["baseline"]["detail"])}
<h3>Duplicate IDs — endline</h3>
{df_table_html(dup["endline"]["detail"])}

{treatment_section}

<footer>Generated by AGEVAL Harmonization v1.0 — {esc(project_name)}</footer>
</body>
</html>
"""
    return html


# ── Streamlit page ───────────────────────────────────────────────────────────
def render(standalone: bool = False):
    if standalone:
        st.set_page_config(page_title="AGEVAL Harmonization - Quality Check", layout="wide")

    st.title("AGEVAL Harmonization - Step 4 — Baseline/Endline Quality Check")
    st.caption(
        "Checks the STRUCTURE of the harmonized panel and treatment assignment — matched/attrition/new "
        "respondents, duplicate IDs, and treatment coverage/consistency. This does not check individual "
        "survey variables (ranges, outliers, missingness)."
    )

    # ── Input ──────────────────────────────────────────────────────────────
    st.subheader("1. Long database")
    default_path = os.path.join(DATA_HARMONIZATION_DIR, DEFAULT_LONG_TREATMENT_FILENAME)
    if not os.path.exists(default_path):
        default_path = os.path.join(DATA_HARMONIZATION_DIR, DEFAULT_LONG_FILENAME)
    source, name = pick_or_upload(
        DATA_HARMONIZATION_DIR, "Long database (with treatment, from Step 3 — or plain, from Step 2)",
        "qc_long_db", default_path=default_path,
    )
    if source is None:
        st.info("Select or upload the long database to continue.")
        return

    df = read_table(source)
    st.success(f"Loaded `{name}` — {df.shape[0]} rows, {df.shape[1]} columns.")

    if ID_VAR not in df.columns or PERIOD_VAR not in df.columns:
        st.error(f"This file is missing `{ID_VAR}` and/or `{PERIOD_VAR}` — it doesn't look like a harmonized long database.")
        return

    has_treatment = TREATMENT_VAR in df.columns and TREATMENT_LABEL in df.columns

    with st.sidebar.expander("Quality thresholds", expanded=False):
        attrition_warn = st.number_input("Attrition rate — warn at (%)", 0, 100, 10)
        attrition_err = st.number_input("Attrition rate — error at (%)", 0, 100, 25)

    st.markdown("---")

    # ── Panel structure ────────────────────────────────────────────────────
    st.subheader("2. Panel structure")
    structure = panel_structure(df, ID_VAR, PERIOD_VAR, BASELINE_VALUE, ENDLINE_VALUE)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total respondents", structure["n_total_ids"])
    c2.metric("Matched (both rounds)", f"{structure['n_matched']} ({structure['pct_matched']:.1f}%)")
    c3.metric("Baseline only (attrition)", f"{structure['n_baseline_only']} ({structure['attrition_rate']:.1f}%)")
    c4.metric("Endline only (new)", f"{structure['n_endline_only']} ({structure['new_entrant_rate']:.1f}%)")

    attrition_status = status_of(structure["attrition_rate"], attrition_warn, attrition_err)
    if attrition_status == "ERR":
        st.error(f"Attrition rate is {structure['attrition_rate']:.1f}% — above the {attrition_err}% error threshold.")
    elif attrition_status == "WARN":
        st.warning(f"Attrition rate is {structure['attrition_rate']:.1f}% — above the {attrition_warn}% warning threshold.")
    else:
        st.success(f"Attrition rate is {structure['attrition_rate']:.1f}% — within the expected range.")

    n_rows_expected = structure["n_matched"] * 2 + structure["n_baseline_only"] + structure["n_endline_only"]
    st.caption(
        f"Total rows in database: {len(df)}. Expected rows if every ID appears at most once per round: "
        f"{n_rows_expected}. {'Matches.' if len(df) == n_rows_expected else 'Difference is explained by duplicate IDs (see below).'}"
    )

    fig_panel = chart_panel_structure(structure)
    st.pyplot(fig_panel, width="content")
    chart_panel_b64 = fig_to_base64(chart_panel_structure(structure))

    st.markdown("---")

    # ── Duplicates ──────────────────────────────────────────────────────────
    st.subheader("3. Duplicates")
    dup = duplicates_by_round(df, ID_VAR, PERIOD_VAR, BASELINE_VALUE, ENDLINE_VALUE)
    d1, d2 = st.columns(2)
    for col, label in [(d1, "baseline"), (d2, "endline")]:
        info = dup[label]
        with col:
            st.markdown(f"**{label.capitalize()}**")
            if info["n_duplicate_ids"] > 0:
                st.error(f"{info['n_duplicate_ids']} duplicated ID(s) — {info['n_duplicate_rows']} rows affected.")
                st.dataframe(info["detail"], hide_index=True, width="stretch")
            else:
                st.success(f"No duplicate {ID_VAR} found in {label} ({info['n_unique_ids']} unique IDs, {info['n_rows']} rows).")

    st.markdown("---")

    # ── Treatment assignment ─────────────────────────────────────────────────
    st.subheader("4. Treatment assignment")
    ct = ct_pct = mismatches = coverage_df = attrition_df = None
    n_both_periods = 0
    chart_balance_b64 = chart_attr_b64 = None

    if not has_treatment:
        st.info(f"No `{TREATMENT_VAR}`/`{TREATMENT_LABEL}` columns found — run Step 3 (Treatment Assignment) first.")
    else:
        ct, ct_pct = treatment_crosstab(df, PERIOD_VAR, TREATMENT_LABEL, BASELINE_VALUE, ENDLINE_VALUE)
        coverage_df = treatment_coverage(df, PERIOD_VAR, TREATMENT_LABEL, BASELINE_VALUE, ENDLINE_VALUE)
        mismatches, n_both_periods, n_mismatches = treatment_consistency(
            df, ID_VAR, PERIOD_VAR, TREATMENT_LABEL, BASELINE_VALUE, ENDLINE_VALUE
        )
        attrition_df = attrition_by_treatment(df, ID_VAR, PERIOD_VAR, TREATMENT_LABEL, BASELINE_VALUE, ENDLINE_VALUE)

        st.markdown("**Group sizes by round**")
        g1, g2 = st.columns(2)
        g1.dataframe(ct, width="stretch")
        g2.dataframe(ct_pct.round(1), width="stretch")

        fig_balance = chart_treatment_balance(ct, BASELINE_VALUE, ENDLINE_VALUE)
        if fig_balance:
            st.pyplot(fig_balance, width="content")
            chart_balance_b64 = fig_to_base64(chart_treatment_balance(ct, BASELINE_VALUE, ENDLINE_VALUE))

        st.markdown("**Coverage — is treatment assigned to every observation?**")
        for _, r in coverage_df.iterrows():
            if r["n_missing_treatment"] == 0:
                st.success(f"{r['round'].capitalize()}: treatment assigned to all {int(r['n_obs'])} observations.")
            elif r["round"] == "baseline":
                st.error(f"{r['round'].capitalize()}: {int(r['n_missing_treatment'])}/{int(r['n_obs'])} observations "
                         f"({r['pct_missing']:.1f}%) missing treatment.")
            else:
                st.warning(f"{r['round'].capitalize()}: {int(r['n_missing_treatment'])}/{int(r['n_obs'])} observations "
                           f"({r['pct_missing']:.1f}%) missing treatment.")

        st.markdown("**Consistency — same treatment in baseline and endline?**")
        if n_mismatches == 0:
            st.success(f"All {n_both_periods} respondents observed in both rounds have a consistent treatment value.")
        else:
            st.error(f"{n_mismatches}/{n_both_periods} respondents have a different treatment value between baseline and endline.")
            st.dataframe(mismatches, hide_index=True, width="stretch")

        st.markdown("**Attrition by treatment arm**")
        st.caption("A large gap between arms here suggests differential attrition, which can bias impact estimates.")
        st.dataframe(attrition_df.round(1), hide_index=True, width="stretch")
        fig_attr = chart_attrition_by_treatment(attrition_df, TREATMENT_LABEL)
        if fig_attr:
            st.pyplot(fig_attr, width="content")
            chart_attr_b64 = fig_to_base64(chart_attrition_by_treatment(attrition_df, TREATMENT_LABEL))

    st.markdown("---")

    # ── Export ────────────────────────────────────────────────────────────
    st.subheader("5. Export report")

    badge_counts = {"OK": 0, "WARN": 0, "ERR": 0}
    badge_counts[attrition_status] += 1
    badge_counts["ERR" if dup["baseline"]["n_duplicate_ids"] > 0 else "OK"] += 1
    badge_counts["ERR" if dup["endline"]["n_duplicate_ids"] > 0 else "OK"] += 1
    if has_treatment:
        for _, r in coverage_df.iterrows():
            if r["n_missing_treatment"] > 0:
                badge_counts["ERR" if r["round"] == "baseline" else "WARN"] += 1
            else:
                badge_counts["OK"] += 1
        badge_counts["ERR" if len(mismatches) > 0 else "OK"] += 1

    batch_name = os.path.splitext(name)[0]
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    report_html = build_html_report(
        PROJECT_NAME, batch_name, generated_at, structure, dup, ct, ct_pct, mismatches, n_both_periods,
        coverage_df, attrition_df, TREATMENT_LABEL, ID_VAR, has_treatment, badge_counts,
        chart_panel_b64, chart_balance_b64, chart_attr_b64,
    )

    out_filename = st.text_input("Report filename", value=f"quality_report_{batch_name}.html")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("💾 Save report to project (outputs_quality folder)", type="primary"):
            os.makedirs(OUTPUTS_QUALITY_DIR, exist_ok=True)
            out_path = os.path.join(OUTPUTS_QUALITY_DIR, out_filename)
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(report_html)
            st.success(f"Saved report to `{os.path.abspath(out_path)}`.")
    with c2:
        st.download_button(
            label="⬇️ Download HTML report",
            data=report_html.encode("utf-8"),
            file_name=out_filename or f"quality_report_{batch_name}.html",
            mime="text/html",
        )

    with st.expander("Preview exported report", expanded=False):
        st.iframe(report_html, height=1200, width="stretch")


if __name__ == "__main__":
    render(standalone=True)
