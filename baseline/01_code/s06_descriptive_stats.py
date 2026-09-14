"""
AGEVAL Step 6 — Descriptive Statistics Generator
Run: python scripts/s06_descriptive_stats.py --data data_clean/cleaned_data.csv
Run: python baseline/01_code/s06_descriptive_stats.py --data baseline/04_data/dataClean/test_data_honduras_n2052_clean.xlsx 

Reads cleaned data and generates charts + summary HTML for all
variables flagged descriptive_include=1 in the dictionary.

Chart type is now derived directly from each variable's `var_type`:
  - numerical               -> histogram
  - categorical / dummy     -> bar chart

The report is organized using the `sections` sheet of variables_master.xlsx,
which defines the display order and labels for each topic and subtopic.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use("Agg")
import os
import yaml
import argparse
from datetime import datetime
from jinja2 import Template


# Read config
CFG = os.path.join(os.path.dirname(__file__), "s00_config.yaml")
with open(CFG, "r") as f:
    config = yaml.safe_load(f)
    cfg = yaml.safe_load(f)
SURVEY_ROUND = config["project"]["survey_round"]

# Paths
BASE       = os.path.normpath(os.path.join(os.path.dirname(CFG), config["paths"]["base"]))
MASTER     = os.path.join(BASE, config["paths"]["dictionary_master"])
DICT       = os.path.join(BASE, config["paths"]["dictionary_personalized"].format(survey_round=SURVEY_ROUND))
DATARAW   = os.path.join(BASE, config["paths"]["data_raw"].format(survey_round=SURVEY_ROUND))
CORRECTIONS = os.path.join(BASE, config["paths"]["correction_files"].format(survey_round=SURVEY_ROUND))
DATACLEAN   = os.path.join(BASE, config["paths"]["data_clean"].format(survey_round=SURVEY_ROUND))
OUTPLOTS = os.path.join(BASE, config["paths"]["outputs_plots"].format(survey_round=SURVEY_ROUND))
OUTSTATS = os.path.join(BASE, config["paths"]["outputs_stats"].format(survey_round=SURVEY_ROUND))


# ── Plot style ─────────────────────────────────────────────────────────────────
BLUE   = "#1D4ED8"
TEAL   = "#0F6E56"
CORAL  = "#D85A30"
GRAY   = "#94A3B8"
plt.rcParams.update({
    "figure.facecolor": "white",
    "axes.facecolor":   "white",
    "axes.spines.top":  False,
    "axes.spines.right":False,
    "axes.grid":        True,
    "grid.alpha":       0.3,
    "font.size":        10,
})

# ── Chart generators ───────────────────────────────────────────────────────────
def make_histogram(series, varname, label, out_path):
    fig, ax = plt.subplots(figsize=(6, 3.5))
    vals = pd.to_numeric(series, errors="coerce").dropna()
    ax.hist(vals, bins=20, color=BLUE, edgecolor="white", linewidth=0.5)
    ax.set_xlabel(label, fontsize=10)
    ax.set_ylabel("Frequency", fontsize=10)
    ax.set_title(f"{varname}\nn={len(vals)}, mean={vals.mean():.2f}, median={vals.median():.2f}", fontsize=10)
    fig.tight_layout()
    fig.savefig(out_path, dpi=120, bbox_inches="tight")
    plt.close(fig)

def make_bar(series, varname, label, out_path):
    fig, ax = plt.subplots(figsize=(6, 3.5))
    counts = series.dropna().value_counts().sort_index()
    colors = [BLUE, TEAL, CORAL, GRAY, "#7F77DD", "#D4537E"]
    bars = ax.bar(range(len(counts)), counts.values,
                  color=[colors[i % len(colors)] for i in range(len(counts))],
                  edgecolor="white", linewidth=0.5)
    ax.set_xticks(range(len(counts)))
    ax.set_xticklabels(counts.index, rotation=30, ha="right", fontsize=9)
    ax.set_ylabel("Count", fontsize=10)
    ax.set_title(f"{varname}\nn={len(series.dropna())}", fontsize=10)
    # Value labels on bars
    for bar in bars:
        h = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, h + 0.3, str(int(h)), ha="center", fontsize=8)
    fig.tight_layout()
    fig.savefig(out_path, dpi=120, bbox_inches="tight")
    plt.close(fig)


def chart_type_for(var_type):
    """Map a dictionary var_type to the chart function to use.

    numerical              -> histogram
    categorical / dummy    -> bar
    anything else          -> falls back to bar (safer default for
                               non-numeric data than a histogram)
    """
    vt = str(var_type).strip().lower()
    if vt == "numerical":
        return "histogram"
    if vt in ("categorical", "dummy"):
        return "bar"
    return "bar"


# ── Summary stats table ────────────────────────────────────────────────────────
def summary_stats(series, qtype):
    vals = pd.to_numeric(series, errors="coerce") if qtype in ("integer", "decimal") else series
    n     = series.notna().sum()
    miss  = series.isna().sum()
    stats = {"n": n, "missing": miss}

    if qtype in ("integer", "decimal"):
        numeric = vals.dropna()
        stats.update({
            "mean":   round(numeric.mean(), 3) if len(numeric) else "",
            "median": round(numeric.median(), 3) if len(numeric) else "",
            "sd":     round(numeric.std(), 3) if len(numeric) else "",
            "min":    round(numeric.min(), 3) if len(numeric) else "",
            "max":    round(numeric.max(), 3) if len(numeric) else "",
            "p25":    round(numeric.quantile(.25), 3) if len(numeric) else "",
            "p75":    round(numeric.quantile(.75), 3) if len(numeric) else "",
        })
    else:
        vc = series.dropna().value_counts()
        stats["mode"] = vc.index[0] if len(vc) else ""
        stats["n_unique"] = series.nunique()

    return stats


# ── Sections (topic / subtopic ordering & labels) ──────────────────────────────
def load_sections(sections_path):
    """Read the `sections` sheet and build lookup tables that let us:
       - order topics and subtopics for the report
       - display a human readable label instead of the raw key
       - know which topic a subtopic belongs to

    Expected columns: level ("topic"/"subtopic"), key, order,
    related_topic (topic *order*, for subtopic rows), label_spanish,
    label_english.
    """
    sections = pd.read_excel(sections_path, sheet_name="sections")

    topics_df    = sections[sections["level"] == "topic"].copy()
    subtopics_df = sections[sections["level"] == "subtopic"].copy()

    topic_order = {row.key: row.order for row in topics_df.itertuples()}
    topic_label = {row.key: row.label_spanish for row in topics_df.itertuples()}

    # related_topic stores the *order* of the parent topic, not its key
    order_to_topic_key = {row.order: row.key for row in topics_df.itertuples()}

    subtopic_order = {row.key: row.order for row in subtopics_df.itertuples()}
    subtopic_label = {row.key: row.label_spanish for row in subtopics_df.itertuples()}
    subtopic_topic = {
        row.key: order_to_topic_key.get(row.related_topic)
        for row in subtopics_df.itertuples()
    }

    return {
        "topic_order":     topic_order,
        "topic_label":     topic_label,
        "subtopic_order":  subtopic_order,
        "subtopic_label":  subtopic_label,
        "subtopic_topic":  subtopic_topic,
    }


def load_data(data_path):
    """Read the survey data regardless of whether it's CSV or Excel."""
    ext = os.path.splitext(data_path)[1].lower()
    if ext in (".xlsx", ".xls", ".xlsm"):
        return pd.read_excel(data_path)
    return pd.read_csv(data_path)


# ── HTML summary report ────────────────────────────────────────────────────────
HTML_TMPL = """
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>AGEVAL Descriptive Statistics</title>
<style>
  body  { font-family: -apple-system, sans-serif; max-width: 1100px; margin: 2rem auto; color: #1a1a1a; }
  h1    { font-size: 1.5rem; font-weight: 600; border-bottom: 2px solid #1D4ED8; padding-bottom:.5rem; }
  h2    { font-size: 1rem; font-weight: 600; margin: 2.5rem 0 .25rem; color: #1D4ED8; }
  .var  { display: grid; grid-template-columns: 1fr 1fr; gap: 1.5rem; align-items: start; margin-bottom: 2rem; }
  img   { width: 100%; border-radius: 6px; border: 1px solid #E2E8F0; }
  table { width: 100%; border-collapse: collapse; font-size: .8rem; }
  th    { background: #1D4ED8; color: white; padding: .4rem .6rem; text-align: left; }
  td    { padding: .35rem .6rem; border-bottom: 1px solid #E2E8F0; }
  tr:nth-child(even) { background: #F8FAFC; }
  .topic-header    { background: #EFF6FF; padding: .5rem 1rem; border-radius: 6px;
                      font-weight: 600; font-size: 1rem; color: #1E40AF; margin: 2.5rem 0 .5rem; }
  .subtopic-header { padding: .3rem .75rem; border-left: 3px solid #1D4ED8;
                      font-weight: 600; font-size: .85rem; color: #334155; margin: 1.25rem 0 .5rem; }
  footer { margin-top: 3rem; font-size: .75rem; color: #94A3B8; }
</style>
</head>
<body>
<h1>AGEVAL — Descriptive Statistics</h1>
<p>Dataset: <strong>{{ data_path }}</strong> &nbsp;|&nbsp; Generated: {{ generated }} &nbsp;|&nbsp; n = <strong>{{ n_records }}</strong></p>

{% set ns = namespace(current_topic=None, current_subtopic=None) %}
{% for var in variables %}
  {% if var.topic != ns.current_topic %}
    {% set ns.current_topic = var.topic %}
    {% set ns.current_subtopic = None %}
    <div class="topic-header">{{ var.topic_label }}</div>
  {% endif %}
  {% if var.subtopic != ns.current_subtopic %}
    {% set ns.current_subtopic = var.subtopic %}
    <div class="subtopic-header">{{ var.subtopic_label }}</div>
  {% endif %}
  <h2>{{ var.varname }} &nbsp;<small style="font-weight:400;color:#64748B">{{ var.label }}</small></h2>
  <div class="var">
    <img src="{{ var.chart_path }}" alt="{{ var.varname }} chart">
    <table>
      <tr><th>Statistic</th><th>Value</th></tr>
      {% for k, v in var.stats.items() %}
      <tr><td>{{ k }}</td><td>{{ v }}</td></tr>
      {% endfor %}
    </table>
  </div>
{% endfor %}

<footer>Generated by AGEVAL v1.0 — CIAT | {{ generated }}</footer>
</body>
</html>
"""

# ── Main ───────────────────────────────────────────────────────────────────────
def run_descriptive(data_path, dict_path=None):
    with open(CFG) as f:
        cfg = yaml.safe_load(f)

    dict_df   = pd.read_csv(dict_path or DICT)
    desc_vars = dict_df[dict_df["descriptive_include"] == 1].copy()
    sections  = load_sections(MASTER)
    data      = load_data(data_path)
    n_records = len(data)

    os.makedirs(OUTSTATS, exist_ok=True)
    charts_dir = OUTPLOTS
    os.makedirs(charts_dir, exist_ok=True)

    variables_rendered = []

    for _, row in desc_vars.iterrows():
        vname    = row["variable_name"]
        qtype    = row.get("surv_type", "text")
        label    = row.get("label_spanish", vname)
        var_type = row.get("var_type", "")
        topic    = row.get("topic", "")
        subtopic = row.get("subtopic", "")

        chart_type = chart_type_for(var_type)

        # Prefer standardised column if available
        col = row.get("surv_calculation_output", vname)
        col = col if (pd.notna(col) and col in data.columns) else vname

        if col not in data.columns:
            print(f"  ⚠  Column '{col}' not found — skipping {vname}")
            continue

        series     = data[col]
        chart_path = os.path.join(charts_dir, f"{vname}.png")

        try:
            if chart_type == "histogram":
                make_histogram(series, vname, label, chart_path)
            else:
                make_bar(series, vname, label, chart_path)
        except Exception as e:
            print(f"  ⚠  Chart error for {vname}: {e}")
            continue

        stats = summary_stats(series, qtype)
        variables_rendered.append({
            "varname":         vname,
            "label":           label,
            "topic":           topic,
            "subtopic":        subtopic,
            "topic_order":     sections["topic_order"].get(topic, float("inf")),
            "subtopic_order":  sections["subtopic_order"].get(subtopic, float("inf")),
            "topic_label":     sections["topic_label"].get(topic, str(topic).upper()),
            "subtopic_label":  sections["subtopic_label"].get(subtopic, str(subtopic)),
            "chart_path":      os.path.relpath(chart_path, OUTSTATS),
            "stats":           stats,
        })

    # Sort by section order (topic -> subtopic), falling back to name for ties
    variables_rendered.sort(
        key=lambda x: (x["topic_order"], x["subtopic_order"], x["varname"])
    )

    # Render HTML
    tmpl = Template(HTML_TMPL)
    html = tmpl.render(
        data_path=os.path.basename(data_path),
        generated=datetime.now().strftime("%Y-%m-%d %H:%M"),
        n_records=n_records,
        variables=variables_rendered,
    )

    ts       = datetime.now().strftime("%Y%m%d_%H%M")
    out_path = os.path.join(OUTSTATS, f"descriptive_stats_{ts}.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅  Descriptive stats report saved: {out_path}")
    print(f"    Variables rendered: {len(variables_rendered)}")
    return out_path, variables_rendered

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="AGEVAL Descriptive Stats")
    parser.add_argument("--data", required=True, help="Path to cleaned CSV data file")
    parser.add_argument("--dict", default=None, help="Path to personalized dictionary CSV (default: config path)")
    args = parser.parse_args()
    run_descriptive(args.data, dict_path=args.dict)