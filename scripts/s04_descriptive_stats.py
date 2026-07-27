"""
AGEVAL Step 4 — Descriptive Statistics Generator
Run: python scripts/s04_descriptive_stats.py --data data_clean/cleaned_data.csv

Reads cleaned data and generates charts + summary HTML for all
variables flagged descriptive_include=1 in the dictionary.
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

BASE   = os.path.dirname(os.path.abspath(__file__))
ROOT   = os.path.join(BASE, "..")
CFG    = os.path.join(ROOT, "config", "config.yaml")
DICT   = os.path.join(ROOT, "dictionary", "variables_personalized.csv")
OUTDIR = os.path.join(ROOT, "outputs", "stats")

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

def make_boxplot(series, varname, label, out_path, group_series=None):
    fig, ax = plt.subplots(figsize=(6, 3.5))
    vals = pd.to_numeric(series, errors="coerce").dropna()

    if group_series is not None and len(group_series) > 0:
        groups = group_series.reindex(vals.index).dropna().unique()
        data   = [pd.to_numeric(series[group_series == g], errors="coerce").dropna()
                  for g in groups]
        ax.boxplot(data, labels=groups, patch_artist=True,
                   boxprops=dict(facecolor=BLUE + "44", color=BLUE),
                   medianprops=dict(color=CORAL, linewidth=2))
        ax.set_xticklabels(groups, rotation=30, ha="right", fontsize=9)
    else:
        ax.boxplot([vals], patch_artist=True,
                   boxprops=dict(facecolor=BLUE + "44", color=BLUE),
                   medianprops=dict(color=CORAL, linewidth=2))
        ax.set_xticklabels([varname], fontsize=9)

    ax.set_ylabel(label, fontsize=10)
    ax.set_title(f"{varname}\nn={len(vals)}, median={vals.median():.2f}", fontsize=10)
    fig.tight_layout()
    fig.savefig(out_path, dpi=120, bbox_inches="tight")
    plt.close(fig)

def make_pie(series, varname, label, out_path):
    fig, ax = plt.subplots(figsize=(5, 4))
    counts = series.dropna().value_counts()
    colors = [BLUE, TEAL, CORAL, GRAY, "#7F77DD", "#D4537E"]
    ax.pie(counts.values, labels=counts.index,
           colors=[colors[i % len(colors)] for i in range(len(counts))],
           autopct="%1.1f%%", startangle=90,
           wedgeprops=dict(edgecolor="white", linewidth=1))
    ax.set_title(f"{varname} (n={len(series.dropna())})", fontsize=10)
    fig.tight_layout()
    fig.savefig(out_path, dpi=120, bbox_inches="tight")
    plt.close(fig)

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
  .topic-header { background: #EFF6FF; padding: .5rem 1rem; border-radius: 6px;
                  font-weight: 600; font-size: .9rem; color: #1E40AF; margin: 2rem 0 .5rem; }
  footer { margin-top: 3rem; font-size: .75rem; color: #94A3B8; }
</style>
</head>
<body>
<h1>AGEVAL — Descriptive Statistics</h1>
<p>Dataset: <strong>{{ data_path }}</strong> &nbsp;|&nbsp; Generated: {{ generated }} &nbsp;|&nbsp; n = <strong>{{ n_records }}</strong></p>

{% set ns = namespace(current_topic="") %}
{% for var in variables %}
  {% if var.topic != ns.current_topic %}
    {% set ns.current_topic = var.topic %}
    <div class="topic-header">{{ var.topic | upper }}</div>
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
def run_descriptive(data_path):
    with open(CFG) as f:
        cfg = yaml.safe_load(f)

    dict_df = pd.read_csv(DICT)
    desc_vars = dict_df[dict_df["descriptive_include"] == 1].copy()
    data = pd.read_csv(data_path)
    n_records = len(data)

    os.makedirs(OUTDIR, exist_ok=True)
    charts_dir = os.path.join(OUTDIR, "charts")
    os.makedirs(charts_dir, exist_ok=True)

    variables_rendered = []

    for _, row in desc_vars.iterrows():
        vname      = row["variable_name"]
        qtype      = row.get("question_type", "text")
        label      = row.get("label_spanish", vname)
        chart_type = row.get("descriptive_chart", "histogram")
        group_var  = row.get("descriptive_group_by")

        # Prefer standardised column if available
        col = row.get("odk_calculate_output", vname)
        col = col if (pd.notna(col) and col in data.columns) else vname

        if col not in data.columns:
            print(f"  ⚠  Column '{col}' not found — skipping {vname}")
            continue

        series      = data[col]
        chart_path  = os.path.join(charts_dir, f"{vname}.png")
        group_series = data[group_var] if (pd.notna(group_var) and group_var in data.columns) else None

        try:
            if chart_type == "histogram":
                make_histogram(series, vname, label, chart_path)
            elif chart_type == "bar":
                make_bar(series, vname, label, chart_path)
            elif chart_type == "boxplot":
                make_boxplot(series, vname, label, chart_path, group_series)
            elif chart_type == "pie":
                make_pie(series, vname, label, chart_path)
            else:
                make_histogram(series, vname, label, chart_path)
        except Exception as e:
            print(f"  ⚠  Chart error for {vname}: {e}")
            continue

        stats = summary_stats(series, qtype)
        variables_rendered.append({
            "varname":    vname,
            "label":      label,
            "topic":      row.get("topic", ""),
            "chart_path": os.path.relpath(chart_path, OUTDIR),
            "stats":      stats,
        })

    # Sort by topic
    variables_rendered.sort(key=lambda x: (x["topic"], x["varname"]))

    # Render HTML
    tmpl = Template(HTML_TMPL)
    html = tmpl.render(
        data_path=os.path.basename(data_path),
        generated=datetime.now().strftime("%Y-%m-%d %H:%M"),
        n_records=n_records,
        variables=variables_rendered,
    )

    ts       = datetime.now().strftime("%Y%m%d_%H%M")
    out_path = os.path.join(OUTDIR, f"descriptive_stats_{ts}.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅  Descriptive stats report saved: {out_path}")
    print(f"    Variables rendered: {len(variables_rendered)}")
    return out_path

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="AGEVAL Descriptive Stats")
    parser.add_argument("--data", required=True, help="Path to cleaned CSV data file")
    args = parser.parse_args()
    run_descriptive(args.data)
