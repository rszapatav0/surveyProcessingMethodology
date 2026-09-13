"""
AGEVAL Step 3 — Data Quality Checker
Run: python scripts/s03_quality_check.py --data data_raw/collected_data.xlsx
Run: python scripts/s03_quality_check.py --data data_raw/test_data_honduras_n2051.xlsx

Reads a collected ODK/Kobo export (XLSX, with the general/main sheet plus any
number of loop/repeat sheets — or, for backward compatibility, a single flat
CSV) and applies quality rules from the dictionary. Produces an HTML quality
report per batch, with separate sections for:
  1. Duplicates              (general/main sheet, respondent-level)
  2. General variables       (general/main sheet, sample = respondents)
  3. Each additional loop/repeat sheet (sample = observations in that sheet;
     a respondent may contribute more than one observation)

For XLSX input, the general/main sheet is the one named in the YAML config
under `project.name`; every other sheet is treated as a loop/repeat sheet and
is linked back to its parent respondent using the standard Kobo/ODK
parent/submission metadata columns (`_submission__uuid` / `_uuid`,
`_parent_index` / `_index`, etc.), so no identifiers are hard-coded.
"""
 
import pandas as pd
import numpy as np
import os
import re
import yaml
import argparse
from datetime import datetime
from jinja2 import Template
 
BASE   = os.path.dirname(os.path.abspath(__file__))
ROOT   = os.path.join(BASE, "..")
CFG    = os.path.join(ROOT, "config", "config.yaml")
DICT   = os.path.join(ROOT, "dictionary", "variables_personalized.csv")
OUTDIR = os.path.join(ROOT, "outputs", "quality")
 
# ── HTML report template ───────────────────────────────────────────────────────
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>AGEVAL Quality Report - {{ batch }}</title>
<style>
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
  footer { margin-top: 3rem; font-size: .75rem; color: #94A3B8; }
</style>
</head>
<body>
<h1>AGEVAL Quality Report</h1>
<p>Batch: <strong>{{ batch }}</strong> &nbsp;|&nbsp; Generated: {{ generated }} &nbsp;|&nbsp; Records checked: <strong>{{ n_records }}</strong></p>
 
<div class="summary-grid">
  <div class="card"><div class="val ok">{{ n_records }}</div><div class="lbl">Total records</div></div>
  <div class="card"><div class="val warn">{{ n_warnings }}</div><div class="lbl">Warnings</div></div>
  <div class="card"><div class="val err">{{ n_errors }}</div><div class="lbl">Errors (out of range)</div></div>
  <div class="card"><div class="val">{{ n_vars_checked }}</div><div class="lbl">Variables checked</div></div>
</div>
 
{% macro render_flagged(flagged) %}
{% if flagged %}
<details>
  <summary style="cursor:pointer; margin-top:.5rem; font-size:.8rem; color:#1D4ED8;">
    Show {{ flagged|length }} flagged records
  </summary>
  <table style="margin-top:.5rem;">
    <tr><th>ID</th><th>Value</th><th>Issue</th></tr>
    {% for rec in flagged %}
    <tr><td>{{ rec.id }}</td><td>{{ rec.value }}</td><td>{{ rec.issue }}</td></tr>
    {% endfor %}
  </table>
</details>
{% endif %}
{% endmacro %}

{% macro render_var_section(vname, result) %}
<h3>{{ vname }} <small style="font-weight:400; color:#64748B;">— {{ result.label }}</small></h3>
<table>
  <tr>
    <th>Check</th><th>Status</th><th>Detail</th>
  </tr>
  {% for check in result.checks %}
  <tr>
    <td>{{ check.name }}</td>
    <td><span class="badge badge-{{ check.status }}">{{ check.status.upper() }}</span></td>
    <td>{{ check.detail }}</td>
  </tr>
  {% endfor %}
</table>
{{ render_flagged(result.flagged_ids) }}
{% endmacro %}

<h2>1. Duplicates <span class="sample">— {{ n_records }} respondents ({{ main_sheet_name }})</span></h2>
<table>
  <tr>
    <th>Check</th><th>Status</th><th>Detail</th>
  </tr>
  {% for check in duplicates.checks %}
  <tr>
    <td>{{ check.name }}</td>
    <td><span class="badge badge-{{ check.status }}">{{ check.status.upper() }}</span></td>
    <td>{{ check.detail }}</td>
  </tr>
  {% endfor %}
</table>
{{ render_flagged(duplicates.flagged_ids) }}

<h2>2. General variables <span class="sample">— {{ n_records }} respondents ({{ main_sheet_name }})</span></h2>
{% if general_results %}
{% for vname, result in general_results.items() %}
{{ render_var_section(vname, result) }}
{% endfor %}
{% else %}
<p><em>No general variables checked.</em></p>
{% endif %}

{% for sheet in loop_sections %}
<h2>3. {{ sheet.name }} <span class="sample">— {{ sheet.n_records }} observations (loop/repeat sheet)</span></h2>
{% if sheet.results %}
{% for vname, result in sheet.results.items() %}
{{ render_var_section(vname, result) }}
{% endfor %}
{% else %}
<p><em>No variables from the dictionary matched this sheet.</em></p>
{% endif %}
{% endfor %}

<footer>Generated by AGEVAL v1.0 — CIAT</footer>
</body>
</html>
"""
 
# ── Load config and functions  ─────────────────────────────────────────────────
cfg = yaml.safe_load(open(CFG))
# Extract global thresholds from config
cfg_duplicate_pct = cfg["quality"]["duplicate_pct"]
cfg_method        = cfg["quality"]["outlier_method"]
cfg_missing_warn  = cfg["quality"]["missing_warn"]
cfg_missing_err   = cfg["quality"]["missing_err"]
 
# ── Duplicate detection ─────────────────────────────────────────────────────────
def check_duplicates(report_data, id_col, compare_cols, cfg_duplicate_pct):
    checks = []
    flagged = []
 
    extra_cols = [c for c in ["respondent_id", "start", "end", "deviceid"] if c in report_data.columns]
 
    # 1. Exact respondent_id duplicates
    id_dup_mask = report_data[id_col].duplicated(keep=False) & report_data[id_col].notna()
    id_dups = report_data[id_dup_mask]
    for idx, r in id_dups.iterrows():
        flagged.append({
            "id": r.get(id_col, idx),
            "value": ", ".join(f"{c}={r.get(c, '')}" for c in extra_cols),
            "issue": "Duplicate ID",
        })
    checks.append({
        "name":   "Duplicate respondent_id",
        "status": "err" if len(id_dups) > 0 else "ok",
        "detail": f"{len(id_dups)} records with duplicate respondent_id",
    })
 
    # 2. Similarity duplicates — pairwise comparison across previously checked variables
    valid_compare_cols = [c for c in compare_cols if c in report_data.columns]
    n_cols = len(valid_compare_cols)
    sim_flagged = []
 
    if n_cols > 0:
        records = report_data[valid_compare_cols].reset_index(drop=True)
        ids     = report_data[id_col].reset_index(drop=True)
        extras  = report_data[extra_cols].reset_index(drop=True) if extra_cols else None
        n = len(records)
 
        for i in range(n):
            row_i = records.iloc[i]
            for j in range(i + 1, n):
                row_j = records.iloc[j]
                matches = (row_i == row_j) | (row_i.isna() & row_j.isna())
                pct = matches.sum() / n_cols * 100
 
                if pct >= cfg_duplicate_pct:
                    detail_extra = ""
                    if extras is not None:
                        detail_extra = " | ".join(
                            f"{c}: {extras.iloc[i][c]} / {extras.iloc[j][c]}" for c in extra_cols
                        )
                    sim_flagged.append({
                        "id":    f"{ids[i]} / {ids[j]}",
                        "value": f"{round(pct, 1)}%" + (f" — {detail_extra}" if detail_extra else ""),
                        "issue": "Potential duplicate based on similarity",
                    })
 
    checks.append({
        "name":   f"Similarity duplicates (≥{cfg_duplicate_pct}%)",
        "status": "warn" if len(sim_flagged) > 0 else "ok",
        "detail": f"{len(sim_flagged)} record pairs flagged out of {n_cols} compared variables",
    })
 
    flagged.extend(sim_flagged)
    return checks, flagged
 
 
# ── Shared ODK expression translation (used by relevance and constraint checks) ─
def odk_to_python_expr(expr, dot_col=None):
    """
    Convert an ODK-style expression (as used in `relevant` and `constraint`
    columns of XLSForms — stored here as `surv_relevant` / `surv_constraint`)
    into a Python-evaluable expression string that operates on a `row` dict.
 
    Supported translations:
      ${var}                -> row.get('var')
      .                      -> row.get(dot_col)  (ODK's "current field value"
                                placeholder, used in constraint expressions;
                                pass dot_col=<the field's own column name>)
      single '=' comparison -> '=='  (leaves ==, !=, <=, >= untouched)
      selected(${var},'x')  -> membership test against space-separated values
      and / or / not        -> pass through (valid Python keywords already)
    Returns None if the expression is empty/NaN.
    """
    if pd.isna(expr) or not str(expr).strip():
        return None
 
    e = str(expr).strip()
 
    # selected(${var}, 'value')  -> handled before ${..} substitution so we
    # can reference the raw variable name twice safely.
    def _selected_repl(m):
        var, val = m.group(1), m.group(2)
        return f"('{val}' in str(row.get('{var}') or '').split())"
    e = re.sub(r"selected\(\s*\$\{([a-zA-Z0-9_]+)\}\s*,\s*'([^']*)'\s*\)", _selected_repl, e)
 
    # ${var} -> row.get('var')
    e = re.sub(r"\$\{([a-zA-Z0-9_]+)\}", r"row.get('\1')", e)
 
    # standalone '.' (ODK's current-value placeholder, e.g. ". > 0") -> row.get(dot_col)
    # The lookaround avoids matching decimal points (e.g. "0.5", which are
    # always preceded by a digit) or dots that are part of another token.
    if dot_col is not None:
        e = re.sub(r"(?<![\w.])\.(?![\w.])", f"row.get('{dot_col}')", e)
 
    # single '=' to '==' without touching !=, <=, >=, == already present
    e = re.sub(r"(?<![=!<>])=(?!=)", "==", e)
 
    # ODK string literals use double quotes sometimes; Python is fine with both.
    return e
 
 
def coerce_numeric(value):
    """Try to coerce a value to float so that >, <, >=, <= comparisons in
    relevance/constraint expressions behave numerically instead of falling
    back to (incorrect) lexicographic string comparison. Non-numeric values
    (including NaN/None) are returned unchanged."""
    if value is None:
        return value
    try:
        if pd.isna(value):
            return value
    except (TypeError, ValueError):
        pass
    try:
        return float(value)
    except (TypeError, ValueError):
        return value
 
 
def _coerced_row_dict(row):
    return {k: coerce_numeric(v) for k, v in row.to_dict().items()}


def _extract_odk_vars(expr):
    """Return the ${var} names referenced in an ODK expression (relevant/constraint)."""
    if pd.isna(expr):
        return []
    return re.findall(r"\$\{([a-zA-Z0-9_]+)\}", str(expr))


def _missing_dependency_breakdown(dep_vars, row_dicts):
    """
    Given the list of ${var} names referenced by an expression and the row
    dicts for the records that could not be evaluated, count — per
    dependency variable — how many of those records had that variable
    missing (NaN/None or absent from the data). Returns a detail string
    fragment like "area_total_ha missing in 52 records", or "" if no
    dependency variable was found to be missing (e.g. the failure came from
    something else, such as a type mismatch).
    """
    if not dep_vars or not row_dicts:
        return ""
    counts = {}
    for var in dep_vars:
        n_missing = sum(
            1 for rd in row_dicts
            if var not in rd or rd.get(var) is None or (isinstance(rd.get(var), float) and pd.isna(rd.get(var)))
        )
        if n_missing > 0:
            counts[var] = n_missing
    if not counts:
        return ""
    return "; ".join(f"'{v}' missing in {n} of them" for v, n in counts.items())
 
 
def evaluate_relevance(expr_py, row_dict):
    """Evaluate a translated relevance expression against a row dict.
    Returns True/False, or None if it can't be evaluated (e.g. a dependency
    variable is not present in the data)."""
    if expr_py is None:
        return None
    try:
        return bool(eval(expr_py, {"__builtins__": {}}, {"row": row_dict}))
    except Exception:
        return None
 
 
def check_relevance(report_data, id_col, col, relevant_expr):
    """
    Flags:
      - records with a value in `col` even though `relevant_expr` evaluates
        to False (the field should have been skipped by the form logic).
      - records where `relevant_expr` evaluates to True but `col` is missing
        (the field was expected to be filled in — surface this alongside the
        existing missing-value rule so it can be reviewed together).
    """
    checks = []
    flagged = []
 
    expr_py = odk_to_python_expr(relevant_expr)
    if expr_py is None or col not in report_data.columns:
        return checks, flagged
 
    n_unexpected_value = 0
    n_unexpected_missing = 0
    n_unevaluable = 0
    unevaluable_rows = []
    dep_vars = _extract_odk_vars(relevant_expr)
 
    for idx, r in report_data.iterrows():
        row_dict = _coerced_row_dict(r)
        is_relevant = evaluate_relevance(expr_py, row_dict)
        rid = r.get(id_col, idx)
 
        if is_relevant is None:
            n_unevaluable += 1
            unevaluable_rows.append(row_dict)
            continue
 
        value = r.get(col)
        has_value = pd.notna(value) and str(value).strip() != ""
 
        if not is_relevant and has_value:
            n_unexpected_value += 1
            flagged.append({
                "id":    rid,
                "value": value,
                "issue": f"Value present but relevance condition not met ({relevant_expr})",
            })
        elif is_relevant and not has_value:
            n_unexpected_missing += 1
            flagged.append({
                "id":    rid,
                "value": "(missing)",
                "issue": f"Relevance condition met ({relevant_expr}) but value is missing — review against quality rules",
            })
 
    checks.append({
        "name":   "Relevance: value present when not relevant",
        "status": "err" if n_unexpected_value > 0 else "ok",
        "detail": f"{n_unexpected_value} records have a value despite condition '{relevant_expr}' not being met",
    })
    checks.append({
        "name":   "Relevance: missing value when relevant",
        "status": "warn" if n_unexpected_missing > 0 else "ok",
        "detail": f"{n_unexpected_missing} records missing a value despite condition '{relevant_expr}' being met",
    })
    if n_unevaluable > 0:
        breakdown = _missing_dependency_breakdown(dep_vars, unevaluable_rows)
        detail = f"{n_unevaluable} records could not be evaluated (missing dependency variable(s) in '{relevant_expr}')"
        if breakdown:
            detail += f" — {breakdown}"
        checks.append({
            "name":   "Relevance: condition could not be evaluated",
            "status": "warn",
            "detail": detail,
        })
 
    return checks, flagged
 
 
# ── Constraint check (based on surv_constraint / ODK "constraint" expressions) ──
def check_constraint(report_data, id_col, col, constraint_expr):
    """
    ODK evaluates a field's `constraint` only when the field has a value,
    using '.' as the placeholder for that field's own value (e.g. ". > 0",
    ". >= 1 and . <= 10", ". > 0 and . <= ${area_total_farm_raw}").
 
    Flags every record whose value does NOT satisfy its constraint.
    """
    checks = []
    flagged = []
 
    if pd.isna(constraint_expr) or not str(constraint_expr).strip() or col not in report_data.columns:
        return checks, flagged
 
    expr_py = odk_to_python_expr(constraint_expr, dot_col=col)
    if expr_py is None:
        return checks, flagged
 
    n_violations = 0
    n_unevaluable = 0
    unevaluable_rows = []
    dep_vars = _extract_odk_vars(constraint_expr)
 
    for idx, r in report_data.iterrows():
        value = r.get(col)
        has_value = pd.notna(value) and str(value).strip() != ""
        if not has_value:
            continue  # constraint only applies when a value is present, as in ODK
 
        rid = r.get(id_col, idx)
        row_dict = _coerced_row_dict(r)
 
        try:
            satisfied = bool(eval(expr_py, {"__builtins__": {}}, {"row": row_dict}))
        except Exception:
            n_unevaluable += 1
            unevaluable_rows.append(row_dict)
            continue
 
        if not satisfied:
            n_violations += 1
            flagged.append({
                "id":    rid,
                "value": value,
                "issue": f"Value {value!r} violates constraint ({constraint_expr})",
            })
 
    checks.append({
        "name":   f"Constraint ({constraint_expr})",
        "status": "err" if n_violations > 0 else "ok",
        "detail": f"{n_violations} records violate constraint '{constraint_expr}'",
    })
    if n_unevaluable > 0:
        breakdown = _missing_dependency_breakdown(dep_vars, unevaluable_rows)
        detail = f"{n_unevaluable} records could not be evaluated (missing dependency variable(s) in '{constraint_expr}')"
        if breakdown:
            detail += f" — {breakdown}"
        checks.append({
            "name":   "Constraint: could not be evaluated",
            "status": "warn",
            "detail": detail,
        })
 
    return checks, flagged
 
 
# ── Core quality checks ────────────────────────────────────────────────────────
def check_variable(series, row, id_series, reference_series=None):
    checks = []
    flagged = []
 
    if reference_series is None:
        reference_series = series
 
    # Extract variable-specific thresholds from the dictionary row
    vname   = row["variable_name"]
    label   = row.get("label_spanish", vname)
    vmin    = row.get("quality_min")
    vmax    = row.get("quality_max")
    sd_thr  = row.get("quality_outlier_sd")
 
    valid = series.dropna()
    n_total   = len(series)
    missing_mask = series.isna() | series.isin([555, 666, 777, 888, 999])
    n_missing = missing_mask.sum()

    # 1. Missing values (reported range only)
    pct_missing = round(n_missing / n_total * 100, 1) if n_total > 0 else 0
    for idx in series[missing_mask].index:
        val = series.get(idx)
        flagged.append({
            "id":    id_series.get(idx, idx),
            "value": "(missing)" if pd.isna(val) else val,
            "issue": "Missing value",
        })
    checks.append({
        "name":   "Missing values",
        "status": "err" if pct_missing > cfg_missing_warn else ("warn" if pct_missing > cfg_missing_warn else "ok"),
        "detail": f"{n_missing}/{n_total} missing ({pct_missing}%)",
    })
 
    # 2. Range check (fixed thresholds — flagged within reported range only)
    if pd.notna(vmin) and pd.notna(vmax):
        numeric = pd.to_numeric(valid, errors="coerce").dropna()
        numeric = numeric[~numeric.isin([555, 666, 777, 888, 999])]
        out_of_range = numeric[(numeric < vmin) | (numeric > vmax)]
        for idx in out_of_range.index:
            flagged.append({"id": id_series.get(idx, idx), "value": numeric[idx], "issue": f"Out of range [{vmin}, {vmax}]"})
        checks.append({
            "name":   f"Range [{vmin} – {vmax}]",
            "status": "err" if len(out_of_range) > 0 else "ok",
            "detail": f"{len(out_of_range)} values out of range",
        })
 
    # 3. Outlier check — reference (mean/sd) from reference_series,
    #    flags raised only among the reported range.
    reference_numeric = pd.to_numeric(reference_series.dropna(), errors="coerce").dropna()
    reference_numeric = reference_numeric[~reference_numeric.isin([555, 666, 777, 888, 999])]
    if len(reference_numeric) > 4:
        numeric = pd.to_numeric(valid, errors="coerce").dropna()
        numeric = numeric[~numeric.isin([888, 999])]
 
        if cfg_method == "sd" and pd.notna(sd_thr) and sd_thr > 0:
            mean = reference_numeric.mean()
            sd = reference_numeric.std()
            lower = mean - sd_thr * sd
            upper = mean + sd_thr * sd
            check_name = f"Outlier (>{sd_thr} SD)"
            issue = f"Outlier (>{sd_thr} SD from mean {mean:.2f})"
            detail = (
                f"mean={mean:.2f}, sd={sd:.2f}, "
                f"reference n={len(reference_numeric)}"
            )
 
        elif cfg_method == "iqr":
            q1 = reference_numeric.quantile(0.25)
            q3 = reference_numeric.quantile(0.75)
            iqr = q3 - q1
            lower = q1 - 1.5 * iqr
            upper = q3 + 1.5 * iqr
            check_name = "Outlier (IQR)"
            issue = f"Outlier (outside [{lower:.2f}, {upper:.2f}])"
            detail = (
                f"Q1={q1:.2f}, Q3={q3:.2f}, IQR={iqr:.2f}, "
                f"reference n={len(reference_numeric)}"
            )
 
        outliers = numeric[(numeric < lower) | (numeric > upper)]
        for idx in outliers.index:
            flagged.append({
                "id": id_series.get(idx, idx),
                "value": round(numeric[idx], 2),
                "issue": issue,
            })
        checks.append({
            "name": check_name,
            "status": "warn" if len(outliers) else "ok",
            "detail": f"{len(outliers)} outliers detected ({detail})",
        })
 
    return checks, flagged, label
 
 
# ── Date filtering ──────────────────────────────────────────────────────────────
def filter_by_date(data, date_start=None, date_end=None, date_col="surveyDate"):
    if date_col not in data.columns:
        return data
    if not pd.api.types.is_datetime64_any_dtype(data[date_col]):
        data = data.copy()
        data[date_col] = pd.to_datetime(data[date_col], errors="coerce")
    if date_start is not None:
        data = data[data[date_col] >= pd.to_datetime(date_start)]
    if date_end is not None:
        data = data[data[date_col] <= pd.to_datetime(date_end)]
    return data
 
def get_available_dates(data_path, date_col="surveyDate", sheet_name=None):
    """
    Return the sorted list of distinct survey dates found in `date_col`.
    For XLSX input, only the general/main sheet is inspected (loop sheets
    inherit their date from the parent respondent, so the main sheet already
    reflects the full set of available dates). `sheet_name` can be passed
    explicitly (e.g. by a Streamlit app that already knows it); otherwise it
    is resolved the same way as in `run_quality_check` (via `project.name`
    in the YAML config, falling back to the first sheet).
    """
    try:
        if is_excel_file(data_path):
            if sheet_name is None:
                sheet_name = get_main_sheet_name(data_path, cfg)
            data = pd.read_excel(data_path, sheet_name=sheet_name, usecols=lambda c: c == date_col)
        else:
            data = pd.read_csv(data_path, usecols=lambda c: c == date_col)
    except (ValueError, FileNotFoundError, KeyError):
        return []
    if date_col not in data.columns:
        return []
    dates = pd.to_datetime(data[date_col], errors="coerce").dropna().dt.date
    return sorted(dates.unique())


# ── XLSX (multi-sheet) support ──────────────────────────────────────────────────
def is_excel_file(data_path):
    return os.path.splitext(str(data_path))[1].lower() in (".xlsx", ".xlsm", ".xls")


def get_main_sheet_name(data_path, cfg):
    """
    The general/main sheet is the one whose name matches `project.name` in
    the YAML config. If that name isn't found among the workbook's sheets
    (or isn't configured), fall back to the first sheet — this keeps a
    single-sheet workbook working with no config changes required.
    """
    sheet_names = pd.ExcelFile(data_path).sheet_names
    project_name = (cfg.get("project") or {}).get("name")
    if project_name in sheet_names:
        return project_name
    return sheet_names[0]


def load_sheets(data_path, cfg):
    """
    Load every sheet of an XLSX workbook, or — for backward compatibility —
    a single flat CSV. Returns (sheets, main_sheet_name) where `sheets` is an
    ordered {sheet_name: DataFrame} dict. For CSV input there is only one
    "sheet" (no loop/repeat sheets).
    """
    if is_excel_file(data_path):
        xls = pd.ExcelFile(data_path)
        main_sheet_name = get_main_sheet_name(data_path, cfg)
        sheets = {sn: pd.read_excel(xls, sheet_name=sn) for sn in xls.sheet_names}
    else:
        main_sheet_name = os.path.basename(data_path).rsplit(".", 1)[0]
        sheets = {main_sheet_name: pd.read_csv(data_path)}
    return sheets, main_sheet_name


# Kobo/ODK metadata column pairs that link a loop/repeat sheet row back to its
# parent record in the general/main sheet, tried in order of preference.
LOOP_PARENT_LINK_COLUMNS = [
    ("_submission__uuid", "_uuid"),
    ("_parent_index", "_index"),
    ("_submission__id", "_id"),
]


def link_loop_to_main(loop_df, main_df, id_col="respondent_id", date_col="surveyDate"):
    """
    Attach the respondent identifier and survey date from the general/main
    sheet onto each row of a loop/repeat sheet, using the standard Kobo/ODK
    parent/submission metadata columns (no hard-coded identifiers). Also adds
    a unique per-row observation id (`_obs_id`, from the sheet's own `_index`
    when available) and a composite `_report_id` combining the two, so each
    repeated observation can be traced back to its respondent in reports.
    """
    loop_df = loop_df.copy()

    key_loop = key_main = None
    for lk, mk in LOOP_PARENT_LINK_COLUMNS:
        if lk in loop_df.columns and mk in main_df.columns:
            key_loop, key_main = lk, mk
            break

    if key_loop is None:
        # No recognizable parent link — respondent_id/date can't be derived.
        if id_col not in loop_df.columns:
            loop_df[id_col] = np.nan
        if date_col not in loop_df.columns:
            loop_df[date_col] = pd.NaT
    else:
        lookup_cols = [c for c in (id_col, date_col) if c in main_df.columns]
        lookup = main_df.set_index(key_main)[lookup_cols]
        lookup = lookup[~lookup.index.duplicated(keep="first")]
        if id_col in lookup_cols:
            loop_df[id_col] = loop_df[key_loop].map(lookup[id_col])
        elif id_col not in loop_df.columns:
            loop_df[id_col] = np.nan
        if date_col in lookup_cols:
            loop_df[date_col] = loop_df[key_loop].map(lookup[date_col])
        elif date_col not in loop_df.columns:
            loop_df[date_col] = pd.NaT

    # Unique observation id within this loop sheet (Kobo/ODK's own `_index`
    # for the repeat-group record when present; otherwise row position).
    if "_index" in loop_df.columns:
        loop_df["_obs_id"] = loop_df["_index"]
    else:
        loop_df["_obs_id"] = range(1, len(loop_df) + 1)

    # Parent record index in the general/main sheet (Kobo/ODK's own
    # `_parent_index` column when present; otherwise fall back to whatever
    # key was used above to link back to the main sheet).
    if "_parent_index" in loop_df.columns:
        parent_idx = loop_df["_parent_index"]
    elif key_loop is not None:
        parent_idx = loop_df[key_loop]
    else:
        parent_idx = pd.NA

    loop_df["_report_id"] = (
        loop_df[id_col].astype(str)
        + " (parent index: " + pd.Series(parent_idx, index=loop_df.index).astype(str)
        + ", index: " + loop_df["_obs_id"].astype(str) + ")"
    )
    return loop_df
 
 
# ── Per-sheet variable checks (shared by the general sheet and every loop) ──────
def run_variable_checks(report_data, reference_data, qc_vars, id_col):
    """
    Run the existing range/outlier/relevance/constraint checks (unchanged
    logic — see `check_variable`, `check_relevance`, `check_constraint`) for
    every dictionary variable whose column is present in `report_data`. Used
    once for the general/main sheet and once per loop/repeat sheet, so a
    variable is automatically scoped to whichever sheet(s) actually contain
    its column — no per-sheet variable list needs to be hard-coded.
    """
    results = {}
    total_warnings = 0
    total_errors   = 0

    id_series = report_data[id_col] if id_col in report_data.columns else report_data.index.to_series()

    for _, row in qc_vars.iterrows():
        vname = row["variable_name"]
        # Use calculated output column if available
        col = row.get("surv_calculation_output", vname)
        col = col if (pd.notna(col) and col in report_data.columns) else vname

        if col not in report_data.columns:
            continue

        reference_col = col if (reference_data is not None and col in reference_data.columns) else None
        checks, flagged, label = check_variable(
            report_data[col], row, id_series,
            reference_series=reference_data[reference_col] if reference_col is not None else None,
        )

        # ── Relevance check ─────────────────────────────────────────────────
        relevant_expr = row.get("surv_relevant")
        if pd.notna(relevant_expr) and str(relevant_expr).strip():
            rel_checks, rel_flagged = check_relevance(report_data, id_col, col, relevant_expr)
            checks.extend(rel_checks)
            flagged.extend(rel_flagged)

        # ── Constraint check ─────────────────────────────────────────────────
        constraint_expr = row.get("surv_constraint")
        if pd.notna(constraint_expr) and str(constraint_expr).strip():
            con_checks, con_flagged = check_constraint(report_data, id_col, col, constraint_expr)
            checks.extend(con_checks)
            flagged.extend(con_flagged)

        results[vname] = {"label": label, "checks": checks, "flagged_ids": flagged}

        for c in checks:
            if c["status"] == "warn":
                total_warnings += 1
            elif c["status"] == "err":
                total_errors += 1

    return results, total_warnings, total_errors


# ── Main ───────────────────────────────────────────────────────────────────────
def run_quality_check(data_path, batch_name=None, date_start=None, date_end=None, date_col="surveyDate"):
    with open(CFG) as f:
        cfg = yaml.safe_load(f)
 
    dict_df = pd.read_csv(DICT)
    qc_vars = dict_df[dict_df["quality_include"] == 1]
 
    batch = batch_name or os.path.basename(data_path).rsplit(".", 1)[0]

    # `sheets` holds every sheet in the XLSX workbook ({name: DataFrame}), or
    # — for backward-compatible CSV input — a single {batch: DataFrame} entry.
    sheets, main_sheet_name = load_sheets(data_path, cfg)
    main_data = sheets[main_sheet_name]

    # Reference: all records up to (and including) the selected end date,
    # regardless of start date — used only to compute statistical references.
    main_reference = filter_by_date(main_data, None, date_end, date_col)
    # Reported: records within the full selected range — these are the ones
    # actually checked/flagged and shown in the report.
    main_report = filter_by_date(main_data, date_start, date_end, date_col)
    n_records = len(main_report)
 
    # ID column for reporting (general/main sheet — respondent-level)
    id_col = "respondent_id" if "respondent_id" in main_report.columns else main_report.columns[0]

    total_warnings = 0
    total_errors   = 0

    # ── Section 1: Duplicates (general/main sheet, respondent-level) ───────────
    compare_cols = [
        row.get("surv_calculation_output", row["variable_name"])
        if pd.notna(row.get("surv_calculation_output")) and row.get("surv_calculation_output") in main_report.columns
        else row["variable_name"]
        for _, row in qc_vars.iterrows()
    ]
    compare_cols = [c for c in compare_cols if c in main_report.columns]

    dup_checks, dup_flagged = check_duplicates(main_report, id_col, compare_cols, cfg_duplicate_pct)
    duplicates_result = {"label": "Duplicate detection", "checks": dup_checks, "flagged_ids": dup_flagged}

    for c in dup_checks:
        if c["status"] == "warn":
            total_warnings += 1
        elif c["status"] == "err":
            total_errors += 1

    # ── Section 2: General variables (general/main sheet) ──────────────────────
    general_results, gw, ge = run_variable_checks(main_report, main_reference, qc_vars, id_col)
    total_warnings += gw
    total_errors   += ge

    # ── Section 3: Each additional loop/repeat sheet ────────────────────────────
    # Sample here is the observations in that sheet (not respondents) — a
    # respondent can contribute multiple observations to a loop sheet.
    loop_sections = []
    for sheet_name, raw_sheet_df in sheets.items():
        if sheet_name == main_sheet_name:
            continue

        linked = link_loop_to_main(raw_sheet_df, main_data, id_col=id_col, date_col=date_col)
        loop_reference = filter_by_date(linked, None, date_end, date_col)
        loop_report    = filter_by_date(linked, date_start, date_end, date_col)

        # Loop-level results identify both respondent_id and a unique
        # observation id (combined in `_report_id`, added by link_loop_to_main).
        loop_results, lw, le = run_variable_checks(loop_report, loop_reference, qc_vars, "_report_id")
        total_warnings += lw
        total_errors   += le

        loop_sections.append({
            "name":      sheet_name,
            "n_records": len(loop_report),
            "results":   loop_results,
        })

    n_vars_checked = len(general_results) + sum(len(s["results"]) for s in loop_sections)

    # ── Render HTML report ─────────────────────────────────────────────────────
    tmpl = Template(HTML_TEMPLATE)
    html = tmpl.render(
        batch=batch,
        generated=datetime.now().strftime("%Y-%m-%d %H:%M"),
        n_records=n_records,
        n_warnings=total_warnings,
        n_errors=total_errors,
        n_vars_checked=n_vars_checked,
        main_sheet_name=main_sheet_name,
        duplicates=duplicates_result,
        general_results=general_results,
        loop_sections=loop_sections,
    )
 
    os.makedirs(OUTDIR, exist_ok=True)
    out_path  = os.path.join(OUTDIR, f"quality_report_{batch}.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
 
    print(f"✅  Quality report saved: {out_path}")
    print(f"    General sheet ('{main_sheet_name}') records: {n_records}")
    for sheet in loop_sections:
        print(f"    Loop sheet '{sheet['name']}' observations: {sheet['n_records']}")
    print(f"    Variables checked: {n_vars_checked}")
    print(f"    Warnings:          {total_warnings}")
    print(f"    Errors:            {total_errors}")
    return out_path
 
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="AGEVAL Quality Check")
    parser.add_argument("--data",  required=True, help="Path to collected data file (.xlsx with general + loop/repeat sheets, or legacy flat .csv)")
    parser.add_argument("--batch", default=None,  help="Batch name (optional label)")
    parser.add_argument("--date-start", default=None, help="Filter: only records on/after this date (YYYY-MM-DD)")
    parser.add_argument("--date-end",   default=None, help="Filter: only records on/before this date (YYYY-MM-DD)")
    parser.add_argument("--date-col",   default="surveyDate", help="Column name holding the survey date")
    args = parser.parse_args()
    run_quality_check(args.data, args.batch, args.date_start, args.date_end, args.date_col)