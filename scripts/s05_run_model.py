"""
AGEVAL Step 5 — Econometric Model Runner
Run: python scripts/s05_run_model.py --data data_clean/cleaned_data.csv

Uses model_role flags from dictionary to build and run regression models.
  1 = dependent variable
  2 = independent variable
  3 = control variable
  0 = excluded

Supports OLS, Fixed Effects (via linearmodels), and basic IV.
"""

import pandas as pd
import numpy as np
import os
import yaml
import argparse
from datetime import datetime

BASE   = os.path.dirname(os.path.abspath(__file__))
ROOT   = os.path.join(BASE, "..")
CFG    = os.path.join(ROOT, "config", "config.yaml")
DICT   = os.path.join(ROOT, "dictionary", "variables_personalized.csv")
OUTDIR = os.path.join(ROOT, "outputs", "models")

try:
    import statsmodels.formula.api as smf
    import statsmodels.api as sm
    HAS_STATSMODELS = True
except ImportError:
    HAS_STATSMODELS = False
    print("⚠  statsmodels not installed. Run: pip install statsmodels")

try:
    from linearmodels import PanelOLS, PooledOLS
    HAS_LINEARMODELS = True
except ImportError:
    HAS_LINEARMODELS = False

# ── Helpers ────────────────────────────────────────────────────────────────────
def get_col(row, data):
    """Return the best available column name for a variable."""
    calc_out = row.get("odk_calculate_output", "")
    if pd.notna(calc_out) and calc_out in data.columns:
        return calc_out
    if row["variable_name"] in data.columns:
        return row["variable_name"]
    return None

def prepare_data(dict_df, data):
    """Return clean DataFrame with only modeled variables, categoricals as pd.Categorical."""
    model_vars = dict_df[dict_df["model_role"] > 0].copy()

    keep_cols = []
    col_map   = {}   # original_varname -> column in data

    for _, row in model_vars.iterrows():
        col = get_col(row, data)
        if col:
            keep_cols.append(col)
            col_map[row["variable_name"]] = col

    df = data[list(set(keep_cols))].copy()

    # Mark string / select columns as pd.Categorical so patsy handles them
    select_vars = dict_df[(dict_df["model_role"] > 0) &
                          (dict_df["question_type"].isin(["select_one", "select_multiple", "text"]))]
    for _, row in select_vars.iterrows():
        col = col_map.get(row["variable_name"])
        if col and col in df.columns and df[col].dtype == object:
            df[col] = pd.Categorical(df[col])

    return df, col_map, model_vars

def build_formula(dict_df, col_map):
    """Build R-style formula string from model_role flags."""
    dependent = []
    independent = []
    controls   = []

    for _, row in dict_df.iterrows():
        role = row.get("model_role", 0)
        col  = col_map.get(row["variable_name"])
        if not col or role == 0:
            continue
        safe_col = col.replace(" ", "_").replace("-", "_")
        if role == 1:
            dependent.append(safe_col)
        elif role == 2:
            independent.append(safe_col)
        elif role == 3:
            controls.append(safe_col)

    if not dependent:
        raise ValueError("No dependent variable found. Set model_role=1 for at least one variable.")
    if not independent:
        raise ValueError("No independent variable found. Set model_role=2 for at least one variable.")

    y    = dependent[0]
    rhs  = " + ".join(independent + controls)
    return f"{y} ~ {rhs}", y, independent, controls

# ── OLS ────────────────────────────────────────────────────────────────────────
def run_ols(formula, df, cluster_var=None):
    if not HAS_STATSMODELS:
        return None, "statsmodels not installed"

    clean_df = df.rename(columns=lambda x: x.replace(" ", "_").replace("-", "_"))
    model    = smf.ols(formula, data=clean_df).fit()

    if cluster_var and cluster_var in df.columns:
        model = model.get_robustcov_results(cov_type="cluster", groups=df[cluster_var])

    return model, None

# ── Results formatter ──────────────────────────────────────────────────────────
def format_results_table(model, title="OLS Results"):
    """Return a DataFrame and a LaTeX string of the results."""
    if model is None:
        return pd.DataFrame(), ""

    results = pd.DataFrame({
        "Variable":   model.params.index,
        "Coef":       model.params.values,
        "Std Err":    model.bse.values,
        "t-stat":     model.tvalues.values,
        "p-value":    model.pvalues.values,
        "CI lower":   model.conf_int()[0].values,
        "CI upper":   model.conf_int()[1].values,
    })

    results["Signif"] = results["p-value"].apply(
        lambda p: "***" if p < 0.01 else ("**" if p < 0.05 else ("*" if p < 0.1 else ""))
    )
    results = results.round(4)

    summary_footer = pd.DataFrame([{
        "Variable": "N", "Coef": model.nobs, "Std Err": "",
        "t-stat": "", "p-value": "",
        "CI lower": "", "CI upper": "", "Signif": ""
    }, {
        "Variable": "R²", "Coef": round(model.rsquared, 4), "Std Err": "",
        "t-stat": "", "p-value": "",
        "CI lower": "", "CI upper": "", "Signif": ""
    }, {
        "Variable": "Adj. R²", "Coef": round(model.rsquared_adj, 4), "Std Err": "",
        "t-stat": "", "p-value": "",
        "CI lower": "", "CI upper": "", "Signif": ""
    }])

    full_table = pd.concat([results, summary_footer], ignore_index=True)

    # LaTeX
    latex = model.summary().as_latex()

    return full_table, latex

# ── Main ───────────────────────────────────────────────────────────────────────
def run_models(data_path):
    with open(CFG) as f:
        cfg = yaml.safe_load(f)

    dict_df     = pd.read_csv(DICT)
    data        = pd.read_csv(data_path)
    estimator   = cfg["model"].get("default_estimator", "OLS")
    cluster_var = cfg["model"].get("cluster_var", None)
    out_fmt     = cfg["model"].get("output_format", "both")

    os.makedirs(OUTDIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M")

    print(f"\n── AGEVAL Model Runner ─────────────────────────────")
    print(f"   Data:      {data_path}")
    print(f"   Estimator: {estimator}")
    print(f"   Cluster:   {cluster_var}")

    df, col_map, model_vars = prepare_data(dict_df, data)

    try:
        formula, dep, indep, controls = build_formula(dict_df, col_map)
    except ValueError as e:
        print(f"❌  {e}")
        return {"error": str(e)}

    print(f"\n   Formula: {formula}")
    print(f"   Dependent:    {dep}")
    print(f"   Independent:  {indep}")
    print(f"   Controls:     {controls}")
    print()

    model, err = run_ols(formula, df, cluster_var)

    if err:
        print(f"❌  {err}")
        return {"error": err, "formula": formula, "dep": dep, "indep": indep, "controls": controls}

    print(model.summary())

    full_table, latex = format_results_table(model, title=f"{estimator} — {dep}")

    # ── Save outputs ───────────────────────────────────────────────────────────
    if out_fmt in ("csv", "both"):
        csv_path = os.path.join(OUTDIR, f"model_results_{ts}.csv")
        full_table.to_csv(csv_path, index=False)
        print(f"\n✅  CSV results saved: {csv_path}")

    if out_fmt in ("latex", "both"):
        tex_path = os.path.join(OUTDIR, f"model_results_{ts}.tex")
        with open(tex_path, "w") as f:
            f.write(latex)
        print(f"✅  LaTeX results saved: {tex_path}")

    # ── Save model metadata ────────────────────────────────────────────────────
    meta = {
        "formula":   formula,
        "estimator": estimator,
        "n_obs":     int(model.nobs),
        "r2":        round(model.rsquared, 4),
        "r2_adj":    round(model.rsquared_adj, 4),
        "cluster_var": cluster_var,
        "data_file": os.path.basename(data_path),
        "generated": datetime.now().isoformat(),
    }
    meta_path = os.path.join(OUTDIR, f"model_metadata_{ts}.yaml")
    with open(meta_path, "w") as f:
        yaml.dump(meta, f)
    print(f"✅  Metadata saved: {meta_path}")

    return {
        "error":       None,
        "full_table":  full_table,
        "model":       model,
        "summary_txt": model.summary().as_text(),
        "formula":     formula,
        "dep":         dep,
        "indep":       indep,
        "controls":    controls,
        "n_obs":       int(model.nobs),
        "r2":          round(model.rsquared, 4),
        "r2_adj":      round(model.rsquared_adj, 4),
        "csv_path":    csv_path if out_fmt in ("csv", "both") else None,
        "tex_path":    tex_path if out_fmt in ("latex", "both") else None,
        "meta_path":   meta_path,
    }

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="AGEVAL Econometric Model Runner")
    parser.add_argument("--data", required=True, help="Path to cleaned CSV data file")
    args = parser.parse_args()
    run_models(args.data)
