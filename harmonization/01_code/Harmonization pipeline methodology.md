# AGEVAL Harmonization Pipeline — Methodology Notes

Companion Streamlit pipeline to the individual baseline/endline pipeline, for combining and analyzing baseline+endline data together. Six steps, each a standalone-runnable `s0N_*.py` module also wired into a unified `app_harmonization.py`. All paths/settings come from `s00_harmonizationConfig.yaml`.

## Step 1 — Dictionary Merge (`s01_dictionary_merge.py`)
Combines the master dictionary with the personalized baseline/endline dictionaries into one harmonization dictionary. `baseline`/`endline` membership is auto-detected from each round's personalized dictionary. Analyst-editable checkbox flags per variable (pre-loaded from the master dictionary, editable in the app):
- `plots` — include in Step 5 comparison plots
- `outcome_var` — eligible regression outcome
- `control_var` — eligible baseline control
- `logarithm_var` — also offer this variable in log form
- `fe_var` — eligible fixed-effect variable
- `cluster_var` — eligible clustering variable

`surv_type`/`surv_choices` (carried through from master) identify select_one/select_multiple variables and their response options for later steps, since `var_type` alone can mistag a select_multiple as `numerical`.

## Step 2 — Combine into Long Database (`s02_combine_long.py`)
Appends baseline+endline vertically into one long database with a `period` indicator (0=baseline, 1=endline), keeping the participant ID unchanged and reconciling variable types/names across rounds per the dictionary. No merge/quality checks here.

## Step 3 — Treatment Assignment (`s03_treatment_assignment.py`)
Left-joins a treatment-assignment file onto the long database on user-chosen merge key(s) (supports individual- or community/group-level assignment), validated `m:1` so no existing row is duplicated/dropped. Keeps a numeric `treatment` column and a `treatment_label` column, available for both periods.

## Step 4 — Quality Check (`s04_quality_check.py`)
Panel/treatment-structure checks only (not variable-level quality): matched/attrition/new observations, duplicate IDs, treatment assignment and consistency by period/group. Exports an HTML report styled after the existing AGEVAL quality report.

## Step 5 — Comparison Plots (`s05_comparison_plots.py`)
Treatment × period comparison chart per variable flagged `plots=1`: numerical/dummy show mean + 95% CI + N; categorical (select_one/select_multiple) show counts per option, not stacked. A variable available in only one round is still plotted, with the missing round marked. Classification priority: `var_type=="dummy"` always wins; otherwise `surv_type`/`surv_choices` take priority over `var_type` (corrects select_multiple mistagged as `numerical` in the master dictionary).

## Step 6 — RCT Regressions (`s06_regressions.py`)
Reshapes the long database into a respondent-level wide database using the Step 1 flags, then runs OLS regressions.

**Wide reshape rules:**
- Outcome variables get two columns: `<var>__bl` (baseline, usable as a lagged/ANCOVA control) and `<var>__el` (endline, the actual regression outcome).
- Control variables use the baseline value only (falls back to endline if baseline is entirely missing for that variable, with a note).
- Fixed-effect and cluster variables are treated as time-invariant: baseline value, falling back to endline if baseline is missing.
- Logarithm-flagged variables get an additional `ln_<col>` column (non-positive values become missing, noted).
- Both raw and log forms are offered separately wherever applicable (as separate selectable "outcomes" or "controls" in the UI).

**Specification builder:** each specification independently sets treatment coding (categorical arm-dummies vs. a chosen reference, or the numeric treatment variable as a continuous regressor), whether to include the baseline value of the outcome as a control, baseline controls, treatment×control interactions (auto-adds the interacted variable as a main-effect control if not already selected), fixed effects (one or more `C(var)` dummy sets), and the standard-error type (classical, robust/HC1, or clustered by a flagged cluster variable). Each specification is then run against one or more chosen outcomes, so the same specification can be replicated across many outcomes, and different specifications can be compared side-by-side for the same outcome.

**Output:** results are grouped by outcome and laid out as standard side-by-side regression tables (via `statsmodels.iolib.summary2.summary_col`) — coefficients with significance stars, SEs in parentheses, N, R², and descriptive rows (`Controls: Yes/No`, `Baseline outcome control: Yes/No`, `<FE var> FE: Yes/No`, `SE type`, `Treatment`). Everything (regression tables + a specification-definitions table) exports to one HTML file and one compilable `.tex` file (validated by actually compiling with `pdflatex`).

**Dependency:** Step 6 requires `statsmodels` (not needed by Steps 1–5). No new config keys were needed — it reuses the existing `paths.outputs_regressions` and `model.output_format` keys from `s00_harmonizationConfig.yaml`.

## Cross-cutting conventions
- `pick_or_upload()` pattern: dropdown of files in the configured project folder + file uploader override (upload always wins), with a `default_path` auto-load fallback.
- Consistent HTML export styling (badges, striped tables, `#2563EB` header rule) across Steps 4/5/6.
- Validated colorblind-safe categorical palette (`#2a78d6, #eb6834, #1baf7a, #eda100, #e87ba4, #008300, #4a3aa7, #e34948`) assigned in fixed order per treatment arm, reused across every chart.
- All steps tested headlessly via `streamlit.testing.v1.AppTest` plus direct unit calls to internal functions before delivery.