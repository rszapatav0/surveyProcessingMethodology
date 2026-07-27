# AGEVAL — Agricultural Impact Evaluation Pipeline

End-to-end pipeline manager for agricultural impact evaluations.
Built for CIAT by Federico Ceballos.

## Philosophy

One file (`variables_master.csv`) is the single source of truth.
Each row is a variable. Each column controls one stage of the data pipeline.
You configure with 1s and 0s; the pipeline does the rest.

---

## Folder structure

```
ageval/
├── config/
│   └── config.yaml              ← Project settings (name, language, estimator...)
├── dictionary/
│   └── variables_master.csv     ← THE single source of truth
├── forms/
│   └── *.xlsx                   ← Generated ODK XLS forms
├── data_raw/
│   └── *.csv                    ← KoboCollect/ODK CSV exports
├── data_clean/
│   └── *.csv                    ← Cleaned data after QC
├── outputs/
│   ├── quality/                 ← HTML quality reports
│   ├── stats/                   ← Descriptive stat charts + HTML
│   └── models/                  ← Regression tables (CSV + LaTeX)
└── scripts/
    ├── s01_dictionary_selector.py   ← Streamlit app to toggle variables
    ├── s02_generate_odk_form.py     ← Builds ODK XLSForm
    ├── s03_quality_check.py         ← QC report per batch
    ├── s04_descriptive_stats.py     ← Charts + summary HTML
    ├── s05_run_model.py             ← OLS / FE / IV regressions
    ├── generate_test_data.py       ← Synthetic data for testing
    └── run_pipeline.py             ← Master runner
```

---

## Quick start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Edit project settings
nano config/config.yaml

# 3. Launch the variable selector (browser UI)
#streamlit run scripts/01_dictionary_selector.py
python -m streamlit run scripts/s01_dictionary_selector.py

# 4. Generate ODK XLS form
python scripts/s02_generate_odk_form.py

# 5. Generate synthetic test data (or drop your real CSV in data_raw/)
#python scripts/generate_test_data.py

# 6. Run quality check on collected data
python scripts/s03_quality_check.py --data data_raw/test_data_honduras_n200.csv --batch round1

# 7. Run descriptive statistics
python scripts/s04_descriptive_stats.py --data data_raw/test_data_honduras_n200.csv

# 8. Run econometric model
python scripts/s05_run_model.py --data data_raw/test_data_honduras_n200.csv

# Or run everything at once (steps 2–5)
python scripts/run_pipeline.py --data data_raw/test_data_honduras_n200.csv --steps 2 3 4 5
```

---

## Dictionary schema (`variables_master.csv`)

| Column | Description |
|--------|-------------|
| `variable_name` | Snake_case variable name — the canonical identifier |
| `label_english` | English question label |
| `label_spanish` | Spanish question label |
| `topic` | Thematic group (household, farm, production, market, ...) |
| `question_type` | ODK type: integer, decimal, text, select_one, select_multiple |
| `choices` | For select types: `value:Label;value:Label` (semicolon-separated) |
| `required` | 1 = required in ODK |
| `constraint` | ODK constraint expression (e.g. `. >= 18 and . <= 100`) |
| `constraint_message` | Message shown when constraint fails |
| **Pipeline flags** | |
| `questionnaire_include` | 1 = include in ODK form |
| `odk_calculate_expr` | ODK calculate expression (unit conversion, aggregation) |
| `odk_calculate_output` | Output variable name with semantic suffix (e.g. `area_ha`) |
| `odk_calculate_include` | 1 = generate calculate field |
| `quality_include` | 1 = include in QC check |
| `quality_min` | Minimum valid value |
| `quality_max` | Maximum valid value |
| `quality_outlier_sd` | Flag if value > N standard deviations from mean |
| `quality_report` | 1 = include in periodic QC report |
| `descriptive_include` | 1 = generate chart |
| `descriptive_chart` | Chart type: histogram, bar, boxplot, pie |
| `descriptive_group_by` | Variable to disaggregate by (optional) |
| `model_role` | 0=excluded, 1=dependent, 2=independent, 3=control |
| `model_notes` | Human notes on modeling rationale |

---

## Adding a new variable

1. Open `dictionary/variables_master.csv`
2. Add a new row with the variable name, label, type, and pipeline flags
3. Re-run whichever steps are affected (no other files need editing)

---

## Variable naming conventions

- Snake_case always
- Suffixes encode semantics:
  - `_raw` = as captured in the field (original units)
  - `_ha` = standardized to hectares
  - `_usd` = standardized to USD
  - `_kg` = standardized to kilograms
  - `_by_farmer` = aggregated at farmer level
  - `_by_farm` = aggregated at farm level
  - `_by_plot` = at plot level

---

## Supported estimators (Step 5)

Set in `config/config.yaml` under `model.default_estimator`:
- `OLS` — Ordinary Least Squares (statsmodels)
- `FE` — Fixed Effects Panel (linearmodels, requires panel structure)
- `IV` — Instrumental Variables (statsmodels iv2sls)

All produce CSV and LaTeX output tables.

---

## Credits

AGEVAL — CIAT / Alliance Bioversity-CIAT
Lead: Federico Ceballos (f.ceballos@cgiar.org)
