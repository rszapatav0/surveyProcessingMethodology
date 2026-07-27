"""
AGEVAL Step 2 — ODK XLS Form Generator
Run: python scripts/s02_generate_odk_form.py

Reads variables_master.csv and produces a valid ODK XLSForm
ready to upload to KoboCollect or ODK Central.
"""

import pandas as pd
import os
import yaml
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from datetime import datetime

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE   = os.path.dirname(os.path.abspath(__file__))
ROOT   = os.path.join(BASE, "..")
CFG    = os.path.join(ROOT, "config", "config.yaml")
DICT   = os.path.join(ROOT, "dictionary", "variables_personalized.csv")
OUTDIR = os.path.join(ROOT, "forms")

def load_config():
    with open(CFG) as f:
        return yaml.safe_load(f)

def load_dict():
    df = pd.read_csv(DICT)
    # Only include rows flagged for questionnaire
    return df[df["questionnaire_include"] == 1].copy()

# ── Build survey sheet rows ────────────────────────────────────────────────────
def build_survey(df, cfg):
    lang = cfg["project"]["language"]
    label_col = "label_spanish" if lang == "spanish" else "label_english"
    rows = []

    # Form metadata
    rows.append({"type": "start", "name": "start", "label::Spanish (es)": "", "label::English (en)": ""})
    rows.append({"type": "end",   "name": "end",   "label::Spanish (es)": "", "label::English (en)": ""})
    rows.append({"type": "today", "name": "today", "label::Spanish (es)": "", "label::English (en)": ""})
    rows.append({"type": "deviceid", "name": "deviceid", "label::Spanish (es)": "", "label::English (en)": ""})

    # Section: Household
    topics_order = ["quality_meta", "household", "farm", "production", "market", "finance", "inputs", "technology", "wtp", "geospatial"]
    section_labels = {
        "quality_meta":  ("Información del Encuestador", "Enumerator Information"),
        "household":     ("Información del Hogar",       "Household Information"),
        "farm":          ("Información de la Finca",      "Farm Information"),
        "production":    ("Producción de Café",           "Coffee Production"),
        "market":        ("Mercado y Ventas",              "Market & Sales"),
        "finance":       ("Finanzas y Crédito",            "Finance & Credit"),
        "inputs":        ("Insumos",                       "Inputs"),
        "technology":    ("Tecnología",                    "Technology Adoption"),
        "geospatial":    ("Ubicación GPS",                 "GPS Location"),
    }

    for topic in topics_order:
        topic_rows = df[df["topic"] == topic]
        if topic_rows.empty:
            continue

        label_es, label_en = section_labels.get(topic, (topic.upper(), topic.upper()))

        # Begin group
        rows.append({
            "type":               f"begin_group",
            "name":               f"grp_{topic}",
            "label::Spanish (es)": label_es,
            "label::English (en)": label_en,
            "appearance":          "field-list",
        })

        for _, row in topic_rows.iterrows():
            vname = row["variable_name"]
            qtype = row["question_type"]

            # select_one / select_multiple get a list name
            if qtype in ("select_one", "select_multiple"):
                qtype_full = f"{qtype} {vname}_choices"
            else:
                qtype_full = qtype

            survey_row = {
                "type":                qtype_full,
                "name":                vname,
                "label::Spanish (es)": row.get("label_spanish", vname),
                "label::English (en)": row.get("label_english", vname),
                "required":            "yes" if row.get("required", 0) == 1 else "no",
            }

            # Constraint
            if pd.notna(row.get("constraint")) and str(row["constraint"]).strip():
                survey_row["constraint"]         = row["constraint"]
                survey_row["constraint_message"] = row.get("constraint_message", "Invalid value")

            rows.append(survey_row)

            # Add ODK calculate immediately after raw variable
            if row.get("odk_calculate_include", 0) == 1:
                expr   = row.get("odk_calculate_expr", "")
                output = row.get("odk_calculate_output", f"{vname}_calc")
                if pd.notna(expr) and str(expr).strip():
                    rows.append({
                        "type":                "calculate",
                        "name":                output,
                        "label::Spanish (es)": f"[calc] {output}",
                        "label::English (en)": f"[calc] {output}",
                        "calculation":         expr,
                    })

        rows.append({"type": "end_group", "name": f"grp_{topic}"})

    return pd.DataFrame(rows)

# ── Build choices sheet ────────────────────────────────────────────────────────
def build_choices(df):
    rows = []
    choice_rows = df[df["question_type"].isin(["select_one", "select_multiple"])]
    for _, row in choice_rows.iterrows():
        list_name = f"{row['variable_name']}_choices"
        choices_raw = row.get("choices", "")
        if not pd.notna(choices_raw) or not str(choices_raw).strip():
            continue
        for pair in str(choices_raw).split(";"):
            pair = pair.strip()
            if ":" not in pair:
                continue
            val, lbl = pair.split(":", 1)
            rows.append({
                "list_name":           list_name,
                "name":                val.strip(),
                "label::Spanish (es)": lbl.strip(),
                "label::English (en)": lbl.strip(),
            })
    return pd.DataFrame(rows)

# ── Build settings sheet ───────────────────────────────────────────────────────
def build_settings(cfg):
    return pd.DataFrame([{
        "form_title":            cfg["project"]["name"],
        "form_id":               cfg["project"]["form_id"],
        "version":               cfg["project"]["form_version"],
        "default_language":      cfg["odk"]["default_language"],
        "style":                 "pages",
        "instance_name":         "concat(${respondent_id}, '_', today())",
    }])

# ── Style helpers ──────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill("solid", fgColor="1D4ED8")
CALC_FILL   = PatternFill("solid", fgColor="D1FAE5")
GROUP_FILL  = PatternFill("solid", fgColor="E0E7FF")

def style_sheet(ws):
    for cell in ws[1]:
        cell.fill  = HEADER_FILL
        cell.font  = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=2):
        type_val = str(row[0].value or "")
        if "calculate" in type_val:
            for cell in row:
                cell.fill = CALC_FILL
        elif "group" in type_val:
            for cell in row:
                cell.fill = GROUP_FILL
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)

# ── Main ───────────────────────────────────────────────────────────────────────
def generate_form():
    cfg = load_config()
    df  = load_dict()

    survey   = build_survey(df, cfg)
    choices  = build_choices(df)
    settings = build_settings(cfg)

    ts       = datetime.now().strftime("%Y%m%d_%H%M")
    form_id  = cfg["project"]["form_id"]
    out_path = os.path.join(OUTDIR, f"{form_id}_{ts}.xlsx")
    os.makedirs(OUTDIR, exist_ok=True)

    wb = Workbook()

    def write_sheet(wb, name, df, first=False):
        ws = wb.active if first else wb.create_sheet(name)
        ws.title = name
        ws.append(list(df.columns))
        for _, row in df.iterrows():
            ws.append([row.get(c, "") for c in df.columns])
        style_sheet(ws)
        return ws

    write_sheet(wb, "survey",   survey,   first=True)
    write_sheet(wb, "choices",  choices)
    write_sheet(wb, "settings", settings)

    wb.save(out_path)
    print(f"✅  ODK XLS form saved: {out_path}")
    print(f"    Survey rows:  {len(survey)}")
    print(f"    Choice rows:  {len(choices)}")
    return out_path

if __name__ == "__main__":
    generate_form()
