"""
AGEVAL Step 5 — Apply Corrections
Run: python baseline/01_code/s05_apply_corrections.py --data baseline/04_data/dataRaw/test_data_honduras_n2052.xlsx --template baseline/04_data/correctionFiles/correction_template_test_data_honduras_n2052.xlsx

Reads a correction template (produced by s04_correction_template.py) after it
has been reviewed by the survey supervision team, and applies the completed
corrections back onto the original dataset — now across every sheet (general/
main sheet plus every loop/repeat sheet), and recalculating any calculated
variable that depends on a corrected value.

  * The template can have multiple tabs, one per data sheet. Each tab's
    title row states which original sheet it belongs to; that's used to map
    the tab back to the right sheet, regardless of Excel's 31-character tab
    name truncation.
  * Matches each template row to its observation using the survey ID; for
    loop/repeat sheets, the observation id (`_obs_id`, Kobo/ODK's own
    `_index` for that repeat instance) narrows it down to the exact row. On
    the rare occasions the plain survey ID alone is ambiguous on the general
    sheet (e.g. a flagged duplicate respondent_id), start/end/deviceid are
    used to disambiguate, same as before.
  * For each variable, reads only the "Corrected value" column — the
    "Correction reason" column is documentation only and is never applied.
    The one exception is the "Duplicate detection" group: its "Corrected
    value" isn't a data variable but a fix to the survey ID column itself
    (e.g. a flagged duplicate `respondent_id` where the reviewer determined
    one of the two records is actually a different respondent). That rename
    is applied first, directly on the general sheet, before loop sheets are
    linked to it — so the correction cascades to that respondent's
    loop/repeat rows too, and every other correction for that same record
    lines up against the corrected ID.
  * A variable's original value is left untouched unless a corrected value
    was actually entered.
  * No columns are added, removed, or renamed — only existing variable
    values are overwritten, in place, using the original column's dtype
    whenever possible.
  * After corrections are applied, every calculated variable
    (`surv_type == "calculate"`) is rebuilt from `surv_calculation`, in
    dependency order, so that e.g. a corrected plot-level `area_total`
    automatically refreshes `area_total_ha`, and in turn the respondent-level
    `farm_area_total_ha` aggregate that sums it across the loop sheet — all
    driven generically by the dictionary, never hard-coded per variable.
  * The result is saved as a new file (original name + suffix, "_clean" by
    default) so the original dataset is never modified. XLSX input produces
    an XLSX output with the same sheets/columns/order; legacy flat CSV input
    produces a CSV output.

Structured as a small set of pure functions (parse layout -> match rows ->
coerce value -> apply -> recalculate) so it can be called directly from a
future Streamlit app (e.g. one button per step: generate template / upload
completed template / download cleaned dataset) without any changes.
"""

import re
import pandas as pd
import numpy as np
import os
import sys
import argparse
import yaml
from openpyxl import load_workbook
import s04_correction_template as tmpl

# ── Reuse config/dictionary paths and template layout constants from s04 ──────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

qc = tmpl.qc  # s03_quality_check, re-exported by s04

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
CORRECTIONS    = os.path.join(BASE, config["paths"]["correction_files"].format(survey_round=SURVEY_ROUND))
DATACLEAN   = os.path.join(BASE, config["paths"]["data_clean"].format(survey_round=SURVEY_ROUND))


GROUP_HEADER_ROW = tmpl.GROUP_HEADER_ROW
COL_HEADER_ROW   = tmpl.COL_HEADER_ROW
DATA_START_ROW   = tmpl.DATA_START_ROW
SUBCOLS          = tmpl.SUBCOLS
CORRECTED_VALUE_OFFSET = SUBCOLS.index("Corrected value")
DUPLICATES_LABEL = "Duplicate detection"  # its "Corrected value" fixes id_col itself, not a data variable — see resolve_duplicate_ids

DISPLAY_TO_INTERNAL = {v: k for k, v in tmpl.ID_DISPLAY_NAMES.items()}
SHEET_LABEL_PATTERN = re.compile(r"Sheet:\s*(.*?)\s*\|")


# ── Template layout parsing ──────────────────────────────────────────────────
def extract_sheet_label(ws):
    """Recovers the original data-sheet name from a template tab's title row
    (written verbatim by s04, unlike the tab title itself, which may be
    truncated/suffixed to fit Excel's 31-character sheet-name limit)."""
    title_text = ws.cell(row=1, column=1).value or ""
    m = SHEET_LABEL_PATTERN.search(str(title_text))
    return m.group(1) if m else None


def detect_id_block_width(ws):
    """The identifying columns run from column 1 up to (but not including)
    the first variable group's 'Original value' sub-header."""
    for c in range(1, ws.max_column + 1):
        if ws.cell(row=COL_HEADER_ROW, column=c).value is not None:
            return c - 1
    return ws.max_column


def parse_sheet_layout(ws):
    """
    Inspects one template tab and returns:
      id_col_idx: {internal_id_col_name: column_index}
      var_groups: {var_key: corrected_value_column_index}
      dup_corrected_col: column index of the "Duplicate detection" group's
        "Corrected value" sub-column, or None if that group isn't on this
        tab (or has no such sub-column). This one is pulled out separately
        from var_groups — a duplicate-ID fix isn't a data-variable
        correction, it's a fix to the survey ID column itself, and is
        applied by a dedicated pass (see resolve_duplicate_ids) rather than
        the generic per-variable loop.
    """
    n_id_cols = detect_id_block_width(ws)
    id_col_idx = {}
    for i in range(1, n_id_cols + 1):
        display_name = ws.cell(row=GROUP_HEADER_ROW, column=i).value
        internal_name = DISPLAY_TO_INTERNAL.get(display_name, display_name)
        if internal_name:
            id_col_idx[internal_name] = i

    var_groups = {}
    dup_corrected_col = None
    for c in range(n_id_cols + 1, ws.max_column + 1):
        subheader = ws.cell(row=COL_HEADER_ROW, column=c).value
        if subheader != "Corrected value":
            continue

        group_header = ws.cell(row=GROUP_HEADER_ROW, column=c - CORRECTED_VALUE_OFFSET).value
        if not group_header or not str(group_header).strip():
            continue
        group_header = str(group_header).strip()

        if group_header == DUPLICATES_LABEL:
            dup_corrected_col = c
            continue

        # Group headers are written as "<var_key> — <label>" by s04; fall
        # back to using the whole header as the key if no separator is found.
        var_key = group_header.split(" — ", 1)[0].strip() if " — " in group_header else group_header
        var_groups[var_key] = c

    return id_col_idx, var_groups, dup_corrected_col


# ── Row matching ─────────────────────────────────────────────────────────────
def match_rows(df, id_col, id_values):
    """
    Returns the list of `df` index labels corresponding to a template row's
    identifying columns (`id_values`, keyed by internal id column name).

    IMPORTANT: `_obs_id` (Kobo/ODK's `_index` for a repeat instance) is only
    unique *within* a respondent's submission -- e.g. every respondent's
    first plot is `_index == 1` -- so it is NEVER matched on its own. It is
    always combined with `respondent_id` first, narrowing to that
    respondent's rows before picking the exact repeat instance.

    The general sheet is matched on the survey ID, disambiguated using
    whatever other identifying columns the template carries for that sheet
    (e.g. `surveyDate`, or `start`/`end`/`deviceid` if present) in case the
    same ID happens to appear more than once (e.g. a flagged duplicate
    record).
    """
    resp_id = id_values.get("respondent_id")
    obs_id = id_values.get("_obs_id")

    if resp_id not in (None, "") and id_col in df.columns:
        resp_candidates = df.index[df[id_col].astype(str) == str(resp_id)]
    elif id_col in df.columns:
        resp_candidates = pd.Index([])
    else:
        resp_candidates = df.index

    if obs_id not in (None, "") and "_obs_id" in df.columns:
        sub = df.loc[resp_candidates]
        narrowed = sub.index[sub["_obs_id"].astype(str) == str(obs_id)]
        return list(narrowed)

    if resp_id in (None, "") or id_col not in df.columns:
        return []

    candidates = resp_candidates
    if len(candidates) <= 1:
        return list(candidates)

    # Disambiguate using any other identifying column the template captured
    # for this row (surveyDate, start, end, deviceid, ...) that also exists
    # in the data -- generic, rather than a hardcoded column-name list.
    sub = df.loc[candidates]
    mask = pd.Series(True, index=sub.index)
    for col_name, val in id_values.items():
        if col_name in ("respondent_id", "_obs_id"):
            continue
        if val in (None, "") or col_name not in sub.columns:
            continue
        mask &= sub[col_name].apply(lambda x: _id_values_match(x, val))

    narrowed = sub.index[mask]
    return list(narrowed) if len(narrowed) > 0 else list(candidates)


def _id_values_match(a, b):
    """Compares two identifying-column values for row disambiguation.
    Falls back to a same-day comparison for date/datetime values, since a
    template column like 'surveyDate' (date-only) and a raw column of the
    same name (which may carry a midnight timestamp) would otherwise never
    string-match; plain values fall back to a (truncated) string compare,
    same as before, to tolerate millisecond-precision timestamp noise."""
    if is_blank(a) or is_blank(b):
        return False
    try:
        ta, tb = pd.to_datetime(a), pd.to_datetime(b)
        if pd.notna(ta) and pd.notna(tb):
            return ta == tb or ta.date() == tb.date()
    except (ValueError, TypeError):
        pass
    return str(a)[:19] == str(b)[:19]


# ── Value coercion ───────────────────────────────────────────────────────────
def coerce_to_dtype(value, reference_series):
    """Casts a corrected value to match the dtype of the column being updated,
    falling back to the raw value if it can't be cleanly converted."""
    if isinstance(value, str):
        value = value.strip()

    dtype = reference_series.dtype

    if pd.api.types.is_datetime64_any_dtype(dtype):
        try:
            return pd.to_datetime(value)
        except (ValueError, TypeError):
            return value

    if pd.api.types.is_bool_dtype(dtype):
        if isinstance(value, bool):
            return value
        return str(value).strip().lower() in ("1", "true", "yes", "y", "si", "sí")

    if pd.api.types.is_integer_dtype(dtype):
        try:
            f = float(value)
            return int(f) if f.is_integer() else f
        except (TypeError, ValueError):
            return value

    if pd.api.types.is_float_dtype(dtype):
        try:
            return float(value)
        except (TypeError, ValueError):
            return value

    return str(value)


def is_blank(value):
    return value is None or (isinstance(value, str) and value.strip() == "")


# ── Duplicate-ID resolution ───────────────────────────────────────────────────
def resolve_duplicate_ids(wb, main_df, main_sheet_name, id_col):
    """
    Applies "Duplicate detection" corrections: unlike every other group, its
    "Corrected value" isn't a data variable — it's a fix to the survey ID
    column itself (e.g. two records were flagged as duplicate `respondent_id
    == h1`, and the reviewer determined the second one is actually a
    different respondent, `h2`).

    Must run before loop sheets are linked to the main sheet (and before any
    other correction is matched), so the corrected ID is what everything
    downstream — loop linkage, row matching, recalculation — sees. Mutates
    `main_df` in place.

    Only the tab matching `main_sheet_name` is considered: duplicate IDs are
    a respondent-level concept, and loop/repeat rows are already
    disambiguated by their own `_obs_id`, not by the survey ID.

    Returns:
      applied: list of (old_id, new_id) pairs actually applied, for the
        summary — a row whose "Corrected value" repeats its own original ID
        is not counted (no-op).
      row_match_cache: {(worksheet_title, row_number): target_rows}, the
        match_rows() result for every identifying template row on the main
        tab, computed BEFORE any rename above. The main correction loop
        reuses this for the same (tab, row) instead of matching again
        afterwards — re-matching a duplicate-corrected row by its *original*
        ID once the rename has moved it out of that ID's group would
        otherwise silently land on the wrong surviving row, if that same
        template row also carries an ordinary variable correction.
    """
    applied = []
    row_match_cache = {}
    for ws in wb.worksheets:
        if extract_sheet_label(ws) != main_sheet_name:
            continue

        id_col_idx, _, dup_corrected_col = parse_sheet_layout(ws)
        if not id_col_idx:
            continue

        for r in range(DATA_START_ROW, ws.max_row + 1):
            id_values = {name: ws.cell(row=r, column=c).value for name, c in id_col_idx.items()}
            if is_blank(id_values.get("respondent_id")) and is_blank(id_values.get("_obs_id")):
                continue

            target_rows = match_rows(main_df, id_col, id_values)
            row_match_cache[(ws.title, r)] = target_rows

            if dup_corrected_col is None or not target_rows:
                continue

            corrected_val = ws.cell(row=r, column=dup_corrected_col).value
            if is_blank(corrected_val):
                continue

            coerced = coerce_to_dtype(corrected_val, main_df[id_col])
            old_id = id_values.get("respondent_id")
            if str(coerced) == str(old_id):
                continue  # corrected value just repeats the original ID

            for idx in target_rows:
                main_df.at[idx, id_col] = coerced
            applied.append((str(old_id), str(coerced)))

    return applied, row_match_cache


# ── Calculated-variable recalculation ────────────────────────────────────────
# Aggregate functions ODK/XLSForm formulas commonly wrap a cross-sheet (loop)
# reference in, e.g. "sum(${area_total_ha})". Generic — not tied to any one
# calculated variable — so any future dictionary formula using one of these
# over a loop-sheet column is supported automatically.
AGG_FUNCS = {
    "sum":   lambda vals: float(np.sum(vals)) if vals else 0.0,
    "mean":  lambda vals: float(np.mean(vals)) if vals else float("nan"),
    "avg":   lambda vals: float(np.mean(vals)) if vals else float("nan"),
    "min":   lambda vals: float(np.min(vals)) if vals else float("nan"),
    "max":   lambda vals: float(np.max(vals)) if vals else float("nan"),
    "count": lambda vals: len(vals),
}
AGG_PATTERN = re.compile(r"\b(sum|mean|avg|min|max|count)\(\s*\$\{([a-zA-Z0-9_]+)\}\s*\)")
VAR_PATTERN = re.compile(r"\$\{([a-zA-Z0-9_]+)\}")


def _split_top_level_commas(s):
    parts, depth, current = [], 0, []
    for ch in s:
        if ch in "([":
            depth += 1
            current.append(ch)
        elif ch in ")]":
            depth -= 1
            current.append(ch)
        elif ch == "," and depth == 0:
            parts.append("".join(current))
            current = []
        else:
            current.append(ch)
    parts.append("".join(current))
    return parts


def convert_if_calls(expr):
    """
    Recursively converts ODK/XLSForm `if(cond, a, b)` calls — including
    nested ones — into native Python conditional expressions
    `((a) if (cond) else (b))`. This preserves short-circuit evaluation
    (e.g. a division-by-zero guard like `if(${a}>0, ${b}/${a}, 0)`), which a
    naive `IF(cond, a, b)` *function* call would not: Python only evaluates
    the selected branch of a conditional expression, so an unselected
    branch referencing an unavailable variable never raises.
    """
    s = expr
    search_from = 0
    while True:
        idx = s.find("if(", search_from)
        if idx == -1:
            break
        if idx > 0 and (s[idx - 1].isalnum() or s[idx - 1] == "_"):
            search_from = idx + 1
            continue

        start = idx + 3
        depth = 1
        i = start
        while i < len(s) and depth > 0:
            if s[i] == "(":
                depth += 1
            elif s[i] == ")":
                depth -= 1
            i += 1
        end = i - 1  # index of the matching ")"

        parts = _split_top_level_commas(s[start:end])
        if len(parts) != 3:
            search_from = idx + 3  # malformed; skip past and keep looking
            continue

        cond, a, b = (convert_if_calls(p.strip()) for p in parts)
        replacement = f"(({a}) if ({cond}) else ({b}))"
        s = s[:idx] + replacement + s[end + 1:]
        search_from = idx  # re-scan from here in case of adjacent if()s

    return s


def prep_calc_expr(expr):
    e = convert_if_calls(str(expr).strip())
    e = re.sub(r"(?<![=!<>])=(?!=)", "==", e)  # single '=' -> '==' (ODK style)
    return e


def build_var_to_sheet(dict_df, sheets):
    var_to_sheet = {}
    for vname in dict_df["variable_name"].dropna().unique():
        for sheet_name, df in sheets.items():
            if vname in df.columns:
                var_to_sheet[vname] = sheet_name
                break
    return var_to_sheet


def topo_sort_calc_vars(calc_vars, calc_lookup, type_lookup):
    """Orders calculated variables so that any calculated variable a formula
    depends on is recalculated first (dependencies before dependents)."""
    order, visited, in_progress = [], set(), set()

    def visit(v):
        if v in visited or v in in_progress:
            return
        in_progress.add(v)
        deps = qc._extract_odk_vars(calc_lookup.get(v)) if pd.notna(calc_lookup.get(v)) else []
        for d in deps:
            if d in calc_vars and tmpl._is_calculated(d, type_lookup):
                visit(d)
        in_progress.discard(v)
        visited.add(v)
        order.append(v)

    for v in calc_vars:
        visit(v)
    return order


def recalculate_variables(working, dict_df, id_col):
    """
    Rebuilds every calculated variable (`surv_type == "calculate"`) in
    dependency order, generically from `surv_calculation` — no variable
    names are hard-coded. Works across sheets: a same-sheet formula (e.g. a
    plot-level `area_total_ha`) is evaluated row by row; a formula
    aggregating a loop sheet's column into a general-sheet value (e.g.
    `sum(${area_total_ha})`) is evaluated once per respondent, gathering
    that respondent's rows from the dependency's own sheet.

    Returns a list of (variable_name, row_index) pairs that could not be
    recalculated (left at their prior value) — e.g. a formula referencing a
    variable that no longer exists in the data.
    """
    type_lookup = dict_df.set_index("variable_name")["surv_type"].to_dict()
    calc_lookup = dict_df.set_index("variable_name")["surv_calculation"].to_dict()

    var_to_sheet = build_var_to_sheet(dict_df, working)

    calc_vars = {
        v for v in dict_df["variable_name"].dropna().unique()
        if tmpl._is_calculated(v, type_lookup) and pd.notna(calc_lookup.get(v)) and v in var_to_sheet
    }
    order = topo_sort_calc_vars(calc_vars, calc_lookup, type_lookup)

    # Row positions per sheet grouped by respondent, for aggregate lookups.
    respondent_positions = {
        sn: (df.groupby(df[id_col].astype(str)).groups if id_col in df.columns else {})
        for sn, df in working.items()
    }

    skipped = []
    for vname in order:
        sheet_name = var_to_sheet[vname]
        df = working[sheet_name]
        expr_template = prep_calc_expr(calc_lookup[vname])

        for idx, row in df.iterrows():
            expr = expr_template

            def _agg_repl(m):
                func, var = m.group(1), m.group(2)
                dep_sheet = var_to_sheet.get(var)
                if dep_sheet is None:
                    return "None"
                if dep_sheet == sheet_name:
                    values = [row.get(var)]
                else:
                    rid = str(row.get(id_col))
                    positions = respondent_positions.get(dep_sheet, {}).get(rid, [])
                    dep_df = working[dep_sheet]
                    values = [dep_df.at[p, var] for p in positions] if var in dep_df.columns else []
                numeric = pd.to_numeric(pd.Series(values), errors="coerce").dropna().tolist()
                if not numeric:
                    return "None"
                return repr(AGG_FUNCS[func](numeric))

            expr = AGG_PATTERN.sub(_agg_repl, expr)

            def _var_repl(m):
                var = m.group(1)
                dep_sheet = var_to_sheet.get(var)
                if dep_sheet is not None and dep_sheet != sheet_name:
                    return "None"  # cross-sheet ref without aggregation: unresolvable
                val = qc.coerce_numeric(row.get(var))
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    return "None"
                return repr(val)

            expr = VAR_PATTERN.sub(_var_repl, expr)

            try:
                result = eval(expr, {"__builtins__": {}}, {})
            except Exception:
                skipped.append((vname, idx))
                continue

            if isinstance(result, float) and not pd.isna(result):
                result = round(result, 6)  # avoid float round-trip noise (e.g. 1e-16) on unaffected rows

            df.at[idx, vname] = result

        working[sheet_name] = df

    return skipped


# ── Main ─────────────────────────────────────────────────────────────────────
def apply_corrections(data_path, template_path, id_col=None, output_suffix="_clean", output_path=None, dict_path=None):
    """
    Applies a completed correction template onto the original dataset,
    recalculates every calculated variable, and saves the result as a new
    file. Returns (output_path, summary dict).
    """
    dict_df = pd.read_csv(dict_path or DICT)

    sheets_raw, main_sheet_name = qc.load_sheets(data_path, qc.cfg)
    if id_col is None:
        id_col = "respondent_id" if "respondent_id" in sheets_raw[main_sheet_name].columns else list(sheets_raw[main_sheet_name].columns)[0]

    # Preserve every original column (dictionary-tracked or not), in order,
    # so the output never gains or loses a column.
    original_columns = {sn: list(df.columns) for sn, df in sheets_raw.items()}

    wb = load_workbook(template_path, data_only=True)

    # Duplicate-ID corrections run first and directly on the raw main sheet
    # (see resolve_duplicate_ids), so loop-sheet linkage below already keys
    # off the corrected ID rather than the flagged duplicate.
    duplicate_id_updates, main_row_match_cache = resolve_duplicate_ids(
        wb, sheets_raw[main_sheet_name], main_sheet_name, id_col
    )

    # Working copies. Loop/repeat sheets get respondent_id/_obs_id attached
    # (same helper columns s04 uses) purely to support matching/recalculation
    # — these are dropped again before saving.
    working = {main_sheet_name: sheets_raw[main_sheet_name].copy()}
    for sheet_name, raw_df in sheets_raw.items():
        if sheet_name == main_sheet_name:
            continue
        working[sheet_name] = qc.link_loop_to_main(raw_df, working[main_sheet_name], id_col=id_col, date_col="surveyDate")

    updated_cells = 0
    updated_records = set()
    unmatched = []            # (sheet_label, id_values) that couldn't be matched
    skipped_vars = set()      # var_key referenced in the template but not found in the data

    for ws in wb.worksheets:
        sheet_label = extract_sheet_label(ws)
        if sheet_label is None or sheet_label not in working:
            continue  # not a recognizable data sheet tab (or sheet no longer exists)

        df = working[sheet_label]
        id_col_idx, var_groups, _dup_corrected_col = parse_sheet_layout(ws)
        if not var_groups:
            continue  # no corrigible variable groups on this tab

        for r in range(DATA_START_ROW, ws.max_row + 1):
            id_values = {name: ws.cell(row=r, column=c).value for name, c in id_col_idx.items()}
            if is_blank(id_values.get("respondent_id")) and is_blank(id_values.get("_obs_id")):
                continue

            target_rows = (
                main_row_match_cache.get((ws.title, r))
                if sheet_label == main_sheet_name
                else match_rows(df, id_col, id_values)
            )
            if target_rows is None:
                target_rows = match_rows(df, id_col, id_values)  # fallback; cache should always have this row
            if not target_rows:
                unmatched.append((sheet_label, id_values.get("respondent_id") or id_values.get("_obs_id")))
                continue

            for var_key, corrected_col in var_groups.items():
                corrected_val = ws.cell(row=r, column=corrected_col).value
                if is_blank(corrected_val):
                    continue

                if var_key not in df.columns:
                    skipped_vars.add(f"{sheet_label}:{var_key}")
                    continue

                coerced = coerce_to_dtype(corrected_val, df[var_key])
                for idx in target_rows:
                    df.at[idx, var_key] = coerced

                updated_cells += 1
                updated_records.add((sheet_label, str(id_values.get("respondent_id") or id_values.get("_obs_id"))))

        working[sheet_label] = df

    skipped_calc = recalculate_variables(working, dict_df, id_col)

    if output_path is None:
        base, ext = os.path.splitext(os.path.basename(data_path))
        ext = ext if ext else (".xlsx" if qc.is_excel_file(data_path) else ".csv")
        output_path = os.path.join(DATACLEAN, f"{base}{output_suffix}{ext}")
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

    if qc.is_excel_file(data_path):
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for sheet_name, cols in original_columns.items():
                working[sheet_name][cols].to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        working[main_sheet_name][original_columns[main_sheet_name]].to_csv(output_path, index=False)

    summary = {
        "output_path": output_path,
        "records_updated": len(updated_records),
        "cells_updated": updated_cells,
        "duplicate_ids_resolved": duplicate_id_updates,
        "unmatched": unmatched,
        "skipped_variables": sorted(skipped_vars),
        "skipped_calculations": skipped_calc,
    }
    return output_path, summary


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="AGEVAL Apply Corrections")
    parser.add_argument("--data",     required=True, help="Path to the original collected data file (.xlsx with general + loop/repeat sheets, or legacy flat .csv)")
    parser.add_argument("--template", required=True, help="Path to the completed correction template (.xlsx)")
    parser.add_argument("--id-col",   default=None,  help="Survey ID column name (default: respondent_id)")
    parser.add_argument("--suffix",   default="_clean", help="Suffix appended to the output file name")
    parser.add_argument("--output",   default=None,  help="Optional explicit output path (overrides --suffix)")
    parser.add_argument("--dict",     default=None,  help="Path to personalized dictionary CSV (default: config path)")
    args = parser.parse_args()

    out_path, summary = apply_corrections(
        args.data, args.template, id_col=args.id_col, output_suffix=args.suffix, output_path=args.output,
        dict_path=args.dict,
    )

    print(f"✅  Cleaned dataset saved: {out_path}")
    print(f"    Records updated:        {summary['records_updated']}")
    print(f"    Cells updated:          {summary['cells_updated']}")
    print(f"    Calculated variables recalculated (with any per-row exceptions noted below)")
    if summary["duplicate_ids_resolved"]:
        pairs = ", ".join(f"{old} → {new}" for old, new in summary["duplicate_ids_resolved"])
        print(f"    🔁 Duplicate IDs resolved: {pairs}")
    if summary["unmatched"]:
        print(f"    ⚠️  Template rows not found in original dataset: {summary['unmatched']}")
    if summary["skipped_variables"]:
        print(f"    ⚠️  Variables skipped (no matching column in dataset): {', '.join(summary['skipped_variables'])}")
    if summary["skipped_calculations"]:
        by_var = {}
        for vname, idx in summary["skipped_calculations"]:
            by_var.setdefault(vname, 0)
            by_var[vname] += 1
        detail = ", ".join(f"{v} ({n} rows)" for v, n in by_var.items())
        print(f"    ⚠️  Could not recalculate: {detail}")