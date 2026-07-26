"""
AGEVAL — Generate synthetic test data
Run: python scripts/generate_test_data.py

Creates a realistic synthetic dataset that mimics a KoboCollect CSV export
so you can test Steps 3, 4, and 5 without real data.
"""

import pandas as pd
import numpy as np
import os

BASE    = os.path.dirname(os.path.abspath(__file__))
ROOT    = os.path.join(BASE, "..")
OUTDIR  = os.path.join(ROOT, "data_raw")
os.makedirs(OUTDIR, exist_ok=True)

np.random.seed(42)
N = 200  # number of simulated farmers

def generate():
    # Metadata
    respondent_id     = [f"RES{str(i).zfill(4)}" for i in range(1, N+1)]
    enumerator_id     = np.random.choice(["E01", "E02", "E03", "E04"], N)
    interview_duration_min = np.random.normal(45, 12, N).clip(10, 120).astype(int)

    # Household
    respondent_age    = np.random.normal(45, 12, N).clip(18, 80).astype(int)
    respondent_sex    = np.random.choice(["male", "female"], N, p=[0.72, 0.28])
    household_size    = np.random.randint(2, 9, N)
    years_farming     = np.random.normal(18, 8, N).clip(1, 50).astype(int)
    education_level   = np.random.choice(
        ["none", "primary", "secondary", "technical", "university"],
        N, p=[0.10, 0.45, 0.30, 0.10, 0.05]
    )

    # Farm
    farm_count        = np.random.choice([1, 2, 3], N, p=[0.70, 0.22, 0.08])
    area_unit         = np.random.choice(["manzana", "ha"], N, p=[0.80, 0.20])
    area_planted_coffee_raw = np.random.lognormal(0.5, 0.7, N).clip(0.1, 20)

    # Standardize to ha
    area_planted_coffee_ha = np.where(
        area_unit == "manzana",
        area_planted_coffee_raw * 0.7,
        area_planted_coffee_raw
    ).round(3)

    area_total_farm_raw = area_planted_coffee_raw * np.random.uniform(1.2, 3.0, N)
    area_total_farm_ha  = np.where(
        area_unit == "manzana",
        area_total_farm_raw * 0.7,
        area_total_farm_raw
    ).round(3)

    # Production
    coffee_variety    = np.random.choice(
        ["catuai", "caturra", "lempira", "ihcafe90", "bourbon", "other"],
        N, p=[0.30, 0.25, 0.15, 0.15, 0.10, 0.05]
    )
    uses_improved_variety = np.random.choice(["yes", "no"], N, p=[0.55, 0.45])

    # Yield (kg/ha) — improved variety produces more
    base_yield = np.random.normal(800, 300, N)
    yield_bonus = np.where(uses_improved_variety == "yes", 200, 0)
    coffee_yield_kg = (base_yield + yield_bonus + area_planted_coffee_ha * 50).clip(100, 3000).round(1)

    yield_unit = np.full(N, "kg")

    # Market
    coffee_price_received_usd = np.random.normal(130, 30, N).clip(60, 280).round(2)
    sells_to_cooperative      = np.random.choice(["yes", "no"], N, p=[0.60, 0.40])

    # Finance
    has_formal_credit    = np.random.choice(["yes", "no"], N, p=[0.35, 0.65])
    credit_amount_usd    = np.where(
        has_formal_credit == "yes",
        np.random.lognormal(7, 1, N).clip(100, 20000),
        0
    ).round(2)

    # Inputs
    uses_fertilizer      = np.random.choice(["yes", "no"], N, p=[0.75, 0.25])
    fertilizer_cost_usd  = np.where(
        uses_fertilizer == "yes",
        np.random.normal(400, 150, N).clip(50, 1500),
        0
    ).round(2)

    # WTP
    wtp_certification_usd = np.random.lognormal(3.5, 1.0, N).clip(0, 300).round(2)

    # GPS (Honduras bounding box)
    gps_latitude  = np.random.uniform(13.5, 15.5, N).round(6)
    gps_longitude = np.random.uniform(-89.0, -83.0, N).round(6)

    # Introduce some data quality issues for testing
    # ~3% missing on yield
    missing_yield_idx = np.random.choice(N, int(N * 0.03), replace=False)
    coffee_yield_kg_series = pd.Series(coffee_yield_kg)
    coffee_yield_kg_series.iloc[missing_yield_idx] = np.nan

    # ~2 outliers on credit
    credit_series = pd.Series(credit_amount_usd)
    credit_series.iloc[0] = 99999   # obvious outlier

    df = pd.DataFrame({
        "respondent_id":           respondent_id,
        "enumerator_id":           enumerator_id,
        "interview_duration_min":  interview_duration_min,
        "respondent_age":          respondent_age,
        "respondent_sex":          respondent_sex,
        "household_size":          household_size,
        "years_farming":           years_farming,
        "education_level":         education_level,
        "farm_count":              farm_count,
        "area_unit":               area_unit,
        "area_planted_coffee_raw": area_planted_coffee_raw.round(2),
        "area_planted_coffee_ha":  area_planted_coffee_ha,
        "area_total_farm_raw":     area_total_farm_raw.round(2),
        "area_total_farm_ha":      area_total_farm_ha,
        "coffee_variety":          coffee_variety,
        "uses_improved_variety":   uses_improved_variety,
        "coffee_yield_raw":        coffee_yield_kg_series.round(1),
        "coffee_yield_kg":         coffee_yield_kg_series.round(1),
        "yield_unit":              yield_unit,
        "coffee_price_received_usd": coffee_price_received_usd,
        "sells_to_cooperative":    sells_to_cooperative,
        "has_formal_credit":       has_formal_credit,
        "credit_amount_usd":       credit_series,
        "uses_fertilizer":         uses_fertilizer,
        "fertilizer_cost_usd":     fertilizer_cost_usd,
        "wtp_certification_usd":   wtp_certification_usd,
        "gps_latitude":            gps_latitude,
        "gps_longitude":           gps_longitude,
    })

    out_path = os.path.join(OUTDIR, "test_data_honduras_n200.csv")
    df.to_csv(out_path, index=False)
    print(f"✅  Synthetic test data saved: {out_path}")
    print(f"    Rows: {len(df)} | Columns: {len(df.columns)}")
    return out_path

if __name__ == "__main__":
    generate()
