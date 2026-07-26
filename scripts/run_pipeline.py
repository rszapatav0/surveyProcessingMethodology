"""
AGEVAL Master Pipeline Runner
Run: python scripts/run_pipeline.py --data data_raw/collected_data.csv --steps 2 3 4 5

Orchestrates all pipeline steps in sequence.
Steps:
  1 — Dictionary selector (Streamlit, launched separately)
  2 — Generate ODK XLS form
  3 — Quality check
  4 — Descriptive statistics
  5 — Econometric model
"""

import argparse
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

def main():
    parser = argparse.ArgumentParser(description="AGEVAL Pipeline Runner")
    parser.add_argument("--data",  default=None, help="Path to CSV data file (required for steps 3–5)")
    parser.add_argument("--steps", nargs="+", type=int, default=[2, 3, 4, 5],
                        help="Which steps to run (2=ODK, 3=QC, 4=Stats, 5=Model)")
    parser.add_argument("--batch", default=None, help="Batch name for quality report")
    args = parser.parse_args()

    print("\n╔══════════════════════════════════════════════╗")
    print("║          AGEVAL — Impact Eval Pipeline       ║")
    print("╚══════════════════════════════════════════════╝\n")

    if 2 in args.steps:
        print("── Step 2: Generate ODK XLS form ───────────────")
        from scripts.s02_generate_odk_form import generate_form
        generate_form()
        print()

    if 3 in args.steps:
        if not args.data:
            print("⚠  --data required for Step 3. Skipping.")
        else:
            print("── Step 3: Quality check ────────────────────────")
            from scripts.s03_quality_check import run_quality_check
            run_quality_check(args.data, args.batch)
            print()

    if 4 in args.steps:
        if not args.data:
            print("⚠  --data required for Step 4. Skipping.")
        else:
            print("── Step 4: Descriptive statistics ───────────────")
            from scripts.s04_descriptive_stats import run_descriptive
            run_descriptive(args.data)
            print()

    if 5 in args.steps:
        if not args.data:
            print("⚠  --data required for Step 5. Skipping.")
        else:
            print("── Step 5: Econometric model ─────────────────────")
            from scripts.s05_run_model import run_models
            run_models(args.data)
            print()

    print("\n✅  Pipeline complete.\n")

if __name__ == "__main__":
    main()
