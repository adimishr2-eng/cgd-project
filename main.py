"""
================================================================================
PNG DEMAND FORECASTING AND NETWORK EXPANSION PLANNING MODEL
================================================================================

Project Name:
    PNG Demand Forecasting and Network Expansion Planning Model

Description:
    This project forecasts Piped Natural Gas (PNG) demand across urban zones,
    prioritizes zones for network expansion based on multi-criteria scoring,
    and performs comprehensive techno-economic analysis including NPV, IRR,
    and payback period calculations for City Gas Distribution (CGD) networks.

    The model runs in three sequential parts:
        Part 1 – Demand Forecasting:
            Projects residential, commercial, and industrial PNG demand over
            a configurable planning horizon using population growth, penetration
            curves, and consumption benchmarks.

        Part 2 – Zone Prioritization:
            Ranks geographical zones for network expansion using weighted
            scoring across parameters such as demand density, infrastructure
            readiness, revenue potential, and execution feasibility.

        Part 3 – Techno-Economic Analysis:
            Builds a financial model for each prioritized zone, computing
            capital expenditure (CAPEX), operating expenditure (OPEX),
            revenue projections, net present value (NPV), internal rate of
            return (IRR), and discounted payback period.

How to Run:
    python main.py

How to Change Inputs:
    Edit the CSV files located in the inputs/ folder:

    - city_data.csv:
        Change city-level parameters such as total population, annual
        population growth rate, household size, existing PNG connections,
        and target penetration rates.

    - zone_data.csv:
        Add or modify individual zones with their characteristics including
        area, number of households, building types, road conditions, and
        distance from existing pipeline infrastructure.

    - financial_params.csv:
        Adjust financial assumptions such as discount rate, project lifetime,
        gas purchase cost, consumer tariff, pipeline cost per km, CAPEX
        components, and OPEX percentages.

Output Files:
    All outputs are generated in the outputs/ folder:
    - outputs/demand_forecast.xlsx   : Yearly demand projections by category
    - outputs/zone_ranking.xlsx      : Prioritized zone rankings with scores
    - outputs/financial_model.xlsx   : Techno-economic results per zone
    - outputs/charts/                : Publication-quality visualisation charts

Required Packages:
    Install all dependencies with:
        pip install -r requirements.txt

    Core dependencies: pandas, numpy, matplotlib, openpyxl
================================================================================
"""

# =============================================================================
# STANDARD LIBRARY IMPORTS
# =============================================================================
import os      # For file-system path operations
import sys     # For system-level operations (e.g., exit on fatal error)
import time    # For timing each stage of the pipeline

# =============================================================================
# PROJECT MODULE IMPORTS
# =============================================================================
# Each module exposes a single public entry-point function that accepts
# the project base_path and returns a DataFrame of results.
from models.demand_forecast import run_demand_forecast
from models.zone_prioritization import run_zone_prioritization
from models.techno_economic import run_techno_economic


# =============================================================================
# HELPER UTILITIES
# =============================================================================

def _format_elapsed(seconds: float) -> str:
    """Return a human-readable string for an elapsed duration.

    Args:
        seconds: Elapsed time in fractional seconds.

    Returns:
        A string like '2.34 seconds' or '1 minute 12.5 seconds'.
    """
    if seconds < 60:
        return f"{seconds:.2f} seconds"
    minutes = int(seconds // 60)
    remaining = seconds % 60
    return f"{minutes} minute{'s' if minutes > 1 else ''} {remaining:.1f} seconds"


# =============================================================================
# MAIN PIPELINE FUNCTION
# =============================================================================

def main() -> None:
    """Execute the full PNG demand-forecasting and network-expansion pipeline.

    This function orchestrates three sequential analysis stages:
        1. Demand Forecasting
        2. Zone Prioritization
        3. Techno-Economic Analysis

    Each stage is wrapped in its own try/except block so that a failure in
    one part does not prevent the remaining parts from executing.  Timing
    information is printed for every stage and for the overall run.
    """

    # -------------------------------------------------------------------------
    # Determine the base directory (where main.py lives).
    # All relative paths for inputs/ and outputs/ are resolved from here.
    # -------------------------------------------------------------------------
    base_path = os.path.dirname(os.path.abspath(__file__))

    # Record the wall-clock start time for the entire pipeline
    pipeline_start = time.time()

    # =========================================================================
    # HEADER
    # =========================================================================
    print("=" * 80)
    print("PNG DEMAND FORECASTING AND NETWORK EXPANSION PLANNING MODEL")
    print("City Gas Distribution Analytics Framework")
    print("=" * 80)
    print()  # blank separator line

    # =========================================================================
    # CREATE OUTPUT DIRECTORIES
    # =========================================================================
    # Ensure the outputs/charts directory tree exists before any module tries
    # to write files.  exist_ok=True prevents errors if the dirs already exist.
    outputs_dir = os.path.join(base_path, "outputs")
    charts_dir = os.path.join(base_path, "outputs", "charts")
    os.makedirs(charts_dir, exist_ok=True)
    print(f"[INFO] Output directory ready : {outputs_dir}")
    print(f"[INFO] Charts directory ready  : {charts_dir}")

    # =========================================================================
    # PART 1: DEMAND FORECASTING
    # =========================================================================
    demand_df = None  # Will hold the returned DataFrame (or stay None on error)

    print("\n" + "-" * 60)
    print("PART 1: DEMAND FORECASTING")
    print("-" * 60)

    try:
        part1_start = time.time()

        # Run the demand-forecasting model; it reads inputs/city_data.csv and
        # writes outputs/demand_forecast.xlsx plus charts.
        demand_df = run_demand_forecast(base_path)

        part1_elapsed = time.time() - part1_start
        print(f"\n[OK] PART 1: DEMAND FORECASTING COMPLETE  "
              f"[{_format_elapsed(part1_elapsed)}]\n")

    except Exception as exc:
        # Log the error but allow the pipeline to continue with Part 2.
        print(f"\n[X] PART 1 FAILED: {exc}\n")
        print("[WARNING] Continuing to Part 2 despite the error above.")

    # =========================================================================
    # PART 2: ZONE PRIORITIZATION
    # =========================================================================
    zone_df = None  # Will hold the zone-ranking DataFrame (or stay None)

    print("\n" + "-" * 60)
    print("PART 2: ZONE PRIORITIZATION")
    print("-" * 60)

    try:
        part2_start = time.time()

        # Run zone prioritization; reads inputs/zone_data.csv and writes
        # outputs/zone_ranking.xlsx plus charts.
        zone_df = run_zone_prioritization(base_path)

        part2_elapsed = time.time() - part2_start
        print(f"\n[OK] PART 2: ZONE PRIORITIZATION COMPLETE  "
              f"[{_format_elapsed(part2_elapsed)}]\n")

    except Exception as exc:
        print(f"\n[X] PART 2 FAILED: {exc}\n")
        print("[WARNING] Continuing to Part 3 despite the error above.")

    # =========================================================================
    # PART 3: TECHNO-ECONOMIC ANALYSIS
    # =========================================================================
    financial_df = None  # Will hold the financial-model DataFrame (or stay None)

    print("\n" + "-" * 60)
    print("PART 3: TECHNO-ECONOMIC ANALYSIS")
    print("-" * 60)

    try:
        part3_start = time.time()

        # If Part 2 succeeded we pass its zone-ranking DataFrame into the
        # techno-economic model so it can align financial projections with
        # the prioritised expansion sequence.  If Part 2 failed, we call
        # the function without zone data and let it use its own defaults.
        if zone_df is not None:
            financial_df = run_techno_economic(base_path, zone_ranking_df=zone_df)
        else:
            print("[INFO] Zone ranking unavailable – running techno-economic "
                  "analysis with default zone assumptions.")
            financial_df = run_techno_economic(base_path)

        part3_elapsed = time.time() - part3_start
        print(f"\n[OK] PART 3: TECHNO-ECONOMIC ANALYSIS COMPLETE  "
              f"[{_format_elapsed(part3_elapsed)}]\n")

    except Exception as exc:
        print(f"\n[X] PART 3 FAILED: {exc}\n")

    # =========================================================================
    # FINAL SUMMARY
    # =========================================================================
    pipeline_elapsed = time.time() - pipeline_start

    print("\n" + "=" * 80)
    print("MODEL EXECUTION COMPLETE")
    print("=" * 80)
    print(f"  Total pipeline time : {_format_elapsed(pipeline_elapsed)}")
    print(f"  Outputs saved to    : /outputs/ folder")
    print(f"  Charts saved to     : /outputs/charts/ folder")
    print(f"  Excel files         : demand_forecast.xlsx, "
          f"zone_ranking.xlsx, financial_model.xlsx")
    print("=" * 80)


# =============================================================================
# SCRIPT ENTRY POINT
# =============================================================================
# When this file is executed directly (python main.py), invoke the main()
# pipeline.  When imported as a module, main() can be called explicitly.
if __name__ == "__main__":
    main()
