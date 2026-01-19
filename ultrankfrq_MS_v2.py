#!/usr/bin/env python3
"""
Revised ultrankfrq.py
- More robust I/O (creates output folder)
- Handles duplicate/unsorted distance values by grouping and averaging
- Avoids divide-by-zero in score calculation
- Saves representativeness scores to CSV
- Skips sheets/files with no valid numeric data
- Uses pathlib for path handling and clearer logging
"""
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d
from pathlib import Path
import sys

# Config
common_distance_step = 0.5  # meters
interp_kind = 'linear'      # interpolation kind
input_folder = Path(".")
output_folder = Path("output_profiles")
y_axis_label = "Mean MS Response 10^-5 SI"
score_epsilon = 1e-8        # prevents division by zero when computing score

# Prepare folders
output_folder.mkdir(parents=True, exist_ok=True)

# Find Excel files (ignore temporary ~ files)
excel_files = sorted([p for p in input_folder.glob("*.xlsx") if not p.name.startswith("~$")])

if not excel_files:
    print("No .xlsx files found in input folder. Exiting.")
    sys.exit(0)

for excel_path in excel_files:
    print(f"\nProcessing file: {excel_path.name}")
    try:
        excel = pd.ExcelFile(excel_path)
    except Exception as e:
        print(f"  ERROR: Could not open {excel_path.name}: {e}")
        continue

    output_data = {}
    representativeness_scores = {}

    # Iterate sheets
    for sheet_name in excel.sheet_names:
        try:
            df = pd.read_excel(excel, sheet_name=sheet_name)
        except Exception as e:
            print(f"  Skipping sheet {sheet_name}: could not read sheet ({e})")
            continue

        if df.shape[1] < 2:
            print(f"  Skipping sheet {sheet_name}: expected at least 2 columns (distance + lines).")
            continue

        # Convert first column (distance) to numeric
        distance = pd.to_numeric(df.iloc[:, 0], errors='coerce').values
        # Convert remaining columns to numeric where possible
        line_data = df.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')

        if np.all(np.isnan(distance)):
            print(f"  Skipping sheet {sheet_name}: all distance values are NaN.")
            continue

        # Determine valid min/max for constructing common axis
        min_dist = np.nanmin(distance)
        max_dist = np.nanmax(distance)
        if not np.isfinite(min_dist) or not np.isfinite(max_dist) or (max_dist - min_dist) < common_distance_step:
            print(f"  Skipping sheet {sheet_name}: invalid or too-short distance range.")
            continue

        common_dist = np.arange(min_dist, max_dist + common_distance_step, common_distance_step)

        # Interpolate each numeric column
        interpolated_lines = {}
        for col in line_data.columns:
            y = line_data[col].values
            mask = ~np.isnan(distance) & ~np.isnan(y)
            if np.sum(mask) < 2:
                # Not enough valid points
                continue

            # Aggregate duplicates by distance: take mean for equal distance values
            d_valid = distance[mask]
            y_valid = y[mask]
            df_xy = pd.DataFrame({"d": d_valid, "y": y_valid})
            df_grouped = df_xy.groupby("d", as_index=False).mean()

            if df_grouped.shape[0] < 2:
                continue

            # Ensure ascending distances
            df_grouped = df_grouped.sort_values("d")
            xs = df_grouped["d"].values
            ys = df_grouped["y"].values

            # Build interpolator
            try:
                f = interp1d(xs, ys, kind=interp_kind, bounds_error=False, fill_value=np.nan, assume_sorted=True)
                interpolated = f(common_dist)
                interpolated_lines[col] = interpolated
            except Exception as e:
                print(f"    Warning: interpolation failed for column '{col}' in sheet '{sheet_name}': {e}")
                continue

        if not interpolated_lines:
            print(f"  No valid numeric lines to interpolate in sheet {sheet_name}.")
            continue

        interpolated_df = pd.DataFrame(interpolated_lines, index=common_dist)

        # Representative profile (mean across lines) and std
        rep_prof = interpolated_df.mean(axis=1, skipna=True)
        std_prof = interpolated_df.std(axis=1, skipna=True)

        # Metrics and score (safe divide)
        mean_std = float(std_prof.mean()) if std_prof.size > 0 else float("nan")
        amplitude = float(np.nanmax(rep_prof) - np.nanmin(rep_prof)) if rep_prof.size > 0 else float("nan")
        score = amplitude / max(mean_std, score_epsilon) if np.isfinite(amplitude) else 0.0

        representativeness_scores[sheet_name] = {
            "mean_std": mean_std,
            "amplitude": amplitude,
            "score": score
        }

        # Save output DataFrame for this sheet (Distance first column)
        out_df = pd.DataFrame({"Distance (m)": common_dist, f"{sheet_name}_mean": rep_prof.values})
        output_data[sheet_name] = out_df

        # Plot representative profile if not all NaN
        if not np.all(np.isnan(rep_prof.values)):
            plt.plot(common_dist, rep_prof.values, label=f"{sheet_name} (score={score:.2f})")

    # After processing sheets for this file: write outputs if any
    base_name = excel_path.stem
    if not output_data:
        print(f"  No valid sheets processed for {excel_path.name}. Nothing to write.")
        plt.clf()
        continue

    # Save interpolated data to Excel
    out_excel = output_folder / f"{base_name}_interpolated.xlsx"
    try:
        with pd.ExcelWriter(out_excel) as writer:
            for sheet_name, df_out in output_data.items():
                df_out.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"  Wrote interpolated data to: {out_excel}")
    except Exception as e:
        print(f"  ERROR: Failed to write Excel output for {excel_path.name}: {e}")

    # Save representativeness scores to CSV for machine consumption
    try:
        scores_df = pd.DataFrame.from_dict(representativeness_scores, orient="index")
        scores_df.index.name = "sheet"
        scores_out = output_folder / f"{base_name}_representativeness_scores.csv"
        scores_df.to_csv(scores_out)
        print(f"  Wrote representativeness scores to: {scores_out}")
    except Exception as e:
        print(f"  Warning: could not write representativeness scores CSV: {e}")

    # Save plot
    try:
        plt.xlabel("Distance (m)")
        plt.ylabel(y_axis_label)
        plt.title(f"Representative profiles: {base_name}")
        plt.legend()
        plt.grid(True)
        plt.tight_layout()
        out_plot = output_folder / f"{base_name}_representative_profiles.png"
        plt.savefig(out_plot, dpi=300)
        print(f"  Wrote plot to: {out_plot}")
    except Exception as e:
        print(f"  Warning: could not save plot for {excel_path.name}: {e}")
    finally:
        plt.clf()

    # Print ranking to console
    print(f"\nFrequency Ranking by Representativeness Score for {base_name}:")
    ranking = sorted(representativeness_scores.items(), key=lambda x: x[1]["score"], reverse=True)
    for i, (freq, metrics) in enumerate(ranking, 1):
        print(f"{i}. {freq}")
        print(f"   Score     : {metrics['score']:.2f}")
        print(f"   Mean Std  : {metrics['mean_std']:.4g}")
        print(f"   Amplitude : {metrics['amplitude']:.4g}")
    if ranking:
        best = ranking[0]
        print(f"\nBest Frequency: {best[0]} with Score: {best[1]['score']:.2f}\n")
    else:
        print("\nNo valid frequency sheets were processed for this file.\n")