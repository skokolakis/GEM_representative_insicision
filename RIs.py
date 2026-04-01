"""
Unified ultrankfrq controller
- Option 1: EC mode
- Option 2: MS mode
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path
from scipy.interpolate import (
    Akima1DInterpolator,
    CubicSpline,
    PchipInterpolator,
    make_interp_spline,
)
from numpy.polynomial.polynomial import polyfit, polyval
import sys

# -----------------------------
# Shared configuration
# -----------------------------
COMMON_DISTANCE_STEP = 0.5
INPUT_FOLDER = Path(".")
OUTPUT_FOLDER = Path("output_profiles")
SCORE_EPSILON = 1e-8

OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

ALL_INTERP_METHODS = ["linear", "cubic", "nearest", "quadratic", "pchip", "akima", "polynomial"]

_METHOD_MIN_POINTS = {
    "linear": 2,
    "nearest": 1,
    "pchip": 2,
    "quadratic": 3,
    "polynomial": 3,
    "akima": 5,
    "cubic": 4,
}


# -----------------------------
# Interpolation helper
# -----------------------------
def _interpolate_with_method(
    xp: np.ndarray, yp: np.ndarray, target_x: np.ndarray, method: str
) -> np.ndarray:
    """Interpolate yp at xp onto target_x using the named method."""
    if method == "linear":
        return np.interp(target_x, xp, yp)

    if method == "nearest":
        idx = np.searchsorted(xp, target_x, side="left")
        idx_left = np.clip(idx - 1, 0, len(xp) - 1)
        idx_right = np.clip(idx, 0, len(xp) - 1)
        nearest = np.where(
            np.abs(target_x - xp[idx_left]) <= np.abs(target_x - xp[idx_right]),
            idx_left,
            idx_right,
        )
        return yp[nearest]

    if method == "pchip":
        return PchipInterpolator(xp, yp)(target_x)

    if method == "quadratic":
        return make_interp_spline(xp, yp, k=2)(target_x)

    if method == "polynomial":
        degree = min(len(xp) - 1, 5)
        coeffs = polyfit(xp, yp, degree)
        return polyval(target_x, coeffs)

    if method == "akima":
        return Akima1DInterpolator(xp, yp)(target_x)

    if method == "cubic":
        return CubicSpline(xp, yp)(target_x)

    raise ValueError(f"Unknown interpolation method: {method!r}")


# -----------------------------
# Core processing engine
# -----------------------------
def run_ultrankfrq(mode="EC", interp_kind="linear"):

    if mode == "EC":
        y_axis_label = "Mean EC Response S/m"
        mode_tag = "EC"
    elif mode == "MS":
        y_axis_label = "Mean MS Response 10^-5 SI"
        mode_tag = "MS"
    else:
        raise ValueError("Mode must be 'EC' or 'MS'")

    print(f"\n=== Running UltraRank Frequency Analysis [{mode_tag}] | Method: {interp_kind} ===")

    excel_files = sorted([p for p in INPUT_FOLDER.glob("*.xlsx") if not p.name.startswith("~$")])

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

        for sheet_name in excel.sheet_names:
            try:
                df = pd.read_excel(excel, sheet_name=sheet_name)
            except Exception as e:
                print(f"  Skipping sheet {sheet_name}: {e}")
                continue

            if df.shape[1] < 2:
                print(f"  Skipping sheet {sheet_name}: not enough columns.")
                continue

            distance = pd.to_numeric(df.iloc[:, 0], errors="coerce").values
            line_data = df.iloc[:, 1:].apply(pd.to_numeric, errors="coerce")

            if np.all(np.isnan(distance)):
                continue

            min_dist = np.nanmin(distance)
            max_dist = np.nanmax(distance)

            if not np.isfinite(min_dist) or not np.isfinite(max_dist) or (max_dist - min_dist) < COMMON_DISTANCE_STEP:
                continue

            n_points = int(round((max_dist - min_dist) / COMMON_DISTANCE_STEP)) + 1
            common_dist = np.linspace(min_dist, max_dist, n_points)

            interpolated_lines = {}

            for col in line_data.columns:
                y = line_data[col].values
                mask = ~np.isnan(distance) & ~np.isnan(y)

                if np.sum(mask) < 2:
                    continue

                df_xy = pd.DataFrame({"d": distance[mask], "y": y[mask]})
                df_grouped = (
                    df_xy.groupby("d", as_index=False)
                    .mean()
                    .sort_values("d")
                    .drop_duplicates(subset="d")
                )

                if df_grouped.shape[0] < 2:
                    continue

                if not df_grouped["d"].is_monotonic_increasing:
                    print(f"  Warning: column '{col}' distance not monotonic after dedup, skipped.")
                    continue

                xp = df_grouped["d"].values
                yp = df_grouped["y"].values

                min_pts = _METHOD_MIN_POINTS.get(interp_kind, 2)
                if len(xp) < min_pts:
                    print(f"  Warning: column '{col}' has {len(xp)} points, need >={min_pts} for {interp_kind}, skipped.")
                    continue

                try:
                    interpolated_lines[col] = _interpolate_with_method(xp, yp, common_dist, interp_kind)
                except Exception as e:
                    print(f"  Warning: interpolation failed for column '{col}': {e}")
                    continue

            if not interpolated_lines:
                continue

            interpolated_df = pd.DataFrame(interpolated_lines, index=common_dist)

            rep_prof = interpolated_df.mean(axis=1, skipna=True)
            std_prof = interpolated_df.std(axis=1, skipna=True, ddof=0)

            mean_std = float(std_prof.mean())
            amplitude = float(np.nanmax(rep_prof) - np.nanmin(rep_prof))
            if np.isnan(mean_std) or mean_std < SCORE_EPSILON:
                window = max(5, len(rep_prof) // 10)
                smoothed = rep_prof.rolling(window=window, center=True, min_periods=1).mean()
                intra_noise = float((rep_prof - smoothed).std())
                mean_std = intra_noise
                score = amplitude / intra_noise if intra_noise >= SCORE_EPSILON else amplitude
            else:
                score = amplitude / mean_std

            representativeness_scores[sheet_name] = {
                "mean_std": mean_std,
                "amplitude": amplitude,
                "score": score
            }

            out_df = pd.DataFrame({
                "Distance (m)": common_dist,
                f"{sheet_name}_mean": rep_prof.values
            })

            output_data[sheet_name] = out_df

            if not np.all(np.isnan(rep_prof.values)):
                plt.plot(common_dist, rep_prof.values, label=f"{sheet_name} (score={score:.2f})")

        if not output_data:
            plt.clf()
            continue

        base_name = excel_path.stem

        out_excel = OUTPUT_FOLDER / f"{base_name}_{mode_tag}_{interp_kind}_interpolated.xlsx"
        with pd.ExcelWriter(out_excel) as writer:
            for sheet_name, df_out in output_data.items():
                df_out.to_excel(writer, sheet_name=sheet_name, index=False)

        scores_df = pd.DataFrame.from_dict(representativeness_scores, orient="index")
        scores_df.index.name = "sheet"
        scores_df.to_csv(OUTPUT_FOLDER / f"{base_name}_{mode_tag}_{interp_kind}_representativeness_scores.csv")

        plt.xlabel("Distance (m)")
        plt.ylabel(y_axis_label)
        plt.title(f"Representative profiles [{mode_tag}] [{interp_kind}] — {base_name}")
        plt.legend()
        plt.grid(True)
        plt.tight_layout()
        plt.savefig(OUTPUT_FOLDER / f"{base_name}_{mode_tag}_{interp_kind}_representative_profiles.png", dpi=300)
        plt.clf()

        print(f"\nFrequency Ranking [{mode_tag}] — {base_name}")
        ranking = sorted(representativeness_scores.items(), key=lambda x: x[1]["score"], reverse=True)

        for i, (freq, metrics) in enumerate(ranking, 1):
            print(f"{i}. {freq} | Score={metrics['score']:.2f} | Std={metrics['mean_std']:.4g} | Amp={metrics['amplitude']:.4g}")

        if ranking:
            print(f"\nBest Frequency: {ranking[0][0]} (Score={ranking[0][1]['score']:.2f})\n")


# -----------------------------
# Control panel
# -----------------------------
def _prompt_interp_method() -> str:
    """Prompt the user to select an interpolation method and return the choice."""
    print("\n  Interpolation method:")
    for i, m in enumerate(ALL_INTERP_METHODS, 1):
        print(f"    {i}) {m}")
    while True:
        raw = input("  Select method [1-7, default=1 linear]: ").strip()
        if raw == "":
            return "linear"
        if raw.isdigit() and 1 <= int(raw) <= len(ALL_INTERP_METHODS):
            return ALL_INTERP_METHODS[int(raw) - 1]
        print(f"  Invalid choice. Enter a number between 1 and {len(ALL_INTERP_METHODS)}.")


def control_panel():
    while True:
        print("\n===== Script Control Panel =====")
        print("1) Run EC data script")
        print("2) Run MS data script")
        print("Q) Quit")

        choice = input("Select option: ").strip().lower()

        if choice == "1":
            interp_kind = _prompt_interp_method()
            run_ultrankfrq("EC", interp_kind)
        elif choice == "2":
            interp_kind = _prompt_interp_method()
            run_ultrankfrq("MS", interp_kind)
        elif choice in ("q", "quit", "exit"):
            print("Exiting.")
            break
        else:
            print("Invalid option.")


if __name__ == "__main__":
    control_panel()
