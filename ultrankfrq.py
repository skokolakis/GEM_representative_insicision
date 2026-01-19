import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d
import os

## config ##
common_distance_step = 0.5  # Common distance step for interpolation, in meters
interp_kind = 'linear'      # Interpolation method, can be 'linear', 'nearest', 'zero', 'slinear', 'quadratic', 'cubic'
input_folder = "."
output_folder = "output_profiles"
y_axis_label = "Mean EC Response S/m" 

excel_files = [
    f for f in os.listdir(input_folder)
    if f.endswith('.xlsx') and not f.startswith('~$')
]

for excel_file in excel_files:
    print(f"Processing file: {excel_file}")
    excel = pd.ExcelFile(excel_file)
    output_data = {}
    representativeness_scores = {}

    ## loop over each sheet ##
    for sheet_name in excel.sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet_name)
        distance = df.iloc[:, 0].values
        line_data = df.iloc[:, 1:]

        ## check distance column ##
        if np.all(np.isnan(distance)):
            print(f"Skipping sheet {sheet_name}: all distance values are NaN.")
            continue

        min_dist = np.nanmin(distance)
        max_dist = np.nanmax(distance)

        if max_dist - min_dist < common_distance_step:
            print(f"Skipping sheet {sheet_name}: too short or invalid distance range.")
            continue

        ## interpolation axis ##
        common_dist = np.arange(min_dist, max_dist + common_distance_step, common_distance_step)

        print(f"[DEBUG] Sheet: {sheet_name}")
        print(f"[DEBUG] Distance column:\n{df.iloc[:, 0]}")
        print(f"[DEBUG] Line data columns:\n{line_data.columns}")

        ## interpolation ##
        interpolated_lines = {}
        for col in line_data.select_dtypes(include=[np.number]).columns:
            y = line_data[col].values
            mask = ~np.isnan(distance) & ~np.isnan(y)
            if np.sum(mask) < 2:
                continue

            f = interp1d(distance[mask], y[mask], kind=interp_kind, bounds_error=False, fill_value=np.nan)
            interpolated_lines[col] = f(common_dist)

        interpolated_df = pd.DataFrame(interpolated_lines, index=common_dist)

        ## representative profile ##
        rep_prof = interpolated_df.mean(axis=1, skipna=True)
        std_prof = interpolated_df.std(axis=1, skipna=True)

        ## metrics ##
        mean_std = std_prof.mean()
        amplitude = rep_prof.max() - rep_prof.min()
        score = amplitude / mean_std if mean_std != 0 else 0
        representativeness_scores[sheet_name] = {
            'mean_std': mean_std,
            'amplitude': amplitude,
            'score': score
        }

        ## save to dict ##
        output_df = pd.DataFrame({
            'Distance (m)': common_dist,
            f'{sheet_name}_mean': rep_prof
        })
        output_data[sheet_name] = output_df

        ## 4theplot ##
        plt.plot(common_dist, rep_prof, label=f'{sheet_name} (score={score:.2f})')

    ## Save Excel Output ##
    file_base = os.path.splitext(excel_file)[0]
    out_excel = os.path.join(output_folder, f"{file_base}_interpolated.xlsx")
    with pd.ExcelWriter(out_excel) as writer:
        for sheet_name, df_out in output_data.items():
            df_out.to_excel(writer, sheet_name=sheet_name, index=False)

    ## save_4theplot ##
    plt.xlabel('Distance (m)')
    plt.ylabel(y_axis_label)
    plt.title(f'Representative incision from : {file_base}')
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    out_plot = os.path.join(output_folder, f"{file_base}_representative_profiles.png")
    plt.savefig(out_plot, dpi=300)
    plt.close()

    ## ranking ##
    print(f"\nFrequency Ranking by Representativeness Score for {file_base}:")
    ranking = sorted(representativeness_scores.items(), key=lambda x: x[1]['score'], reverse=True)
    for i, (freq, metrics) in enumerate(ranking, 1):
        print(f"{i}. {freq}")
        print(f"   Score     : {metrics['score']:.2f}")
        print(f"   Mean Std  : {metrics['mean_std']:.2f}")
        print(f"   Amplitude : {metrics['amplitude']:.2f}")
    if ranking:
        best_ranking = ranking[0]
        print(f"\nBest Frequency: {best_ranking[0]} with Score: {best_ranking[1]['score']:.2f}\n")
    else:
        print("\nNo valid frequency sheets were processed for this file.\n")
