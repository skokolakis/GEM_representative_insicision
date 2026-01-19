# ultrankfrq v2

A small utility to load one or more Excel (.xlsx) files containing distance-series measurements (e.g., electrical conductivity profiles), interpolate them to a common distance axis, compute a representative profile per sheet, and rank sheets by a simple "representativeness" score.

This README documents what the script does, expected input format, produced outputs, configuration options, and troubleshooting tips.

## Overview

The script:
- Reads all `.xlsx` files in the current folder (ignores temporary files that start with `~$`).
- For each sheet in each Excel file:
  - Treats the first column as distance (meters) and remaining columns as measured lines.
  - Converts values to numeric and skips non-numeric entries.
  - Aggregates duplicate distance values (by averaging).
  - Interpolates each valid line onto a common distance grid.
  - Computes a representative profile (mean across lines) and per-distance standard deviation.
  - Computes a representativeness score: amplitude / mean(std), with protection against division-by-zero.
  - Saves interpolated data, a CSV of representativeness scores, and a PNG plot of representative profiles to an output folder.
- Skips sheets with insufficient or invalid data.

## Requirements

- Python 3.8+ (should work on 3.8, 3.9, 3.10, 3.11)
- Packages:
  - pandas
  - numpy
  - scipy
  - matplotlib
  - openpyxl (for reading/writing .xlsx)

Install with pip:
```bash
pip install pandas numpy scipy matplotlib openpyxl
```

## Files

- Script (as provided by you): e.g. `ultrankfrq v2.py`
  - Note: If your filename contains spaces, quote it when running: `python "ultrankfrq v2.py"`
- Output (created under `output_profiles/` by default):
  - `{inputfilename}_interpolated.xlsx` — one sheet per original sheet with two columns: Distance (m) and <sheetname>_mean
  - `{inputfilename}_representativeness_scores.csv` — CSV with columns: mean_std, amplitude, score (indexed by sheet name)
  - `{inputfilename}_representative_profiles.png` — plot of representative profiles for the file

## Input format

- Each Excel sheet should have at least 2 columns:
  - Column 1: distance (meters) — numeric values. NaNs are permitted; rows with NaN distances are ignored.
  - Columns 2..N: measurements (one line per column) — numeric values. Non-numeric values are converted to NaN and ignored.
- There is no strict requirement on headers: the script accesses columns by position. If headers are present they are used as the column names in outputs.
- Duplicate distance values are allowed; the script groups equal distances and averages the measured values before interpolation.
- Sheets where all distance values are NaN, or where there are too few unique distance points (< 2) after grouping, are skipped.

## How the representativeness score is computed

- Representative profile = mean across interpolated lines (per distance point).
- Per-distance standard deviation is computed across lines.
- mean_std = mean of the per-distance standard deviation (ignoring NaNs).
- amplitude = max(rep_profile) - min(rep_profile) (ignoring NaNs).
- score = amplitude / max(mean_std, epsilon) where epsilon is a small constant to avoid divide-by-zero.
- Higher score indicates a cleaner, higher-amplitude representative profile relative to within-profile variability.

## Configuration

The script exposes a few straightforward configuration variables at the top of the file:

- `common_distance_step` (float): interpolation spacing (meters). Default: `0.5`
- `interp_kind` (str): interpolation kind passed to `scipy.interpolate.interp1d`. E.g. `'linear'`, `'nearest'`, `'cubic'`. Default: `'linear'`
- `input_folder` (Path): folder to search for `.xlsx` files. Default: current directory (`Path(".")`)
- `output_folder` (Path): folder where outputs are written. Default: `Path("output_profiles")` (created automatically)
- `y_axis_label` (str): used for the saved plot y-axis label.
- `score_epsilon` (float): small number used to avoid division-by-zero in score calculation. Default `1e-8`.

Modify these variables at the top of the script to change behavior, or wrap the script with a short launcher if you want command-line flags.

## Usage

Place your `.xlsx` files in the same folder as the script (or change `input_folder`), then run:

```bash
python "ultrankfrq v2.py"
```

After running, inspect `output_profiles/` for the generated Excel, CSV and PNG files.

Example output files for an input file named `survey.xlsx`:
- `output_profiles/survey_interpolated.xlsx`
- `output_profiles/survey_representativeness_scores.csv`
- `output_profiles/survey_representative_profiles.png`

## Example sheet layout

A minimal sheet could look like:

| Distance | Line A | Line B | Line C |
|----------|--------|--------|--------|
| 0.0      | 0.12   | 0.11   | 0.13   |
| 0.5      | 0.14   | 0.13   | 0.15   |
| 1.0      | 0.20   | 0.19   | 0.22   |

- Duplicate distances (e.g., two rows with distance 0.5) are averaged per line before interpolation.

## Troubleshooting

- "No .xlsx files found in input folder. Exiting." — Place one or more `.xlsx` files in the same directory or change `input_folder`.
- Script fails to open an Excel file — confirm the file is a valid `.xlsx` and not password protected. Temporary files starting with `~$` are ignored.
- No valid numeric lines in sheet — check that the first column contains numeric distances and subsequent columns contain numeric measurements.
- Plot or Excel write errors — confirm you have permission to create files in the `output_folder` or change `output_folder`.
- If interpolation complains about insufficient points, ensure each measured line has at least two unique distance points after grouping duplicates.

## Suggested improvements

- Add a command-line interface (argparse / click) to control input/output paths and options at runtime.
- Add more interpolation options or smoothing.
- Export the full interpolated matrix (all lines) in addition to the representative mean.
- Add unit tests and example test data.

## License

Choose a license appropriate for your project (e.g., MIT, BSD, Apache). This README does not include a license; add one if you intend to publish.

## Contact / Next steps

If you want, I can:
- Create a requirements.txt
- Add a simple CLI wrapper that accepts input/output paths and options
- Produce an example `.xlsx` test file

Just tell me which of the above you'd like next.