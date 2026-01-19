# GEM Representative insicision Script

The `RIs_v1.py` is a unified controller for generating representative profiles from geophysical (EC/MS) line measurements.

The script:
- Reads .xlsx files containing distance vs. measurement traces (one or more traces per sheet),
- Interpolates all traces onto a common distance grid,
- Computes a representative (mean) profile per sheet,
- Computes a representativeness score per sheet (amplitude / mean standard deviation),
- Writes interpolated profiles, scores, and representative-profile plots to an output folder.

---

## Table of contents

- [Requirements](#requirements)
- [Installation](#installation)
- [Input format](#input-format)
- [Configuration](#configuration)
- [Usage](#usage)
  - [Interactive mode (default)](#interactive-mode-default)
  - [Programmatic / headless usage](#programmatic--headless-usage)
- [Outputs](#outputs)
- [How the representativeness score is computed](#how-the-representativeness-score-is-computed)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License & contact](#license--contact)

---

## Requirements

- Python 3.8+
- Packages:
  - pandas
  - numpy
  - matplotlib
  - scipy
  - openpyxl

You can install the required packages with:

```bash
python -m pip install pandas numpy matplotlib scipy openpyxl
```

(or create a virtual environment first)

---

## Installation

1. Clone the repository (if not already present):
   ```bash
   git clone https://github.com/skokolakis/GEM_representative_insicision.git
   cd GEM_representative_insicision
   ```

2. Install requirements as shown above.

---

## Input format

- Place one or more Excel files (`*.xlsx`) in the working directory (or change the INPUT_FOLDER in the script).
- Each Excel file may contain multiple sheets.
- Each sheet should have:
  - First column: distance values (numeric), in meters.
  - Remaining columns: measurement traces (one trace per column). These typically are different frequencies or channels.
- The script ignores temporary Excel files whose names start with `~$`.

Notes:
- Rows with NaN distance or NaN measurement values will be handled (masked) during interpolation.
- If a trace has fewer than 2 valid points it will be skipped.

---

## Configuration

Configuration variables are defined at the top of `RIs_v1.py`:

- `COMMON_DISTANCE_STEP` (default: 0.5) — step used for the common distance grid (meters).
- `INTERP_KIND` (default: `"linear"`) — interpolation kind passed to `scipy.interpolate.interp1d` (e.g., `"linear"`, `"nearest"`, `"cubic"`).
- `INPUT_FOLDER` (default: `Path(".")`) — folder to search for `.xlsx` files.
- `OUTPUT_FOLDER` (default: `Path("output_profiles")`) — folder where outputs will be written.
- `SCORE_EPSILON` (default: `1e-8`) — small value to avoid division by zero in score computation.

Edit those variables directly in `RIs_v1.py` if you want different behavior.

---

## Usage

### Interactive mode (default)
Run the script directly:

```bash
python RIs_v1.py
```

You'll get a simple control panel:

1) Run EC data script  
2) Run MS data script  
Q) Quit

Pick `1` or `2` to process files in EC or MS mode respectively.

### Programmatic / headless usage

You can call the core function from another script or from a one-liner:

```bash
python -c "from RIs_v1 import run_ultrankfrq; run_ultrankfrq('EC')"
```

Or in a Python interpreter:

```python
from RIs_v1 import run_ultrankfrq
run_ultrankfrq("MS")
```

If running on a headless server where matplotlib has no display, ensure the `Agg` backend is used before import/plots (or modify the script):

```python
import matplotlib
matplotlib.use("Agg")
from RIs_v1 import run_ultrankfrq
run_ultrankfrq("EC")
```

---

## Outputs

For each processed Excel file, the script writes to `OUTPUT_FOLDER`:

- `{base_name}_{mode}_interpolated.xlsx`  
  - One sheet per original sheet, containing:
    - `Distance (m)` and `<sheet_name>_mean` (representative profile interpolated onto the common grid).
- `{base_name}_{mode}_representativeness_scores.csv`  
  - CSV with columns `mean_std`, `amplitude`, `score` indexed by `sheet`.
- `{base_name}_{mode}_representative_profiles.png`  
  - A plot showing representative profiles for all sheets (legend includes score).

Printed console output will show a ranking of sheets by score and the best frequency/sheet.

---

## How the representativeness score is computed

For each sheet:
1. Interpolate each trace to the same regular distance grid.
2. Compute the representative profile as the pointwise mean across traces.
3. Compute pointwise standard deviation across traces and take the mean of that standard-deviation profile (`mean_std`).
4. Compute amplitude of the representative profile: max(rep_prof) - min(rep_prof) (`amplitude`).
5. Score = amplitude / max(mean_std, SCORE_EPSILON).

Interpretation: higher score means the representative profile has a larger amplitude relative to inter-trace variability, i.e., more "distinct" and consistent signal.

---

## Troubleshooting

- "No .xlsx files found in input folder"  
  - Ensure Excel files exist in the current working directory, or change `INPUT_FOLDER`.

- Excel open errors when writing (`PermissionError`)  
  - Make sure output files are not open in Excel while the script runs.

- Missing backends / plotting errors on servers  
  - Use the `Agg` backend (see above).

- Interpolation returns NaNs or no interpolated lines  
  - Check that sheets have at least two valid (distance, value) pairs per trace, and that distance spans more than `COMMON_DISTANCE_STEP`.

- If you need to support additional Excel engines, ensure `openpyxl` or other engines are installed.

---

## Contributing

- If you want to add features (e.g., CLI arguments, more robust file discovery, unit tests, or alternative scoring), please open an issue or submit a PR.
- Suggested enhancements:
  - Add CLI flags (argparse) for `INPUT_FOLDER`, `OUTPUT_FOLDER`, `COMMON_DISTANCE_STEP`, `INTERP_KIND`, and mode.
  - Allow saving all per-trace interpolated results in the output Excel for debugging.
  - Add unit tests for interpolation and scoring logic.

---

## License & contact

- License: MIT (you can add a LICENSE file to the repo if desired).
- Author / maintainer: skokolakis — open an issue or reach out via GitHub: [skokolakis](https://github.com/skokolakis)
