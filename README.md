# RIs.py — UltraRank Frequency Analysis (CLI)

## Overview

**RIs.py** is a command-line tool for analysing multi-frequency electromagnetic induction (EMI) survey data from legacy multi-sheet XLSX files. It ranks frequencies by signal-to-noise score to identify the most representative operating frequency for EC or MS profiling.

It is the command-line predecessor to **RIs_v2.py** (Streamlit web app) and shares the same core scoring logic. Use it when you need a quick, scriptable, no-browser workflow.

---

## Requirements

- Python 3.9+
- `pandas`, `numpy`, `matplotlib`, `scipy`, `openpyxl`

```bash
pip install pandas numpy matplotlib scipy openpyxl
```

---

## Usage

```bash
python RIs.py
```

The interactive control panel presents three options:

```
===== Script Control Panel =====
1) Run EC data script
2) Run MS data script
Q) Quit
```

After selecting a mode you are prompted to choose an interpolation method:

```
  Interpolation method:
    1) linear
    2) cubic
    3) nearest
    4) quadratic
    5) pchip
    6) akima
    7) polynomial
  Select method [1-7, default=1 linear]:
```

Press **Enter** to accept the default (linear). The script then processes every `.xlsx` file found in the current working directory and writes results to `output_profiles/`.

---

## Input format

RIs.py expects **legacy multi-sheet XLSX files** placed in the same directory as the script:

| Element | Requirement |
|---|---|
| File location | Same folder as `RIs.py` (current working directory) |
| File format | `.xlsx` (files beginning with `~$` are ignored) |
| Sheet structure | One frequency per sheet |
| Column 0 | Distance along transect (metres, numeric) |
| Columns 1+ | One column per survey line / trace (numeric) |

Multiple `.xlsx` files are processed in a single run.

---

## Output files

All outputs are written to `output_profiles/` (created automatically).

| File | Content |
|---|---|
| `{stem}_{mode}_{method}_interpolated.xlsx` | Interpolated mean profiles — one sheet per frequency, with `Distance (m)` and `{freq}_mean` columns |
| `{stem}_{mode}_{method}_representativeness_scores.csv` | Per-frequency scores: `mean_std`, `amplitude`, `score` |
| `{stem}_{mode}_{method}_representative_profiles.png` | Overview plot of all mean profiles, labelled with scores (300 dpi) |

`{stem}` is the input filename without extension, `{mode}` is `EC` or `MS`, and `{method}` is the chosen interpolation method.

---

## Scoring

Each frequency receives a representativeness score:

$$\text{Score} = \frac{A}{\sigma_{\text{noise}}}$$

| Symbol | Meaning |
|---|---|
| **A** | Amplitude — peak-to-trough range of the mean profile: max(p̄) − min(p̄) |
| **σ_noise** | Noise — see below |

**Noise estimation:**

- **Multi-trace (≥ 2 lines):** mean of the point-wise population standard deviation across all traces at each distance step (ddof = 0, because the traces are the full ensemble of survey passes, not a sample).
- **Single-trace fallback:** residual std after subtracting a rolling-mean smoother (window = max(5, N/10)). This decomposes the signal into a geological trend and a high-frequency noise component.
- **Near-zero noise:** score falls back to raw amplitude A to avoid numerical instability.

A higher score means the frequency resolves large subsurface contrasts clearly above the noise level — the standard geophysical signal-to-noise criterion.

---

## Interpolation methods

All traces are resampled onto a common evenly-spaced distance grid (`np.linspace`) before averaging.

| Method | Min. points | Notes |
|---|---|---|
| **linear** | 2 | Default. Piecewise linear; conservative, no overshoot. |
| **nearest** | 1 | Step-function; useful for sparse or categorical-style data. |
| **quadratic** | 3 | Quadratic B-spline; smoother than linear with modest curvature. |
| **cubic** | 4 | Cubic spline (continuous 2nd derivative); may overshoot sharp boundaries. |
| **pchip** | 2 | Shape-preserving, monotone within each interval; avoids cubic overshoot. |
| **akima** | 5 | Local spline using neighbouring slopes; robust to isolated outliers. |
| **polynomial** | 3 | Global least-squares polynomial (degree ≤ 5); avoid for long profiles. |

Columns with fewer than the required minimum points are skipped and a warning is printed.

---

## Console output

For each processed file the script prints a ranked frequency table:

```
Frequency Ranking [EC] — H4Keratos_GEM_FREQ_EC
1. 4525Hz  | Score=18.34 | Std=0.0023 | Amp=0.0421
2. 9225Hz  | Score=12.71 | Std=0.0031 | Amp=0.0394
3. 20025Hz | Score= 8.05 | Std=0.0048 | Amp=0.0386

Best Frequency: 4525Hz (Score=18.34)
```

---

## Configuration

Hard-coded constants at the top of the file can be edited directly:

| Constant | Default | Effect |
|---|---|---|
| `COMMON_DISTANCE_STEP` | `0.5` m | Grid spacing for interpolation |
| `INPUT_FOLDER` | `.` (current dir) | Where `.xlsx` files are read from |
| `OUTPUT_FOLDER` | `output_profiles/` | Where results are written |
| `SCORE_EPSILON` | `1e-8` | Guard against division by zero in scoring |

---

## Relationship to other scripts

| Script | Interface | Format support | Interpolation |
|---|---|---|---|
| `RIs_v1.py` | CLI | Legacy XLSX only | Linear only |
| **`RIs.py`** | CLI | Legacy XLSX only | All 7 methods |
| `RIs_v2.py` | Streamlit web app | Legacy XLSX + GEM CSV/XLSX | All 7 methods |

`RIs.py` is the CLI-equivalent of `RIs_v2.py` for legacy-format files, adding the full interpolation method menu absent from `RIs_v1.py`.

---

## References

- Won, I.J. et al. (1996). GEM-2: A new multifrequency electromagnetic sensor. *J. Environ. Eng. Geophys.*, **1**(2), 129–137. https://doi.org/10.4133/JEEG1.2.129
- McNeill, J.D. (1980). *Electromagnetic terrain conductivity measurement at low induction numbers*. Technical Note TN-6, Geonics Limited.
- Callegary, J.B., Ferré, T.P.A. & Groom, R.W. (2007). Vertical spatial sensitivity and exploration depth of low-induction-number EMI instruments. *Vadose Zone J.*, **6**(1), 158–167. https://doi.org/10.2136/vzj2006.0120
- Sheriff, R.E. & Geldart, L.P. (1995). *Exploration Seismology* (2nd ed.). Cambridge University Press.
