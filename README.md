# Representative Incision Tool — GEM Multi-Frequency Analysis

## Overview

**RIs_v2.py** is a Streamlit web application for analyzing multi-frequency electromagnetic induction (EMI) survey data. It automatically identifies the **most representative frequencies** (EC and MS channels) for geophysical profiling by scoring them based on signal-to-noise ratio.

### Key Features

- 🔄 **Dual-format support**: Upload GEM instrument data as `.csv` or `.xlsx` (full precision)
- 📊 **Intelligent ranking**: Frequencies ranked by signal-to-noise score (amplitude ÷ noise)
- 🎨 **Interactive graph editor**: Customize axis limits, line styles, titles, and visibility
- 📥 **Multi-format downloads**: Export interpolated profiles (XLSX), scores (CSV), and plots (PNG)
- 🔍 **Per-frequency detail plots**: Individual traces with mean ± 1σ envelope
- ⚙️ **Flexible interpolation**: Linear, cubic, nearest-neighbor, or quadratic methods
- ⚠️ **Data quality warnings**: Auto-detects and warns about precision loss in CSV exports

---

## Installation

### Requirements
- Python 3.9+
- Dependencies: `streamlit`, `pandas`, `numpy`, `scipy`, `matplotlib`, `openpyxl`

### Setup

```bash
# Clone the repository
git clone <repo-url>
cd GEM_representative_insicision

# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run RIs_v2.py
```

The app opens in your browser at `http://localhost:8501`.

---

## Usage

### Uploading Data

**GEM instrument format** (recommended):
- Single file (CSV or XLSX) with all frequencies and lines in one table
- Columns: `Line`, `Y`, `EC4525Hz[mS/m]`, `EC20025Hz[mS/m]`, ..., `MSusc4525Hz[1/1000]`, ...
- Automatically detects and shows both **EC** and **MS** results in separate tabs

**Legacy format** (multi-sheet XLSX):
- One frequency per Excel sheet
- Column 0 = distance (meters), Columns 1+ = one per line/trace
- Select measurement mode (EC or MS) from sidebar

### Understanding the Score

$$\text{Score} = \frac{\text{Amplitude}}{\sigma_{\text{noise}}}$$

- **Amplitude** = max − min of the representative (mean) profile
- **σ_noise** depends on the data type:
  - **Multi-trace files** → mean between-trace standard deviation at each distance point
  - **Single-trace files (GEM)** → within-profile noise from a rolling-window smoother

**Higher score = cleaner, larger-contrast signal** → better for detailed profiling.

### Customization

1. **Sidebar settings**:
   - Interpolation method (affects noise estimation and smoothness)
   - Distance step (m) — controls interpolation grid density

2. **Graph editor** (per mode):
   - Toggle visibility of individual frequencies
   - Set axis limits (auto or manual)
   - Adjust line width and style
   - Edit axis labels and plot title
   - Show/hide individual traces and ±1σ envelope

3. **Downloads**:
   - **Interpolated profiles** — XLSX with distance column + mean for each frequency
   - **Scores** — CSV with amplitude, noise, and score for each frequency
   - **Overview plot** — PNG showing all frequencies on one axes

---

## Data Quality Notes

### CSV vs XLSX Precision

The GEM instrument exports different precision levels:

| Format | EC precision | MS precision | Note |
|---|---|---|---|
| `.csv` | Integer (no decimals) | 1 d.p. | **Reduced precision** — scores may differ slightly |
| `.xlsx` | 3 d.p. | 4+ d.p. | **Full instrument precision** — recommended for analysis |

The app **automatically detects and warns** when a CSV has reduced precision. For quantitative frequency comparison, always use the XLSX file.

### Single-Trace Files

Files with only one line (one trace per frequency) still produce valid scores:
- Between-trace std = NaN → fallback to intra-profile SNR
- Noise is estimated as residual variation after removing the low-frequency signal trend
- Ranking still works correctly; high score = large amplitude with low residual noise

---

## Interpolation Methods

All methods create a common distance grid and interpolate each trace onto it. Differences:

| Method | Speed | Smoothness | Best for |
|---|---|---|---|
| **Linear** | Fast | Piecewise linear | Quick preview, small datasets |
| **Nearest** | Fastest | Step function | Testing, sparse data |
| **Quadratic** | Medium | Smooth polynomial | Smooth profiles, ≥3 points/trace |
| **Cubic** | Medium | Cubic spline | High-fidelity representation, ≥4 points |

---

## Output Files

### Interpolated profiles (`.xlsx`)
One sheet per frequency with columns:
- `Distance (m)` — common interpolation grid
- `{Frequency}_mean` — representative (mean across lines) profile

### Scores (`.csv`)
Row per frequency:
- `mean_std` — noise level (σ) in same units as measurement
- `amplitude` — dynamic range of the profile
- `score` — final ranking metric (amplitude ÷ noise)

### Overview plot (`.png`)
All frequencies on a single axes with:
- Legend showing score for each frequency
- Score-based color gradient (green = high, red = low)
- User-customized axis limits and labels

---

## References

1. Won, I.J., Keiswetter, D.A., Fields, G.R.A. & Sutton, L.C. (1996). GEM-2: A new multifrequency electromagnetic sensor. *Journal of Environmental and Engineering Geophysics*, **1**(2), 129–137. https://doi.org/10.4133/JEEG1.2.129

2. McNeill, J.D. (1980). *Electromagnetic terrain conductivity measurement at low induction numbers*. Technical Note TN-6, Geonics Limited, Mississauga, Canada.

3. Delefortrie, S., Saey, T., Van De Vijver, E., De Smedt, P., Missiaen, T., Demerre, I. & Van Meirvenne, M. (2014). Frequency domain electromagnetic induction survey in the intertidal zone: data acquisition and correction procedures. *Journal of Applied Geophysics*, **100**, 119–130. https://doi.org/10.1016/j.jappgeo.2013.10.017

4. Callegary, J.B., Ferré, T.P.A. & Groom, R.W. (2007). Vertical spatial sensitivity and exploration depth of low-induction-number electromagnetic induction instruments. *Vadose Zone Journal*, **6**(1), 158–167. https://doi.org/10.2136/vzj2006.0120

5. Reynolds, J.M. (2011). *An Introduction to Applied and Environmental Geophysics* (2nd ed.). Wiley-Blackwell.

---

## Architecture

```
RIs_v2.py (37 KB)
├── GEM format detection (is_gem_format, parse_gem_dataframe, pivot_gem_frequency)
├── Data processing (process_sheet, process_file, process_gem_file)
├── Plotting (make_overview_figure, make_sheet_figure)
├── Export (build_excel_download, build_scores_csv, fig_to_png)
├── UI components (render_graph_editor, _render_mode_section)
├── Dispatch (render_legacy_results, render_gem_results)
└── Streamlit app (main)
```

**No external APIs or databases** — all processing is local and deterministic.

---

## Troubleshooting

### "No usable sheets found"
- Reduce `Distance step` in sidebar (currently: 0.5 m by default)
- Ensure first column is numeric distance
- Check that traces have ≥2 valid points after NaN removal

### "Reduced precision detected" warning (CSV only)
- This is expected — GEM CSV exports round EC values
- Use the XLSX file for full-precision analysis
- Scores may differ by <5% between CSV and XLSX due to rounding

### Different interpolation methods give different scores
- Cubic/quadratic smooth out high-frequency noise differently
- Linear interpolation preserves noise; quadratic reduces it
- Choose method based on your noise tolerance


## License

See `LICENSE` file for terms.

---

## Citation

See `CITATION.cff` file.
