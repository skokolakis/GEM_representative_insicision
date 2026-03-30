"""
UltraRank Frequency — Streamlit UI
Refactored from RIs_v1.py

Run with:
    streamlit run RIs_v2.py
"""
from __future__ import annotations

import io
import logging
import re
from dataclasses import dataclass, field
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st
from scipy.interpolate import CubicSpline, make_interp_spline

# ---------------------------------------------------------------------------
# Configuration (all in one place, easily overridden via Streamlit widgets)
# ---------------------------------------------------------------------------
DEFAULT_DISTANCE_STEP = 0.5
DEFAULT_INTERP_KIND = "linear"
SCORE_EPSILON = 1e-8
EXCEL_SHEET_NAME_MAX = 31

MODES = {
    "EC": "Mean EC Response (S/m)",
    "MS": "Mean MS Response (10⁻⁵ SI)",
}

LINE_STYLES = {
    "Solid": "-",
    "Dashed": "--",
    "Dash-dot": "-.",
    "Dotted": ":",
}

# GEM instrument format detection
GEM_EC_PATTERN = re.compile(r'^EC(\d+)Hz\[mS/m\]$')
GEM_MS_PATTERN = re.compile(r'^MSusc(\d+)Hz\[1/1000\]$')
GEM_REQUIRED_COLS = {'Line', 'Y'}

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Graph editor options
# ---------------------------------------------------------------------------

@dataclass
class GraphOptions:
    """Visual settings collected from the graph editor panel."""
    selected_sheets: list[str] = field(default_factory=list)  # empty → show all
    x_lim: tuple[float, float] | None = None
    y_lim: tuple[float, float] | None = None
    show_grid: bool = True
    line_width: float = 2.0
    line_style: str = "-"
    show_traces: bool = True
    show_envelope: bool = True
    plot_title: str = ""   # empty → use auto-generated default
    x_label: str = ""      # empty → "Distance (m)"
    y_label: str = ""      # empty → MODES[mode]


# ---------------------------------------------------------------------------
# GEM format detection & pivoting
# ---------------------------------------------------------------------------

def is_gem_format(df: pd.DataFrame) -> bool:
    """Return True if *df* has the GEM instrument column structure."""
    cols = set(str(c) for c in df.columns)
    if not GEM_REQUIRED_COLS.issubset(cols):
        return False
    has_ec = any(GEM_EC_PATTERN.match(str(c)) for c in df.columns)
    has_ms = any(GEM_MS_PATTERN.match(str(c)) for c in df.columns)
    return has_ec or has_ms


def pivot_gem_frequency(
    df: pd.DataFrame,
    value_col: str,
    distance_col: str = "Y",
    line_col: str = "Line",
) -> pd.DataFrame:
    """
    Pivot a GEM dataframe for one frequency column into process_sheet format.

    Returns a DataFrame with column 0 = distance (Y values) and
    columns 1+ = one column per unique Line value.
    """
    subset = df[[distance_col, line_col, value_col]].copy()
    subset[distance_col] = pd.to_numeric(subset[distance_col], errors="coerce")
    subset[value_col] = pd.to_numeric(subset[value_col], errors="coerce")
    subset = subset.dropna(subset=[distance_col, value_col])

    pivoted = subset.pivot_table(
        index=distance_col,
        columns=line_col,
        values=value_col,
        aggfunc="mean",
    )

    # Rename columns to "Line_0", "Line_1", etc.
    pivoted.columns = [f"Line_{int(c)}" for c in pivoted.columns]

    # Reset index so column 0 = distance (what process_sheet expects)
    pivoted = pivoted.reset_index()
    return pivoted


def parse_gem_dataframe(
    df: pd.DataFrame,
) -> dict[str, dict[str, pd.DataFrame]]:
    """
    Parse a GEM-format DataFrame into pivoted DataFrames per mode and frequency.

    Returns
    -------
    {"EC": {"4525Hz": pivoted_df, ...}, "MS": {"4525Hz": pivoted_df, ...}}
    """
    result: dict[str, dict[str, pd.DataFrame]] = {"EC": {}, "MS": {}}

    for col in df.columns:
        col_str = str(col)

        ec_match = GEM_EC_PATTERN.match(col_str)
        if ec_match:
            freq = ec_match.group(1)
            label = f"{freq}Hz"
            result["EC"][label] = pivot_gem_frequency(df, value_col=col_str)
            continue

        ms_match = GEM_MS_PATTERN.match(col_str)
        if ms_match:
            freq = ms_match.group(1)
            label = f"{freq}Hz"
            result["MS"][label] = pivot_gem_frequency(df, value_col=col_str)

    return result


@st.cache_data(show_spinner=False)
def process_gem_file(
    file_bytes: bytes,
    file_name: str,
    distance_step: float = DEFAULT_DISTANCE_STEP,
    interp_kind: str = DEFAULT_INTERP_KIND,
) -> tuple[dict[str, dict[str, pd.DataFrame]], dict[str, dict[str, dict]], list[str]]:
    """
    Process a GEM-format file (CSV or XLSX).

    Returns
    -------
    output_data : {"EC": {freq: interp_df, ...}, "MS": {freq: interp_df, ...}}
    scores      : {"EC": {freq: score_dict, ...}, "MS": {freq: score_dict, ...}}
    warnings    : list of warning strings
    """
    warnings: list[str] = []

    # Read raw data based on extension
    if file_name.lower().endswith(".csv"):
        try:
            raw_df = pd.read_csv(io.BytesIO(file_bytes))
        except Exception as exc:
            warnings.append(f"Could not read CSV: {exc}")
            return {"EC": {}, "MS": {}}, {"EC": {}, "MS": {}}, warnings
    else:
        try:
            raw_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0)
        except Exception as exc:
            warnings.append(f"Could not read Excel: {exc}")
            return {"EC": {}, "MS": {}}, {"EC": {}, "MS": {}}, warnings

    # ── Precision check ─────────────────────────────────────────────────────
    # GEM CSV exports often round EC values to integers and MS to 1 d.p.,
    # while the XLSX retains full instrument precision (3+ decimal places).
    # Detect this and warn the user so they know scores may differ.
    if file_name.lower().endswith(".csv"):
        low_prec_cols: list[str] = []
        for col in raw_df.columns:
            col_str = str(col)
            if GEM_EC_PATTERN.match(col_str) or GEM_MS_PATTERN.match(col_str):
                vals = pd.to_numeric(raw_df[col], errors="coerce").dropna()
                # Flag columns where every value is an integer (no fractional part)
                if len(vals) > 0 and (vals % 1 == 0).all():
                    low_prec_cols.append(col_str)
        if low_prec_cols:
            warnings.append(
                "⚠️ Reduced precision detected: the following columns contain only "
                f"integer values — {', '.join(low_prec_cols)}. "
                "GEM CSV exports typically round measurements, which causes scores to "
                "differ slightly from the XLSX equivalent. "
                "Use the XLSX file for full instrument precision."
            )

    gem_data = parse_gem_dataframe(raw_df)

    output_data: dict[str, dict[str, pd.DataFrame]] = {"EC": {}, "MS": {}}
    scores: dict[str, dict[str, dict]] = {"EC": {}, "MS": {}}

    for mode_key in ("EC", "MS"):
        for freq_label, pivoted_df in gem_data[mode_key].items():
            interp_df, score_dict, error, col_warnings = process_sheet(
                pivoted_df, distance_step, interp_kind
            )

            for w in col_warnings:
                warnings.append(f"{mode_key} {freq_label}: {w}")

            if error:
                warnings.append(f"{mode_key} {freq_label}: skipped — {error}")
                continue

            output_data[mode_key][freq_label] = interp_df
            scores[mode_key][freq_label] = score_dict

    return output_data, scores, warnings


# ---------------------------------------------------------------------------
# Core processing (pure functions — no side effects, no sys.exit)
# ---------------------------------------------------------------------------

def process_sheet(
    df: pd.DataFrame,
    distance_step: float = DEFAULT_DISTANCE_STEP,
    interp_kind: str = DEFAULT_INTERP_KIND,
) -> tuple[pd.DataFrame | None, dict | None, str | None, list[str]]:
    """
    Process a single sheet.

    Returns
    -------
    interpolated_df : DataFrame indexed by common distance, columns = original trace columns
    score_dict      : {"mean_std": float, "amplitude": float, "score": float}
    error           : human-readable reason for failure, or None on success
    col_warnings    : list of per-column warning strings surfaced to the UI
    """
    col_warnings: list[str] = []

    if df.shape[1] < 2:
        return None, None, "fewer than 2 columns", col_warnings

    distance = pd.to_numeric(df.iloc[:, 0], errors="coerce").values
    line_data = df.iloc[:, 1:].apply(pd.to_numeric, errors="coerce")

    if np.all(np.isnan(distance)):
        return None, None, "distance column is all NaN", col_warnings

    min_dist = float(np.nanmin(distance))
    max_dist = float(np.nanmax(distance))

    if not (np.isfinite(min_dist) and np.isfinite(max_dist)):
        return None, None, "non-finite distance range", col_warnings

    span = max_dist - min_dist
    if span < distance_step:
        return None, None, f"distance span ({span:.2f} m) < step ({distance_step} m)", col_warnings

    # Use linspace to avoid float accumulation past max_dist
    n_points = int(round(span / distance_step)) + 1
    common_dist = np.linspace(min_dist, max_dist, n_points)

    interpolated_lines: dict[str, np.ndarray] = {}

    for col in line_data.columns:
        y = line_data[col].values
        mask = ~np.isnan(distance) & ~np.isnan(y)

        if mask.sum() < 2:
            log.debug("  Column %s: fewer than 2 valid points, skipped.", col)
            continue

        df_xy = (
            pd.DataFrame({"d": distance[mask], "y": y[mask]})
            .groupby("d", as_index=False)
            .mean()
            .sort_values("d")
            .drop_duplicates(subset="d")
        )

        if not df_xy["d"].is_monotonic_increasing:
            col_warnings.append(f"column '{col}': distance not monotonic after dedup, skipped")
            continue

        if df_xy.shape[0] < 2:
            continue

        xp = df_xy["d"].values
        yp = df_xy["y"].values

        try:
            if interp_kind == "linear":
                interpolated_lines[col] = np.interp(common_dist, xp, yp)
            elif interp_kind == "nearest":
                idx = np.searchsorted(xp, common_dist, side="left")
                idx_left = np.clip(idx - 1, 0, len(xp) - 1)
                idx_right = np.clip(idx, 0, len(xp) - 1)
                nearest = np.where(
                    np.abs(common_dist - xp[idx_left]) <= np.abs(common_dist - xp[idx_right]),
                    idx_left,
                    idx_right,
                )
                interpolated_lines[col] = yp[nearest]
            elif interp_kind == "quadratic":
                if len(xp) < 3:
                    col_warnings.append(
                        f"column '{col}': need ≥3 points for quadratic interpolation, skipped"
                    )
                    continue
                spl = make_interp_spline(xp, yp, k=2)
                interpolated_lines[col] = spl(common_dist)
            elif interp_kind == "cubic":
                spl = CubicSpline(xp, yp)
                interpolated_lines[col] = spl(common_dist)
        except Exception as exc:
            col_warnings.append(f"column '{col}': interpolation failed — {exc}")
            continue

    if not interpolated_lines:
        return None, None, "no columns survived interpolation", col_warnings

    interpolated_df = pd.DataFrame(interpolated_lines, index=common_dist)
    interpolated_df.index.name = "Distance (m)"

    rep_prof = interpolated_df.mean(axis=1, skipna=True)
    std_prof = interpolated_df.std(axis=1, skipna=True)

    mean_std = float(std_prof.mean())
    amplitude = float(np.nanmax(rep_prof) - np.nanmin(rep_prof))

    if np.isnan(mean_std) or mean_std < SCORE_EPSILON:
        # Single trace or near-identical traces: between-trace std is meaningless.
        # Fall back to intra-profile SNR: amplitude / within-profile noise.
        # Noise is estimated as the std of residuals from a rolling-mean smoother,
        # which separates low-frequency signal from high-frequency noise.
        window = max(5, len(rep_prof) // 10)
        smoothed = rep_prof.rolling(window=window, center=True, min_periods=1).mean()
        intra_noise = float((rep_prof - smoothed).std())
        mean_std = intra_noise  # store noise level so it appears in the UI
        score = amplitude / max(intra_noise, SCORE_EPSILON)
    else:
        # Multi-trace: score = signal range / between-trace variability (SNR)
        score = amplitude / mean_std

    score_dict = {"mean_std": mean_std, "amplitude": amplitude, "score": score}
    return interpolated_df, score_dict, None, col_warnings


@st.cache_data(show_spinner=False)
def process_file(
    excel_bytes: bytes,
    mode: str,
    distance_step: float = DEFAULT_DISTANCE_STEP,
    interp_kind: str = DEFAULT_INTERP_KIND,
) -> tuple[dict, dict, list[str]]:
    """
    Process all sheets in an uploaded Excel file (legacy multi-sheet format).

    Returns
    -------
    output_data               : {sheet_name: interpolated_df}
    representativeness_scores : {sheet_name: score_dict}
    warnings                  : list of warning strings
    """
    output_data: dict[str, pd.DataFrame] = {}
    scores: dict[str, dict] = {}
    warnings: list[str] = []

    try:
        excel = pd.ExcelFile(io.BytesIO(excel_bytes))
    except Exception as exc:
        warnings.append(f"Could not open file: {exc}")
        return output_data, scores, warnings

    for sheet_name in excel.sheet_names:
        sheet_name = str(sheet_name)  # guard against integer sheet names (e.g. 9000)
        try:
            df = pd.read_excel(excel, sheet_name=sheet_name)
        except Exception as exc:
            warnings.append(f"Sheet '{sheet_name}': read error — {exc}")
            continue

        interp_df, score_dict, error, col_warnings = process_sheet(df, distance_step, interp_kind)

        for w in col_warnings:
            warnings.append(f"Sheet '{sheet_name}': {w}")

        if error:
            warnings.append(f"Sheet '{sheet_name}': skipped — {error}")
            continue

        output_data[sheet_name] = interp_df
        scores[sheet_name] = score_dict

    return output_data, scores, warnings


# ---------------------------------------------------------------------------
# Plot helpers (return Figure objects, never touch global pyplot state)
# ---------------------------------------------------------------------------

def make_overview_figure(
    output_data: dict[str, pd.DataFrame],
    scores: dict[str, dict],
    mode: str,
    file_name: str,
    opts: GraphOptions | None = None,
) -> plt.Figure:
    """All representative profiles on one axes."""
    if opts is None:
        opts = GraphOptions()

    visible = opts.selected_sheets if opts.selected_sheets else list(output_data.keys())

    fig, ax = plt.subplots(figsize=(10, 5))
    y_label = MODES[mode]

    for sheet_name in visible:
        interp_df = output_data.get(sheet_name)
        if interp_df is None:
            continue
        rep_prof = interp_df.mean(axis=1, skipna=True)
        common_dist = interp_df.index.values
        sc = scores[sheet_name]["score"]
        if not np.all(np.isnan(rep_prof.values)):
            ax.plot(
                common_dist,
                rep_prof.values,
                label=f"{sheet_name} (score={sc:.2f})",
                linewidth=opts.line_width,
                linestyle=opts.line_style,
            )

    ax.set_xlabel(opts.x_label or "Distance (m)")
    ax.set_ylabel(opts.y_label or y_label)
    ax.set_title(opts.plot_title or f"Representative profiles [{mode}] — {file_name}")
    ax.legend(fontsize=8)
    ax.grid(opts.show_grid, alpha=0.4)
    if opts.x_lim is not None:
        ax.set_xlim(opts.x_lim)
    if opts.y_lim is not None:
        ax.set_ylim(opts.y_lim)
    fig.tight_layout()
    return fig


def make_sheet_figure(
    sheet_name: str,
    interp_df: pd.DataFrame,
    mode: str,
    opts: GraphOptions | None = None,
) -> plt.Figure:
    """Per-sheet plot: individual traces + mean +/- 1 sigma envelope."""
    if opts is None:
        opts = GraphOptions()

    fig, ax = plt.subplots(figsize=(10, 4))
    common_dist = interp_df.index.values
    rep_prof = interp_df.mean(axis=1, skipna=True).values
    std_prof = interp_df.std(axis=1, skipna=True).values

    # Individual traces (thin, semi-transparent)
    if opts.show_traces:
        for col in interp_df.columns:
            ax.plot(
                common_dist,
                interp_df[col].values,
                color="steelblue",
                alpha=0.25,
                linewidth=0.8,
                linestyle=opts.line_style,
            )

    # Mean profile
    ax.plot(
        common_dist,
        rep_prof,
        color="navy",
        linewidth=opts.line_width,
        linestyle=opts.line_style,
        label="Mean",
    )

    # +/- 1 sigma envelope
    if opts.show_envelope:
        ax.fill_between(
            common_dist,
            rep_prof - std_prof,
            rep_prof + std_prof,
            alpha=0.2,
            color="navy",
            label="±1σ",
        )

    ax.set_xlabel(opts.x_label or "Distance (m)")
    ax.set_ylabel(opts.y_label or MODES[mode])
    ax.set_title(f"{sheet_name} — individual traces & representative profile")
    ax.legend()
    ax.grid(opts.show_grid, alpha=0.4)
    if opts.x_lim is not None:
        ax.set_xlim(opts.x_lim)
    if opts.y_lim is not None:
        ax.set_ylim(opts.y_lim)
    fig.tight_layout()
    return fig


# ---------------------------------------------------------------------------
# Export helpers
# ---------------------------------------------------------------------------

def build_excel_download(output_data: dict[str, pd.DataFrame]) -> bytes:
    """Pack all interpolated sheets into a single Excel workbook."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, interp_df in output_data.items():
            rep_prof = interp_df.mean(axis=1, skipna=True)
            col_label = f"{sheet_name[:EXCEL_SHEET_NAME_MAX - 5]}_mean"
            out_df = pd.DataFrame(
                {
                    "Distance (m)": interp_df.index.values,
                    col_label: rep_prof.values,
                }
            )
            out_df.to_excel(writer, sheet_name=sheet_name[:EXCEL_SHEET_NAME_MAX], index=False)
    return buf.getvalue()


def build_scores_csv(scores: dict[str, dict]) -> bytes:
    df = pd.DataFrame.from_dict(scores, orient="index")
    df.index.name = "sheet"
    return df.to_csv().encode()


def fig_to_png(fig: plt.Figure) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200)
    buf.seek(0)
    return buf.read()


# ---------------------------------------------------------------------------
# Graph editor UI helper
# ---------------------------------------------------------------------------

def render_graph_editor(
    output_data: dict[str, pd.DataFrame],
    file_key: str,
) -> GraphOptions:
    """
    Render the graph editor expander and return the current GraphOptions.

    Parameters
    ----------
    output_data : processed sheets for this file
    file_key    : unique string used to namespace widget keys per file
    """
    all_sheet_names = list(output_data.keys())

    # Compute data extents for axis-limit defaults
    all_x = np.concatenate([df.index.values for df in output_data.values()])
    all_y_means = np.concatenate(
        [df.mean(axis=1, skipna=True).values for df in output_data.values()]
    )
    valid_y = all_y_means[np.isfinite(all_y_means)]
    x_data_min, x_data_max = float(all_x.min()), float(all_x.max())
    y_data_min = float(valid_y.min()) if len(valid_y) else 0.0
    y_data_max = float(valid_y.max()) if len(valid_y) else 1.0

    opts = GraphOptions()

    with st.expander("✏️ Graph Editor", expanded=False):
        col_sheets, col_style, col_axes = st.columns([2, 1, 2])

        with col_sheets:
            st.markdown("**Frequencies / Sheets**")
            selected = st.multiselect(
                "Visible items",
                options=all_sheet_names,
                default=all_sheet_names,
                key=f"ge_sheets_{file_key}",
                label_visibility="collapsed",
            )
            opts.selected_sheets = selected

            st.markdown("**Detail plots**")
            opts.show_traces = st.checkbox(
                "Show individual traces",
                value=True,
                key=f"ge_traces_{file_key}",
            )
            opts.show_envelope = st.checkbox(
                "Show ±1σ envelope",
                value=True,
                key=f"ge_envelope_{file_key}",
            )

        with col_style:
            st.markdown("**Style**")
            opts.show_grid = st.checkbox(
                "Grid",
                value=True,
                key=f"ge_grid_{file_key}",
            )
            opts.line_width = st.slider(
                "Line width",
                min_value=0.5,
                max_value=5.0,
                value=2.0,
                step=0.5,
                key=f"ge_lw_{file_key}",
            )
            style_label = st.selectbox(
                "Line style",
                options=list(LINE_STYLES.keys()),
                index=0,
                key=f"ge_ls_{file_key}",
            )
            opts.line_style = LINE_STYLES[style_label]

        with col_axes:
            st.markdown("**X axis**")
            x_auto = st.checkbox("Auto", value=True, key=f"ge_xauto_{file_key}")
            if not x_auto:
                xc1, xc2 = st.columns(2)
                x_min = xc1.number_input(
                    "Min", value=x_data_min, key=f"ge_xmin_{file_key}", format="%.2f"
                )
                x_max = xc2.number_input(
                    "Max", value=x_data_max, key=f"ge_xmax_{file_key}", format="%.2f"
                )
                if x_min < x_max:
                    opts.x_lim = (x_min, x_max)
                else:
                    st.caption("⚠️ X min must be < X max")

            st.markdown("**Y axis**")
            y_auto = st.checkbox("Auto", value=True, key=f"ge_yauto_{file_key}")
            if not y_auto:
                yc1, yc2 = st.columns(2)
                y_min = yc1.number_input(
                    "Min", value=y_data_min, key=f"ge_ymin_{file_key}", format="%.4g"
                )
                y_max = yc2.number_input(
                    "Max", value=y_data_max, key=f"ge_ymax_{file_key}", format="%.4g"
                )
                if y_min < y_max:
                    opts.y_lim = (y_min, y_max)
                else:
                    st.caption("⚠️ Y min must be < Y max")

        st.divider()
        st.markdown("**Labels & Title**")
        lc1, lc2, lc3 = st.columns(3)
        opts.plot_title = lc1.text_input(
            "Overview plot title",
            value="",
            placeholder="Leave blank for default",
            key=f"ge_title_{file_key}",
        )
        opts.x_label = lc2.text_input(
            "X axis label",
            value="",
            placeholder="Distance (m)",
            key=f"ge_xlabel_{file_key}",
        )
        opts.y_label = lc3.text_input(
            "Y axis label",
            value="",
            placeholder=f"{next(iter(MODES.values()))} …",
            key=f"ge_ylabel_{file_key}",
        )

    return opts


# ---------------------------------------------------------------------------
# Reusable rendering section (shared by legacy and GEM paths)
# ---------------------------------------------------------------------------

def _render_mode_section(
    output_data: dict[str, pd.DataFrame],
    scores: dict[str, dict],
    mode: str,
    file_name: str,
    file_key: str,
) -> None:
    """
    Render ranking table, graph editor, plots, and downloads for one mode.

    Parameters
    ----------
    output_data : {freq_or_sheet_name: interpolated_df}
    scores      : {freq_or_sheet_name: score_dict}
    mode        : "EC" or "MS"
    file_name   : original uploaded filename (for plot titles)
    file_key    : unique string for Streamlit widget key namespacing
    """
    if not output_data:
        st.error("No usable frequencies / sheets found.")
        return

    # ── Ranking table ───────────────────────────────────────────────────
    ranking = sorted(scores.items(), key=lambda x: x[1]["score"], reverse=True)

    if not ranking:
        st.error("No sheets could be scored.")
        return

    stem = Path(file_name).stem

    rank_df = pd.DataFrame(
        [
            {
                "Rank": i,
                "Frequency / Sheet": name,
                "Score": round(m["score"], 2),
                "Amplitude": round(m["amplitude"], 6),
                "Noise (σ)": round(m["mean_std"], 6),
            }
            for i, (name, m) in enumerate(ranking, 1)
        ]
    )

    col_table, col_best = st.columns([3, 1])
    with col_table:
        st.markdown("### 🏆 Frequency Ranking")
        st.dataframe(
            rank_df.style.background_gradient(subset=["Score"], cmap="RdYlGn"),
            use_container_width=True,
            hide_index=True,
        )
    with col_best:
        best_name, best_metrics = ranking[0]
        st.metric(
            "Best frequency",
            best_name,
            f"score {best_metrics['score']:.2f}",
            help=(
                "Score = amplitude / noise (σ). "
                "For multi-trace files, noise is the mean between-trace std. "
                "For single-trace files, noise is the intra-profile residual std "
                "after a rolling-window smoother. Higher score = cleaner, larger signal."
            ),
        )
        st.metric("Amplitude", f"{best_metrics['amplitude']:.4g}")
        st.metric("Noise (σ)", f"{best_metrics['mean_std']:.4g}")

    # ── Graph editor ─────────────────────────────────────────────────────
    opts = render_graph_editor(output_data, file_key=file_key)

    # ── Overview plot ────────────────────────────────────────────────────
    st.markdown("### 📈 All representative profiles")
    overview_fig = make_overview_figure(output_data, scores, mode, file_name, opts)
    st.pyplot(overview_fig, use_container_width=True)
    plt.close(overview_fig)

    # ── Per-sheet detail ─────────────────────────────────────────────────
    with st.expander("🔍 Per-frequency detail plots", expanded=False):
        visible = opts.selected_sheets if opts.selected_sheets else list(output_data.keys())
        for sheet_name in visible:
            interp_df = output_data.get(sheet_name)
            if interp_df is None:
                continue
            sc = scores[sheet_name]
            st.markdown(
                f"**{sheet_name}** — score `{sc['score']:.2f}` | "
                f"amp `{sc['amplitude']:.4g}` | std `{sc['mean_std']:.4g}`"
            )
            sheet_fig = make_sheet_figure(sheet_name, interp_df, mode, opts)
            st.pyplot(sheet_fig, use_container_width=True)
            plt.close(sheet_fig)

    # ── Downloads ────────────────────────────────────────────────────────
    st.markdown("### 💾 Downloads")
    dl1, dl2, dl3 = st.columns(3)

    with dl1:
        st.download_button(
            label="📥 Interpolated profiles (.xlsx)",
            data=build_excel_download(output_data),
            file_name=f"{stem}_{mode}_interpolated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_xlsx_{file_key}",
        )

    with dl2:
        st.download_button(
            label="📥 Scores (.csv)",
            data=build_scores_csv(scores),
            file_name=f"{stem}_{mode}_scores.csv",
            mime="text/csv",
            key=f"dl_csv_{file_key}",
        )

    with dl3:
        download_fig = make_overview_figure(output_data, scores, mode, file_name, opts)
        png_bytes = fig_to_png(download_fig)
        plt.close(download_fig)
        st.download_button(
            label="📥 Overview plot (.png)",
            data=png_bytes,
            file_name=f"{stem}_{mode}_overview.png",
            mime="image/png",
            key=f"dl_png_{file_key}",
        )


# ---------------------------------------------------------------------------
# Per-format rendering dispatchers
# ---------------------------------------------------------------------------

def render_legacy_results(
    file_bytes: bytes,
    file_name: str,
    mode: str,
    distance_step: float,
    interp_kind: str,
) -> None:
    """Process and render a legacy multi-sheet Excel file."""
    stem = Path(file_name).stem

    with st.spinner("Processing…"):
        output_data, scores, warnings = process_file(
            file_bytes, mode,
            distance_step=distance_step,
            interp_kind=interp_kind,
        )

    if warnings:
        with st.expander(f"⚠️ {len(warnings)} warning(s)", expanded=False):
            for w in warnings:
                st.warning(w)

    if not output_data:
        st.error("No usable sheets found in this file.")
        if distance_step > 1.0:
            st.info("Try reducing the distance step in the sidebar.")
        return

    _render_mode_section(output_data, scores, mode, file_name, file_key=stem)


def render_gem_results(
    file_bytes: bytes,
    file_name: str,
    distance_step: float,
    interp_kind: str,
) -> None:
    """Process and render a GEM instrument file (CSV or XLSX)."""
    stem = Path(file_name).stem

    with st.spinner("Processing GEM file…"):
        output_data, scores, warnings = process_gem_file(
            file_bytes, file_name,
            distance_step=distance_step,
            interp_kind=interp_kind,
        )

    if warnings:
        with st.expander(f"⚠️ {len(warnings)} warning(s)", expanded=False):
            for w in warnings:
                st.warning(w)

    # Determine which modes have data
    available_modes = [m for m in ("EC", "MS") if output_data.get(m)]

    if not available_modes:
        st.error("No usable frequencies found in this GEM file.")
        if distance_step > 1.0:
            st.info("Try reducing the distance step in the sidebar.")
        return

    st.caption("GEM format detected — showing all frequencies for both EC and MS")

    tabs = st.tabs([f"{m} ({len(output_data[m])} frequencies)" for m in available_modes])

    for tab, mode_key in zip(tabs, available_modes):
        with tab:
            _render_mode_section(
                output_data[mode_key],
                scores[mode_key],
                mode_key,
                file_name,
                file_key=f"{stem}_{mode_key}",
            )


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

def main():
    st.set_page_config(
        page_title="UltraRank Frequency",
        page_icon="📡",
        layout="wide",
    )

    st.title("📡 Representative Incision Tool")
    st.caption("Geophysical representative profile builder — EC / MS modes")

    # ── Sidebar controls ────────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ Settings")

        mode = st.radio(
            "Measurement mode",
            list(MODES.keys()),
            horizontal=True,
            help="For legacy multi-sheet files. GEM files show both EC and MS automatically.",
        )

        distance_step = st.number_input(
            "Distance step (m)",
            min_value=0.01,
            max_value=100.0,
            value=DEFAULT_DISTANCE_STEP,
            step=0.1,
            format="%.2f",
        )

        interp_kind = st.selectbox(
            "Interpolation method",
            ["linear", "cubic", "nearest", "quadratic"],
            index=0,
        )

        if distance_step > 10.0:
            st.warning(
                f"Distance step is {distance_step:.2f} m. "
                "If your data spans less than this, all sheets will be skipped."
            )

        st.divider()
        st.markdown("**Output files are available for download after processing.**")

        st.divider()
        with st.expander("ℹ️ About", expanded=False):
            st.markdown(
                """
**Representative Incision Tool** — v2.0
Geophysical profile builder for multi-frequency EMI surveys.

---

#### How it works
1. **Upload** a GEM instrument file (`.csv` or `.xlsx`) or a legacy multi-sheet Excel file.
2. Each measurement frequency is interpolated onto a common distance grid.
3. Frequencies are **ranked** by a signal-to-noise score:

$$\\text{Score} = \\frac{\\text{Amplitude}}{\\sigma_{\\text{noise}}}$$

- **Multi-trace files** — σ is the mean between-trace standard deviation (trace-to-trace consistency).
- **Single-trace files** — σ is the intra-profile residual noise estimated via a rolling-window smoother.

A **higher score** means larger geophysical contrast relative to noise — i.e. a more *representative* and reliable frequency for that incision.

---

#### Supported instruments & formats
| Format | Notes |
|---|---|
| GEM-2 `.csv` | Integer-rounded EC values (reduced precision) |
| GEM-2 `.xlsx` | Full instrument precision — **recommended** |
| Legacy multi-sheet `.xlsx` | One sheet per frequency, col 0 = distance |

> ⚠️ GEM `.csv` exports round EC values to integers, which causes scores to differ slightly from the `.xlsx` equivalent. Always prefer `.xlsx` for quantitative comparison.

---

#### References
- Won, I.J., Keiswetter, D.A., Fields, G.R.A. & Sutton, L.C. (1996). GEM-2: A new multifrequency electromagnetic sensor. *Journal of Environmental and Engineering Geophysics*, **1**(2), 129–137. https://doi.org/10.4133/JEEG1.2.129
- McNeill, J.D. (1980). *Electromagnetic terrain conductivity measurement at low induction numbers*. Technical Note TN-6, Geonics Limited, Mississauga, Canada.
- Delefortrie, S., Saey, T., Van De Vijver, E., De Smedt, P., Missiaen, T., Demerre, I. & Van Meirvenne, M. (2014). Frequency domain electromagnetic induction survey in the intertidal zone: data acquisition and correction procedures. *Journal of Applied Geophysics*, **100**, 119–130. https://doi.org/10.4133/JEEG1.2.129
- Callegary, J.B., Ferré, T.P.A. & Groom, R.W. (2007). Vertical spatial sensitivity and exploration depth of low-induction-number electromagnetic induction instruments. *Vadose Zone Journal*, **6**(1), 158–167. https://doi.org/10.2136/vzj2006.0120
- Reynolds, J.M. (2011). *An Introduction to Applied and Environmental Geophysics* (2nd ed.). Wiley-Blackwell.
                """
            )

    # ── File uploader ───────────────────────────────────────────────────────
    uploaded_files = st.file_uploader(
        "Upload one or more data files (.xlsx or .csv)",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("Upload at least one `.xlsx` or `.csv` file to get started.")
        return

    # ── Process each file ───────────────────────────────────────────────────
    for uploaded_file in uploaded_files:
        st.divider()
        st.subheader(f"📄 {uploaded_file.name}")

        file_bytes = uploaded_file.getvalue()
        file_name = uploaded_file.name

        # Probe for GEM format (read only 5 rows for speed)
        is_gem = False
        try:
            if file_name.lower().endswith(".csv"):
                probe = pd.read_csv(io.BytesIO(file_bytes), nrows=5)
            else:
                probe = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, nrows=5)
            is_gem = is_gem_format(probe)
        except Exception:
            pass

        if is_gem:
            render_gem_results(file_bytes, file_name, distance_step, interp_kind)
        else:
            render_legacy_results(file_bytes, file_name, mode, distance_step, interp_kind)


if __name__ == "__main__":
    main()
