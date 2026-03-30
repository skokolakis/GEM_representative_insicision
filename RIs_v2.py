"""
UltraRank Frequency — Streamlit UI
Refactored from RIs_v1.py

Run with:
    streamlit run RIs_v2.py
"""
from __future__ import annotations

import io
import logging
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
    score = amplitude / max(mean_std, SCORE_EPSILON)

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
    Process all sheets in an uploaded Excel file.

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
    """Per-sheet plot: individual traces + mean ± 1σ envelope."""
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

    # ±1σ envelope
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
            st.markdown("**Sheets**")
            selected = st.multiselect(
                "Visible sheets",
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

        mode = st.radio("Measurement mode", list(MODES.keys()), horizontal=True)

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

    # ── File uploader ───────────────────────────────────────────────────────
    uploaded_files = st.file_uploader(
        "Upload one or more `.xlsx` files",
        type=["xlsx"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("Upload at least one `.xlsx` file to get started.")
        return

    # ── Process each file ───────────────────────────────────────────────────
    for uploaded_file in uploaded_files:
        st.divider()
        st.subheader(f"📄 {uploaded_file.name}")
        stem = Path(uploaded_file.name).stem

        with st.spinner("Processing…"):
            output_data, scores, warnings = process_file(
                uploaded_file.getvalue(),
                mode,
                distance_step=distance_step,
                interp_kind=interp_kind,
            )

        # Show any warnings
        if warnings:
            with st.expander(f"⚠️ {len(warnings)} warning(s)", expanded=False):
                for w in warnings:
                    st.warning(w)

        if not output_data:
            st.error("No usable sheets found in this file.")
            if distance_step > 1.0:
                st.info("Try reducing the distance step in the sidebar.")
            continue

        # ── Ranking table ───────────────────────────────────────────────────
        ranking = sorted(scores.items(), key=lambda x: x[1]["score"], reverse=True)

        if not ranking:
            st.error("No sheets could be scored.")
            continue

        rank_df = pd.DataFrame(
            [
                {
                    "Rank": i,
                    "Frequency / Sheet": name,
                    "Score": round(m["score"], 2),
                    "Amplitude": round(m["amplitude"], 6),
                    "Mean Std": round(m["mean_std"], 6),
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
                help="Score = amplitude / mean_std. Higher means a large signal range relative to noise.",
            )
            st.metric("Amplitude", f"{best_metrics['amplitude']:.4g}")
            st.metric("Mean Std", f"{best_metrics['mean_std']:.4g}")

        # ── Graph editor ─────────────────────────────────────────────────────
        opts = render_graph_editor(output_data, file_key=stem)

        # ── Overview plot ────────────────────────────────────────────────────
        st.markdown("### 📈 All representative profiles")
        overview_fig = make_overview_figure(output_data, scores, mode, uploaded_file.name, opts)
        st.pyplot(overview_fig, use_container_width=True)
        plt.close(overview_fig)

        # ── Per-sheet detail ─────────────────────────────────────────────────
        with st.expander("🔍 Per-sheet detail plots", expanded=False):
            visible_sheets = opts.selected_sheets if opts.selected_sheets else list(output_data.keys())
            for sheet_name in visible_sheets:
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
            )

        with dl2:
            st.download_button(
                label="📥 Scores (.csv)",
                data=build_scores_csv(scores),
                file_name=f"{stem}_{mode}_scores.csv",
                mime="text/csv",
            )

        with dl3:
            # Always regenerate a fresh figure for download (overview_fig is already closed above)
            download_fig = make_overview_figure(output_data, scores, mode, uploaded_file.name, opts)
            png_bytes = fig_to_png(download_fig)
            plt.close(download_fig)
            st.download_button(
                label="📥 Overview plot (.png)",
                data=png_bytes,
                file_name=f"{stem}_{mode}_overview.png",
                mime="image/png",
            )


if __name__ == "__main__":
    main()
