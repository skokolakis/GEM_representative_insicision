"""
UltraRank Frequency — Streamlit UI
Refactored from RIs_v1.py

Run with:
    streamlit run RIs_v2_streamlit.py
"""

import io
import logging
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st
from scipy.interpolate import interp1d

# ---------------------------------------------------------------------------
# Configuration (all in one place, easily overridden via Streamlit widgets)
# ---------------------------------------------------------------------------
DEFAULT_DISTANCE_STEP = 0.5
DEFAULT_INTERP_KIND = "linear"
SCORE_EPSILON = 1e-8

MODES = {
    "EC": "Mean EC Response (S/m)",
    "MS": "Mean MS Response (10⁻⁵ SI)",
}

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Core processing (pure functions — no side effects, no sys.exit)
# ---------------------------------------------------------------------------

def process_sheet(
    df: pd.DataFrame,
    distance_step: float = DEFAULT_DISTANCE_STEP,
    interp_kind: str = DEFAULT_INTERP_KIND,
) -> tuple[pd.DataFrame | None, dict | None, str | None]:
    """
    Process a single sheet.

    Returns
    -------
    interpolated_df : DataFrame indexed by common distance, columns = original trace columns
    score_dict      : {"mean_std": float, "amplitude": float, "score": float}
    error           : human-readable reason for failure, or None on success
    """
    if df.shape[1] < 2:
        return None, None, "fewer than 2 columns"

    distance = pd.to_numeric(df.iloc[:, 0], errors="coerce").values
    line_data = df.iloc[:, 1:].apply(pd.to_numeric, errors="coerce")

    if np.all(np.isnan(distance)):
        return None, None, "distance column is all NaN"

    min_dist = float(np.nanmin(distance))
    max_dist = float(np.nanmax(distance))

    if not (np.isfinite(min_dist) and np.isfinite(max_dist)):
        return None, None, "non-finite distance range"

    span = max_dist - min_dist
    if span < distance_step:
        return None, None, f"distance span ({span:.2f} m) < step ({distance_step} m)"

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
        )

        if df_xy.shape[0] < 2:
            continue

        try:
            f = interp1d(
                df_xy["d"].values,
                df_xy["y"].values,
                kind=interp_kind,
                bounds_error=False,
                fill_value=np.nan,
                assume_sorted=True,
            )
            interpolated_lines[col] = f(common_dist)
        except Exception as exc:
            log.warning("  Interpolation failed for column %s: %s", col, exc)
            continue

    if not interpolated_lines:
        return None, None, "no columns survived interpolation"

    interpolated_df = pd.DataFrame(interpolated_lines, index=common_dist)
    interpolated_df.index.name = "Distance (m)"

    rep_prof = interpolated_df.mean(axis=1, skipna=True)
    std_prof = interpolated_df.std(axis=1, skipna=True)

    mean_std = float(std_prof.mean())
    amplitude = float(np.nanmax(rep_prof) - np.nanmin(rep_prof))
    score = amplitude / max(mean_std, SCORE_EPSILON)

    score_dict = {"mean_std": mean_std, "amplitude": amplitude, "score": score}
    return interpolated_df, score_dict, None


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
    output_data            : {sheet_name: interpolated_df}
    representativeness_scores : {sheet_name: score_dict}
    warnings               : list of warning strings
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
        try:
            df = pd.read_excel(excel, sheet_name=sheet_name)
        except Exception as exc:
            warnings.append(f"Sheet '{sheet_name}': read error — {exc}")
            continue

        interp_df, score_dict, error = process_sheet(df, distance_step, interp_kind)

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
) -> plt.Figure:
    """All representative profiles on one axes."""
    fig, ax = plt.subplots(figsize=(10, 5))
    y_label = MODES[mode]

    for sheet_name, interp_df in output_data.items():
        rep_prof = interp_df.mean(axis=1, skipna=True)
        common_dist = interp_df.index.values
        sc = scores[sheet_name]["score"]
        if not np.all(np.isnan(rep_prof.values)):
            ax.plot(common_dist, rep_prof.values, label=f"{sheet_name} (score={sc:.2f})")

    ax.set_xlabel("Distance (m)")
    ax.set_ylabel(y_label)
    ax.set_title(f"Representative profiles [{mode}] — {file_name}")
    ax.legend(fontsize=8)
    ax.grid(True, alpha=0.4)
    fig.tight_layout()
    return fig


def make_sheet_figure(
    sheet_name: str,
    interp_df: pd.DataFrame,
    mode: str,
) -> plt.Figure:
    """Per-sheet plot: individual traces + mean ± 1σ envelope."""
    fig, ax = plt.subplots(figsize=(10, 4))
    common_dist = interp_df.index.values
    rep_prof = interp_df.mean(axis=1, skipna=True).values
    std_prof = interp_df.std(axis=1, skipna=True).values

    # Individual traces (thin, semi-transparent)
    for col in interp_df.columns:
        ax.plot(common_dist, interp_df[col].values, color="steelblue", alpha=0.25, linewidth=0.8)

    # Mean profile
    ax.plot(common_dist, rep_prof, color="navy", linewidth=2, label="Mean")

    # ±1σ envelope
    ax.fill_between(
        common_dist,
        rep_prof - std_prof,
        rep_prof + std_prof,
        alpha=0.2,
        color="navy",
        label="±1σ",
    )

    ax.set_xlabel("Distance (m)")
    ax.set_ylabel(MODES[mode])
    ax.set_title(f"{sheet_name} — individual traces & representative profile")
    ax.legend()
    ax.grid(True, alpha=0.4)
    fig.tight_layout()
    return fig


# ---------------------------------------------------------------------------
# Export helpers
# ---------------------------------------------------------------------------

def build_excel_download(output_data: dict[str, pd.DataFrame], mode: str) -> bytes:
    """Pack all interpolated sheets into a single Excel workbook."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, interp_df in output_data.items():
            rep_prof = interp_df.mean(axis=1, skipna=True)
            out_df = pd.DataFrame(
                {
                    "Distance (m)": interp_df.index.values,
                    f"{sheet_name[:28]}_mean": rep_prof.values,  # cap at 28 chars (Excel limit 31)
                }
            )
            out_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
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

        with st.spinner("Processing…"):
            output_data, scores, warnings = process_file(
                uploaded_file.read(),
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
            continue

        # ── Ranking table ───────────────────────────────────────────────────
        ranking = sorted(scores.items(), key=lambda x: x[1]["score"], reverse=True)
        rank_df = pd.DataFrame(
            [
                {
                    "Rank": i,
                    "Frequency / Sheet": name,
                    "Score": f"{m['score']:.2f}",
                    "Amplitude": f"{m['amplitude']:.4g}",
                    "Mean Std": f"{m['mean_std']:.4g}",
                }
                for i, (name, m) in enumerate(ranking, 1)
            ]
        )

        col_table, col_best = st.columns([3, 1])
        with col_table:
            st.markdown("### 🏆 Frequency Ranking")
            st.dataframe(rank_df, use_container_width=True, hide_index=True)
        with col_best:
            best_name, best_metrics = ranking[0]
            st.metric("Best frequency", best_name, f"score {best_metrics['score']:.2f}")
            st.metric("Amplitude", f"{best_metrics['amplitude']:.4g}")
            st.metric("Mean Std", f"{best_metrics['mean_std']:.4g}")

        # ── Overview plot ────────────────────────────────────────────────────
        st.markdown("### 📈 All representative profiles")
        overview_fig = make_overview_figure(output_data, scores, mode, uploaded_file.name)
        st.pyplot(overview_fig, use_container_width=True)
        plt.close(overview_fig)

        # ── Per-sheet detail ─────────────────────────────────────────────────
        with st.expander("🔍 Per-sheet detail plots", expanded=False):
            for sheet_name, interp_df in output_data.items():
                sc = scores[sheet_name]
                st.markdown(
                    f"**{sheet_name}** — score `{sc['score']:.2f}` | "
                    f"amp `{sc['amplitude']:.4g}` | std `{sc['mean_std']:.4g}`"
                )
                sheet_fig = make_sheet_figure(sheet_name, interp_df, mode)
                st.pyplot(sheet_fig, use_container_width=True)
                plt.close(sheet_fig)

        # ── Downloads ────────────────────────────────────────────────────────
        st.markdown("### 💾 Downloads")
        stem = Path(uploaded_file.name).stem
        dl1, dl2, dl3 = st.columns(3)

        with dl1:
            st.download_button(
                label="📥 Interpolated profiles (.xlsx)",
                data=build_excel_download(output_data, mode),
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
            png_bytes = fig_to_png(overview_fig if not overview_fig.axes else make_overview_figure(
                output_data, scores, mode, uploaded_file.name
            ))
            st.download_button(
                label="📥 Overview plot (.png)",
                data=png_bytes,
                file_name=f"{stem}_{mode}_overview.png",
                mime="image/png",
            )


if __name__ == "__main__":
    main()
