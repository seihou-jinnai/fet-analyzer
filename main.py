# main.py
# ======================================================================================
# FET Mobility Analyzer (PySide6 GUI)
# ======================================================================================
# This script provides a GUI to:
#   - Load one or more Keithley output Excel files (.xls/.xlsx/.xlsm)
#   - Let the user select a sheet and pick the two key columns:
#       * V_G  (gate voltage)
#       * I_SD (source-drain current)
#   - Configure device parameters (W, L, C) and a fitting window
#   - Run an analysis based on the saturation-regime equation
#   - Export:
#       * One PNG plot per item in the Execution list
#       * One CSV summary containing the inputs + fitted outputs
#         (mobility, Vth, fit range, R2, status)
#
# IMPORTANT UPDATE IN THIS VERSION:
#   The fitting logic has been improved so that only physically meaningful transfer
#   trends are accepted:
#
#   - p-type:
#       As V_G becomes smaller (more negative), |I_D| must become larger.
#       Equivalently, if data are sorted by increasing V_G, sqrt(|I_D|) must decrease.
#
#   - n-type:
#       As V_G becomes larger (more positive), |I_D| must become larger.
#       Equivalently, if data are sorted by increasing V_G, sqrt(|I_D|) must increase.
#
#   This prevents "successful" fitting on data that do not actually show a valid FET
#   transfer characteristic in the expected direction.
# ======================================================================================

from __future__ import annotations  # Enable forward references in type hints (Python 3.7+)

# --------------------------------------------------------------------------------------
# Standard library imports
# --------------------------------------------------------------------------------------
import os          # File system checks (e.g., os.path.isfile)
import re          # Regular expressions (numeric parsing, filename sanitizing, range parsing)
import sys         # argv + Qt application event loop
import math        # sqrt, isfinite, etc.
import csv         # Write the result summary CSV file
from dataclasses import dataclass  # Lightweight structured record types
from datetime import datetime      # Timestamp for result CSV filename (fet-results_YYYYMMDD.csv)
from pathlib import Path           # Safer filesystem path operations
from typing import Dict, List, Optional, Any, Literal, Tuple  # Type hints for clarity & correctness

# --------------------------------------------------------------------------------------
# Third-party imports
# --------------------------------------------------------------------------------------
import pandas as pd  # Reading Excel sheets into DataFrames, robust parsing, numeric conversion

# ---- Matplotlib (non-interactive) ----
# This GUI app should not pop up matplotlib windows. We set a non-interactive backend
# ("Agg") so all plotting is done off-screen and saved directly as PNG files.
import matplotlib
matplotlib.use("Agg")  # important: no GUI popups
import matplotlib.pyplot as plt  # Standard plotting interface

# ---- PySide6 (Qt) imports ----
# These provide:
#   - Core Qt constants/types (Qt, QModelIndex)
#   - Model/view base class for tables (QAbstractTableModel)
#   - GUI widgets and layouts (QMainWindow, QTableView, etc.)
from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex
from PySide6.QtGui import QAction, QColor, QPalette, QBrush
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenuBar,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QSizePolicy,
    QTableView,
    QVBoxLayout,
    QWidget,
    QTextEdit,
)

# ======================================================================================
# UI Theme
# ======================================================================================
# We apply a bright, "natural white" Fusion theme to improve readability and give a
# consistent look across Windows/macOS/Linux.
# ======================================================================================

# -----------------------------
# Light theme (natural white)
# -----------------------------
def apply_light_fusion_theme(app: QApplication) -> None:
    # Force Qt to use the Fusion style for consistent, modern widgets.
    app.setStyle("Fusion")

    # QPalette controls application-wide colors.
    pal = QPalette()

    # Window background and widget base colors
    pal.setColor(QPalette.Window, QColor("#FAFAFA"))         # window background
    pal.setColor(QPalette.Base, QColor("#FFFFFF"))           # input fields background
    pal.setColor(QPalette.AlternateBase, QColor("#F3F3F3"))  # alternating rows in tables

    # Primary text colors
    pal.setColor(QPalette.WindowText, QColor("#111111"))
    pal.setColor(QPalette.Text, QColor("#111111"))
    pal.setColor(QPalette.ButtonText, QColor("#111111"))
    pal.setColor(QPalette.PlaceholderText, QColor("#888888"))

    # Button and selection highlight colors
    pal.setColor(QPalette.Button, QColor("#F0F0F0"))
    pal.setColor(QPalette.Highlight, QColor("#2F6FED"))      # selection highlight
    pal.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))

    # Disabled-state colors (gray)
    pal.setColor(QPalette.Disabled, QPalette.Text, QColor("#9A9A9A"))
    pal.setColor(QPalette.Disabled, QPalette.ButtonText, QColor("#9A9A9A"))

    # Apply palette to the whole app.
    app.setPalette(pal)


# ======================================================================================
# FET analysis helpers (ported from fet-anal.py)
# ======================================================================================
# The block below is the core analysis logic. It:
#   1) Parses numeric inputs from the GUI (W, L, C, fit window)
#   2) Splits forward/return sweeps (if the data includes a sweep reversal)
#   3) Performs a linear fit on sqrt(|I_D|) vs V_G in a chosen window
#   4) Calculates:
#       - Vth (x-intercept of the fit)
#       - mobility (from slope using saturation-regime equation)
#       - R^2 of the fit
#   5) Generates two-panel plots and saves them as PNG
#
# IMPORTANT DESIGN UPDATE:
#   In this version, fit candidates are accepted only if they show the expected
#   transfer direction for the selected device type:
#
#   p-type:
#       smaller / more negative V_G  -> larger |I_D|
#       therefore, after sorting by increasing V_G:
#           sqrt(|I_D|) should decrease
#           fitted slope should be negative
#
#   n-type:
#       larger / more positive V_G -> larger |I_D|
#       therefore, after sorting by increasing V_G:
#           sqrt(|I_D|) should increase
#           fitted slope should be positive
#
#   This prevents false fits on data that do not actually exhibit the expected
#   field-effect trend.
# ======================================================================================

FitMode = Literal["span", "range"]  # Fit window selection mode


@dataclass(frozen=True)
class FitWindowSpec:
    # Container describing how to choose the fitting region.
    # mode="span": search for best contiguous region spanning span_v volts
    # mode="range": use fixed vmin..vmax
    mode: FitMode
    span_v: float | None = None
    vmin: float | None = None
    vmax: float | None = None


# Regular expression that accepts typical float and scientific notation, with optional sign.
# Examples accepted:
#   1000
#   -10
#   1.15E-08
#   +3.2e+5
_NUM_RE = re.compile(r"^[+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$")


def _parse_float_token(tok: str) -> float:
    """
    Parse a numeric string token into a float with sanity checks.

    - Strips whitespace
    - Removes commas (e.g., "1,000" -> "1000")
    - Validates against a numeric regex
    - Ensures the value is finite (not NaN/inf)

    Raises:
        ValueError if parsing fails.
    """
    s = (tok or "").strip()
    if s == "":
        raise ValueError("Empty numeric value.")

    # Permit comma separators in user input.
    s = s.replace(",", "")

    # If the regex does not match, try float() anyway to give a robust fallback.
    if not _NUM_RE.match(s):
        try:
            v = float(s)
        except ValueError as e:
            raise ValueError(f"Invalid number: '{tok}'") from e

    v = float(s)

    # Finite check:
    # - (v == v) is False only for NaN
    # - inf/-inf explicitly excluded
    if not (v == v) or v in (float("inf"), float("-inf")):
        raise ValueError(f"Invalid finite number: '{tok}'")

    return v


# Regex for parsing a two-number range "a-b" (allowing scientific notation and signs)
# Examples accepted:
#   "20-30"
#   "-5-15"
#   "-10--2"
_RANGE2_RE = re.compile(
    r"^\s*([+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)\s*-\s*([+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)\s*$"
)


def parse_fit_window_gui(text: str) -> FitWindowSpec:
    """
    Parse the "fit window" GUI input into a structured FitWindowSpec.

    Supported input forms:
      - "10" or "-10" => span mode (span_v = abs(value))
      - "20-30" or "-5-15" or "-10--2" => range mode ([vmin, vmax])

    Returns:
      FitWindowSpec describing the requested window type.

    Raises:
      ValueError if empty or invalid.
    """
    # Remove spaces to make inputs like "20 - 30" behave nicely.
    s = (text or "").strip().replace(" ", "")
    if not s:
        raise ValueError("Fit window is empty.")

    # If it matches the range pattern, interpret as fixed range.
    m = _RANGE2_RE.match(s)
    if m:
        a = _parse_float_token(m.group(1))
        b = _parse_float_token(m.group(2))

        # Normalize ordering
        vmin, vmax = (a, b) if a <= b else (b, a)

        # Reject zero width (meaningless fit region)
        if vmin == vmax:
            raise ValueError("Fit window range must have non-zero width (e.g., 20-30).")

        return FitWindowSpec(mode="range", vmin=vmin, vmax=vmax)

    # Otherwise interpret as a span (voltage width) for the best-window search.
    span = abs(_parse_float_token(s))
    if span <= 0:
        raise ValueError("Fit window span must be positive (e.g., 10).")
    return FitWindowSpec(mode="span", span_v=span)


def _split_forward_return(vg: List[float], y: List[float]) -> Tuple[List[float], List[float], List[float], List[float], int]:
    """
    Attempt to split a sweep into forward and return segments.

    Strategy:
      - Identify initial direction by looking for first non-zero delta(VG)
      - Find the first index where the direction flips -> "turn"
      - Return:
          vg_forward, y_forward, vg_return, y_return, turn_index

    Notes:
      - If the direction never flips, return segment is empty.
      - If VG is constant or too short, treat as no-split.
    """
    n = len(vg)
    if n < 3:
        # Not enough points to reliably detect a turn.
        return vg[:], y[:], [], [], n

    # Determine initial direction: +1 for increasing VG, -1 for decreasing.
    direction = 0
    for i in range(1, n):
        dv = vg[i] - vg[i - 1]
        if dv > 0:
            direction = 1
            break
        if dv < 0:
            direction = -1
            break

    # If all deltas are zero, the sweep direction is undefined.
    if direction == 0:
        return vg[:], y[:], [], [], n

    # Find first point where direction changes sign.
    turn = n
    for i in range(2, n):
        dv = vg[i] - vg[i - 1]
        if dv == 0:
            # Ignore flat segments.
            continue
        if (dv > 0 and direction < 0) or (dv < 0 and direction > 0):
            turn = i
            break

    # Forward sweep: start -> turn-1
    vg_f = vg[:turn]
    y_f = y[:turn]

    # Return sweep: include the turning point for continuity (turn-1 -> end)
    vg_r = vg[turn - 1:] if turn - 1 >= 0 else vg[:]
    y_r = y[turn - 1:] if turn - 1 >= 0 else y[:]

    return vg_f, y_f, vg_r, y_r, turn


def _linfit_r2(x: List[float], y: List[float]) -> Tuple[float, float, float]:
    """
    Ordinary least squares linear regression y = a*x + b with R^2.

    Returns:
      a (slope), b (intercept), r2

    Raises:
      ValueError if insufficient points or degenerate x.
    """
    n = len(x)
    if n < 2:
        raise ValueError("Need at least 2 points for linear fit.")

    # Mean values
    x_mean = sum(x) / n
    y_mean = sum(y) / n

    # Variance of x
    sxx = sum((xi - x_mean) ** 2 for xi in x)
    if sxx == 0:
        # All x are identical -> slope undefined
        raise ValueError("All X values are identical; cannot fit.")

    # Covariance of x and y
    sxy = sum((x[i] - x_mean) * (y[i] - y_mean) for i in range(n))

    # Fit coefficients
    a = sxy / sxx
    b = y_mean - a * x_mean

    # Compute predictions for R^2
    y_hat = [a * xi + b for xi in x]
    ss_res = sum((y[i] - y_hat[i]) ** 2 for i in range(n))
    ss_tot = sum((yi - y_mean) ** 2 for yi in y)

    # R^2 definition with zero-variance guard
    if ss_tot == 0:
        r2 = 1.0 if ss_res == 0 else 0.0
    else:
        r2 = 1.0 - ss_res / ss_tot

    return a, b, r2


def _is_valid_transfer_trend(vg: List[float], ysqrt: List[float], dev_type: Literal["p", "n"]) -> bool:
    """
    Check whether the selected region shows the physically expected transfer trend.

    Why this is needed:
      A purely mathematical line fit can still succeed even when the dataset does not
      show a real field-effect transfer characteristic in the expected direction.
      For example:
        - a p-type device whose |ID| becomes smaller as VG becomes more negative
        - an n-type device whose |ID| becomes smaller as VG becomes more positive

      Those cases should be rejected, even if R^2 happens to be high.

    Expected trend:
      - p-type:
            as VG becomes smaller (more negative), |ID| should increase
            -> if we sort VG in ascending order, sqrt(|ID|) should decrease
            -> slope should be negative
      - n-type:
            as VG becomes larger (more positive), |ID| should increase
            -> if we sort VG in ascending order, sqrt(|ID|) should increase
            -> slope should be positive

    The check is intentionally simple and robust:
      1) Sort points by VG so the interpretation is independent of sweep direction.
      2) Perform a linear fit on the sorted data.
      3) Require both:
         - expected slope sign
         - expected endpoint direction

    Returns:
      True  -> acceptable transfer trend
      False -> invalid / opposite / indeterminate trend
    """
    if len(vg) < 2 or len(ysqrt) < 2:
        return False

    # Sort by VG so the trend rule is evaluated in a consistent voltage direction.
    xy = sorted(zip(vg, ysqrt), key=lambda t: t[0])
    xs = [t[0] for t in xy]
    ys = [t[1] for t in xy]

    try:
        a, b, r2 = _linfit_r2(xs, ys)
    except ValueError:
        return False

    # Net endpoint trend:
    #   positive -> overall increase with increasing VG
    #   negative -> overall decrease with increasing VG
    dy = ys[-1] - ys[0]

    if dev_type == "p":
        # For p-type, increasing VG (less negative) should reduce |ID|:
        # therefore y should decrease with increasing VG.
        return (a < 0) and (dy < 0)
    else:
        # For n-type, increasing VG should increase |ID|:
        # therefore y should increase with increasing VG.
        return (a > 0) and (dy > 0)


def _best_window_fit_span(
    vg: List[float],
    ysqrt: List[float],
    *,
    span_v: float,
    dev_type: Literal["p", "n"],
) -> Dict[str, Any]:
    """
    Search for the best contiguous window that spans at least span_v volts in VG,
    and select the one that maximizes R^2 of the linear fit.

    IMPORTANT:
      Not every mathematically linear region is physically acceptable.
      Therefore, each candidate window is first tested by _is_valid_transfer_trend().
      Only windows with the expected p-type / n-type transfer direction are allowed.

    Returns dict with:
      start, end, a, b, r2, xw, yw

    Raises:
      ValueError if no valid window meets criteria.
    """
    n = len(vg)
    if n < 3:
        raise ValueError("Not enough points for window fit.")

    best: Dict[str, Any] | None = None

    # Consider each possible window start index
    for start in range(0, n - 2):
        vmin = vg[start]
        vmax = vg[start]
        end = None

        # Expand end index until voltage span requirement is met
        for j in range(start + 1, n):
            vmin = min(vmin, vg[j])
            vmax = max(vmax, vg[j])
            if (vmax - vmin) >= span_v:
                end = j
                break

        # If we can't achieve the span anymore, stop searching.
        if end is None:
            break

        # Slice the candidate window
        xw = vg[start:end + 1]
        yw = ysqrt[start:end + 1]

        # Reject windows that do not show the expected transfer direction.
        # This is the key improvement preventing false fits on "wrong-way" data.
        if not _is_valid_transfer_trend(xw, yw, dev_type):
            continue

        # Fit and compute R^2
        try:
            a, b, r2 = _linfit_r2(xw, yw)
        except ValueError:
            # Skip windows that produce invalid fits (degenerate x, etc.)
            continue

        cand = {"start": start, "end": end, "a": a, "b": b, "r2": r2, "xw": xw, "yw": yw}

        # Choose best by R^2
        if best is None or cand["r2"] > best["r2"]:
            best = cand

    if best is None:
        raise ValueError(
            f"Failed to find a valid >= {span_v:g} V-span window with the expected "
            f"{'p-type' if dev_type == 'p' else 'n-type'} transfer trend."
        )

    return best


def _fit_fixed_range(
    vg: List[float],
    ysqrt: List[float],
    *,
    vmin: float,
    vmax: float,
    dev_type: Literal["p", "n"],
) -> Dict[str, Any]:
    """
    Fit using all points where vmin <= VG <= vmax.

    IMPORTANT:
      The specified range is not accepted automatically.
      Even if the points are numerous enough, the range is rejected if it does not
      show the expected transfer direction for the selected device type.

    Returns dict with:
      a, b, r2, xw, yw, vmin, vmax

    Raises:
      ValueError if fewer than 2 points exist in the range, or if the trend is invalid.
    """
    # Normalize ordering
    if vmin > vmax:
        vmin, vmax = vmax, vmin

    xw: List[float] = []
    yw: List[float] = []

    # Filter points in range
    for x, y in zip(vg, ysqrt):
        if vmin <= x <= vmax:
            xw.append(x)
            yw.append(y)

    # At least two points are needed for linear regression
    if len(xw) < 2:
        raise ValueError(f"Not enough points in fit range [{vmin:g}, {vmax:g}] V (need >= 2).")

    # Reject the range if it does not show the expected transfer trend.
    if not _is_valid_transfer_trend(xw, yw, dev_type):
        raise ValueError(
            f"The selected fit range [{vmin:g}, {vmax:g}] V does not show the expected "
            f"{'p-type' if dev_type == 'p' else 'n-type'} transfer trend."
        )

    a, b, r2 = _linfit_r2(xw, yw)
    return {"a": a, "b": b, "r2": r2, "xw": xw, "yw": yw, "vmin": vmin, "vmax": vmax}


def _format_wlc_note(w_um: float, l_um: float, c_fcm2: float) -> str:
    """
    Create a LaTeX-like string describing the device geometry and gate dielectric capacitance.
    This text is shown under the left plot.
    """
    return rf"$W$ = {w_um:g} um   $L$ = {l_um:g} um   $C$ = {c_fcm2:g} F/cm$^2$"


def analyze_fet_and_save_figure(
    vg: List[float],
    isd: List[float],
    *,
    w_um: float,
    l_um: float,
    c_fcm2: float,
    dev_type: Literal["p", "n"],
    fit_spec: FitWindowSpec,
    title: str,
    comment: str,  # kept for CSV, not plotted
    out_png: Path,
) -> Dict[str, Any]:
    """
    Analyze one VG/ID dataset, save a 2-panel plot as PNG, and return fitted results.

    Outputs returned (as a dict):
      - mobility: mobility in cm^2/(V*s) (displayed as cm/Vs in the plot text)
      - vth: threshold voltage (x-intercept where sqrt(|ID|)=0)
      - r2: coefficient of determination for the linear fit
      - fit_vmin/fit_vmax: actual fitted window limits used
      - slope/intercept: linear fit parameters (a, b) for y = a*VG + b

    Plot content:
      Left: log-scale ID vs VG (forward/return separated by color)
      Right: sqrt(|ID|) vs VG with fitted dotted line

    Note about Comment:
      - Comment is intentionally NOT plotted,
        but it is still carried along for CSV export in the calling code.
    """
    # Basic input validation: arrays must be same length and non-empty.
    if len(vg) != len(isd) or len(vg) == 0:
        raise ValueError("Invalid input arrays.")

    # For p-type devices, conventional plotting uses -ID so the current is positive.
    # For n-type devices, ID is already positive in accumulation; we keep sign.
    sign = -1.0 if dev_type == "p" else 1.0
    id_plot = [sign * i for i in isd]

    # Split the ID trace into forward and return sweeps for plotting
    vg_f, id_f, vg_r, id_r, _ = _split_forward_return(vg, id_plot)

    # Prepare sqrt(|ID|) for fitting.
    # Fitting is done on the square root of the absolute current magnitude.
    ysqrt_all = [math.sqrt(abs(v)) for v in id_plot]
    vg_f2, y_f2, vg_r2, y_r2, _ = _split_forward_return(vg, ysqrt_all)

    # -------------------------
    # Determine the fit window
    # -------------------------
    if fit_spec.mode == "span":
        # Search the forward sweep for the best window (by R^2),
        # but only among windows showing the expected p/n transfer direction.
        if fit_spec.span_v is None:
            raise ValueError("Internal error: span_v is None.")
        fit = _best_window_fit_span(vg_f2, y_f2, span_v=fit_spec.span_v, dev_type=dev_type)
        x_used = fit["xw"]
        a = float(fit["a"])
        b = float(fit["b"])
        r2 = float(fit["r2"])
        fit_vmin = float(min(x_used))
        fit_vmax = float(max(x_used))
    else:
        # Use a fixed range, but reject it if the trend direction is physically wrong.
        if fit_spec.vmin is None or fit_spec.vmax is None:
            raise ValueError("Internal error: vmin/vmax is None.")
        fit = _fit_fixed_range(vg_f2, y_f2, vmin=fit_spec.vmin, vmax=fit_spec.vmax, dev_type=dev_type)
        x_used = fit["xw"]
        a = float(fit["a"])
        b = float(fit["b"])
        r2 = float(fit["r2"])
        fit_vmin = float(fit["vmin"])
        fit_vmax = float(fit["vmax"])

    # Threshold voltage is the x-intercept of the fitted line: 0 = a*Vth + b -> Vth = -b/a
    vth = -(b) / a if a != 0 else float("nan")

    # -------------------------
    # Mobility calculation
    # -------------------------
    # Convert micrometers to centimeters:
    #   1 um = 1e-4 cm
    w_cm = w_um * 1e-4
    l_cm = l_um * 1e-4

    # Saturation-regime formula in sqrt form:
    #   sqrt(ID) = sqrt( (W/(2L)) * mu * C * (VG - Vth)^2 ) = a*VG + b
    # From the slope a, derive mobility:
    #   mu = 2L/(W*C) * a^2
    mobility = 2.0 * l_cm / w_cm / c_fcm2 * a * a

    # ---- Fit line domain: from fitted location to x-intercept ----
    # The dotted line is drawn only in a meaningful range:
    #   from the fitted region edge to where the line crosses sqrt(ID)=0.
    x_left = float(min(x_used)) if x_used else (float(min(vg_f2)) if vg_f2 else float(min(vg)))
    x_right = (-b / a) if a != 0 else float("nan")

    # Create x/y arrays for the dotted fit line
    x_fit: List[float] = []
    y_fit: List[float] = []
    if a != 0 and (x_right == x_right) and math.isfinite(x_right):
        # Ensure x1 <= x2 for line generation
        x1, x2 = x_left, float(x_right)
        if x2 < x1:
            x1, x2 = x2, x1

        # Uniform sampling points for a smooth dotted line
        steps = 100
        dx = (x2 - x1) / (steps - 1) if steps > 1 else 0.0
        x_fit = [x1 + i * dx for i in range(steps)]
        y_fit = [a * x + b for x in x_fit]

    # -------------------------
    # Styling by device type
    # -------------------------
    # Requested colors:
    #   p-type: forward #17489C, return #8BA4CE
    #   n-type: forward #D61900, return #EB8C80
    if dev_type == "p":
        col_fwd = "#17489C"
        col_ret = "#8BA4CE"
    else:
        col_fwd = "#D61900"
        col_ret = "#EB8C80"

    # Plot line width (controls both forward/return in both panels)
    lw = 2.5

    # Create a figure with 2 subplots in a single row.
    # figsize tuned for clear export and comfortable read.
    fig, (axL, axR) = plt.subplots(1, 2, figsize=(10.8, 5.6))

    # Figure title (file and sheet)
    if title:
        fig.suptitle(title, fontsize=12)

    # Make both axes square-ish for consistent appearance
    axL.set_box_aspect(1)
    axR.set_box_aspect(1)

    # -------------------------
    # Left panel: log(ID) vs VG
    # -------------------------
    def _log_y(vals: List[float]) -> List[float]:
        # Avoid log of non-positive values by mapping them to NaN.
        # Matplotlib won't plot NaNs in log scale.
        return [v if v > 0 else float("nan") for v in vals]

    # Draw return first so forward can appear "on top" if needed
    if vg_r and id_r:
        axL.plot(vg_r, _log_y(id_r), linestyle="-", linewidth=lw, color=col_ret)
    if vg_f and id_f:
        axL.plot(vg_f, _log_y(id_f), linestyle="-", linewidth=lw, color=col_fwd)

    axL.set_xlabel(r"$V_G$ (V)")
    axL.set_ylabel(r"$-I_D$ (A)" if dev_type == "p" else r"$I_D$ (A)")
    axL.set_yscale("log")
    axL.grid(True, which="both")  # show grid on major+minor ticks

    # -------------------------
    # Right panel: sqrt(ID) vs VG
    # -------------------------
    if vg_r2 and y_r2:
        axR.plot(vg_r2, y_r2, linestyle="-", linewidth=lw, color=col_ret)
    if vg_f2 and y_f2:
        axR.plot(vg_f2, y_f2, linestyle="-", linewidth=lw, color=col_fwd)

    # Plot the fitted line as a dotted black line if it was constructed successfully
    if x_fit and y_fit:
        axR.plot(x_fit, y_fit, linestyle=":", linewidth=2, color="black")

    axR.set_xlabel(r"$V_G$ (V)")
    axR.set_ylabel(r"$(-I_D)^{1/2}$ (A$^{1/2}$)" if dev_type == "p" else r"$(I_D)^{1/2}$ (A$^{1/2}$)")
    axR.grid(True)

    # ---- Bottom notes (Comment removed) ----
    # These annotations appear below each subplot, giving:
    #   left: W/L/C settings
    #   right: mobility and Vth
    left_note = _format_wlc_note(w_um, l_um, c_fcm2)
    right_note = rf"mobility = {mobility:.2E} cm/Vs   Vth = {vth:.4g} V"

    axL.text(0.5, -0.25, left_note, transform=axL.transAxes, ha="center", va="top")
    axR.text(0.5, -0.25, right_note, transform=axR.transAxes, ha="center", va="top")

    # Reduce bottom margin now that comment is gone
    fig.subplots_adjust(bottom=0.28, top=0.90, wspace=0.28)

    # Ensure output directory exists (creates intermediate folders if needed)
    out_png.parent.mkdir(parents=True, exist_ok=True)

    # Save to PNG with a fairly high DPI for readability in papers/presentations
    fig.savefig(out_png, dpi=220)

    # Close the figure to free memory (important when exporting many rows)
    plt.close(fig)

    # Return results for CSV and GUI summary
    return {
        "mobility": mobility,
        "vth": vth,
        "r2": r2,
        "fit_vmin": fit_vmin,
        "fit_vmax": fit_vmax,
        "slope": a,
        "intercept": b,
    }


def _sanitize_filename(s: str, max_len: int = 120) -> str:
    """
    Make a filesystem-safe filename component.

    - Replace problematic characters with "_"
    - Collapse multiple spaces
    - Truncate to max_len to avoid OS/path limitations
    """
    s2 = re.sub(r"[^\w\-. ]+", "_", s, flags=re.UNICODE).strip()
    s2 = re.sub(r"\s+", " ", s2)
    if not s2:
        s2 = "output"
    if len(s2) > max_len:
        s2 = s2[:max_len].rstrip()
    return s2


def _coerce_numeric_series(series: pd.Series) -> pd.Series:
    """
    Convert a pandas Series to numeric values safely.

    Handling:
      - If dtype is object (strings), strip and remove commas
      - Convert empty/'nan'/'None' to missing values
      - Use pandas.to_numeric(..., errors="coerce") to turn invalids into NaN
    """
    if series.dtype == object:
        series = series.astype(str).str.replace(",", "", regex=False).str.strip()
        series = series.replace({"": None, "nan": None, "None": None})
    return pd.to_numeric(series, errors="coerce")


# ======================================================================================
# Models (Qt Model/View)
# ======================================================================================
# Qt's Model/View architecture:
#   - The model provides data (cells, headers, formatting)
#   - The view (QTableView) renders it and handles user interactions (selection/click)
#
# We implement two models:
#   1) PreviewTableModel: shows a sheet preview and highlights selected columns
#   2) ExecutionListModel: shows the list of analyses to be executed
# ======================================================================================

class PreviewTableModel(QAbstractTableModel):
    """
    Table model used for previewing the selected sheet.

    Features:
      - Displays the first N rows (here, 50) of the sheet
      - Supports "picked column" UX highlighting
      - Supports committed roles:
          * I-SD column (blue highlight)
          * V-G column (green highlight)
      - Updates header text to mark selected columns
    """
    def __init__(self):
        super().__init__()

        # Column names and preview data (stored as list of lists for simplicity)
        self._headers: List[str] = []
        self._rows: List[List[object]] = []

        # UI state for selections
        self.selected_col: Optional[int] = None  # last clicked column (temporary pick)
        self.isd_col: Optional[int] = None       # committed I-SD selection
        self.vg_col: Optional[int] = None        # committed V-G selection

        # Background brushes used for highlighting columns
        self._brush_selected = QBrush(QColor("#FFF6D6"))  # pale yellow (current pick)
        self._brush_isd = QBrush(QColor("#E8F1FF"))       # pale blue (I-SD)
        self._brush_vg = QBrush(QColor("#E9FAEE"))        # pale green (V-G)
        self._brush_both = QBrush(QColor("#F0E9FF"))      # pale purple (shouldn't happen)

    def set_data(self, headers: List[str], rows: List[List[object]]) -> None:
        """
        Replace the entire preview dataset.

        This triggers a model reset which tells Qt:
          - "All indexes may have changed"
          - Rebuild the view based on new headers/rows
        """
        self.beginResetModel()
        self._headers = headers
        self._rows = rows

        # Reset selection state when a new sheet is loaded
        self.selected_col = None
        self.isd_col = None
        self.vg_col = None

        self.endResetModel()

    def set_selected_col(self, col: Optional[int]) -> None:
        """
        Set the "picked" column (soft selection highlight).
        This does not commit it as I-SD or V-G; it just marks what the user clicked.
        """
        self.selected_col = col

        # Emit dataChanged so view redraws cell backgrounds
        if self.columnCount() > 0 and self.rowCount() > 0:
            self.dataChanged.emit(
                self.index(0, 0),
                self.index(self.rowCount() - 1, self.columnCount() - 1),
                [Qt.BackgroundRole],
            )

        # Emit header change so header markers update
        self.headerDataChanged.emit(Qt.Horizontal, 0, max(0, self.columnCount() - 1))

    def set_isd_col(self, col: Optional[int]) -> None:
        """
        Commit a column as I-SD.
        This affects:
          - Background highlighting
          - Header "(I-SD)" tag
        """
        self.isd_col = col
        if self.columnCount() > 0 and self.rowCount() > 0:
            self.dataChanged.emit(
                self.index(0, 0),
                self.index(self.rowCount() - 1, self.columnCount() - 1),
                [Qt.BackgroundRole],
            )
        self.headerDataChanged.emit(Qt.Horizontal, 0, max(0, self.columnCount() - 1))

    def set_vg_col(self, col: Optional[int]) -> None:
        """
        Commit a column as V-G.
        This affects:
          - Background highlighting
          - Header "(V-G)" tag
        """
        self.vg_col = col
        if self.columnCount() > 0 and self.rowCount() > 0:
            self.dataChanged.emit(
                self.index(0, 0),
                self.index(self.rowCount() - 1, self.columnCount() - 1),
                [Qt.BackgroundRole],
            )
        self.headerDataChanged.emit(Qt.Horizontal, 0, max(0, self.columnCount() - 1))

    def clear_roles(self) -> None:
        """
        Clear the picked column and both committed column roles.
        Used by the "Clear" button in the UI.
        """
        self.selected_col = None
        self.isd_col = None
        self.vg_col = None
        if self.columnCount() > 0 and self.rowCount() > 0:
            self.dataChanged.emit(
                self.index(0, 0),
                self.index(self.rowCount() - 1, self.columnCount() - 1),
                [Qt.BackgroundRole],
            )
        self.headerDataChanged.emit(Qt.Horizontal, 0, max(0, self.columnCount() - 1))

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return len(self._rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return len(self._headers)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        """
        Provide data for each cell depending on role:

        - DisplayRole/EditRole: return the cell text
        - BackgroundRole: return a QBrush to color the background
        """
        if not index.isValid():
            return None

        r, c = index.row(), index.column()

        if role in (Qt.DisplayRole, Qt.EditRole):
            try:
                v = self._rows[r][c]
            except Exception:
                v = ""
            return "" if v is None else str(v)

        if role == Qt.BackgroundRole:
            # Committed roles take priority over soft selection
            if self.isd_col is not None and self.vg_col is not None and c == self.isd_col == self.vg_col:
                return self._brush_both
            if self.isd_col is not None and c == self.isd_col:
                return self._brush_isd
            if self.vg_col is not None and c == self.vg_col:
                return self._brush_vg
            if self.selected_col is not None and c == self.selected_col:
                return self._brush_selected

        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        """
        Provide header labels.

        Horizontal header:
          - Column name
          - Plus "(I-SD)" or "(V-G)" tag for committed selections

        Vertical header:
          - Row number (1-based)
        """
        if role != Qt.DisplayRole:
            return None
        if orientation != Qt.Horizontal:
            return str(section + 1)

        if not (0 <= section < len(self._headers)):
            return str(section + 1)

        name = self._headers[section]
        tag = ""
        if self.isd_col == section:
            tag += " (I-SD)"
        if self.vg_col == section:
            tag += " (V-G)"
        return f"{name}{tag}"


@dataclass
class ExecRow:
    file_name: str
    sheet_name: str
    w: str
    l: str
    c: str
    fit_window_v: str
    pn: str
    i_sd: str
    v_g: str
    comment: str


class ExecutionListModel(QAbstractTableModel):
    """
    Table model for the "Execution list" section.

    - Displays rows added via the "add" button.
    - First column "#" is auto-generated from the row position.
    - The actual stored data is in self._rows, which the Execute action reads directly.
    """
    HEADERS = ["#", "File name", "Sheet name", "W", "L", "C", "FitWin(V)", "P/N", "I-SD", "V-G", "Comment"]

    def __init__(self):
        super().__init__()
        self._rows: List[ExecRow] = []

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return len(self._rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return len(self.HEADERS)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid() or role != Qt.DisplayRole:
            return None

        row_idx = index.row()
        r = self._rows[row_idx]

        values = [
            str(row_idx + 1),
            r.file_name,
            r.sheet_name,
            r.w,
            r.l,
            r.c,
            r.fit_window_v,
            r.pn,
            r.i_sd,
            r.v_g,
            r.comment,
        ]
        return values[index.column()]

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self.HEADERS[section]
        return str(section + 1)

    def add_row(self, row: ExecRow) -> None:
        self.beginInsertRows(QModelIndex(), len(self._rows), len(self._rows))
        self._rows.append(row)
        self.endInsertRows()

    def remove_rows(self, rows_to_remove: List[int]) -> None:
        uniq = sorted(set(r for r in rows_to_remove if 0 <= r < len(self._rows)), reverse=True)
        if not uniq:
            return
        for r in uniq:
            self.beginRemoveRows(QModelIndex(), r, r)
            del self._rows[r]
            self.endRemoveRows()


# ======================================================================================
# Drag & drop line edit
# ======================================================================================

class DropLineEdit(QLineEdit):
    def __init__(self, placeholder: str = ""):
        super().__init__()
        self.setPlaceholderText(placeholder)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        paths = [u.toLocalFile() for u in urls if u.isLocalFile()]

        parent = self.parent()
        while parent is not None and not hasattr(parent, "open_files"):
            parent = parent.parent()

        if parent is not None:
            parent.open_files(paths)  # type: ignore[arg-type]

        event.acceptProposedAction()


@dataclass
class OpenedBook:
    path: str
    excel: pd.ExcelFile
    sheet_names: List[str]


# ======================================================================================
# Main Window (GUI)
# ======================================================================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("FET Mobility Analyzer")
        self.resize(1020, 820)
        self.setAcceptDrops(True)

        self.books: Dict[str, OpenedBook] = {}
        self.current_path: Optional[str] = None

        self.exec_model = ExecutionListModel()
        self.preview_model = PreviewTableModel()

        self._build_menu()
        self._build_ui()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        paths = [u.toLocalFile() for u in urls if u.isLocalFile()]
        if paths:
            self.open_files(paths)
        event.acceptProposedAction()

    def _build_menu(self) -> None:
        menubar = QMenuBar(self)
        self.setMenuBar(menubar)

        m_open = menubar.addMenu("Open")
        act_open = QAction("Open Excel…", self)
        m_open.addAction(act_open)
        act_open.triggered.connect(self.on_open_dialog)

        m_help = menubar.addMenu("How to use")
        act_help = QAction("How to use", self)
        m_help.addAction(act_help)

        howto_text = (
            "This software analyze FET devices using the saturation-regime equation.\n\n"
            "How to use:\n"
            "[1] Drag and drop the Keithley output file(s) (xls).\n"
            "[2] Select a file from “Select file.”\n"
            "[3] Select a sheet from “Select sheet.”\n"
            "[4] Enter a comment (optional).\n"
            "[5] Input W, L, C, fit window, and Type.\n"
            "[6] From the sheet preview, select the columns for I(S–D) and V(G), then click "
            "“Set as I-SD” and “Set as V-G,” respectively.\n"
            "[7] Click “Add” to include in the analysis list.\n"
            "[8] Select a folder for exporting, then click “Execute.”\n\n"
            "Hope you get excellent results!\n"
            "S"
        )

        act_help.triggered.connect(
            lambda: QMessageBox.information(
                self,
                "How to use",
                howto_text,
            )
        )

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)

        outer = QVBoxLayout(root)
        outer.setContentsMargins(14, 14, 14, 14)
        outer.setSpacing(12)

        gb_input = QGroupBox("(1) Input")
        outer.addWidget(gb_input)

        input_layout = QHBoxLayout(gb_input)
        input_layout.setSpacing(28)
        input_layout.setContentsMargins(12, 10, 12, 12)

        left = QVBoxLayout()
        left.setSpacing(10)
        input_layout.addLayout(left, stretch=3)

        row_open = QHBoxLayout()
        row_open.setSpacing(10)
        left.addLayout(row_open)

        lbl_open = QLabel("Open")
        lbl_open.setFixedWidth(78)
        row_open.addWidget(lbl_open)

        self.le_drop = DropLineEdit("(drag and drop)")
        row_open.addWidget(self.le_drop, stretch=1)

        self.btn_open_xls = QPushButton("Open xls")
        self.btn_open_xls.setFixedWidth(120)
        self.btn_open_xls.clicked.connect(self.on_open_dialog)
        row_open.addWidget(self.btn_open_xls)

        row_file = QHBoxLayout()
        row_file.setSpacing(10)
        left.addLayout(row_file)

        lbl_file = QLabel("Select file")
        lbl_file.setFixedWidth(78)
        row_file.addWidget(lbl_file)

        self.cb_file = QComboBox()
        self.cb_file.currentIndexChanged.connect(self.on_file_changed)
        row_file.addWidget(self.cb_file, stretch=1)

        row_sheet = QHBoxLayout()
        row_sheet.setSpacing(10)
        left.addLayout(row_sheet)

        lbl_sheet = QLabel("Select sheet")
        lbl_sheet.setFixedWidth(78)
        row_sheet.addWidget(lbl_sheet)

        self.cb_sheet = QComboBox()
        self.cb_sheet.currentIndexChanged.connect(self.on_sheet_changed)
        row_sheet.addWidget(self.cb_sheet, stretch=1)

        row_comment = QHBoxLayout()
        row_comment.setSpacing(10)
        left.addLayout(row_comment)

        lbl_comment = QLabel("Comment")
        lbl_comment.setFixedWidth(78)
        lbl_comment.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        row_comment.addWidget(lbl_comment)

        self.te_comment = QTextEdit()
        self.te_comment.setPlaceholderText("Optional notes…")
        fm = self.te_comment.fontMetrics()
        self.te_comment.setFixedHeight(int(fm.lineSpacing() * 3.2) + 18)
        row_comment.addWidget(self.te_comment, stretch=1)

        right = QVBoxLayout()
        right.setSpacing(10)
        input_layout.addLayout(right, stretch=2)

        def form_row(label: str) -> tuple[QHBoxLayout, QLineEdit]:
            r = QHBoxLayout()
            r.setSpacing(10)
            lab = QLabel(label)
            lab.setFixedWidth(80)
            r.addWidget(lab)
            le = QLineEdit()
            le.setMinimumWidth(150)
            r.addWidget(le, stretch=1)
            return r, le

        r_w, self.le_w = form_row("W (µm)")
        r_l, self.le_l = form_row("L (µm)")
        r_c, self.le_c = form_row("C (F/cm²)")
        r_fw, self.le_fitwin = form_row("fit window (V)")

        self.le_w.setPlaceholderText("e.g., 1000")
        self.le_l.setPlaceholderText("e.g., 30")
        self.le_c.setPlaceholderText("e.g., 1.15E-08")
        self.le_fitwin.setPlaceholderText("e.g., 10, or 20-30")

        right.addLayout(r_w)
        right.addLayout(r_l)
        right.addLayout(r_c)
        right.addLayout(r_fw)

        type_row = QHBoxLayout()
        type_row.setSpacing(10)
        lab_type = QLabel("Type")
        lab_type.setFixedWidth(105)
        type_row.addWidget(lab_type)

        self.rb_p = QRadioButton("p-type")
        self.rb_n = QRadioButton("n-type")
        self.rb_p.setChecked(True)

        type_row.addWidget(self.rb_p)
        type_row.addWidget(self.rb_n)
        type_row.addStretch(1)
        right.addLayout(type_row)

        gb_sel = QGroupBox("(2) Select column")
        outer.addWidget(gb_sel, stretch=1)
        sel_layout = QVBoxLayout(gb_sel)
        sel_layout.setSpacing(10)

        ctrl_row = QHBoxLayout()
        ctrl_row.setSpacing(10)
        sel_layout.addLayout(ctrl_row)

        self.lbl_pick = QLabel("Click a column header (or any cell) to pick a column.")
        self.lbl_pick.setStyleSheet("color: #444;")
        ctrl_row.addWidget(self.lbl_pick, stretch=1)

        self.btn_set_isd = QPushButton("Set as I-SD")
        self.btn_set_vg = QPushButton("Set as V-G")
        self.btn_clear_cols = QPushButton("Clear")

        self.btn_set_isd.setFixedWidth(120)
        self.btn_set_vg.setFixedWidth(120)
        self.btn_clear_cols.setFixedWidth(90)

        ctrl_row.addWidget(self.btn_set_isd)
        ctrl_row.addWidget(self.btn_set_vg)
        ctrl_row.addWidget(self.btn_clear_cols)

        self.tbl_preview = QTableView()
        self.tbl_preview.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.tbl_preview.setFrameShape(QFrame.Box)
        self.tbl_preview.setAlternatingRowColors(True)
        self.tbl_preview.setModel(self.preview_model)
        self.tbl_preview.horizontalHeader().setStretchLastSection(True)
        self.tbl_preview.verticalHeader().setVisible(False)
        self.tbl_preview.setSelectionBehavior(QTableView.SelectItems)
        self.tbl_preview.setSelectionMode(QTableView.SingleSelection)
        sel_layout.addWidget(self.tbl_preview, stretch=1)

        self.tbl_preview.horizontalHeader().sectionClicked.connect(self.on_preview_column_clicked)
        self.tbl_preview.clicked.connect(self.on_preview_cell_clicked)

        self.btn_set_isd.clicked.connect(self.on_set_isd)
        self.btn_set_vg.clicked.connect(self.on_set_vg)
        self.btn_clear_cols.clicked.connect(self.on_clear_cols)

        add_row = QHBoxLayout()
        add_row.addStretch(1)
        self.btn_add = QPushButton("add")
        self.btn_add.setFixedWidth(120)
        add_row.addWidget(self.btn_add)
        sel_layout.addLayout(add_row)
        self.btn_add.clicked.connect(self._on_add_clicked)

        gb_exec = QGroupBox("(3) Execution list")
        outer.addWidget(gb_exec, stretch=1)
        exec_layout = QVBoxLayout(gb_exec)
        exec_layout.setSpacing(10)

        top_exec_row = QHBoxLayout()
        top_exec_row.addStretch(1)
        self.btn_clear_line = QPushButton("Clear line")
        self.btn_clear_line.setFixedWidth(120)
        top_exec_row.addWidget(self.btn_clear_line)
        exec_layout.addLayout(top_exec_row)
        self.btn_clear_line.clicked.connect(self._on_clear_line_clicked)

        self.tbl_exec = QTableView()
        self.tbl_exec.setModel(self.exec_model)
        self.tbl_exec.setFrameShape(QFrame.Box)
        self.tbl_exec.setAlternatingRowColors(True)
        self.tbl_exec.horizontalHeader().setStretchLastSection(True)
        self.tbl_exec.verticalHeader().setVisible(False)
        self.tbl_exec.setSelectionBehavior(QTableView.SelectRows)
        self.tbl_exec.setSelectionMode(QTableView.ExtendedSelection)

        exec_layout.addWidget(self.tbl_exec, stretch=1)

        exec_btn_row = QHBoxLayout()
        exec_btn_row.setSpacing(10)

        lbl_export = QLabel("Export folder")
        lbl_export.setFixedWidth(90)
        exec_btn_row.addWidget(lbl_export)

        self.le_export = QLineEdit()
        self.le_export.setPlaceholderText("Select export folder…")
        exec_btn_row.addWidget(self.le_export, stretch=1)

        self.btn_browse_export = QPushButton("Browse…")
        self.btn_browse_export.setFixedWidth(100)
        exec_btn_row.addWidget(self.btn_browse_export)
        self.btn_browse_export.clicked.connect(self._on_browse_export_clicked)

        exec_btn_row.addStretch(0)

        self.btn_execute = QPushButton("execute")
        self.btn_execute.setFixedWidth(140)
        exec_btn_row.addWidget(self.btn_execute)

        exec_layout.addLayout(exec_btn_row)
        self.btn_execute.clicked.connect(self._on_execute_clicked)

        self._refresh_file_dropdown()

    def on_preview_column_clicked(self, col: int) -> None:
        if self.preview_model.columnCount() <= 0:
            return
        self.preview_model.set_selected_col(col)
        name = self.preview_model._headers[col] if 0 <= col < len(self.preview_model._headers) else f"Col {col+1}"
        self.lbl_pick.setText(f"Picked: [{col+1}] {name}")

    def on_preview_cell_clicked(self, idx: QModelIndex) -> None:
        if not idx.isValid():
            return
        self.on_preview_column_clicked(idx.column())

    def _ensure_pick(self) -> Optional[int]:
        col = self.preview_model.selected_col
        if col is None:
            QMessageBox.information(self, "Pick a column", "Please click a column header (or any cell) first.")
        return col

    def on_set_isd(self) -> None:
        col = self._ensure_pick()
        if col is None:
            return
        if self.preview_model.vg_col == col:
            QMessageBox.warning(self, "Conflict", "This column is already set as V-G. Choose another column.")
            return
        self.preview_model.set_isd_col(col)

    def on_set_vg(self) -> None:
        col = self._ensure_pick()
        if col is None:
            return
        if self.preview_model.isd_col == col:
            QMessageBox.warning(self, "Conflict", "This column is already set as I-SD. Choose another column.")
            return
        self.preview_model.set_vg_col(col)

    def on_clear_cols(self) -> None:
        self.preview_model.clear_roles()
        self.lbl_pick.setText("Click a column header (or any cell) to pick a column.")

    def on_open_dialog(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Open Excel files",
            "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*)",
        )
        if paths:
            self.open_files(paths)

    def open_files(self, paths: List[str]) -> None:
        norm_paths: List[str] = []
        for p in paths:
            if not p:
                continue
            try:
                pp = str(Path(p).expanduser().resolve())
            except Exception:
                continue
            if os.path.isfile(pp):
                norm_paths.append(pp)

        if not norm_paths:
            return

        opened_any = False
        errors: List[str] = []

        for path in norm_paths:
            if path in self.books:
                opened_any = True
                continue
            try:
                engine = self._engine_for(path)
                xls = pd.ExcelFile(path, engine=engine)
                self.books[path] = OpenedBook(path=path, excel=xls, sheet_names=list(xls.sheet_names))
                opened_any = True
            except Exception as e:
                errors.append(f"{Path(path).name}: {e}")

        if opened_any:
            self.le_drop.setText(norm_paths[0] if len(norm_paths) == 1 else f"{len(norm_paths)} files loaded")
            sel = self.current_path or norm_paths[0]
            self._refresh_file_dropdown(select_path=sel)

        if errors:
            QMessageBox.warning(self, "Some files failed to open", "Failed:\n" + "\n".join(errors))

    def _engine_for(self, path: str) -> str:
        ext = Path(path).suffix.lower()
        return "xlrd" if ext == ".xls" else "openpyxl"

    def _refresh_file_dropdown(self, select_path: Optional[str] = None) -> None:
        self.cb_file.blockSignals(True)
        self.cb_file.clear()

        paths = sorted(self.books.keys(), key=lambda p: Path(p).name.lower())
        for p in paths:
            self.cb_file.addItem(Path(p).name, userData=p)

        self.cb_file.blockSignals(False)

        if paths:
            if select_path and select_path in self.books:
                idx = self._find_file_index(select_path)
                self.cb_file.setCurrentIndex(idx if idx is not None else 0)
            else:
                self.cb_file.setCurrentIndex(0)

            self.on_file_changed()
        else:
            self.current_path = None
            self.cb_sheet.clear()
            self.preview_model.set_data([], [])
            self.preview_model.clear_roles()

    def _find_file_index(self, path: str) -> Optional[int]:
        for i in range(self.cb_file.count()):
            if self.cb_file.itemData(i) == path:
                return i
        return None

    def on_file_changed(self) -> None:
        path = self.cb_file.currentData()
        if not path or path not in self.books:
            self.current_path = None
            self.cb_sheet.clear()
            self.preview_model.set_data([], [])
            self.preview_model.clear_roles()
            return

        self.current_path = path
        book = self.books[path]

        self.cb_sheet.blockSignals(True)
        self.cb_sheet.clear()
        self.cb_sheet.addItems(book.sheet_names)
        self.cb_sheet.blockSignals(False)

        if book.sheet_names:
            self.cb_sheet.setCurrentIndex(0)
            self.on_sheet_changed()

    def on_sheet_changed(self) -> None:
        if not self.current_path:
            return
        sheet = self.cb_sheet.currentText()
        if not sheet:
            self.preview_model.set_data([], [])
            self.preview_model.clear_roles()
            return

        try:
            book = self.books[self.current_path]
            df = book.excel.parse(sheet_name=sheet, nrows=50)
        except Exception as e:
            QMessageBox.warning(self, "Read failed", f"Could not read sheet:\n{e}")
            self.preview_model.set_data([], [])
            self.preview_model.clear_roles()
            return

        headers = [str(c) for c in df.columns]
        rows: List[List[object]] = df.fillna("").astype(str).values.tolist()

        self.preview_model.set_data(headers, rows)
        self.lbl_pick.setText("Click a column header (or any cell) to pick a column.")

    def _require_inputs_or_warn(self) -> Optional[dict]:
        w = self.le_w.text().strip()
        l = self.le_l.text().strip()
        c = self.le_c.text().strip()
        fitwin = self.le_fitwin.text().strip()

        missing = []
        if not w:
            missing.append("W")
        if not l:
            missing.append("L")
        if not c:
            missing.append("C")
        if not fitwin:
            missing.append("fit window")

        if missing:
            QMessageBox.warning(
                self,
                "Missing input",
                "Please input required fields:\n  - " + "\n  - ".join(missing),
            )
            return None

        return {"w": w, "l": l, "c": c, "fitwin": fitwin}

    def _find_book_path_by_filename(self, file_name: str) -> Optional[str]:
        target = (file_name or "").strip()
        if not target:
            return None
        for full_path in self.books.keys():
            if Path(full_path).name == target:
                return full_path
        return None

    def _on_add_clicked(self) -> None:
        if not self.current_path:
            QMessageBox.warning(self, "No file", "Please open and select a file first.")
            return
        sheet = self.cb_sheet.currentText()
        if not sheet:
            QMessageBox.warning(self, "No sheet", "Please select a sheet.")
            return

        if self.preview_model.isd_col is None or self.preview_model.vg_col is None:
            QMessageBox.warning(self, "Columns not set", "Please set both I-SD and V-G columns.")
            return

        req = self._require_inputs_or_warn()
        if req is None:
            return

        pn = "P" if self.rb_p.isChecked() else "N"
        comment = self.te_comment.toPlainText().strip()

        self.exec_model.add_row(
            ExecRow(
                file_name=Path(self.current_path).name,
                sheet_name=sheet,
                w=req["w"],
                l=req["l"],
                c=req["c"],
                fit_window_v=req["fitwin"],
                pn=pn,
                i_sd=str(self.preview_model.isd_col + 1),
                v_g=str(self.preview_model.vg_col + 1),
                comment=comment,
            )
        )

    def _on_browse_export_clicked(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Select export folder", "")
        if folder:
            self.le_export.setText(folder)

    def _on_clear_line_clicked(self) -> None:
        sel = self.tbl_exec.selectionModel()
        if sel is None:
            return

        rows = sorted({idx.row() for idx in sel.selectedRows()})
        if not rows:
            QMessageBox.information(self, "Clear line", "Please select row(s) in the execution list.")
            return

        self.exec_model.remove_rows(rows)

    def _on_execute_clicked(self) -> None:
        """
        Execute the analysis for each item in the Execution list.

        For each row:
          - Validate file/sheet exist
          - Parse numeric parameters (W/L/C, fit window)
          - Extract VG and ID columns as numeric series
          - Analyze and export PNG plot (named with index prefix)
          - Append results + status to a CSV summary file

        IMPORTANT in this version:
          Fit may fail intentionally if the data do not show the expected p-type / n-type
          transfer direction. This is desirable behavior.
        """
        n = self.exec_model.rowCount()
        if n == 0:
            QMessageBox.information(self, "Execute", "Execution list is empty.")
            return

        export_dir = self.le_export.text().strip()
        if not export_dir:
            QMessageBox.warning(self, "Export folder", "Please select an export folder before execute.")
            return
        export_path = Path(export_dir)
        if not export_path.exists():
            QMessageBox.warning(self, "Export folder", "The selected export folder does not exist.")
            return

        bad_rows = []
        for i, r in enumerate(self.exec_model._rows, start=1):
            if not (r.w.strip() and r.l.strip() and r.c.strip() and r.fit_window_v.strip()):
                bad_rows.append(i)
        if bad_rows:
            QMessageBox.warning(
                self,
                "Missing input",
                "Some rows have missing W/L/C/fit window.\nRows: " + ", ".join(map(str, bad_rows)),
            )
            return

        today = datetime.now().strftime("%Y%m%d")
        csv_path = export_path / f"fet-results_{today}.csv"

        results_rows: List[Dict[str, Any]] = []
        ok = 0
        failed: List[str] = []
        summary_lines: List[str] = []

        for idx0, r in enumerate(self.exec_model._rows, start=1):
            status = "OK"
            error_msg = ""

            mobility = ""
            vth = ""
            fit_vmin = ""
            fit_vmax = ""
            r2 = ""

            full_path = self._find_book_path_by_filename(r.file_name)
            if full_path is None or full_path not in self.books:
                status = "ERROR"
                error_msg = "file not loaded"
            else:
                book = self.books[full_path]

                if r.sheet_name not in book.sheet_names:
                    status = "ERROR"
                    error_msg = "sheet not found"
                else:
                    try:
                        w_um = _parse_float_token(r.w)
                        l_um = _parse_float_token(r.l)
                        c_fcm2 = _parse_float_token(r.c)
                        if w_um <= 0 or l_um <= 0 or c_fcm2 <= 0:
                            raise ValueError("W/L/C must be positive.")
                        fit_spec = parse_fit_window_gui(r.fit_window_v)
                        dev_type: Literal["p", "n"] = "p" if r.pn.strip().upper() == "P" else "n"
                        col_isd_1b = int(str(r.i_sd).strip())
                        col_vg_1b = int(str(r.v_g).strip())
                    except Exception as e:
                        status = "ERROR"
                        error_msg = f"bad params: {e}"

                    if status == "OK":
                        try:
                            df = book.excel.parse(sheet_name=r.sheet_name)

                            if df.shape[1] < max(col_isd_1b, col_vg_1b):
                                raise ValueError(f"Sheet has only {df.shape[1]} columns.")

                            s_vg = _coerce_numeric_series(df.iloc[:, col_vg_1b - 1])
                            s_id = _coerce_numeric_series(df.iloc[:, col_isd_1b - 1])

                            valid = s_vg.notna() & s_id.notna()
                            vg = s_vg[valid].astype(float).tolist()
                            isd = s_id[valid].astype(float).tolist()

                            if len(vg) < 5:
                                raise ValueError(f"Too few numeric points: {len(vg)}")
                        except Exception as e:
                            status = "ERROR"
                            error_msg = f"read/extract failed: {e}"

                    if status == "OK":
                        stem = _sanitize_filename(Path(r.file_name).stem)
                        sheet_safe = _sanitize_filename(r.sheet_name)
                        out_png = export_path / f"{idx0}_{stem}_{sheet_safe}.png"

                        try:
                            title = f"{r.file_name} :: {r.sheet_name}"

                            res = analyze_fet_and_save_figure(
                                vg,
                                isd,
                                w_um=w_um,
                                l_um=l_um,
                                c_fcm2=c_fcm2,
                                dev_type=dev_type,
                                fit_spec=fit_spec,
                                title=title,
                                comment=r.comment,
                                out_png=out_png,
                            )

                            ok += 1
                            mobility = f"{res['mobility']:.6E}"
                            vth = f"{res['vth']:.6g}"
                            fit_vmin = f"{res['fit_vmin']:.6g}"
                            fit_vmax = f"{res['fit_vmax']:.6g}"
                            r2 = f"{res['r2']:.6g}"

                            summary_lines.append(
                                f"[{idx0}] OK  {r.file_name} :: {r.sheet_name}  mu={mobility} cm/Vs  Vth={vth} V"
                            )
                        except Exception as e:
                            status = "ERROR"
                            error_msg = f"analysis failed: {e}"

            if status != "OK":
                failed.append(f"[{idx0}] {r.file_name} :: {r.sheet_name}  ({error_msg})")

            results_rows.append(
                {
                    "#": idx0,
                    "file_name": r.file_name,
                    "sheet_name": r.sheet_name,
                    "W_um": r.w,
                    "L_um": r.l,
                    "C_Fcm2": r.c,
                    "fit_window_input": r.fit_window_v,
                    "P/N": r.pn,
                    "I-SD_col(1based)": r.i_sd,
                    "V-G_col(1based)": r.v_g,
                    "comment": r.comment,
                    "status": status,
                    "error": error_msg,
                    "mobility_cm_per_Vs": mobility,
                    "Vth_V": vth,
                    "fit_Vmin": fit_vmin,
                    "fit_Vmax": fit_vmax,
                    "R2": r2,
                }
            )

        try:
            fieldnames = [
                "#",
                "file_name",
                "sheet_name",
                "W_um",
                "L_um",
                "C_Fcm2",
                "fit_window_input",
                "P/N",
                "I-SD_col(1based)",
                "V-G_col(1based)",
                "comment",
                "status",
                "error",
                "mobility_cm_per_Vs",
                "Vth_V",
                "fit_Vmin",
                "fit_Vmax",
                "R2",
            ]

            with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
                wcsv = csv.DictWriter(f, fieldnames=fieldnames)
                wcsv.writeheader()
                for row in results_rows:
                    wcsv.writerow(row)
        except Exception as e:
            QMessageBox.warning(self, "CSV export failed", f"Could not write CSV:\n{csv_path}\n\n{e}")

        msg = [f"Export folder:\n{str(export_path)}", "", f"Done: {ok}/{n}", f"CSV: {csv_path.name}"]

        if summary_lines:
            msg += ["", "Results:"] + summary_lines[:30]
            if len(summary_lines) > 30:
                msg += [f"... ({len(summary_lines)-30} more)"]

        if failed:
            msg += ["", "Failed:"] + failed[:30]
            if len(failed) > 30:
                msg += [f"... ({len(failed)-30} more)"]

        QMessageBox.information(self, "Execute", "\n".join(msg))


# ======================================================================================
# App entry point
# ======================================================================================
def main() -> int:
    app = QApplication(sys.argv)
    apply_light_fusion_theme(app)
    w = MainWindow()
    w.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
