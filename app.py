"""
Mold Well Tracker
=================
Streamlit app that lets you upload up to 40 photos of molds (or plates) laid
out as a 15 x 8 grid (120 coordinates), auto-prompts calibration on the first
image, then steps through each image with Next/Previous buttons.

Coordinate naming: columns A-O (15), rows 1-8 (8), e.g. "B3" = column B, row 3.

How it works
------------
1. Upload up to MAX_IMAGES images at once.
2. The app automatically opens the first image and prompts calibration.
3. Click the OUTER TOP-LEFT corner of the grid (just outside A1), then the
   OUTER BOTTOM-RIGHT corner (just outside O8). That rectangle is divided
   evenly into 15 × 8 cells.
4. After calibration the mode switches to "Mark wells". Click ANYWHERE inside
   a cell to toggle it between Present (green) and Empty (red). Every click is
   written to disk immediately.
5. Click "Next →" to move to the next image (or "← Previous" to go back).
6. After all images are reviewed, use the Export tab to download:
   - A heatmap image showing missing-coordinate frequency across all molds.
   - An Excel workbook with two sheets: the heatmap data and a frequency table.
"""

import hashlib
import io
import json
import os
from io import BytesIO

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import numpy as np
import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from streamlit_image_coordinates import streamlit_image_coordinates

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
COLS = 15
ROWS = 8
COL_LABELS = "ABCDEFGHIJKLMNO"
MAX_DISPLAY_WIDTH = 900
MAX_IMAGES = 40

STATE_DIR = "well_data"
os.makedirs(STATE_DIR, exist_ok=True)

PRESENT_FILL = (46, 204, 113)
EMPTY_FILL   = (231, 76,  60)
CALIB_COLOR  = (52,  152, 219)
GRID_LINE    = (0,   0,   0, 180)

st.set_page_config(page_title="Mold Well Tracker", layout="wide")


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def well_ids():
    """All 120 well ids in row-major order."""
    return [f"{COL_LABELS[c]}{r + 1}" for r in range(ROWS) for c in range(COLS)]


def file_signature(uploaded_file) -> str:
    """Unique identity for an uploaded file based on its actual content.

    Reads the first 256 KB + last 256 KB (or the whole file if smaller)
    and hashes those bytes together with the filename.  Fast on large images,
    collision-proof for any realistic set of mold photos — including burst
    shots from the same camera that share identical file sizes.
    """
    data = uploaded_file.getvalue()
    chunk = 256 * 1024  # 256 KB
    if len(data) <= chunk * 2:
        sample = data
    else:
        sample = data[:chunk] + data[-chunk:]
    raw = uploaded_file.name.encode() + sample
    return hashlib.sha256(raw).hexdigest()[:16]


# ---------------------------------------------------------------------------
# STATE PERSISTENCE
# ---------------------------------------------------------------------------

def state_path(sig: str) -> str:
    return os.path.join(STATE_DIR, f"{sig}.json")


def load_state(sig: str) -> dict:
    path = state_path(sig)
    data: dict = {}
    if os.path.exists(path):
        try:
            with open(path) as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError):
            data = {}
    if "wells" not in data or set(data["wells"].keys()) != set(well_ids()):
        data["wells"] = {w: "present" for w in well_ids()}
    # Normalize calibration to the canonical dict form (upgrades legacy
    # [[x1,y1],[x2,y2]] files saved before the gap-calibration feature).
    data["calibration"] = normalize_calibration(data.get("calibration"))
    return data


def save_state(sig: str, data: dict) -> bool:
    try:
        with open(state_path(sig), "w") as f:
            json.dump(data, f, indent=2)
        return True
    except OSError:
        st.warning(
            "Could not write autosave file (read-only filesystem?). "
            "Your changes are kept for this session.",
            icon="⚠️",
        )
        return False


# ---------------------------------------------------------------------------
# IMAGE DECODE — cached, lazy
# ---------------------------------------------------------------------------

# NOTE ON THE FIX:
# st.cache_data excludes any parameter whose name starts with an underscore
# from the cache key hash (this is how you tell Streamlit "don't hash this,
# it's expensive/unhashable"). The original code prefixed BOTH `_file_bytes`
# and `_sig` with underscores, which meant NEITHER contributed to the cache
# key -- only `max_w` did, and `max_w` is always the same constant. Result:
# the very first image decoded got cached, and every later call (for every
# other image) just returned that same cached PNG, regardless of which file
# was actually passed in.
#
# Fix: keep `_file_bytes` unhashed (it's just bytes, hashing it every call
# would be wasteful), but let `sig` (no leading underscore) participate in
# the cache key. `sig` is already a content hash of the uploaded file, so
# it uniquely and cheaply identifies each image for caching purposes.
@st.cache_data(max_entries=MAX_IMAGES, show_spinner=False)
def _decode_and_resize(_file_bytes: bytes, sig: str, max_w: int) -> bytes:
    img = Image.open(BytesIO(_file_bytes))
    if img.width > max_w:
        ratio = max_w / img.width
        img = img.resize((max_w, int(img.height * ratio)), Image.LANCZOS)
    buf = BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def get_working_image(uploaded_file, sig: str) -> Image.Image:
    raw_bytes = uploaded_file.getvalue()
    png_bytes  = _decode_and_resize(raw_bytes, sig, MAX_DISPLAY_WIDTH)
    return Image.open(BytesIO(png_bytes))


# ---------------------------------------------------------------------------
# GRID MATH
# ---------------------------------------------------------------------------
#
# CALIBRATION FORMATS
# --------------------
# "calibration" is stored per-image as either:
#   - None                                          -> not yet calibrated
#   - {"gap_after_row": 0, "points": [[x,y],[x,y]]}  -> simple uniform grid
#       points = [outer top-left, outer bottom-right]
#   - {"gap_after_row": N, "points": [4 points]}     -> gap-aware grid
#       points = [outer top-left,
#                 bottom edge of row N (end of block 1),
#                 top edge of row N+1 (start of block 2),
#                 outer bottom-right]
#       Rows 1..N are spaced evenly within the block-1 band; rows N+1..ROWS
#       are spaced evenly within the block-2 band. Anything between the two
#       bands (the physical gap) maps to no well.
#
# Legacy files saved before the gap feature stored calibration as a plain
# [[x1,y1],[x2,y2]] list -- normalize_calibration() upgrades those on load.

def default_calibration(img_w: int, img_h: int):
    mx = img_w  * 0.04
    my = img_h  * 0.06
    return [[mx, my], [img_w - mx, img_h - my]]


def normalize_calibration(calibration) -> dict | None:
    """Coerce any stored calibration (old list format or new dict format)
    into the canonical {"gap_after_row": int, "points": [...]} shape, or
    None if there's no calibration yet."""
    if not calibration:
        return None
    if isinstance(calibration, list):
        return {"gap_after_row": 0, "points": calibration}
    if isinstance(calibration, dict):
        return {
            "gap_after_row": int(calibration.get("gap_after_row", 0) or 0),
            "points": calibration.get("points", []),
        }
    return None


def required_calibration_points(gap_after_row: int) -> int:
    return 4 if gap_after_row > 0 else 2


def calibration_instructions(gap_after_row: int) -> list[str]:
    """Ordered instruction text, one entry per click still needed."""
    if gap_after_row <= 0:
        return [
            "Click the **outer top-left corner** of the grid (just outside A1).",
            f"Click the **outer bottom-right corner** of the grid "
            f"(just outside {COL_LABELS[-1]}{ROWS}).",
        ]
    nxt = gap_after_row + 1
    return [
        "Click the **outer top-left corner** of the grid (just outside A1).",
        f"Click just **below row {gap_after_row}** — the bottom edge of the "
        f"last row before the gap (e.g. just below A{gap_after_row}).",
        f"Click just **above row {nxt}** — the top edge of the first row "
        f"after the gap (e.g. just above A{nxt}).",
        f"Click the **outer bottom-right corner** of the grid "
        f"(just outside {COL_LABELS[-1]}{ROWS}).",
    ]


def compute_cell_bounds(calibration, img_w: int, img_h: int):
    """Returns (bounds, geometry).

    bounds: dict of well-id -> (x0, y0, x1, y1) pixel rectangle, always 120 entries.
    geometry: dict used by find_cell() -- {"x1","x2","col_w","bands":[...]}
      Each band is {"y1","y2","row_h","row_start","row_count"}. One band for a
      simple grid, two bands (with a gap between them) for a gap-aware grid.
    """
    calib = normalize_calibration(calibration)
    gap_after_row = calib["gap_after_row"] if calib else 0
    pts = calib["points"] if calib else []
    req = required_calibration_points(gap_after_row)

    if len(pts) < req:
        # Not fully calibrated yet -- fall back to a default preview rectangle
        # (uniform, no gap) so the UI has something sane to draw.
        pts = default_calibration(img_w, img_h)
        gap_after_row = 0

    if gap_after_row <= 0:
        (x1, y1), (x2, y2) = pts[0], pts[1]
        x1, x2 = min(x1, x2), max(x1, x2)
        y1, y2 = min(y1, y2), max(y1, y2)
        col_w = (x2 - x1) / COLS
        row_h = (y2 - y1) / ROWS
        bounds = {}
        for r in range(ROWS):
            for c in range(COLS):
                wid  = f"{COL_LABELS[c]}{r + 1}"
                cx0  = x1 + c * col_w
                cy0  = y1 + r * row_h
                bounds[wid] = (cx0, cy0, cx0 + col_w, cy0 + row_h)
        geometry = {
            "x1": x1, "x2": x2, "col_w": col_w,
            "bands": [
                {"y1": y1, "y2": y2, "row_h": row_h, "row_start": 0, "row_count": ROWS}
            ],
        }
        return bounds, geometry

    # ---- Gap-aware grid: two independently-spaced row bands --------------
    p1, p2, p3, p4 = pts
    x1, x2 = min(p1[0], p4[0]), max(p1[0], p4[0])
    y1a, y2a = sorted([p1[1], p2[1]])   # block 1 (rows 1..N): top / bottom
    y1b, y2b = sorted([p3[1], p4[1]])   # block 2 (rows N+1..ROWS): top / bottom

    row_count1 = gap_after_row
    row_count2 = ROWS - gap_after_row
    row_h1 = (y2a - y1a) / row_count1 if row_count1 > 0 else 0
    row_h2 = (y2b - y1b) / row_count2 if row_count2 > 0 else 0
    col_w  = (x2 - x1) / COLS

    bounds = {}
    for r in range(ROWS):
        if r < row_count1:
            y0  = y1a + r * row_h1
            y1c = y0 + row_h1
        else:
            r2  = r - row_count1
            y0  = y1b + r2 * row_h2
            y1c = y0 + row_h2
        for c in range(COLS):
            wid = f"{COL_LABELS[c]}{r + 1}"
            x0  = x1 + c * col_w
            bounds[wid] = (x0, y0, x0 + col_w, y1c)

    geometry = {
        "x1": x1, "x2": x2, "col_w": col_w,
        "bands": [
            {"y1": y1a, "y2": y2a, "row_h": row_h1, "row_start": 0, "row_count": row_count1},
            {"y1": y1b, "y2": y2b, "row_h": row_h2, "row_start": row_count1, "row_count": row_count2},
        ],
    }
    return bounds, geometry


def find_cell(x: float, y: float, geometry: dict) -> str | None:
    """Resolve a click to a well id using the (possibly two-band) geometry.
    Returns None if the click falls outside the grid entirely, OR inside the
    physical gap between the two row blocks (there's no well there)."""
    x1, x2, col_w = geometry["x1"], geometry["x2"], geometry["col_w"]
    if x < x1 or x > x2:
        return None
    col = min(int((x - x1) // col_w), COLS - 1) if col_w > 0 else 0
    for band in geometry["bands"]:
        if band["y1"] <= y <= band["y2"]:
            row_h = band["row_h"]
            local_row = min(int((y - band["y1"]) // row_h), band["row_count"] - 1) if row_h > 0 else 0
            row = band["row_start"] + local_row
            return f"{COL_LABELS[col]}{row + 1}"
    return None  # landed in the gap, or above/below the calibrated grid


# ---------------------------------------------------------------------------
# OVERLAY RENDERING
# ---------------------------------------------------------------------------

def _dashed_hline(draw, y, width, color, dash=10, gap=6):
    x = 0
    while x < width:
        x_end = min(x + dash, width)
        draw.line([(x, y), (x_end, y)], fill=color, width=2)
        x += dash + gap


def draw_grid_overlay(
    base_img: Image.Image,
    bounds: dict,
    wells: dict,
    show_labels: bool,
    calib_points=None,
    gap_after_row: int = 0,
    gap_band=None,
) -> Image.Image:
    """
    gap_after_row: >0 while collecting the 4 gap-calibration points, used to
        shade the in-progress gap band as points 2 and 3 are placed.
    gap_band: (y_top, y_bottom) -- used in normal marking mode (once fully
        calibrated) to shade the finalized gap region for visual context.
    """
    img   = base_img.convert("RGBA")
    layer = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw  = ImageDraw.Draw(layer)
    font  = ImageFont.load_default()

    # Shade the finalized gap band (marking mode) so it's visually obvious
    # why there are no clickable wells in that strip.
    if gap_band is not None:
        y_top, y_bot = sorted(gap_band)
        draw.rectangle([0, y_top, img.width, y_bot], fill=(120, 120, 120, 55))

    for wid, (x0, y0, x1, y1) in bounds.items():
        present = wells.get(wid, "present") == "present"
        fill    = (*PRESENT_FILL, 80) if present else (*EMPTY_FILL, 130)
        draw.rectangle([x0, y0, x1, y1], fill=fill, outline=GRID_LINE)
        if show_labels:
            cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
            draw.text((cx, cy), wid, fill=(0, 0, 0, 255), font=font, anchor="mm")

    if calib_points:
        # While collecting the 2 middle gap points, shade the in-progress
        # gap band so the user can see what they're marking.
        if gap_after_row > 0 and len(calib_points) >= 3:
            y_top, y_bot = sorted([calib_points[1][1], calib_points[2][1]])
            draw.rectangle([0, y_top, img.width, y_bot], fill=(120, 120, 120, 70))

        for i, (x, y) in enumerate(calib_points, start=1):
            r = 9
            _dashed_hline(draw, y, img.width, (*CALIB_COLOR, 140))
            draw.ellipse(
                [x - r, y - r, x + r, y + r],
                outline=(*CALIB_COLOR, 255),
                width=3,
            )
            draw.text((x, y - r - 12), str(i), fill=(*CALIB_COLOR, 255), font=font, anchor="mm")

    return Image.alpha_composite(img, layer).convert("RGB")


# ---------------------------------------------------------------------------
# FREQUENCY / EXPORT DATA
# ---------------------------------------------------------------------------

def compute_frequency(sigs_and_names: list[tuple[str, str]]) -> dict:
    """Count how many images have each well marked empty."""
    freq = {w: 0 for w in well_ids()}
    n_images = len(sigs_and_names)
    for sig, _ in sigs_and_names:
        data = load_state(sig)
        for wid, status in data["wells"].items():
            if status == "empty":
                freq[wid] += 1
    return freq, n_images


# Green -> yellow -> orange -> red, evenly spaced at 0 / .33 / .66 / 1.0.
# These are Excel's own conditional-formatting green/yellow/red plus an
# orange midpoint, so the gradient matches native Excel color scales.
_GRADIENT_CMAP = mcolors.LinearSegmentedColormap.from_list(
    "green_yellow_orange_red",
    ["#63BE7B", "#FFEB84", "#FFA500", "#F8696B"],
)


def gradient_hex(count: int, maximum: int) -> str:
    """Map a count to a hex color along a green -> yellow -> orange -> red
    scale (green = low/good, red = high/bad), matching the reference
    heatmap and frequency-table styling. The Excel export and the
    on-screen heatmap both draw from this same gradient so they stay
    visually consistent."""
    ratio = (count / maximum) if maximum > 0 else 0.0
    ratio = max(0.0, min(1.0, ratio))
    r, g, b, _a = _GRADIENT_CMAP(ratio)
    return f"{int(round(r * 255)):02X}{int(round(g * 255)):02X}{int(round(b * 255)):02X}"


def readable_text_hex(bg_hex: str) -> str:
    """Pick black or white text for legibility against a given hex
    background, based on perceptual luminance (rather than a fixed
    threshold) so it works correctly across the whole green-to-red scale."""
    r, g, b = int(bg_hex[0:2], 16), int(bg_hex[2:4], 16), int(bg_hex[4:6], 16)
    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    return "000000" if luminance > 0.6 else "FFFFFF"


def build_heatmap_figure(freq: dict, n_images: int) -> plt.Figure:
    """Build a matplotlib heatmap of missing-well frequency."""
    grid = np.zeros((ROWS, COLS), dtype=float)
    for r in range(ROWS):
        for c in range(COLS):
            wid = f"{COL_LABELS[c]}{r + 1}"
            grid[r, c] = freq.get(wid, 0)

    fig, ax = plt.subplots(figsize=(14, 6))
    cmap = _GRADIENT_CMAP  # green (low) -> yellow -> orange -> red (high)
    im = ax.imshow(grid, cmap=cmap, aspect="auto",
                   vmin=0, vmax=max(n_images, 1))

    # Annotate each cell with count and %
    for r in range(ROWS):
        for c in range(COLS):
            val = int(grid[r, c])
            pct = (val / n_images * 100) if n_images > 0 else 0
            text_hex = readable_text_hex(gradient_hex(val, max(n_images, 1)))
            color = "white" if text_hex == "FFFFFF" else "black"
            ax.text(c, r, f"{val}\n({pct:.0f}%)", ha="center", va="center",
                    fontsize=7, color=color, fontweight="bold")

    ax.set_xticks(range(COLS))
    ax.set_xticklabels(list(COL_LABELS), fontsize=9)
    ax.set_yticks(range(ROWS))
    ax.set_yticklabels([str(i + 1) for i in range(ROWS)], fontsize=9)
    ax.set_xlabel("Column", fontsize=11)
    ax.set_ylabel("Row", fontsize=11)
    ax.set_title(
        f"Missing Coordinate Frequency Heatmap\n"
        f"{n_images} mold{'s' if n_images != 1 else ''} analyzed",
        fontsize=13, fontweight="bold"
    )

    cbar = fig.colorbar(im, ax=ax, shrink=0.8)
    cbar.set_label("# Times Missing", fontsize=9)
    plt.tight_layout()
    return fig


def build_excel_export(freq: dict, n_images: int) -> bytes:
    """
    Build a two-sheet Excel workbook:
      Sheet 1 – Heatmap (color-coded grid of missing frequency)
      Sheet 2 – Frequency Table (coordinate, count, %)
    """
    wb = openpyxl.Workbook()

    # ---- SHEET 1: Heatmap ------------------------------------------------
    ws_heat = wb.active
    ws_heat.title = "Heatmap"

    thin = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_font  = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    label_font   = Font(name="Arial", bold=True, size=10)
    cell_font    = Font(name="Arial", size=9)
    center       = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Title row
    ws_heat.merge_cells("A1:Q1")
    title_cell = ws_heat["A1"]
    title_cell.value = f"Missing Coordinate Frequency Heatmap — {n_images} Mold(s) Analyzed"
    title_cell.font  = Font(name="Arial", bold=True, size=13)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill("solid", fgColor="2C3E50")
    title_cell.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    ws_heat.row_dimensions[1].height = 30

    # Sub-header row (column letters A-O in columns B-P, row 2)
    ws_heat["A2"].value = "Row \\ Col"
    ws_heat["A2"].font  = header_font
    ws_heat["A2"].fill  = PatternFill("solid", fgColor="34495E")
    ws_heat["A2"].alignment = center
    ws_heat["A2"].border = border

    for c_idx, col_letter in enumerate(COL_LABELS):
        cell = ws_heat.cell(row=2, column=c_idx + 2)
        cell.value = col_letter
        cell.font  = header_font
        cell.fill  = PatternFill("solid", fgColor="34495E")
        cell.alignment = center
        cell.border = border

    ws_heat.row_dimensions[2].height = 22

    # Max frequency for color scaling
    max_freq = max(freq.values()) if freq else 1

    # Data rows (rows 3..10 → grid rows 1..8)
    for r_idx in range(ROWS):
        row_num = r_idx + 3
        ws_heat.row_dimensions[row_num].height = 28

        # Row label
        label_cell = ws_heat.cell(row=row_num, column=1)
        label_cell.value = str(r_idx + 1)
        label_cell.font  = label_font
        label_cell.fill  = PatternFill("solid", fgColor="ECF0F1")
        label_cell.alignment = center
        label_cell.border = border

        for c_idx in range(COLS):
            wid   = f"{COL_LABELS[c_idx]}{r_idx + 1}"
            count = freq.get(wid, 0)
            pct   = (count / n_images * 100) if n_images > 0 else 0
            hex_color = gradient_hex(count, max_freq)

            data_cell = ws_heat.cell(row=row_num, column=c_idx + 2)
            data_cell.value = f"{count} ({pct:.0f}%)"
            data_cell.font  = Font(
                name="Arial", size=8, bold=count > 0,
                color=readable_text_hex(hex_color)
            )
            data_cell.fill  = PatternFill("solid", fgColor=hex_color)
            data_cell.alignment = center
            data_cell.border = border

    # Column widths
    ws_heat.column_dimensions["A"].width = 10
    for c_idx in range(COLS):
        ws_heat.column_dimensions[get_column_letter(c_idx + 2)].width = 10

    # Legend note
    note_row = ROWS + 4
    ws_heat.merge_cells(f"A{note_row}:Q{note_row}")
    note = ws_heat.cell(row=note_row, column=1)
    note.value = (
        f"Each cell shows: count of molds where this coordinate was missing "
        f"(percentage of {n_images} total molds). "
        f"Color follows a green → yellow → orange → red scale: "
        f"green = rarely missing, red = frequently missing."
    )
    note.font = Font(name="Arial", italic=True, size=9, color="555555")
    note.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws_heat.row_dimensions[note_row].height = 30

    # ---- SHEET 2: Frequency Table ----------------------------------------
    ws_table = wb.create_sheet("Frequency Table")

    # Title
    ws_table.merge_cells("A1:E1")
    t = ws_table["A1"]
    t.value = f"Coordinate Flag Frequency Table — {n_images} Mold(s) Analyzed"
    t.font  = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    t.fill  = PatternFill("solid", fgColor="2C3E50")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws_table.row_dimensions[1].height = 30

    # Column headers
    headers = ["Coordinate", "Column", "Row", "Times Missing", "% Missing"]
    header_fills = ["34495E"] * 5
    for col_i, (h, fill) in enumerate(zip(headers, header_fills), start=1):
        cell = ws_table.cell(row=2, column=col_i)
        cell.value = h
        cell.font  = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        cell.fill  = PatternFill("solid", fgColor=fill)
        cell.alignment = center
        cell.border = border
    ws_table.row_dimensions[2].height = 22

    # Data — sorted by frequency descending
    all_wids = well_ids()
    sorted_wids = sorted(all_wids, key=lambda w: -freq.get(w, 0))

    for row_i, wid in enumerate(sorted_wids, start=3):
        count = freq.get(wid, 0)
        pct   = (count / n_images * 100) if n_images > 0 else 0
        col_letter = wid[0]
        row_number = wid[1:]

        hex_color = gradient_hex(count, max_freq)
        text_color = readable_text_hex(hex_color)

        vals = [wid, col_letter, row_number, count, round(pct, 1)]
        for col_i, val in enumerate(vals, start=1):
            cell = ws_table.cell(row=row_i, column=col_i)
            cell.value = val
            cell.border = border
            cell.alignment = center

            if col_i == 4:  # Times Missing
                cell.fill = PatternFill("solid", fgColor=hex_color)
                cell.font = Font(name="Arial", size=10, bold=True, color=text_color)
            elif col_i == 5:  # % Missing
                cell.fill = PatternFill("solid", fgColor=hex_color)
                cell.font = Font(name="Arial", size=10, color=text_color)
                cell.number_format = "0.0%"
                cell.value = pct / 100  # store as fraction for proper % format
            else:
                cell.font = Font(name="Arial", size=10)
                # Alternate row shading
                if row_i % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="F8F9FA")

        ws_table.row_dimensions[row_i].height = 18

    # Column widths
    col_widths = [14, 10, 8, 16, 14]
    for col_i, w in enumerate(col_widths, start=1):
        ws_table.column_dimensions[get_column_letter(col_i)].width = w

    # Summary footer
    footer_row = len(sorted_wids) + 4
    ws_table.merge_cells(f"A{footer_row}:E{footer_row}")
    footer = ws_table.cell(row=footer_row, column=1)
    total_missing = sum(freq.values())
    footer.value = (
        f"Total missing observations across all coordinates: {total_missing}  |  "
        f"Molds analyzed: {n_images}  |  Sorted by frequency (highest first)"
    )
    footer.font = Font(name="Arial", italic=True, size=9, color="555555")
    footer.alignment = Alignment(horizontal="left", vertical="center")
    ws_table.row_dimensions[footer_row].height = 22

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ===========================================================================
# UI
# ===========================================================================

st.title("🧫 Mold Well Tracker")
st.caption(
    f"Upload up to {MAX_IMAGES} mold photos. The app will guide you through "
    "calibrating and marking each image, then export a frequency heatmap."
)

# ---------------------------------------------------------------------------
# File upload
# ---------------------------------------------------------------------------
uploaded_files = st.file_uploader(
    f"Upload mold images (max {MAX_IMAGES})",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("Upload one or more images to get started.")
    st.stop()

if len(uploaded_files) > MAX_IMAGES:
    st.warning(
        f"You uploaded {len(uploaded_files)} images but the limit is {MAX_IMAGES}. "
        f"Only the first {MAX_IMAGES} will be used.",
        icon="⚠️",
    )
    uploaded_files = uploaded_files[:MAX_IMAGES]

# Build per-file signatures.  If two uploaded files produce the same content
# hash (truly identical images), append the index so they remain distinct
# slots — we never want two entries in sig_list to be the same string, because
# that would make the dict lookups below collapse them into one.
_raw_sigs = [file_signature(f) for f in uploaded_files]
sigs_and_names: list[tuple[str, str]] = []
_seen: dict[str, int] = {}
for i, (raw_sig, f) in enumerate(zip(_raw_sigs, uploaded_files)):
    if raw_sig in _seen:
        unique_sig = f"{raw_sig}_{i}"   # make it unique by appending index
    else:
        unique_sig = raw_sig
    _seen[raw_sig] = i
    sigs_and_names.append((unique_sig, f.name))

sig_to_file = {sig: f for (sig, _), f in zip(sigs_and_names, uploaded_files)}
all_states: dict[str, dict] = {sig: load_state(sig) for sig, _ in sigs_and_names}
sig_list  = [sig for sig, _ in sigs_and_names]
name_map  = {sig: name for sig, name in sigs_and_names}

n_total = len(sig_list)

# ---------------------------------------------------------------------------
# Shared calibration: propagate image-0 calibration to all other images
# ---------------------------------------------------------------------------
# Calibration is done exactly once on the first image. Every subsequent image
# automatically inherits those same corner coordinates so the user never has
# to re-calibrate. Images that already have their own saved calibration keep
# it (supports the edge-case where the user manually reset one image).
first_sig   = sig_list[0]
first_calib = all_states[first_sig].get("calibration")  # already normalized dict or None
if first_calib:
    _req = required_calibration_points(first_calib.get("gap_after_row", 0))
    if len(first_calib.get("points", [])) == _req:
        for sig in sig_list[1:]:
            if not all_states[sig].get("calibration"):
                all_states[sig]["calibration"] = {
                    "gap_after_row": first_calib["gap_after_row"],
                    "points": list(first_calib["points"]),  # independent copy
                }
                save_state(sig, all_states[sig])

# ---------------------------------------------------------------------------
# Session state: current image index
# ---------------------------------------------------------------------------
if "current_idx" not in st.session_state:
    st.session_state.current_idx = 0
# Clamp in case files changed
st.session_state.current_idx = max(0, min(st.session_state.current_idx, n_total - 1))

# ---------------------------------------------------------------------------
# Top navigation tabs: Annotation | Export
# ---------------------------------------------------------------------------
tab_annotate, tab_export = st.tabs(["📷 Annotate Images", "📊 Export Results"])

# ===========================================================================
# TAB 1 — ANNOTATE
# ===========================================================================
with tab_annotate:

    idx = st.session_state.current_idx
    active_sig   = sig_list[idx]
    active_file  = sig_to_file[active_sig]
    active_state = all_states[active_sig]

    # Decode active image only
    work_img     = get_working_image(active_file, active_sig)
    img_w, img_h = work_img.size
    bounds, geometry = compute_cell_bounds(active_state["calibration"], img_w, img_h)

    # --- Header row: image progress + nav buttons -------------------------
    h_col1, h_col2, h_col3 = st.columns([1, 4, 1])
    with h_col1:
        if st.button("← Previous", key="btn_prev_top", disabled=(idx == 0), use_container_width=True):
            st.session_state.current_idx -= 1
            st.rerun()
    with h_col2:
        st.markdown(
            f"<div style='text-align:center;font-size:1.05rem;padding-top:6px;'>"
            f"<b>Image {idx + 1} of {n_total}</b> &nbsp;·&nbsp; "
            f"<code>{name_map[active_sig]}</code>"
            f"</div>",
            unsafe_allow_html=True,
        )
    with h_col3:
        if st.button("Next →", key="btn_next_top", disabled=(idx == n_total - 1), use_container_width=True):
            st.session_state.current_idx += 1
            st.rerun()

    st.progress((idx + 1) / n_total)

    # --- Sidebar controls -------------------------------------------------
    with st.sidebar:
        # ---- Full reset: wipe everything and start from scratch ----------
        if st.session_state.get("confirm_full_reset"):
            st.error(
                "This deletes ALL saved calibration and well data for "
                "every image. This cannot be undone."
            )
            rc1, rc2 = st.columns(2)
            with rc1:
                if st.button("✅ Confirm reset", key="btn_confirm_reset", use_container_width=True, type="primary"):
                    for fname in os.listdir(STATE_DIR):
                        try:
                            os.remove(os.path.join(STATE_DIR, fname))
                        except OSError:
                            pass
                    for k in list(st.session_state.keys()):
                        del st.session_state[k]
                    st.rerun()
            with rc2:
                if st.button("Cancel", key="btn_cancel_reset", use_container_width=True):
                    st.session_state["confirm_full_reset"] = False
                    st.rerun()
        else:
            if st.button("🔄 Reset — start from scratch", key="btn_start_reset", use_container_width=True):
                st.session_state["confirm_full_reset"] = True
                st.rerun()

        st.divider()
        st.header("Controls")
        show_labels = st.checkbox("Show cell labels", value=True)

        st.divider()
        present_count = sum(1 for v in active_state["wells"].values() if v == "present")
        empty_count   = len(active_state["wells"]) - present_count
        c1, c2 = st.columns(2)
        c1.metric("Present", present_count)
        c2.metric("Empty",   empty_count)

        st.divider()
        if st.button("Reset all wells → Present", key="btn_reset_wells", use_container_width=True):
            active_state["wells"] = {w: "present" for w in well_ids()}
            save_state(active_sig, active_state)
            st.rerun()
        if st.button("Reset calibration", key="btn_reset_calib", use_container_width=True):
            active_state["calibration"] = None
            save_state(active_sig, active_state)
            st.rerun()

        st.divider()
        st.download_button(
            "Download this image (CSV)",
            pd.DataFrame(
                [{"well": w, "status": s} for w, s in active_state["wells"].items()]
            ).to_csv(index=False).encode("utf-8"),
            file_name=f"mold_{active_sig}_wells.csv",
            mime="text/csv",
            use_container_width=True,
        )
        st.download_button(
            "Download this image (JSON)",
            json.dumps(active_state, indent=2).encode("utf-8"),
            file_name=f"mold_{active_sig}_state.json",
            mime="application/json",
            use_container_width=True,
        )

        restore_file = st.file_uploader(
            "Restore state (JSON)",
            type=["json"],
            key="restore",
        )
        if restore_file is not None:
            try:
                restored = json.load(restore_file)
                if "wells" in restored:
                    active_state["wells"].update(restored["wells"])
                if restored.get("calibration"):
                    active_state["calibration"] = normalize_calibration(restored["calibration"])
                save_state(active_sig, active_state)
                st.success("State restored.")
                st.rerun()
            except (json.JSONDecodeError, KeyError):
                st.error("Invalid state file.")

    # --- Auto-determine mode: calibrate first if not done -----------------
    active_calib = active_state["calibration"]  # normalized dict or None
    needs_calibration = (
        active_calib is None
        or len(active_calib.get("points", [])) < required_calibration_points(
            active_calib.get("gap_after_row", 0)
        )
    )

    if needs_calibration:
        # ---- CALIBRATION MODE ----------------------------------------
        st.subheader("🎯 Step 1 — Calibrate the grid")

        if active_calib is None:
            # ---- Sub-step 0: set where the gap falls ----
            st.info(
                "This image needs calibration. Every mold is laid out as two "
                "row blocks with a physical gap between them — set which row "
                "the gap comes after, then start calibration."
            )

            prior_gap = st.session_state.get("gap_after_row_choice") or (ROWS // 2)
            gap_after_row = st.number_input(
                f"Gap occurs after row # (1–{ROWS - 1})",
                min_value=1,
                max_value=ROWS - 1,
                value=prior_gap,
                step=1,
                key=f"gap_row_{active_sig}",
            )

            st.image(work_img, use_container_width=True, caption="Preview (not yet calibrated)")

            if st.button("Start calibration ▶", key=f"start_calib_{active_sig}", type="primary"):
                gap_after_row = int(gap_after_row)
                st.session_state["gap_after_row_choice"] = gap_after_row
                active_state["calibration"] = {"gap_after_row": gap_after_row, "points": []}
                save_state(active_sig, active_state)
                st.rerun()

        else:
            # ---- Sub-step: collect the required clicks -----------------
            gap_after_row = active_calib.get("gap_after_row", 0)
            calib_pts = active_calib.get("points", [])
            req = required_calibration_points(gap_after_row)
            instructions = calibration_instructions(gap_after_row)

            st.info(f"Point {len(calib_pts) + 1} of {req}: {instructions[len(calib_pts)]}")
            if gap_after_row > 0:
                st.caption(
                    f"Layout: rows 1–{gap_after_row} in the upper block, "
                    f"rows {gap_after_row + 1}–{ROWS} in the lower block, "
                    "with a gap between them."
                )

            overlay = draw_grid_overlay(
                work_img, bounds, active_state["wells"], show_labels,
                calib_points=calib_pts, gap_after_row=gap_after_row,
            )
            click = streamlit_image_coordinates(overlay, key=f"calib_{active_sig}")

            last_key = f"last_calib_click_{active_sig}"
            if click is not None:
                sig_xy = (click.get("x"), click.get("y"))
                if sig_xy != (None, None) and st.session_state.get(last_key) != sig_xy:
                    st.session_state[last_key] = sig_xy
                    pts = list(calib_pts)
                    if len(pts) >= req:
                        pts = []  # safety net: start over if somehow overfull
                    pts.append([sig_xy[0], sig_xy[1]])
                    active_state["calibration"]["points"] = pts
                    save_state(active_sig, active_state)
                    st.rerun()

            with st.expander("Restart calibration"):
                st.caption("Clears the points you've placed so far for this image (keeps the gap setting).")
                if st.button("Clear points and restart", key=f"restart_pts_{active_sig}"):
                    active_state["calibration"] = {"gap_after_row": gap_after_row, "points": []}
                    save_state(active_sig, active_state)
                    st.rerun()
                st.caption("Or change the gap layout entirely for this image:")
                if st.button("Change gap layout", key=f"change_gap_{active_sig}"):
                    active_state["calibration"] = None
                    save_state(active_sig, active_state)
                    st.rerun()

    else:
        # ---- MARK WELLS MODE -----------------------------------------
        st.subheader("Step 2 — Click a cell to toggle Empty / Present")

        _active_gap = active_calib.get("gap_after_row", 0) if active_calib else 0
        col_info, col_nav = st.columns([5, 1])
        with col_info:
            gap_note = f" &nbsp;·&nbsp; gap after row {_active_gap}" if _active_gap > 0 else ""
            st.caption(
                f"🟢 present &nbsp; 🔴 empty &nbsp;·&nbsp; "
                f"Grid: {COLS} cols (A–{COL_LABELS[-1]}) × {ROWS} rows{gap_note}"
            )

        gap_band = None
        if len(geometry["bands"]) == 2:
            gap_band = (geometry["bands"][0]["y2"], geometry["bands"][1]["y1"])

        overlay = draw_grid_overlay(
            work_img, bounds, active_state["wells"], show_labels, gap_band=gap_band
        )
        click = streamlit_image_coordinates(overlay, key=f"mark_{active_sig}")

        last_key = f"last_mark_click_{active_sig}"
        if click is not None:
            sig_xy = (click.get("x"), click.get("y"))
            if sig_xy != (None, None) and st.session_state.get(last_key) != sig_xy:
                st.session_state[last_key] = sig_xy
                wid = find_cell(sig_xy[0], sig_xy[1], geometry)
                if wid:
                    current = active_state["wells"][wid]
                    active_state["wells"][wid] = "empty" if current == "present" else "present"
                    save_state(active_sig, active_state)
                    st.rerun()
                else:
                    st.toast("Click landed outside the calibrated grid.")

        # Quick empty-well list
        with st.expander("Show empty well list"):
            empty_wells = [w for w, s in active_state["wells"].items() if s == "empty"]
            st.write(", ".join(empty_wells) if empty_wells else "None marked empty yet.")

        # Bottom navigation
        st.divider()
        nav1, nav2, nav3 = st.columns([1, 6, 1])
        with nav1:
            if st.button("← Prev", key="btn_prev_bottom", disabled=(idx == 0), use_container_width=True):
                st.session_state.current_idx -= 1
                st.rerun()
        with nav2:
            # Image thumbnail strip
            progress_labels = []
            for i, (s, n) in enumerate(sigs_and_names):
                state_i = all_states[s]
                _c = state_i.get("calibration")
                calib_ok = bool(_c) and len(_c.get("points", [])) >= required_calibration_points(_c.get("gap_after_row", 0))
                icon = "✅" if calib_ok else "⏳"
                progress_labels.append(f"{icon} {i+1}")
            st.caption("  ".join(progress_labels))
        with nav3:
            if st.button("Next →", key="btn_next_bottom", disabled=(idx == n_total - 1), use_container_width=True):
                st.session_state.current_idx += 1
                st.rerun()

            if idx == n_total - 1:
                st.success("All images reviewed! Head to the **Export Results** tab.")


# ===========================================================================
# TAB 2 — EXPORT
# ===========================================================================
with tab_export:

    freq, n_images = compute_frequency(sigs_and_names)

    def _is_calibrated(sig):
        c = all_states[sig].get("calibration")
        if not c:
            return False
        return len(c.get("points", [])) >= required_calibration_points(c.get("gap_after_row", 0))

    n_calibrated = sum(1 for sig, _ in sigs_and_names if _is_calibrated(sig))

    st.subheader(f"📊 Results Summary — {n_images} mold(s) uploaded, {n_calibrated} calibrated")

    if n_images == 0:
        st.info("No images uploaded yet.")
    else:
        # ---- Heatmap preview -------------------------------------------
        st.markdown("### Missing Coordinate Frequency Heatmap")
        st.caption(
            "Color intensity shows how often each coordinate was marked empty "
            "across all analyzed molds."
        )
        fig = build_heatmap_figure(freq, n_images)
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

        # Save heatmap as PNG for download
        heatmap_buf = BytesIO()
        fig2 = build_heatmap_figure(freq, n_images)
        fig2.savefig(heatmap_buf, format="PNG", dpi=150, bbox_inches="tight")
        plt.close(fig2)
        heatmap_buf.seek(0)

        st.download_button(
            "⬇️ Download Heatmap (PNG)",
            heatmap_buf.getvalue(),
            file_name="mold_heatmap.png",
            mime="image/png",
            use_container_width=False,
        )

        st.divider()

        # ---- Frequency table preview ------------------------------------
        st.markdown("### Coordinate Flag Frequency Table")
        rows = []
        for wid in well_ids():
            count = freq.get(wid, 0)
            pct   = (count / n_images * 100) if n_images > 0 else 0
            rows.append({
                "Coordinate": wid,
                "Column": wid[0],
                "Row": wid[1:],
                "Times Missing": count,
                "% Missing": round(pct, 1),
            })

        df_freq = pd.DataFrame(rows).sort_values("Times Missing", ascending=False)
        st.dataframe(
            df_freq,
            use_container_width=True,
            hide_index=True,
            column_config={
                "% Missing": st.column_config.NumberColumn(format="%.1f%%"),
                "Times Missing": st.column_config.ProgressColumn(
                    min_value=0, max_value=n_images, format="%d"
                ),
            },
        )

        st.divider()

        # ---- Excel export -----------------------------------------------
        st.markdown("### Download Excel Report")
        st.caption(
            "The Excel workbook contains two sheets: **Heatmap** (color-coded grid) "
            "and **Frequency Table** (sorted list of all 120 coordinates)."
        )

        excel_bytes = build_excel_export(freq, n_images)
        st.download_button(
            "⬇️ Download Excel Report (.xlsx)",
            excel_bytes,
            file_name="mold_well_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=False,
        )
