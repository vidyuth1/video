"""
Mold Well Tracker
=================
Streamlit app that lets you upload a photo of a mold (or plate) laid out as a
15 x 8 grid (120 coordinates / "wells"), click directly on the image to mark
a well as Empty or Present, and autosaves every click as it happens.

How it works
------------
1. Upload an image.
2. Calibrate the grid ONCE per image: click the center of the first well
   (top-left) then the center of the last well (bottom-right). The 120
   points are evenly interpolated between those two clicks.
3. Switch to "Mark wells" mode and click any well dot to toggle it between
   Present (green) and Empty (red). Every click is written to disk
   immediately -- no save button needed.

All state (which wells are empty + the grid calibration) is stored per-image
(keyed by a hash of the image bytes) in the `well_data/` folder as JSON, so
re-uploading the same image later restores exactly where you left off.
"""

import hashlib
import json
import math
import os
from io import BytesIO

import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from streamlit_image_coordinates import streamlit_image_coordinates

# --------------------------------------------------------------------------
# CONFIG - tweak these to match your mold
# --------------------------------------------------------------------------
ROWS = 8                        # 8 rows    -> labeled A-H
COLS = 15                       # 15 columns -> labeled 1-15  (8 x 15 = 120)
ROW_LABELS = "ABCDEFGH"

MAX_DISPLAY_WIDTH = 900         # working image width in pixels
DEFAULT_CLICK_RADIUS = 18       # px - how close a click must be to count
DOT_RADIUS = 7                  # px - visual size of each well marker

STATE_DIR = "well_data"         # autosave folder (per-image JSON files)
os.makedirs(STATE_DIR, exist_ok=True)

PRESENT_COLOR = (46, 204, 113)  # green
EMPTY_COLOR = (231, 76, 60)     # red
CALIB_COLOR = (52, 152, 219)    # blue

st.set_page_config(page_title="Mold Well Tracker", layout="wide")

# --------------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------------

def well_ids():
    """All 120 well ids in row-major order: A1..A15, B1..B15, ... H1..H15."""
    return [f"{ROW_LABELS[r]}{c + 1}" for r in range(ROWS) for c in range(COLS)]


def image_key(file_bytes: bytes) -> str:
    """Stable id for an uploaded image, used to namespace its saved state."""
    return hashlib.sha256(file_bytes).hexdigest()[:16]


def state_path(key: str) -> str:
    return os.path.join(STATE_DIR, f"{key}.json")


def load_state(key: str) -> dict:
    path = state_path(key)
    data = {}
    if os.path.exists(path):
        try:
            with open(path, "r") as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError):
            data = {}
    if "wells" not in data or set(data["wells"].keys()) != set(well_ids()):
        data["wells"] = {w: "present" for w in well_ids()}
    if "calibration" not in data:
        data["calibration"] = None  # -> [[x1, y1], [x2, y2]] once calibrated
    return data


def save_state(key: str, data: dict) -> bool:
    """Write state to disk immediately (this IS the autosave). Returns
    False (and warns) instead of crashing if the filesystem is read-only."""
    try:
        with open(state_path(key), "w") as f:
            json.dump(data, f, indent=2)
        return True
    except OSError:
        st.warning(
            "Could not write autosave file to disk (read-only filesystem?). "
            "Your changes are still kept for this session — use the "
            "'Download state (JSON)' button in the sidebar to save manually.",
            icon="⚠️",
        )
        return False


def default_calibration(img_w, img_h):
    margin_x = img_w * 0.06
    margin_y = img_h * 0.08
    return [[margin_x, margin_y], [img_w - margin_x, img_h - margin_y]]


def compute_positions(calibration, img_w, img_h):
    """Evenly interpolate 120 (x, y) points between two calibration corners."""
    if not calibration or len(calibration) != 2:
        calibration = default_calibration(img_w, img_h)
    (x1, y1), (x2, y2) = calibration
    col_step = (x2 - x1) / (COLS - 1) if COLS > 1 else 0
    row_step = (y2 - y1) / (ROWS - 1) if ROWS > 1 else 0
    positions = {}
    for r in range(ROWS):
        for c in range(COLS):
            wid = f"{ROW_LABELS[r]}{c + 1}"
            positions[wid] = (x1 + c * col_step, y1 + r * row_step)
    return positions


def nearest_well(x, y, positions, max_dist):
    best_id, best_d = None, None
    for wid, (px, py) in positions.items():
        d = math.hypot(px - x, py - y)
        if best_d is None or d < best_d:
            best_id, best_d = wid, d
    return best_id if best_d is not None and best_d <= max_dist else None


def draw_overlay(base_img, positions, wells, show_labels, calibration_points=None):
    img = base_img.convert("RGB").copy()
    draw = ImageDraw.Draw(img)
    font = ImageFont.load_default()
    for wid, (x, y) in positions.items():
        color = PRESENT_COLOR if wells.get(wid, "present") == "present" else EMPTY_COLOR
        r = DOT_RADIUS
        draw.ellipse([x - r, y - r, x + r, y + r], fill=color, outline=(0, 0, 0))
        if show_labels:
            draw.text((x - r, y - r - 12), wid, fill=(0, 0, 0), font=font)
    if calibration_points:
        for (x, y) in calibration_points:
            r = DOT_RADIUS + 3
            draw.ellipse([x - r, y - r, x + r, y + r], outline=CALIB_COLOR, width=3)
    return img


def resize_working(img: Image.Image, max_w: int) -> Image.Image:
    if img.width <= max_w:
        return img.copy()
    ratio = max_w / img.width
    return img.resize((max_w, int(img.height * ratio)), Image.LANCZOS)


# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------

st.title("🧫 Mold Well Tracker")
st.caption(
    f"Upload a mold photo, calibrate the {ROWS}x{COLS} grid once, then click "
    "any well to mark it Empty / Present. Every click is saved automatically."
)

uploaded = st.file_uploader("Upload mold image", type=["png", "jpg", "jpeg"])

if uploaded is None:
    st.info("Upload an image to get started.")
    st.stop()

file_bytes = uploaded.getvalue()
key = image_key(file_bytes)
state = load_state(key)

raw_img = Image.open(BytesIO(file_bytes))
work_img = resize_working(raw_img, MAX_DISPLAY_WIDTH)
img_w, img_h = work_img.size

# ---- sidebar -------------------------------------------------------------
with st.sidebar:
    st.header("Controls")
    mode = st.radio("Mode", ["Mark wells", "Calibrate grid"], index=0)
    show_labels = st.checkbox("Show well labels", value=False)
    click_radius = st.slider("Click sensitivity (px)", 8, 40, DEFAULT_CLICK_RADIUS)

    st.divider()
    present_count = sum(1 for v in state["wells"].values() if v == "present")
    empty_count = len(state["wells"]) - present_count
    c1, c2 = st.columns(2)
    c1.metric("Present", present_count)
    c2.metric("Empty", empty_count)

    st.divider()
    if st.button("Reset all wells to Present", use_container_width=True):
        state["wells"] = {w: "present" for w in well_ids()}
        save_state(key, state)
        st.rerun()
    if st.button("Reset calibration", use_container_width=True):
        state["calibration"] = None
        save_state(key, state)
        st.rerun()

    st.divider()
    export_df = pd.DataFrame(
        [{"well": w, "status": s} for w, s in state["wells"].items()]
    )
    st.download_button(
        "Download results (CSV)",
        export_df.to_csv(index=False).encode("utf-8"),
        file_name=f"mold_{key}_wells.csv",
        mime="text/csv",
        use_container_width=True,
    )
    st.download_button(
        "Download state (JSON)",
        json.dumps(state, indent=2).encode("utf-8"),
        file_name=f"mold_{key}_state.json",
        mime="application/json",
        use_container_width=True,
    )

    restore_file = st.file_uploader(
        "Restore a previously downloaded state (JSON)", type=["json"], key="restore"
    )
    if restore_file is not None:
        try:
            restored = json.load(restore_file)
            if "wells" in restored:
                state["wells"].update(restored["wells"])
            if restored.get("calibration"):
                state["calibration"] = restored["calibration"]
            save_state(key, state)
            st.success("State restored.")
            st.rerun()
        except (json.JSONDecodeError, KeyError):
            st.error("That file doesn't look like a valid state export.")

positions = compute_positions(state["calibration"], img_w, img_h)

# ---- main panel ------------------------------------------------------
if mode == "Calibrate grid":
    st.subheader("Step 1 — Calibrate the grid")
    st.write(
        "Click the **center of well A1** (top-left), then click the "
        f"**center of the last well** (bottom-right, {ROW_LABELS[-1]}{COLS}). "
        "The 120-point grid will be evenly spaced between those two clicks. "
        "Click again afterwards to re-calibrate from scratch."
    )
    calib_points = state["calibration"] or []
    overlay = draw_overlay(work_img, positions, state["wells"], show_labels, calib_points)
    click = streamlit_image_coordinates(overlay, key=f"calib_{key}")

    last_key = f"last_calib_click_{key}"
    if click is not None:
        click_sig = (click.get("x"), click.get("y"))
        if click_sig != (None, None) and st.session_state.get(last_key) != click_sig:
            st.session_state[last_key] = click_sig
            pts = state["calibration"] or []
            if len(pts) >= 2:
                pts = []
            pts.append([click_sig[0], click_sig[1]])
            state["calibration"] = pts
            save_state(key, state)
            st.rerun()

    if state["calibration"] and len(state["calibration"]) == 2:
        st.success("Calibration complete — switch to 'Mark wells' in the sidebar.")
    elif state["calibration"] and len(state["calibration"]) == 1:
        st.info("First point recorded. Now click the last well (bottom-right).")

else:
    st.subheader("Click a well to toggle Present ⇄ Empty")
    overlay = draw_overlay(work_img, positions, state["wells"], show_labels)
    click = streamlit_image_coordinates(overlay, key=f"mark_{key}")

    last_key = f"last_mark_click_{key}"
    if click is not None:
        click_sig = (click.get("x"), click.get("y"))
        if click_sig != (None, None) and st.session_state.get(last_key) != click_sig:
            st.session_state[last_key] = click_sig
            wid = nearest_well(click_sig[0], click_sig[1], positions, click_radius)
            if wid:
                state["wells"][wid] = "empty" if state["wells"][wid] == "present" else "present"
                save_state(key, state)
                st.rerun()
            else:
                st.toast("No well close enough to that click — try clicking closer to a dot.")

st.caption(
    f"Legend: 🟢 present &nbsp;&nbsp; 🔴 empty &nbsp;&nbsp; "
    f"Grid: {ROWS} rows x {COLS} cols = {ROWS * COLS} wells."
)

with st.expander("Show empty well list"):
    empty_wells = [w for w, s in state["wells"].items() if s == "empty"]
    st.write(", ".join(empty_wells) if empty_wells else "None marked empty yet.")
