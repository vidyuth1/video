"""
Chocolate Mold Inspector — Streamlit App
15 columns (A–O) × 8 rows (1–8) = 120 coordinates per mold frame

Upload frame photos, click coordinates directly on the photo to toggle
missing/present, and export a cumulative Excel heatmap across all frames.

Click-relay mechanism: JS writes the toggled coordinate into a hidden
st.text_input via React's native value setter + input event dispatch.
Streamlit sees the change, reruns, and the Python side persists the toggle.
"""

import base64
import io
import json
import os
import re
import time
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont

# ── Constants ──────────────────────────────────────────────────────────────────
COLS       = list("ABCDEFGHIJKLMNO")   # 15 columns  (A–O)
ROWS       = list(range(1, 9))         # 8 rows       (1–8)
ALL_COORDS = [f"{c}{r}" for r in ROWS for c in COLS]

DATA_FILE   = "mold_data.csv"
IMAGES_DIR  = "frame_images"

MAX_UPLOAD_FILES = 50

PALETTE = {
    "present":  "#2ECC71",
    "empty":    "#E74C3C",
}

os.makedirs(IMAGES_DIR, exist_ok=True)


# ── Data persistence ────────────────────────────────────────────────────────────

def load_data() -> pd.DataFrame:
    if os.path.exists(DATA_FILE):
        df = pd.read_csv(DATA_FILE, dtype=str)
        for coord in ALL_COORDS:
            if coord not in df.columns:
                df[coord] = "1"
        if "image_path" not in df.columns:
            df["image_path"] = ""
        return df
    return pd.DataFrame(
        columns=["frame_id", "frame_name", "timestamp", "image_path"] + ALL_COORDS
    )


def save_data(df: pd.DataFrame) -> None:
    df.to_csv(DATA_FILE, index=False)


def get_frame_dict(df: pd.DataFrame, frame_id: str) -> dict:
    row = df[df["frame_id"] == frame_id]
    if row.empty:
        return {c: True for c in ALL_COORDS}
    r = row.iloc[0]
    return {c: (str(r.get(c, "1")) == "1") for c in ALL_COORDS}


def upsert_frame(
    df: pd.DataFrame,
    frame_id: str,
    frame_name: str,
    coord_dict: dict,
    image_path: str = "",
) -> pd.DataFrame:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row_data = {
        "frame_id": frame_id,
        "frame_name": frame_name,
        "timestamp": ts,
        "image_path": image_path,
    }
    row_data.update({c: ("1" if v else "0") for c, v in coord_dict.items()})
    if frame_id in df["frame_id"].values:
        for k, v in row_data.items():
            df.loc[df["frame_id"] == frame_id, k] = v
    else:
        df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
    return df


def save_frame_image(frame_id: str, uploaded_file) -> str:
    ext  = os.path.splitext(uploaded_file.name)[-1].lower() or ".png"
    path = os.path.join(IMAGES_DIR, f"{frame_id}{ext}")
    with open(path, "wb") as f:
        f.write(uploaded_file.getvalue())
    return path


def load_frame_image(image_path: str):
    if image_path and os.path.exists(image_path):
        return Image.open(image_path).convert("RGB")
    return None


# ── Coordinate-from-click geometry ─────────────────────────────────────────────

def grid_geometry(W: int, H: int) -> dict:
    """Return the grid geometry parameters for a photo of size W×H."""
    margin_l = max(24, int(W * 0.038))
    margin_t = max(20, int(H * 0.055))
    grid_w   = W - margin_l - max(4, int(W * 0.008))
    grid_h   = H - margin_t - max(4, int(H * 0.008))
    cell_w   = grid_w / len(COLS)
    cell_h   = grid_h / len(ROWS)
    return dict(
        margin_l=margin_l, margin_t=margin_t,
        cell_w=cell_w, cell_h=cell_h,
        W=W, H=H,
    )


# ── Image helpers ───────────────────────────────────────────────────────────────

def pil_to_b64(img: Image.Image, fmt="JPEG") -> str:
    buf = io.BytesIO()
    img.save(buf, format=fmt, quality=88)
    return base64.b64encode(buf.getvalue()).decode()


def _try_font(size: int):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            pass
    return ImageFont.load_default()


def render_overlay(
    photo: Image.Image,
    coord_dict: dict,
    opacity: int = 90,
    highlight: str | None = None,
) -> Image.Image:
    """Draw a semi-transparent 15×8 overlay on the mold photo.

    highlight: coordinate to flash with a bright border (last-toggled feedback).
    """
    base = photo.convert("RGBA")
    W, H = base.size
    geo  = grid_geometry(W, H)
    ml, mt  = geo["margin_l"], geo["margin_t"]
    cw, ch  = geo["cell_w"],   geo["cell_h"]

    font_size = max(8, int(min(cw, ch) * 0.28))
    hdr_size  = max(9, int(min(ml, mt) * 0.52))
    fnt       = _try_font(font_size)
    fnt_hdr   = _try_font(hdr_size)

    overlay = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw    = ImageDraw.Draw(overlay)

    for ri, row_num in enumerate(ROWS):
        for ci, col in enumerate(COLS):
            coord   = f"{col}{row_num}"
            present = coord_dict.get(coord, True)
            x0 = ml + ci * cw
            y0 = mt + ri * ch
            x1 = x0 + cw
            y1 = y0 + ch

            alpha = opacity if present else min(opacity + 55, 240)
            fill  = (46, 204, 113, alpha) if present else (231, 76, 60, alpha)
            draw.rectangle([x0 + 1, y0 + 1, x1 - 1, y1 - 1], fill=fill)

            # Border — bright white highlight for last-toggled cell
            if coord == highlight:
                draw.rectangle([x0, y0, x1, y1],
                               outline=(255, 215, 0, 255), width=3)
            else:
                draw.rectangle([x0, y0, x1, y1],
                               outline=(255, 255, 255, 130), width=1)

            txt_col = (10, 10, 10, 255) if present else (255, 255, 255, 255)
            draw.text(
                ((x0 + x1) / 2, (y0 + y1) / 2),
                coord, fill=txt_col, font=fnt, anchor="mm",
            )

    # Column headers
    for ci, col in enumerate(COLS):
        x = ml + (ci + 0.5) * cw
        draw.text((x, mt / 2), col,
                  fill=(255, 255, 255, 220), font=fnt_hdr, anchor="mm")

    # Row headers
    for ri, row_num in enumerate(ROWS):
        y = mt + (ri + 0.5) * ch
        draw.text((ml / 2, y), str(row_num),
                  fill=(255, 255, 255, 220), font=fnt_hdr, anchor="mm")

    return Image.alpha_composite(base, overlay).convert("RGB")


def render_plain_grid(coord_dict: dict,
                      cell_px: int = 68, label_px: int = 34) -> Image.Image:
    """Fallback grid when no photo is uploaded."""
    W = label_px + len(COLS) * cell_px + 2
    H = label_px + len(ROWS) * cell_px + 2
    img  = Image.new("RGB", (W, H), "#1A1A2E")
    draw = ImageDraw.Draw(img)
    fnt_lbl  = _try_font(16)
    fnt_cell = _try_font(13)

    for ci, col in enumerate(COLS):
        x = label_px + ci * cell_px + cell_px // 2
        draw.text((x, label_px // 2), col, fill="#ECF0F1", font=fnt_lbl, anchor="mm")

    for ri, row in enumerate(ROWS):
        y_top = label_px + ri * cell_px
        draw.text((label_px // 2, y_top + cell_px // 2),
                  str(row), fill="#ECF0F1", font=fnt_lbl, anchor="mm")
        for ci, col in enumerate(COLS):
            coord   = f"{col}{row}"
            present = coord_dict.get(coord, True)
            x_left  = label_px + ci * cell_px
            fill    = PALETTE["present"] if present else PALETTE["empty"]
            draw.rounded_rectangle(
                [x_left + 2, y_top + 2, x_left + cell_px - 2, y_top + cell_px - 2],
                radius=6, fill=fill, outline="#1A1A2E", width=2,
            )
            txt_col = "#1A1A2E" if present else "#FFFFFF"
            draw.text(
                (x_left + cell_px // 2, y_top + cell_px // 2),
                coord, fill=txt_col, font=fnt_cell, anchor="mm",
            )
    return img


# ── Excel export ───────────────────────────────────────────────────────────────

def _interpolate_color(count: int, max_count: int) -> str:
    t = min(count / max_count, 1.0) if max_count > 0 else 0.0
    r = int(46  + (231 - 46)  * t)
    g = int(204 + (76  - 204) * t)
    b = int(113 + (60  - 113) * t)
    return f"FF{r:02X}{g:02X}{b:02X}"


def build_cumulative_heatmap_workbook(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Cumulative Heatmap"

    thin   = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    header_fill  = PatternFill("solid", fgColor="FF2C3E50")
    summary_fill = PatternFill("solid", fgColor="FF34495E")
    white_bold   = Font(bold=True, color="FFFFFFFF", name="Arial", size=11)
    cell_font    = Font(name="Arial", size=10)
    cell_font_wh = Font(name="Arial", size=10, color="FFFFFFFF")

    total_frames = len(df)

    missing_counts: dict[str, int] = {}
    for coord in ALL_COORDS:
        if coord in df.columns:
            missing_counts[coord] = int((df[coord].astype(str) == "0").sum())
        else:
            missing_counts[coord] = 0

    max_missing = max(missing_counts.values()) if missing_counts else 1

    # Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS) + 2)
    tc = ws.cell(1, 1,
        f"Chocolate Mold — Cumulative Missing Count  ({total_frames} frame(s) inspected)")
    tc.font = Font(bold=True, color="FFFFFFFF", name="Arial", size=13)
    tc.fill = header_fill; tc.alignment = center
    ws.row_dimensions[1].height = 30

    ws.cell(2, 1, "").fill = header_fill
    for ci, col in enumerate(COLS, start=2):
        c = ws.cell(2, ci, col)
        c.font = white_bold; c.fill = header_fill
        c.alignment = center; c.border = border

    tot_col = len(COLS) + 2
    tc2 = ws.cell(2, tot_col, "Row\nMissing")
    tc2.font = white_bold; tc2.fill = summary_fill
    tc2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    tc2.border = border
    ws.row_dimensions[2].height = 30

    ws.column_dimensions["A"].width = 6
    for ci in range(2, len(COLS) + 2):
        ws.column_dimensions[get_column_letter(ci)].width = 10
    ws.column_dimensions[get_column_letter(tot_col)].width = 11

    for ri, row_num in enumerate(ROWS):
        excel_row = ri + 3
        ws.row_dimensions[excel_row].height = 24

        rh = ws.cell(excel_row, 1, str(row_num))
        rh.font = white_bold; rh.fill = header_fill
        rh.alignment = center; rh.border = border

        row_missing_total = 0
        for ci, col in enumerate(COLS, start=2):
            coord = f"{col}{row_num}"
            count = missing_counts.get(coord, 0)
            row_missing_total += count

            argb = _interpolate_color(count, max_missing)
            fill = PatternFill("solid", fgColor=argb)
            t    = count / max_missing if max_missing > 0 else 0
            font = cell_font_wh if t > 0.45 else cell_font

            cell = ws.cell(excel_row, ci, count)
            cell.fill = fill; cell.font = font
            cell.alignment = center; cell.border = border

        row_total_cell = ws.cell(excel_row, tot_col, row_missing_total)
        row_total_cell.fill = summary_fill; row_total_cell.alignment = center
        row_total_cell.border = border
        row_total_cell.font = Font(bold=True, name="Arial", size=10, color="FFFFFFFF")

    totals_row = len(ROWS) + 3
    ws.row_dimensions[totals_row].height = 24

    col_label = ws.cell(totals_row, 1, "Total")
    col_label.font = Font(bold=True, color="FFFFFFFF", name="Arial", size=10)
    col_label.fill = summary_fill; col_label.alignment = center; col_label.border = border

    grand_total = 0
    for ci, col_letter in enumerate(COLS, start=2):
        col_total = sum(missing_counts.get(f"{col_letter}{r}", 0) for r in ROWS)
        grand_total += col_total
        ct = ws.cell(totals_row, ci, col_total)
        ct.fill = summary_fill; ct.alignment = center; ct.border = border
        ct.font = Font(bold=True, name="Arial", size=10, color="FFFFFFFF")

    gt = ws.cell(totals_row, tot_col, grand_total)
    gt.font = Font(bold=True, name="Arial", size=11, color="FFFFFFFF")
    gt.fill = PatternFill("solid", fgColor="FF1A252F")
    gt.alignment = center; gt.border = border

    legend_row = totals_row + 2
    ws.merge_cells(start_row=legend_row, start_column=1,
                   end_row=legend_row, end_column=len(COLS) + 2)
    leg = ws.cell(legend_row, 1,
        f"Each cell = number of frames where that cavity was missing  |  "
        f"Green = never missing  →  Red = missing in all {total_frames} frame(s)  |  "
        f"Total inspected: {total_frames} frames, {len(ALL_COORDS)} positions each")
    leg.font = Font(italic=True, name="Arial", size=9, color="FF555555")
    leg.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[legend_row].height = 18

    # Frame Summary sheet
    ws_sum = wb.create_sheet(title="Frame Summary", index=1)
    sum_headers = ["Frame", "Missing", "Present", "Total", "% Present", "Timestamp"]
    ws_sum.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(sum_headers))
    t = ws_sum.cell(1, 1, "Chocolate Mold Inspection — Frame Summary")
    t.font = Font(bold=True, color="FFFFFFFF", name="Arial", size=14)
    t.fill = header_fill; t.alignment = center
    ws_sum.row_dimensions[1].height = 30

    for ci, h in enumerate(sum_headers, 1):
        c = ws_sum.cell(2, ci, h)
        c.font = white_bold; c.fill = summary_fill
        c.alignment = center; c.border = border
    ws_sum.row_dimensions[2].height = 22

    for ci, w in enumerate([30, 10, 10, 10, 12, 22], 1):
        ws_sum.column_dimensions[get_column_letter(ci)].width = w

    for ri, (_, rec) in enumerate(df.iterrows(), start=3):
        fname   = str(rec.get("frame_name", rec["frame_id"]))
        missing = sum(1 for c in ALL_COORDS if str(rec.get(c, "1")) == "0")
        present = len(ALL_COORDS) - missing
        ts      = str(rec.get("timestamp", ""))
        total   = len(ALL_COORDS)
        row_vals = [fname, missing, present, total, f"=C{ri}/D{ri}", ts]
        for ci, val in enumerate(row_vals, 1):
            c = ws_sum.cell(ri, ci, val)
            c.alignment = center; c.border = border
            c.font = Font(name="Arial", size=10)
            if ci == 2 and total > 0:
                intensity = min(missing / total, 1.0)
                rv = int(231 * intensity + 46  * (1 - intensity))
                gv = int(76  * intensity + 204 * (1 - intensity))
                bv = int(60  * intensity + 113 * (1 - intensity))
                c.fill = PatternFill("solid", fgColor=f"FF{rv:02X}{gv:02X}{bv:02X}")
            if ci == 5:
                c.number_format = "0.0%"
        ws_sum.row_dimensions[ri].height = 20

    total_data_row = len(df) + 3
    ws_sum.cell(total_data_row, 1, "TOTALS").font = Font(bold=True, name="Arial", size=10)
    if not df.empty:
        last_data = total_data_row - 1
        ws_sum.cell(total_data_row, 2, f"=SUM(B3:B{last_data})").font = Font(bold=True, name="Arial")
        ws_sum.cell(total_data_row, 3, f"=SUM(C3:C{last_data})").font = Font(bold=True, name="Arial")
        ws_sum.cell(total_data_row, 4, f"=SUM(D3:D{last_data})").font = Font(bold=True, name="Arial")
        pct = ws_sum.cell(total_data_row, 5, f"=C{total_data_row}/D{total_data_row}")
        pct.number_format = "0.0%"; pct.font = Font(bold=True, name="Arial")
    for ci in range(1, 7):
        ws_sum.cell(total_data_row, ci).border = border
        ws_sum.cell(total_data_row, ci).alignment = center
    ws_sum.row_dimensions[total_data_row].height = 22

    # Coordinate Frequency sheet
    ws_freq = wb.create_sheet(title="Coordinate Frequency", index=2)
    freq_headers = ["Coordinate", "Times Missing", "Total Frames", "% Flagged", "Rank"]
    ws_freq.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(freq_headers))
    ft = ws_freq.cell(1, 1, f"Coordinate Flag Frequency  —  {total_frames} frame(s) analysed")
    ft.font = Font(bold=True, color="FFFFFFFF", name="Arial", size=13)
    ft.fill = header_fill; ft.alignment = center
    ws_freq.row_dimensions[1].height = 30

    for ci, h in enumerate(freq_headers, 1):
        c = ws_freq.cell(2, ci, h)
        c.font = white_bold; c.fill = summary_fill
        c.alignment = center; c.border = border
    ws_freq.row_dimensions[2].height = 22
    for ci, w in enumerate([14, 16, 14, 12, 8], 1):
        ws_freq.column_dimensions[get_column_letter(ci)].width = w

    sorted_coords = sorted(ALL_COORDS,
                           key=lambda coord: (-missing_counts.get(coord, 0), coord))
    for rank, coord in enumerate(sorted_coords, start=1):
        excel_row = rank + 2
        count = missing_counts.get(coord, 0)
        pct   = count / total_frames if total_frames > 0 else 0.0
        argb  = _interpolate_color(count, max_missing if max_missing > 0 else 1)
        flag_fill = PatternFill("solid", fgColor=argb)
        t_val = count / max_missing if max_missing > 0 else 0
        flag_font = Font(name="Arial", size=10,
                         color="FFFFFFFF" if t_val > 0.45 else "FF000000",
                         bold=(count > 0))

        for ci, val in enumerate([coord, count, total_frames, pct, rank], 1):
            cell = ws_freq.cell(excel_row, ci, val)
            cell.alignment = center; cell.border = border
            cell.font = Font(name="Arial", size=10)
            if ci == 4:
                cell.number_format = "0.0%"
                cell.fill = flag_fill; cell.font = flag_font
            elif ci == 2 and count > 0:
                intensity = min(count / total_frames, 1.0)
                rv = int(231 * intensity + 236 * (1 - intensity))
                gv = int(76  * intensity + 240 * (1 - intensity))
                bv = int(60  * intensity + 241 * (1 - intensity))
                cell.fill = PatternFill("solid", fgColor=f"FF{rv:02X}{gv:02X}{bv:02X}")
        ws_freq.row_dimensions[excel_row].height = 18

    footer_row = len(ALL_COORDS) + 3
    ws_freq.merge_cells(start_row=footer_row, start_column=1,
                        end_row=footer_row, end_column=len(freq_headers))
    total_events = sum(missing_counts.values())
    avg_missing  = total_events / total_frames if total_frames > 0 else 0
    foot = ws_freq.cell(footer_row, 1,
        f"Total missing events: {total_events}   |   "
        f"Avg missing per frame: {avg_missing:.1f} / {len(ALL_COORDS)}   |   "
        f"Coordinates never flagged: {sum(1 for v in missing_counts.values() if v == 0)}")
    foot.font = Font(italic=True, name="Arial", size=9, color="FF555555")
    foot.alignment = Alignment(horizontal="left", vertical="center")
    ws_freq.row_dimensions[footer_row].height = 18

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Session-state helpers ───────────────────────────────────────────────────────

def init_state():
    defaults = {
        "df":               load_data(),
        "active_frame_id":  None,
        "coord_dict":       {c: True for c in ALL_COORDS},
        "frame_image":      None,
        "last_toggled":     None,   # coordinate that was most recently clicked
        "_click_bridge":    "",     # hidden text-input value driven by JS
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def load_frame(frame_id: str):
    st.session_state.active_frame_id = frame_id
    st.session_state.coord_dict = get_frame_dict(st.session_state.df, frame_id)
    st.session_state.last_toggled = None
    row = st.session_state.df[st.session_state.df["frame_id"] == frame_id]
    if not row.empty:
        img_path = str(row.iloc[0].get("image_path", ""))
        st.session_state.frame_image = load_frame_image(img_path)
    else:
        st.session_state.frame_image = None


def persist_frame(frame_id: str, frame_name: str):
    """Save current coord_dict to disk immediately."""
    row = st.session_state.df[st.session_state.df["frame_id"] == frame_id]
    image_path = ""
    if not row.empty:
        image_path = str(row.iloc[0].get("image_path", ""))
    st.session_state.df = upsert_frame(
        st.session_state.df, frame_id, frame_name,
        st.session_state.coord_dict, image_path,
    )
    save_data(st.session_state.df)


def navigate_adjacent(direction: int):
    ids = st.session_state.df["frame_id"].tolist()
    if not ids:
        return
    cur = st.session_state.active_frame_id
    idx = ids.index(cur) if cur in ids else 0
    load_frame(ids[(idx + direction) % len(ids)])


# ── Photo Inspector component ──────────────────────────────────────────────────

def photo_inspector(
    photo: Image.Image,
    coord_dict: dict,
    frame_id: str,
    frame_name: str,
    opacity: int = 90,
):
    """
    Renders the mold photo with a clickable overlay grid.

    Click relay works as follows:
    1. JS listens for clicks on the canvas element.
    2. On each click it resolves which coordinate was hit using the same
       geometry formula used by the Python overlay renderer.
    3. It writes the coordinate string into a hidden <input> element that
       Streamlit owns (a text_input rendered with label_visibility='hidden').
    4. Streamlit detects the change through its normal React event pipeline
       and triggers a rerun.  The Python side reads the value, toggles the
       coord, saves, clears the bridge, and reruns cleanly.
    """
    W, H = photo.size
    geo  = grid_geometry(W, H)

    # ── Resolve any pending click from the bridge ──────────────────────────
    bridge_key = f"bridge_{frame_id}"
    bridge_val = st.session_state.get(bridge_key, "")

    if bridge_val and bridge_val in ALL_COORDS:
        coord = bridge_val
        # Toggle
        coord_dict[coord] = not coord_dict.get(coord, True)
        st.session_state.coord_dict  = coord_dict
        st.session_state.last_toggled = coord
        # Clear bridge
        st.session_state[bridge_key] = ""
        # Persist immediately
        persist_frame(frame_id, frame_name)
        st.rerun()

    # ── Controls strip ─────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns([1, 1, 1, 2])
    with c1:
        if st.button("✅ All Present", use_container_width=True, key="pi_allpres"):
            for c in ALL_COORDS:
                st.session_state.coord_dict[c] = True
            st.session_state.last_toggled = None
            persist_frame(frame_id, frame_name)
            st.rerun()
    with c2:
        if st.button("❌ All Missing", use_container_width=True, key="pi_allmiss"):
            for c in ALL_COORDS:
                st.session_state.coord_dict[c] = False
            st.session_state.last_toggled = None
            persist_frame(frame_id, frame_name)
            st.rerun()
    with c3:
        if st.button("🔄 Invert", use_container_width=True, key="pi_invert"):
            for c in ALL_COORDS:
                st.session_state.coord_dict[c] = not st.session_state.coord_dict.get(c, True)
            st.session_state.last_toggled = None
            persist_frame(frame_id, frame_name)
            st.rerun()
    with c4:
        opacity = st.slider("Overlay opacity", 40, 200, opacity, 10,
                             key="pi_opacity", label_visibility="visible")

    # Missing badges
    missing_list = sorted(c for c, v in coord_dict.items() if not v)
    if missing_list:
        badges = " ".join(
            f"<span style='background:#E74C3C;color:#FFF;"
            f"padding:1px 8px;border-radius:4px;font-size:0.78rem;"
            f"margin:2px;display:inline-block'>{coord}</span>"
            for coord in missing_list
        )
        st.markdown(
            f"<div style='line-height:2.1;padding:4px 0'>"
            f"<strong style='color:#E74C3C'>Missing ({len(missing_list)}):</strong> "
            f"{badges}</div>",
            unsafe_allow_html=True,
        )
    else:
        st.success("🎉 All 120 positions present!")

    st.markdown("<hr style='margin:6px 0'>", unsafe_allow_html=True)

    # ── Render overlay image ───────────────────────────────────────────────
    last_toggled = st.session_state.get("last_toggled")
    overlay_img  = render_overlay(photo, coord_dict,
                                  opacity=opacity, highlight=last_toggled)
    b64 = pil_to_b64(overlay_img, fmt="JPEG")

    # ── Hidden bridge input ────────────────────────────────────────────────
    # Rendered BEFORE the canvas so it exists in the DOM when JS runs.
    # We give it a unique key so Streamlit tracks it independently per frame.
    st.text_input(
        "click_bridge",
        value=st.session_state.get(bridge_key, ""),
        key=bridge_key,
        label_visibility="collapsed",
    )

    # ── Interactive canvas ─────────────────────────────────────────────────
    canvas_id = f"canvas_{re.sub(r'[^a-zA-Z0-9]', '_', frame_id)}"
    cols_json = json.dumps(COLS)
    rows_json = json.dumps([str(r) for r in ROWS])
    state_json = json.dumps(
        {c: (1 if coord_dict.get(c, True) else 0) for c in ALL_COORDS}
    )
    last_toggled_json = json.dumps(last_toggled or "")

    canvas_html = f"""
<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}

#pi_wrap {{
  position: relative;
  width: 100%;
  line-height: 0;
}}

#{canvas_id} {{
  width: 100%;
  height: auto;
  display: block;
  border-radius: 8px;
  cursor: crosshair;
  box-shadow: 0 4px 20px rgba(0,0,0,0.5);
  transition: box-shadow 0.15s;
}}

#{canvas_id}:hover {{
  box-shadow: 0 6px 28px rgba(0,0,0,0.6);
}}

#pi_tooltip {{
  position: fixed;
  background: rgba(10,10,10,0.90);
  color: #fff;
  padding: 6px 14px;
  border-radius: 7px;
  font: 700 13px/1.5 -apple-system, Arial, sans-serif;
  pointer-events: none;
  display: none;
  z-index: 9999;
  white-space: nowrap;
  border: 1px solid rgba(255,255,255,0.15);
  letter-spacing: 0.3px;
}}

#pi_flash {{
  position: fixed;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%) scale(1);
  background: rgba(0,0,0,0.82);
  color: #fff;
  padding: 12px 28px;
  border-radius: 12px;
  font: 700 17px -apple-system, Arial, sans-serif;
  pointer-events: none;
  display: none;
  z-index: 10000;
  text-align: center;
  transition: opacity 0.3s;
}}

#pi_status {{
  font: 13px -apple-system, Arial, sans-serif;
  color: #7F8C8D;
  text-align: center;
  padding: 6px 0 0;
  min-height: 22px;
}}
</style>

<div id="pi_wrap">
  <canvas id="{canvas_id}"></canvas>
  <div id="pi_tooltip"></div>
  <div id="pi_flash"></div>
</div>
<div id="pi_status">
  👆 Click any cavity on the photo to mark it <strong>missing</strong> or restore it as <strong>present</strong>. Changes save instantly.
</div>

<script>
(function () {{
  /* ── Config ───────────────────────────────────────────────────────── */
  const COLS       = {cols_json};
  const ROWS       = {rows_json};
  const state      = {state_json};    // 1 = present, 0 = missing
  const MARGIN_L   = {geo['margin_l']};
  const MARGIN_T   = {geo['margin_t']};
  const CELL_W     = {geo['cell_w']};
  const CELL_H     = {geo['cell_h']};
  const IMG_W      = {W};
  const IMG_H      = {H};
  const CANVAS_ID  = "{canvas_id}";
  const BRIDGE_KEY = "{bridge_key}";
  const LAST_HIT   = {last_toggled_json};  // coordinate highlighted from last rerun

  /* ── DOM refs ─────────────────────────────────────────────────────── */
  const canvas  = document.getElementById(CANVAS_ID);
  const ctx     = canvas.getContext("2d");
  const tooltip = document.getElementById("pi_tooltip");
  const flash   = document.getElementById("pi_flash");
  const status  = document.getElementById("pi_status");

  canvas.width  = IMG_W;
  canvas.height = IMG_H;

  /* ── Load base image ──────────────────────────────────────────────── */
  const baseImg = new Image();
  baseImg.src = "data:image/jpeg;base64,{b64}";
  baseImg.onload = () => ctx.drawImage(baseImg, 0, 0);

  /* ── Geometry helpers ─────────────────────────────────────────────── */
  function getScale() {{
    const r = canvas.getBoundingClientRect();
    return {{ sx: IMG_W / r.width, sy: IMG_H / r.height, rect: r }};
  }}

  function coordFromPixel(px, py) {{
    const ci = Math.floor((px - MARGIN_L) / CELL_W);
    const ri = Math.floor((py - MARGIN_T) / CELL_H);
    if (ci < 0 || ci >= COLS.length || ri < 0 || ri >= ROWS.length) return null;
    return COLS[ci] + ROWS[ri];
  }}

  function cellRect(coord) {{
    const ci = COLS.indexOf(coord[0]);
    const ri = ROWS.indexOf(coord.slice(1));
    return {{
      x0: MARGIN_L + ci * CELL_W,
      y0: MARGIN_T + ri * CELL_H,
      x1: MARGIN_L + (ci + 1) * CELL_W,
      y1: MARGIN_T + (ri + 1) * CELL_H,
    }};
  }}

  /* ── Repaint a single cell (instant local feedback) ──────────────── */
  function repaintCell(coord, isPresent, isHighlighted) {{
    const {{ x0, y0, x1, y1 }} = cellRect(coord);
    const cw = x1 - x0, ch = y1 - y0;

    // Restore background slice from base image
    ctx.drawImage(baseImg, x0, y0, cw, ch, x0, y0, cw, ch);

    // Overlay fill
    ctx.fillStyle = isPresent
      ? "rgba(46,204,113,0.60)"
      : "rgba(231,76,60,0.82)";
    ctx.fillRect(x0 + 1, y0 + 1, cw - 2, ch - 2);

    // Border
    ctx.strokeStyle = isHighlighted
      ? "rgba(255,215,0,1.0)"
      : "rgba(255,255,255,0.65)";
    ctx.lineWidth = isHighlighted ? 3 : 1.5;
    ctx.strokeRect(x0 + 0.75, y0 + 0.75, cw - 1.5, ch - 1.5);

    // Label
    const fontSize = Math.max(8, Math.floor(Math.min(cw, ch) * 0.28));
    ctx.fillStyle = isPresent ? "rgba(10,10,10,0.9)" : "#fff";
    ctx.font      = `bold ${{fontSize}}px Arial, sans-serif`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(coord, x0 + cw / 2, y0 + ch / 2);
  }}

  /* ── Show the flash notification ──────────────────────────────────── */
  let flashTimer = null;
  function showFlash(coord, nowMissing) {{
    flash.textContent = nowMissing
      ? "🔴 Marked MISSING: " + coord
      : "🟢 Restored PRESENT: " + coord;
    flash.style.display = "block";
    flash.style.opacity = "1";
    if (flashTimer) clearTimeout(flashTimer);
    flashTimer = setTimeout(() => {{
      flash.style.opacity = "0";
      setTimeout(() => flash.style.display = "none", 300);
    }}, 900);
  }}

  /* ── Write to Streamlit bridge input ──────────────────────────────── */
  function relayToStreamlit(coord) {{
    // Find the hidden text input Streamlit rendered for our bridge key.
    // Streamlit wraps text_inputs in a div with a data-testid attribute;
    // we locate the actual <input> by searching for one whose current
    // aria-label or data-* attribute contains our key name.
    // Fallback: scan all text inputs and use the last one whose value we
    // can recognise as our bridge (empty → ready to receive).
    let input = null;

    // Strategy 1: aria-label set by label_visibility="collapsed" to the key name
    const allInputs = window.parent.document.querySelectorAll("input[type='text']");
    for (const inp of allInputs) {{
      const label = inp.getAttribute("aria-label") || "";
      if (label === BRIDGE_KEY || label === "click_bridge") {{
        input = inp; break;
      }}
    }}

    // Strategy 2: the input whose value is currently empty (our hidden bridge)
    if (!input) {{
      for (const inp of allInputs) {{
        if (inp.value === "" && inp.style.display !== "none") {{
          input = inp; break;
        }}
      }}
    }}

    if (!input) {{
      // Last resort: try query param approach so at least something happens
      try {{
        const url = new URL(window.parent.location.href);
        url.searchParams.set("click_coord", coord);
        window.parent.history.replaceState({{}}, "", url.toString());
      }} catch(e) {{}}
      return false;
    }}

    // Use React's internal value setter to bypass its synthetic event system
    const nativeInputSetter = Object.getOwnPropertyDescriptor(
      window.HTMLInputElement.prototype, "value"
    ).set;
    nativeInputSetter.call(input, coord);
    input.dispatchEvent(new Event("input", {{ bubbles: true }}));
    return true;
  }}

  /* ── Tooltip on hover ─────────────────────────────────────────────── */
  canvas.addEventListener("mousemove", (e) => {{
    const {{ sx, sy, rect }} = getScale();
    const px = (e.clientX - rect.left)  * sx;
    const py = (e.clientY - rect.top)   * sy;
    const coord = coordFromPixel(px, py);
    if (coord) {{
      const s = state[coord] === 1 ? "🟢 Present" : "🔴 Missing";
      tooltip.textContent = `${{coord}} — ${{s}}  (click to toggle)`;
      tooltip.style.display = "block";
      tooltip.style.left    = (e.clientX + 18) + "px";
      tooltip.style.top     = (e.clientY - 14) + "px";
    }} else {{
      tooltip.style.display = "none";
    }}
  }});
  canvas.addEventListener("mouseleave", () => tooltip.style.display = "none");

  /* ── Click handler ────────────────────────────────────────────────── */
  canvas.addEventListener("click", (e) => {{
    const {{ sx, sy, rect }} = getScale();
    const px = (e.clientX - rect.left) * sx;
    const py = (e.clientY - rect.top)  * sy;
    const coord = coordFromPixel(px, py);
    if (!coord) return;

    // Optimistic toggle in local state
    state[coord] = state[coord] === 1 ? 0 : 1;
    const nowMissing = state[coord] === 0;

    // Instant visual feedback — repaint just the clicked cell
    repaintCell(coord, !nowMissing, true);
    tooltip.style.display = "none";

    // Update status bar
    status.innerHTML = nowMissing
      ? `<span style="color:#E74C3C">🔴 Marked <strong>MISSING</strong>: ${{coord}} — saving…</span>`
      : `<span style="color:#2ECC71">🟢 Restored <strong>PRESENT</strong>: ${{coord}} — saving…</span>`;

    showFlash(coord, nowMissing);

    // Relay coordinate to Streamlit so Python can persist the change
    const ok = relayToStreamlit(coord);
    if (!ok) {{
      status.innerHTML += " <span style='color:#F39C12'>(⚠️ relay failed — try the Grid Editor tab)</span>";
    }}
  }});

  /* ── On first load: highlight last-toggled coord if set ───────────── */
  if (LAST_HIT && LAST_HIT in state) {{
    // Re-highlight from the previous rerun without changing state
    setTimeout(() => repaintCell(LAST_HIT, state[LAST_HIT] === 1, true), 120);
  }}

}})();
</script>
"""

    # Height scales with photo aspect ratio, capped reasonably
    display_h = min(int(H / W * 820), 900)
    components.html(canvas_html, height=display_h + 120, scrolling=False)

    # Download overlay
    buf = io.BytesIO()
    overlay_img.save(buf, format="PNG")
    st.download_button(
        "⬇️ Download overlay image",
        data=buf.getvalue(),
        file_name=f"{re.sub(r'[^a-zA-Z0-9_-]', '_', frame_name)}_overlay.png",
        mime="image/png",
        key=f"dl_overlay_{frame_id}",
    )


# ── Main app ───────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="Chocolate Mold Inspector",
        page_icon="🍫",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    init_state()

    st.markdown("""
    <style>
    [data-testid="stSidebar"] { background: #1A1A2E; }
    h1, h2, h3 { color: #ECF0F1; }
    .stat-box {
        background: #16213E;
        border-radius: 10px;
        padding: 14px 18px;
        text-align: center;
        border: 1px solid #2C3E50;
        margin-bottom: 4px;
    }
    .stat-number { font-size: 1.9rem; font-weight: 700; }
    .stat-label  { font-size: 0.82rem; color: #95A5A6; margin-top: 3px; }
    .green { color: #2ECC71; }
    .red   { color: #E74C3C; }
    .stButton > button { border-radius: 8px; font-weight: 600; }
    .upload-hint {
        background: #16213E;
        border: 1px dashed #3498DB;
        border-radius: 8px;
        padding: 12px 16px;
        font-size: 0.88rem;
        color: #BDC3C7;
        margin-bottom: 8px;
    }
    /* Hide the click bridge input visually but keep it in DOM */
    [data-testid="stTextInput"]:has(input[aria-label="click_bridge"]) {
        height: 0 !important;
        overflow: hidden !important;
        padding: 0 !important;
        margin: 0 !important;
        opacity: 0 !important;
        pointer-events: none !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Sidebar ────────────────────────────────────────────────────────────────
    with st.sidebar:
        st.title("🍫 Mold Inspector")
        st.caption("15 × 8 grid  |  120 positions")
        st.divider()

        st.subheader("📸 Upload Frame Photos")
        st.markdown(
            f'<div class="upload-hint">Upload up to {MAX_UPLOAD_FILES} frame images at once. '
            "A new record is created for each photo automatically.</div>",
            unsafe_allow_html=True,
        )

        uploaded_photos = st.file_uploader(
            "Drop frame images here",
            type=["png", "jpg", "jpeg", "bmp", "tiff", "webp"],
            accept_multiple_files=True,
            label_visibility="collapsed",
            key="batch_uploader",
        )

        if uploaded_photos:
            if len(uploaded_photos) > MAX_UPLOAD_FILES:
                st.warning(f"Only the first {MAX_UPLOAD_FILES} images will be imported.")
                uploaded_photos = uploaded_photos[:MAX_UPLOAD_FILES]

            new_count = 0
            existing_names = (
                st.session_state.df["frame_name"].tolist()
                if not st.session_state.df.empty
                else []
            )
            for uf in uploaded_photos:
                base_name = os.path.splitext(uf.name)[0]
                if base_name in existing_names:
                    continue
                fid   = f"frame_{int(time.time() * 1000)}_{new_count}"
                ipath = save_frame_image(fid, uf)
                fresh = {c: True for c in ALL_COORDS}
                st.session_state.df = upsert_frame(
                    st.session_state.df, fid, base_name, fresh, ipath
                )
                new_count += 1

            if new_count:
                save_data(st.session_state.df)
                st.success(f"Imported {new_count} new frame(s).")
                last_id = st.session_state.df["frame_id"].iloc[-1]
                load_frame(last_id)
                st.rerun()

        st.divider()

        st.subheader("📋 Select Frame")
        df = st.session_state.df

        if df.empty:
            st.info("No frames yet — upload photos above to get started.")
        else:
            frame_names = df["frame_name"].tolist()
            frame_ids   = df["frame_id"].tolist()

            current_idx = 0
            if st.session_state.active_frame_id in frame_ids:
                current_idx = frame_ids.index(st.session_state.active_frame_id)

            selected_idx = st.selectbox(
                "Frame",
                options=range(len(frame_names)),
                format_func=lambda i: frame_names[i],
                index=current_idx,
                label_visibility="collapsed",
            )

            col_load, col_del = st.columns(2)
            if col_load.button("Load", use_container_width=True):
                load_frame(frame_ids[selected_idx])
                st.rerun()
            if col_del.button("🗑 Delete", use_container_width=True):
                fid_del = frame_ids[selected_idx]
                row = st.session_state.df[st.session_state.df["frame_id"] == fid_del]
                if not row.empty:
                    ip = str(row.iloc[0].get("image_path", ""))
                    if ip and os.path.exists(ip):
                        os.remove(ip)
                st.session_state.df = st.session_state.df[
                    st.session_state.df["frame_id"] != fid_del
                ].reset_index(drop=True)
                save_data(st.session_state.df)
                if st.session_state.active_frame_id == fid_del:
                    st.session_state.active_frame_id = None
                    st.session_state.frame_image     = None
                st.rerun()

            st.caption(f"Total frames: **{len(df)}**")

        st.divider()

        st.subheader("📥 Export")
        df = st.session_state.df
        if st.button("⬇️ Download Cumulative Heatmap (.xlsx)",
                     use_container_width=True, disabled=df.empty):
            if df.empty:
                st.warning("No data to export.")
            else:
                xlsx_bytes = build_cumulative_heatmap_workbook(df)
                now        = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    "💾 Save file",
                    data=xlsx_bytes,
                    file_name=f"mold_cumulative_heatmap_{now}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        if not df.empty:
            st.download_button(
                "⬇️ Backup CSV",
                data=df.to_csv(index=False).encode(),
                file_name="mold_data_backup.csv",
                mime="text/csv",
                use_container_width=True,
            )

        st.divider()
        st.subheader("⬆️ Restore from CSV")
        uploaded_csv = st.file_uploader(
            "Upload backup CSV", type="csv", label_visibility="collapsed"
        )
        if uploaded_csv:
            restored = pd.read_csv(uploaded_csv, dtype=str)
            for coord in ALL_COORDS:
                if coord not in restored.columns:
                    restored[coord] = "1"
            if "image_path" not in restored.columns:
                restored["image_path"] = ""
            st.session_state.df = restored
            save_data(restored)
            st.success("Restored!")
            st.rerun()

    # ── Main panel ─────────────────────────────────────────────────────────────
    df = st.session_state.df

    if st.session_state.active_frame_id is None:
        st.markdown("## 🍫 Chocolate Mold Inspector")
        st.info("Upload frame photos in the sidebar, then click **Load** to begin inspection.")
        if not df.empty:
            st.subheader("All Frames — Overview")
            summary = []
            for _, rec in df.iterrows():
                missing = sum(1 for c in ALL_COORDS if str(rec.get(c, "1")) != "1")
                has_img = "✅" if str(rec.get("image_path", "")) else "—"
                summary.append({
                    "Frame":        rec["frame_name"],
                    "Photo":        has_img,
                    "Missing":      missing,
                    "Present":      len(ALL_COORDS) - missing,
                    "% Present":    f"{(len(ALL_COORDS) - missing) / len(ALL_COORDS) * 100:.1f}%",
                    "Last Updated": rec.get("timestamp", ""),
                })
            st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)
        return

    # ── Active frame ───────────────────────────────────────────────────────────
    active_id  = st.session_state.active_frame_id
    df         = st.session_state.df
    active_row = df[df["frame_id"] == active_id].iloc[0]
    frame_name = active_row["frame_name"]
    image_path = str(active_row.get("image_path", ""))
    coord_dict = st.session_state.coord_dict
    photo      = st.session_state.frame_image

    frame_ids    = df["frame_id"].tolist()
    current_idx  = frame_ids.index(active_id) if active_id in frame_ids else 0
    total_frames = len(frame_ids)

    missing_list  = [c for c, v in coord_dict.items() if not v]
    present_count = len(ALL_COORDS) - len(missing_list)

    # ── Header + navigation ────────────────────────────────────────────────────
    h_left, h_right = st.columns([3, 1])
    with h_left:
        st.markdown(f"## 🍫 {frame_name}")
        st.caption(
            f"Frame {current_idx + 1} of {total_frames}  |  "
            f"ID: `{active_id}`  |  Saved: {active_row.get('timestamp', '—')}"
        )
    with h_right:
        nav1, nav2 = st.columns(2)
        with nav1:
            if st.button("◀ Prev", use_container_width=True,
                         disabled=(total_frames <= 1)):
                navigate_adjacent(-1)
                st.rerun()
        with nav2:
            if st.button("Next ▶", use_container_width=True,
                         disabled=(total_frames <= 1), type="primary"):
                navigate_adjacent(+1)
                st.rerun()

    # Stats
    col_a, col_b, col_c, col_d = st.columns(4)
    with col_a:
        st.markdown(
            f'<div class="stat-box"><div class="stat-number green">{present_count}</div>'
            f'<div class="stat-label">Present</div></div>', unsafe_allow_html=True)
    with col_b:
        st.markdown(
            f'<div class="stat-box"><div class="stat-number red">{len(missing_list)}</div>'
            f'<div class="stat-label">Missing</div></div>', unsafe_allow_html=True)
    with col_c:
        pct = present_count / len(ALL_COORDS) * 100
        st.markdown(
            f'<div class="stat-box"><div class="stat-number">{pct:.1f}%</div>'
            f'<div class="stat-label">Fill Rate</div></div>', unsafe_allow_html=True)
    with col_d:
        st.markdown(
            f'<div class="stat-box"><div class="stat-number">{len(ALL_COORDS)}</div>'
            f'<div class="stat-label">Total Positions</div></div>', unsafe_allow_html=True)

    st.divider()

    # ── Per-frame photo upload ─────────────────────────────────────────────────
    with st.expander(
        "📷 " + ("Replace frame photo" if photo else "Attach a photo to this frame"),
        expanded=(photo is None),
    ):
        per_frame_upload = st.file_uploader(
            "Upload mold photo",
            type=["png", "jpg", "jpeg", "bmp", "tiff", "webp"],
            key=f"photo_{active_id}",
            label_visibility="collapsed",
        )
        if per_frame_upload:
            ipath = save_frame_image(active_id, per_frame_upload)
            photo = Image.open(per_frame_upload).convert("RGB")
            st.session_state.frame_image = photo
            st.session_state.df = upsert_frame(
                st.session_state.df, active_id, frame_name, coord_dict, ipath
            )
            save_data(st.session_state.df)
            st.success("Photo saved to this frame.")
            st.rerun()

    # ── Tabs ──────────────────────────────────────────────────────────────────
    if photo:
        tab_photo, tab_grid, tab_quick, tab_vis = st.tabs([
            "🖼️ Photo Inspector", "🔲 Grid Editor", "⌨️ Quick Entry", "📊 Overlay Preview",
        ])
    else:
        tab_photo = None
        tab_grid, tab_quick, tab_vis = st.tabs([
            "🔲 Grid Editor", "⌨️ Quick Entry", "📊 Grid Preview",
        ])

    # ── Photo Inspector tab ────────────────────────────────────────────────────
    if tab_photo and photo:
        with tab_photo:
            photo_inspector(
                photo=photo,
                coord_dict=st.session_state.coord_dict,
                frame_id=active_id,
                frame_name=frame_name,
                opacity=st.session_state.get("pi_opacity", 90),
            )

    # ── Grid Editor tab ───────────────────────────────────────────────────────
    with tab_grid:
        st.markdown(
            "Click a cell to toggle **Present 🟢** ↔ **Missing 🔴**. "
            "Use **Save Frame** to persist changes from this tab."
        )

        tb1, tb2, tb3, tb4 = st.columns(4)
        with tb1:
            if st.button("✅ All Present", use_container_width=True, key="g_all_pres"):
                for c in ALL_COORDS:
                    st.session_state.coord_dict[c] = True
                st.session_state["_grid_dirty"] = True
                st.rerun()
        with tb2:
            if st.button("❌ All Missing", use_container_width=True, key="g_all_miss"):
                for c in ALL_COORDS:
                    st.session_state.coord_dict[c] = False
                st.session_state["_grid_dirty"] = True
                st.rerun()
        with tb3:
            if st.button("🔄 Invert", use_container_width=True, key="g_invert"):
                for c in ALL_COORDS:
                    st.session_state.coord_dict[c] = not st.session_state.coord_dict[c]
                st.session_state["_grid_dirty"] = True
                st.rerun()
        with tb4:
            dirty = st.session_state.get("_grid_dirty", False)
            if st.button("💾 Save Frame", use_container_width=True,
                         type="primary", key="g_save", disabled=not dirty):
                persist_frame(active_id, frame_name)
                st.session_state["_grid_dirty"] = False
                st.success("Saved!")
                st.rerun()

        # Column headers
        hcols = st.columns([0.5] + [1] * len(COLS))
        hcols[0].markdown("**↓**")
        for i, lbl in enumerate(COLS):
            hcols[i + 1].markdown(
                f"<div style='text-align:center;font-weight:700;"
                f"color:#3498DB;font-size:0.8rem'>{lbl}</div>",
                unsafe_allow_html=True,
            )

        for row_num in ROWS:
            rcols = st.columns([0.5] + [1] * len(COLS))
            rcols[0].markdown(
                f"<div style='text-align:center;font-weight:700;color:#3498DB'>{row_num}</div>",
                unsafe_allow_html=True,
            )
            for ci, col_lbl in enumerate(COLS):
                coord   = f"{col_lbl}{row_num}"
                present = st.session_state.coord_dict.get(coord, True)
                with rcols[ci + 1]:
                    if st.button(
                        "🟢" if present else "🔴",
                        key=f"g_{coord}",
                        help=f"{coord}: {'Present' if present else 'MISSING'}",
                        use_container_width=True,
                    ):
                        st.session_state.coord_dict[coord] = not present
                        st.session_state["_grid_dirty"] = True
                        st.rerun()

        if st.session_state.get("_grid_dirty", False):
            st.warning("⚠️ Unsaved changes — click **Save Frame** above.")

    # ── Quick Entry tab ───────────────────────────────────────────────────────
    with tab_quick:
        st.markdown("""
        Paste **missing** coordinates as a comma-separated list.
        All others are assumed **present**.  
        **Examples:** `A1, C3, O8`  or  `B2 D5 F7`
        """)

        missing_list = [c for c, v in st.session_state.coord_dict.items() if not v]
        if missing_list:
            st.markdown(
                f"**Currently missing ({len(missing_list)}):** "
                + ", ".join(f"`{c}`" for c in sorted(missing_list))
            )
        else:
            st.success("All positions currently marked as present.")

        with st.form("quick_entry"):
            raw = st.text_area(
                "Missing coordinates",
                value=", ".join(sorted(missing_list)) if missing_list else "",
                height=90,
                placeholder="e.g. A1, B3, G5, O8",
            )
            ca, cb = st.columns(2)
            apply_btn = ca.form_submit_button(
                "Apply (replace all missing)", use_container_width=True, type="primary"
            )
            add_btn = cb.form_submit_button(
                "Add to existing missing", use_container_width=True
            )

        if apply_btn or add_btn:
            tokens = re.split(r"[\s,;]+", raw.strip().upper())
            valid, invalid = [], []
            for t in tokens:
                if not t:
                    continue
                (valid if t in ALL_COORDS else invalid).append(t)
            if invalid:
                st.error(f"Invalid coordinate(s): {', '.join(invalid)}")
            else:
                if apply_btn:
                    for c in ALL_COORDS:
                        st.session_state.coord_dict[c] = c not in valid
                else:
                    for c in valid:
                        st.session_state.coord_dict[c] = False
                persist_frame(active_id, frame_name)
                n = len([c for c, v in st.session_state.coord_dict.items() if not v])
                st.success(f"Saved! {n} position(s) missing.")
                st.rerun()

        st.subheader("Row-by-row status")
        status_rows = []
        for row_num in ROWS:
            rm = [
                f"{cl}{row_num}"
                for cl in COLS
                if not st.session_state.coord_dict.get(f"{cl}{row_num}", True)
            ]
            status_rows.append({
                "Row": row_num,
                "Missing": len(rm),
                "Missing Coords": ", ".join(rm) if rm else "—",
            })
        st.dataframe(pd.DataFrame(status_rows), use_container_width=True, hide_index=True)

    # ── Overlay Preview tab ───────────────────────────────────────────────────
    with tab_vis:
        if photo:
            st.markdown("Overlay rendered on actual mold photo. 🟢 Present &nbsp; 🔴 Missing")
            opa = st.slider("Overlay opacity", 40, 200, 90, 10, key="vis_opacity")
            ov  = render_overlay(photo, st.session_state.coord_dict, opacity=opa)
            st.image(ov, use_container_width=True, caption=frame_name)
            buf = io.BytesIO()
            ov.save(buf, format="PNG")
            st.download_button(
                "⬇️ Download overlay (.png)",
                data=buf.getvalue(),
                file_name=f"{re.sub(r'[^a-zA-Z0-9_-]', '_', frame_name)}_overlay.png",
                mime="image/png",
            )
        else:
            st.markdown(
                "Synthetic grid (upload a photo to see the real mold). "
                "🟢 Present &nbsp; 🔴 Missing"
            )
            grid_img = render_plain_grid(st.session_state.coord_dict)
            st.image(grid_img, use_container_width=True, caption=frame_name)
            buf = io.BytesIO()
            grid_img.save(buf, format="PNG")
            st.download_button(
                "⬇️ Download grid (.png)",
                data=buf.getvalue(),
                file_name=f"{re.sub(r'[^a-zA-Z0-9_-]', '_', frame_name)}_grid.png",
                mime="image/png",
            )

        missing_list = [c for c, v in st.session_state.coord_dict.items() if not v]
        if missing_list:
            st.subheader(f"🔴 Missing positions ({len(missing_list)})")
            rows_of_missing: dict[str, list] = {}
            for c in sorted(missing_list):
                rows_of_missing.setdefault(c[1:], []).append(c)
            for rn, coords_in_row in sorted(rows_of_missing.items(),
                                            key=lambda x: int(x[0])):
                st.markdown(
                    f"**Row {rn}:** "
                    + " ".join(
                        f"<span style='background:#E74C3C;color:#FFF;"
                        f"padding:2px 6px;border-radius:4px;font-size:0.85rem'>{c}</span>"
                        for c in coords_in_row
                    ),
                    unsafe_allow_html=True,
                )
        else:
            st.success("🎉 All 120 positions are present!")

    # ── Bottom navigation ─────────────────────────────────────────────────────
    st.divider()
    bot_l, bot_m, bot_r = st.columns([1, 2, 1])
    with bot_l:
        if st.button("◀ Previous Frame", use_container_width=True,
                     disabled=(total_frames <= 1)):
            navigate_adjacent(-1)
            st.rerun()
    with bot_m:
        st.markdown(
            f"<div style='text-align:center;color:#95A5A6;padding-top:8px'>"
            f"Frame <strong>{current_idx + 1}</strong> of <strong>{total_frames}</strong></div>",
            unsafe_allow_html=True,
        )
    with bot_r:
        if st.button("Next Frame ▶", use_container_width=True,
                     disabled=(total_frames <= 1), type="primary"):
            navigate_adjacent(+1)
            st.rerun()


if __name__ == "__main__":
    main()
