"""
Chocolate Mold Inspector — Streamlit App
15 columns (A–O) × 8 rows (1–8) = 120 coordinates per mold frame
Upload frame photos, click coordinates to mark them missing,
and export a cumulative Excel heatmap across all frames.
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
COLS       = list("ABCDEFGHIJKLMNO")          # 15 columns  (A–O)
ROWS       = list(range(1, 9))                 # 8 rows       (1–8)
ALL_COORDS = [f"{c}{r}" for r in ROWS for c in COLS]

DATA_FILE   = "mold_data.csv"
IMAGES_DIR  = "frame_images"

MAX_UPLOAD_FILES = 50   # sensible batch limit

PALETTE = {
    "present":  "#2ECC71",
    "empty":    "#E74C3C",
    "selected": "#F39C12",
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
    return pd.DataFrame(columns=["frame_id", "frame_name", "timestamp", "image_path"] + ALL_COORDS)


def save_data(df: pd.DataFrame) -> None:
    df.to_csv(DATA_FILE, index=False)


def get_frame_dict(df: pd.DataFrame, frame_id: str) -> dict:
    row = df[df["frame_id"] == frame_id]
    if row.empty:
        return {c: True for c in ALL_COORDS}
    r = row.iloc[0]
    return {c: (str(r.get(c, "1")) == "1") for c in ALL_COORDS}


def upsert_frame(df: pd.DataFrame, frame_id: str, frame_name: str,
                 coord_dict: dict, image_path: str = "") -> pd.DataFrame:
    ts       = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row_data = {"frame_id": frame_id, "frame_name": frame_name,
                "timestamp": ts, "image_path": image_path}
    row_data.update({c: ("1" if v else "0") for c, v in coord_dict.items()})
    if frame_id in df["frame_id"].values:
        for k, v in row_data.items():
            df.loc[df["frame_id"] == frame_id, k] = v
    else:
        df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
    return df


def save_frame_image(frame_id: str, uploaded_file) -> str:
    """Persist an uploaded image to disk and return its path."""
    ext  = os.path.splitext(uploaded_file.name)[-1].lower() or ".png"
    path = os.path.join(IMAGES_DIR, f"{frame_id}{ext}")
    with open(path, "wb") as f:
        f.write(uploaded_file.getvalue())
    return path


def load_frame_image(image_path: str):
    """Return a PIL Image for the stored path, or None."""
    if image_path and os.path.exists(image_path):
        return Image.open(image_path).convert("RGB")
    return None


# ── Image helpers ───────────────────────────────────────────────────────────────

def pil_to_b64(img: Image.Image, fmt="PNG") -> str:
    buf = io.BytesIO()
    img.save(buf, format=fmt)
    return base64.b64encode(buf.getvalue()).decode()


def render_overlay_on_photo(photo: Image.Image, coord_dict: dict,
                             opacity: int = 60) -> Image.Image:
    """Draw a semi-transparent 15×8 grid overlay on the uploaded mold photo."""
    base = photo.convert("RGBA")
    W, H = base.size

    margin_l = max(24, int(W * 0.038))
    margin_t = max(20, int(H * 0.055))
    grid_w   = W - margin_l - max(4, int(W * 0.008))
    grid_h   = H - margin_t - max(4, int(H * 0.008))

    cell_w = grid_w  / len(COLS)
    cell_h = grid_h  / len(ROWS)

    overlay = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw    = ImageDraw.Draw(overlay)

    font_size = max(9, int(min(cell_w, cell_h) * 0.30))
    try:
        fnt = ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", font_size)
        fnt_hdr = ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            max(10, int(min(margin_l, margin_t) * 0.55)))
    except Exception:
        fnt = fnt_hdr = ImageFont.load_default()

    for ri, row_num in enumerate(ROWS):
        for ci, col in enumerate(COLS):
            coord   = f"{col}{row_num}"
            present = coord_dict.get(coord, True)
            x0 = margin_l + ci * cell_w
            y0 = margin_t + ri * cell_h
            x1 = x0 + cell_w
            y1 = y0 + cell_h

            fill = (46, 204, 113, opacity) if present else (231, 76, 60, min(opacity + 40, 255))
            draw.rectangle([x0 + 1, y0 + 1, x1 - 1, y1 - 1], fill=fill)
            draw.rectangle([x0, y0, x1, y1], outline=(255, 255, 255, 180), width=1)

            txt_col = (10, 10, 10, 255) if present else (255, 255, 255, 255)
            draw.text(((x0 + x1) / 2, (y0 + y1) / 2),
                      coord, fill=txt_col, font=fnt, anchor="mm")

    # Column headers
    for ci, col in enumerate(COLS):
        x = margin_l + (ci + 0.5) * cell_w
        draw.text((x, margin_t / 2), col,
                  fill=(255, 255, 255, 230), font=fnt_hdr, anchor="mm")
    # Row headers
    for ri, row_num in enumerate(ROWS):
        y = margin_t + (ri + 0.5) * cell_h
        draw.text((margin_l / 2, y), str(row_num),
                  fill=(255, 255, 255, 230), font=fnt_hdr, anchor="mm")

    return Image.alpha_composite(base, overlay).convert("RGB")


def render_plain_grid(coord_dict: dict,
                      cell_px: int = 68, label_px: int = 34) -> Image.Image:
    """Fallback grid when no photo is uploaded."""
    W = label_px + len(COLS) * cell_px + 2
    H = label_px + len(ROWS) * cell_px + 2
    img  = Image.new("RGB", (W, H), "#1A1A2E")
    draw = ImageDraw.Draw(img)
    try:
        fnt_lbl  = ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 16)
        fnt_cell = ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 13)
    except Exception:
        fnt_lbl = fnt_cell = ImageFont.load_default()

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
                radius=6, fill=fill, outline="#1A1A2E", width=2)
            txt_col = "#1A1A2E" if present else "#FFFFFF"
            draw.text((x_left + cell_px // 2, y_top + cell_px // 2),
                      coord, fill=txt_col, font=fnt_cell, anchor="mm")
    return img


# ── Excel cumulative heatmap export ────────────────────────────────────────────

def _hex_to_argb(hex_color: str) -> str:
    return "FF" + hex_color.lstrip("#").upper()


def _interpolate_color(count: int, max_count: int) -> str:
    """
    Return an ARGB hex string interpolated from green (0 misses) to red (max misses).
    count=0 → green (#2ECC71), count=max_count → red (#E74C3C).
    """
    if max_count == 0:
        t = 0.0
    else:
        t = min(count / max_count, 1.0)
    r = int(46  + (231 - 46)  * t)   # 46  → 231
    g = int(204 + (76  - 204) * t)   # 204 → 76
    b = int(113 + (60  - 113) * t)   # 113 → 60
    return f"FF{r:02X}{g:02X}{b:02X}"


def build_cumulative_heatmap_workbook(df: pd.DataFrame) -> bytes:
    """
    Build a single-sheet cumulative heatmap.
    Each cell shows how many frames had that coordinate marked missing.
    Color scales from green (0 frames missing) to red (all frames missing).
    """
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

    # Compute missing counts per coordinate
    missing_counts: dict[str, int] = {}
    for coord in ALL_COORDS:
        if coord in df.columns:
            missing_counts[coord] = int((df[coord].astype(str) == "0").sum())
        else:
            missing_counts[coord] = 0

    max_missing = max(missing_counts.values()) if missing_counts else 1

    # ── Title row ──────────────────────────────────────────────────────────────
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS) + 2)
    title_cell = ws.cell(1, 1,
        f"Chocolate Mold — Cumulative Missing Count  ({total_frames} frame(s) inspected)")
    title_cell.font      = Font(bold=True, color="FFFFFFFF", name="Arial", size=13)
    title_cell.fill      = header_fill
    title_cell.alignment = center
    ws.row_dimensions[1].height = 30

    # ── Sub-header: column letters ──────────────────────────────────────────────
    ws.cell(2, 1, "").fill = header_fill
    for ci, col in enumerate(COLS, start=2):
        c = ws.cell(2, ci, col)
        c.font = white_bold; c.fill = header_fill
        c.alignment = center; c.border = border

    # "Total missing" column header
    tot_col = len(COLS) + 2
    tc = ws.cell(2, tot_col, "Row\nMissing")
    tc.font = white_bold; tc.fill = summary_fill
    tc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    tc.border = border
    ws.row_dimensions[2].height = 30

    # Column widths
    ws.column_dimensions["A"].width = 6
    for ci in range(2, len(COLS) + 2):
        ws.column_dimensions[get_column_letter(ci)].width = 10
    ws.column_dimensions[get_column_letter(tot_col)].width = 11

    # ── Data rows ──────────────────────────────────────────────────────────────
    for ri, row_num in enumerate(ROWS):
        excel_row = ri + 3
        ws.row_dimensions[excel_row].height = 24

        # Row header
        rh = ws.cell(excel_row, 1, str(row_num))
        rh.font = white_bold; rh.fill = header_fill
        rh.alignment = center; rh.border = border

        row_missing_total = 0
        for ci, col in enumerate(COLS, start=2):
            coord = f"{col}{row_num}"
            count = missing_counts.get(coord, 0)
            row_missing_total += count

            argb  = _interpolate_color(count, max_missing)
            fill  = PatternFill("solid", fgColor=argb)

            # Choose text colour for legibility
            # Dark text when close to green, white text when closer to red
            t = count / max_missing if max_missing > 0 else 0
            font = cell_font_wh if t > 0.45 else cell_font

            cell = ws.cell(excel_row, ci, count)
            cell.fill      = fill
            cell.font      = font
            cell.alignment = center
            cell.border    = border

        # Row total
        row_total_cell = ws.cell(excel_row, tot_col, row_missing_total)
        row_total_cell.font = Font(bold=True, name="Arial", size=10)
        row_total_cell.fill = summary_fill
        row_total_cell.alignment = center
        row_total_cell.border = border
        # White text for dark background
        row_total_cell.font = Font(bold=True, name="Arial", size=10, color="FFFFFFFF")

    # ── Column totals row ──────────────────────────────────────────────────────
    totals_row = len(ROWS) + 3
    ws.row_dimensions[totals_row].height = 24

    col_label = ws.cell(totals_row, 1, "Total")
    col_label.font = Font(bold=True, color="FFFFFFFF", name="Arial", size=10)
    col_label.fill = summary_fill
    col_label.alignment = center
    col_label.border = border

    grand_total = 0
    for ci, col in enumerate(COLS, start=2):
        coord = f"{col}1"[0] + "x"   # dummy; we iterate by column letter
        # Re-derive: sum all rows for this column
        col_letter = COLS[ci - 2]
        col_total = sum(missing_counts.get(f"{col_letter}{r}", 0) for r in ROWS)
        grand_total += col_total

        ct = ws.cell(totals_row, ci, col_total)
        ct.font      = Font(bold=True, name="Arial", size=10)
        ct.fill      = summary_fill
        ct.alignment = center
        ct.border    = border
        ct.font      = Font(bold=True, name="Arial", size=10, color="FFFFFFFF")

    # Grand total corner
    gt = ws.cell(totals_row, tot_col, grand_total)
    gt.font      = Font(bold=True, name="Arial", size=11, color="FFFFFFFF")
    gt.fill      = PatternFill("solid", fgColor="FF1A252F")
    gt.alignment = center
    gt.border    = border

    # ── Legend row ─────────────────────────────────────────────────────────────
    legend_row = totals_row + 2
    ws.merge_cells(start_row=legend_row, start_column=1,
                   end_row=legend_row, end_column=len(COLS) + 2)
    leg = ws.cell(legend_row, 1,
        f"Each cell = number of frames where that cavity was missing  |  "
        f"Green = never missing  →  Red = missing in all {total_frames} frame(s)  |  "
        f"Total inspected: {total_frames} frames, {len(ALL_COORDS)} positions each")
    leg.font      = Font(italic=True, name="Arial", size=9, color="FF555555")
    leg.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[legend_row].height = 18

    # ── Summary tab ────────────────────────────────────────────────────────────
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
                r_val = int(231 * intensity + 46  * (1 - intensity))
                g_val = int(76  * intensity + 204 * (1 - intensity))
                b_val = int(60  * intensity + 113 * (1 - intensity))
                c.fill = PatternFill("solid", fgColor=f"FF{r_val:02X}{g_val:02X}{b_val:02X}")
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
        pct_cell = ws_sum.cell(total_data_row, 5, f"=C{total_data_row}/D{total_data_row}")
        pct_cell.number_format = "0.0%"
        pct_cell.font = Font(bold=True, name="Arial")
    for ci in range(1, 7):
        ws_sum.cell(total_data_row, ci).border = border
        ws_sum.cell(total_data_row, ci).alignment = center
    ws_sum.row_dimensions[total_data_row].height = 22

    # ── Coordinate Frequency tab ───────────────────────────────────────────────
    ws_freq = wb.create_sheet(title="Coordinate Frequency", index=2)

    freq_headers = ["Coordinate", "Times Missing", "Total Frames", "% Flagged", "Rank"]
    ws_freq.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(freq_headers))
    freq_title = ws_freq.cell(1, 1,
        f"Coordinate Flag Frequency  —  {total_frames} frame(s) analysed")
    freq_title.font      = Font(bold=True, color="FFFFFFFF", name="Arial", size=13)
    freq_title.fill      = header_fill
    freq_title.alignment = center
    ws_freq.row_dimensions[1].height = 30

    for ci, h in enumerate(freq_headers, 1):
        c = ws_freq.cell(2, ci, h)
        c.font = white_bold; c.fill = summary_fill
        c.alignment = center; c.border = border
    ws_freq.row_dimensions[2].height = 22

    # Column widths
    for ci, w in enumerate([14, 16, 14, 12, 8], 1):
        ws_freq.column_dimensions[get_column_letter(ci)].width = w

    # Sort coordinates by missing count descending, then by coord name for ties
    sorted_coords = sorted(
        ALL_COORDS,
        key=lambda coord: (-missing_counts.get(coord, 0), coord)
    )

    for rank, coord in enumerate(sorted_coords, start=1):
        excel_row = rank + 2
        count     = missing_counts.get(coord, 0)
        pct       = count / total_frames if total_frames > 0 else 0.0

        # Color the % Flagged cell: green→red based on percentage
        argb = _interpolate_color(count, max_missing if max_missing > 0 else 1)
        flag_fill = PatternFill("solid", fgColor=argb)
        t_val = count / max_missing if max_missing > 0 else 0
        flag_font = Font(name="Arial", size=10,
                         color="FFFFFFFF" if t_val > 0.45 else "FF000000",
                         bold=(count > 0))

        vals = [coord, count, total_frames, pct, rank]
        for ci, val in enumerate(vals, 1):
            cell = ws_freq.cell(excel_row, ci, val)
            cell.alignment = center
            cell.border    = border
            cell.font      = Font(name="Arial", size=10)

            if ci == 4:   # % Flagged column — colour-coded
                cell.number_format = "0.0%"
                cell.fill = flag_fill
                cell.font = flag_font
            elif ci == 2 and count > 0:  # Times Missing — subtle red tint when non-zero
                intensity = min(count / total_frames, 1.0)
                r_v = int(231 * intensity + 236 * (1 - intensity))
                g_v = int(76  * intensity + 240 * (1 - intensity))
                b_v = int(60  * intensity + 241 * (1 - intensity))
                cell.fill = PatternFill("solid", fgColor=f"FF{r_v:02X}{g_v:02X}{b_v:02X}")

        ws_freq.row_dimensions[excel_row].height = 18

    # Totals / summary footer
    footer_row = len(ALL_COORDS) + 3
    ws_freq.merge_cells(start_row=footer_row, start_column=1,
                        end_row=footer_row, end_column=len(freq_headers))
    total_missing_events = sum(missing_counts.values())
    avg_missing_per_frame = total_missing_events / total_frames if total_frames > 0 else 0
    foot = ws_freq.cell(footer_row, 1,
        f"Total missing events across all frames: {total_missing_events}   |   "
        f"Avg missing positions per frame: {avg_missing_per_frame:.1f} / {len(ALL_COORDS)}   |   "
        f"Coordinates never flagged: {sum(1 for v in missing_counts.values() if v == 0)}")
    foot.font      = Font(italic=True, name="Arial", size=9, color="FF555555")
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
        "dirty":            False,
        "frame_image":      None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def load_frame(frame_id: str):
    st.session_state.active_frame_id = frame_id
    st.session_state.coord_dict      = get_frame_dict(st.session_state.df, frame_id)
    st.session_state.dirty           = False
    row = st.session_state.df[st.session_state.df["frame_id"] == frame_id]
    if not row.empty:
        img_path = str(row.iloc[0].get("image_path", ""))
        st.session_state.frame_image = load_frame_image(img_path)
    else:
        st.session_state.frame_image = None


def auto_save(frame_id: str, frame_name: str, image_path: str = ""):
    if not image_path:
        row = st.session_state.df[st.session_state.df["frame_id"] == frame_id]
        if not row.empty:
            image_path = str(row.iloc[0].get("image_path", ""))
    st.session_state.df = upsert_frame(
        st.session_state.df, frame_id, frame_name,
        st.session_state.coord_dict, image_path)
    save_data(st.session_state.df)
    st.session_state.dirty = False


def navigate_to_adjacent_frame(direction: int):
    """Load the next (+1) or previous (-1) frame relative to the current one."""
    df       = st.session_state.df
    frame_ids = df["frame_id"].tolist()
    if not frame_ids:
        return
    current = st.session_state.active_frame_id
    if current not in frame_ids:
        load_frame(frame_ids[0])
        return
    idx     = frame_ids.index(current)
    new_idx = (idx + direction) % len(frame_ids)
    load_frame(frame_ids[new_idx])


# ── Main app ───────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="Chocolate Mold Inspector",
        page_icon="🍫",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    init_state()

    # ── CSS ────────────────────────────────────────────────────────────────────
    st.markdown("""
    <style>
    [data-testid="stSidebar"] { background: #1A1A2E; }
    h1,h2,h3 { color: #ECF0F1; }
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
    .nav-btn > button {
        background: #2C3E50 !important;
        color: #ECF0F1 !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Sidebar ────────────────────────────────────────────────────────────────
    with st.sidebar:
        st.title("🍫 Mold Inspector")
        st.caption("15 × 8 grid | 120 positions")
        st.divider()

        # ── Batch image upload ─────────────────────────────────────────────────
        st.subheader("📸 Upload Frame Photos")
        st.markdown(
            f'<div class="upload-hint">Upload up to {MAX_UPLOAD_FILES} frame images at once. '
            'A new record is created for each photo automatically.</div>',
            unsafe_allow_html=True)

        uploaded_photos = st.file_uploader(
            "Drop frame images here",
            type=["png", "jpg", "jpeg", "bmp", "tiff", "webp"],
            accept_multiple_files=True,
            label_visibility="collapsed",
            key="batch_uploader",
        )

        if uploaded_photos:
            # Enforce upload cap
            if len(uploaded_photos) > MAX_UPLOAD_FILES:
                st.warning(f"Only the first {MAX_UPLOAD_FILES} images will be imported.")
                uploaded_photos = uploaded_photos[:MAX_UPLOAD_FILES]

            new_count = 0
            for uf in uploaded_photos:
                base_name = os.path.splitext(uf.name)[0]
                existing_names = st.session_state.df["frame_name"].tolist() \
                    if not st.session_state.df.empty else []
                if base_name in existing_names:
                    continue
                fid   = f"frame_{int(time.time() * 1000)}_{new_count}"
                ipath = save_frame_image(fid, uf)
                fresh = {c: True for c in ALL_COORDS}
                st.session_state.df = upsert_frame(
                    st.session_state.df, fid, base_name, fresh, ipath)
                new_count += 1

            if new_count:
                save_data(st.session_state.df)
                st.success(f"Imported {new_count} new frame(s).")
                last_id = st.session_state.df["frame_id"].iloc[-1]
                load_frame(last_id)
                st.rerun()

        st.divider()

        # ── Frame selector ─────────────────────────────────────────────────────
        st.subheader("📋 Select Frame")
        df = st.session_state.df

        if df.empty:
            st.info("No frames yet — upload photos above to get started.")
        else:
            frame_options = df["frame_name"].tolist()
            frame_ids     = df["frame_id"].tolist()

            current_idx = 0
            if st.session_state.active_frame_id in frame_ids:
                current_idx = frame_ids.index(st.session_state.active_frame_id)

            selected_idx = st.selectbox(
                "Frame", options=range(len(frame_options)),
                format_func=lambda i: frame_options[i],
                index=current_idx, label_visibility="collapsed")

            col_load, col_del = st.columns(2)
            if col_load.button("Load", use_container_width=True):
                load_frame(frame_ids[selected_idx])
                st.rerun()
            if col_del.button("🗑 Delete", use_container_width=True):
                fid_to_del = frame_ids[selected_idx]
                row = st.session_state.df[
                    st.session_state.df["frame_id"] == fid_to_del]
                if not row.empty:
                    ip = str(row.iloc[0].get("image_path", ""))
                    if ip and os.path.exists(ip):
                        os.remove(ip)
                st.session_state.df = st.session_state.df[
                    st.session_state.df["frame_id"] != fid_to_del].reset_index(drop=True)
                save_data(st.session_state.df)
                if st.session_state.active_frame_id == fid_to_del:
                    st.session_state.active_frame_id = None
                    st.session_state.frame_image     = None
                st.rerun()

            st.caption(f"Total frames: **{len(df)}**")

        st.divider()

        # ── Export ─────────────────────────────────────────────────────────────
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
                    "💾 Save file", data=xlsx_bytes,
                    file_name=f"mold_cumulative_heatmap_{now}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)

        if not df.empty:
            st.download_button(
                "⬇️ Backup CSV",
                data=df.to_csv(index=False).encode(),
                file_name="mold_data_backup.csv",
                mime="text/csv", use_container_width=True)

        # CSV restore
        st.divider()
        st.subheader("⬆️ Restore from CSV")
        uploaded_csv = st.file_uploader("Upload backup CSV", type="csv",
                                        label_visibility="collapsed")
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
                    "% Present":    f"{(len(ALL_COORDS)-missing)/len(ALL_COORDS)*100:.1f}%",
                    "Last Updated": rec.get("timestamp", ""),
                })
            st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)
        return

    # ── Active frame ───────────────────────────────────────────────────────────
    active_id   = st.session_state.active_frame_id
    df          = st.session_state.df
    active_row  = df[df["frame_id"] == active_id].iloc[0]
    frame_name  = active_row["frame_name"]
    image_path  = str(active_row.get("image_path", ""))
    coord_dict  = st.session_state.coord_dict
    photo       = st.session_state.frame_image

    frame_ids   = df["frame_id"].tolist()
    current_idx = frame_ids.index(active_id) if active_id in frame_ids else 0
    total_frames = len(frame_ids)

    missing_list  = [c for c, v in coord_dict.items() if not v]
    present_count = len(ALL_COORDS) - len(missing_list)

    # ── Header + navigation ────────────────────────────────────────────────────
    header_left, header_right = st.columns([3, 1])
    with header_left:
        st.markdown(f"## 🍫 {frame_name}")
        st.caption(
            f"Frame {current_idx + 1} of {total_frames}  |  "
            f"ID: `{active_id}`  |  Saved: {active_row.get('timestamp','—')}"
        )
    with header_right:
        nav_c1, nav_c2 = st.columns(2)
        with nav_c1:
            if st.button("◀ Prev", use_container_width=True,
                         disabled=(total_frames <= 1),
                         help="Go to previous frame"):
                navigate_to_adjacent_frame(-1)
                st.rerun()
        with nav_c2:
            if st.button("Next ▶", use_container_width=True,
                         disabled=(total_frames <= 1),
                         help="Go to next frame",
                         type="primary"):
                navigate_to_adjacent_frame(+1)
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

    # ── Photo upload for THIS frame ────────────────────────────────────────────
    with st.expander(
        "📷 " + ("Replace frame photo" if photo else "Attach a photo to this frame"),
        expanded=(photo is None)):

        per_frame_upload = st.file_uploader(
            "Upload mold photo for this frame",
            type=["png", "jpg", "jpeg", "bmp", "tiff", "webp"],
            key=f"photo_{active_id}",
            label_visibility="collapsed",
        )
        if per_frame_upload:
            ipath  = save_frame_image(active_id, per_frame_upload)
            photo  = Image.open(per_frame_upload).convert("RGB")
            st.session_state.frame_image = photo
            st.session_state.df = upsert_frame(
                st.session_state.df, active_id, frame_name, coord_dict, ipath)
            save_data(st.session_state.df)
            st.success("Photo saved to this frame.")
            st.rerun()

    # ── Tabs ──────────────────────────────────────────────────────────────────
    if photo:
        tab_photo, tab_grid, tab_quick, tab_vis = st.tabs([
            "🖼️ Photo Inspector", "🔲 Grid Editor", "⌨️ Quick Entry", "📊 Overlay Preview"])
    else:
        tab_grid, tab_quick, tab_vis = st.tabs([
            "🔲 Grid Editor", "⌨️ Quick Entry", "📊 Grid Preview"])
        tab_photo = None

    # ── Tab: Photo Inspector ───────────────────────────────────────────────────
    if tab_photo and photo:
        with tab_photo:
            # ── Process any pending canvas click from query params ─────────────
            qp = st.query_params
            pending_click = qp.get("click", "")
            if pending_click and pending_click in ALL_COORDS:
                st.session_state.coord_dict[pending_click] = \
                    not st.session_state.coord_dict.get(pending_click, True)
                auto_save(active_id, frame_name)
                st.query_params.clear()
                st.rerun()

            # ── Top control strip ─────────────────────────────────────────────
            ctrl_c1, ctrl_c2, ctrl_c3, ctrl_c4, ctrl_c5 = st.columns([1, 1, 1, 1, 3])
            with ctrl_c1:
                if st.button("✅ All Present", use_container_width=True, key="pi_all_pres"):
                    for c in ALL_COORDS:
                        st.session_state.coord_dict[c] = True
                    auto_save(active_id, frame_name)
                    st.rerun()
            with ctrl_c2:
                if st.button("❌ All Missing", use_container_width=True, key="pi_all_miss"):
                    for c in ALL_COORDS:
                        st.session_state.coord_dict[c] = False
                    auto_save(active_id, frame_name)
                    st.rerun()
            with ctrl_c3:
                if st.button("🔄 Invert", use_container_width=True, key="pi_invert"):
                    for c in ALL_COORDS:
                        st.session_state.coord_dict[c] = not st.session_state.coord_dict[c]
                    auto_save(active_id, frame_name)
                    st.rerun()
            with ctrl_c4:
                opacity = st.slider("Overlay opacity", 60, 220, 60, 10,
                                    key="overlay_opacity", label_visibility="collapsed")
                st.caption("Opacity")
            with ctrl_c5:
                if missing_list:
                    badges = " ".join(
                        f"<span style='background:#E74C3C;color:#FFF;"
                        f"padding:1px 7px;border-radius:4px;font-size:0.78rem;"
                        f"margin:1px;display:inline-block'>{coord}</span>"
                        for coord in sorted(missing_list)
                    )
                    st.markdown(
                        f"<div style='line-height:2.0;padding-top:2px'>"
                        f"<strong style='color:#E74C3C'>Missing ({len(missing_list)}):</strong> "
                        f"{badges}</div>",
                        unsafe_allow_html=True)
                else:
                    st.success("All 120 positions present!")

            st.markdown("<hr style='margin:4px 0 8px'>", unsafe_allow_html=True)

            # ── Build overlay image and pass geometry to JS ────────────────────
            cur_opacity = st.session_state.get("overlay_opacity", 60)
            overlay_img = render_overlay_on_photo(
                photo, st.session_state.coord_dict, opacity=cur_opacity)

            W_img, H_img = photo.size
            margin_l = max(24, int(W_img * 0.038))
            margin_t = max(20, int(H_img * 0.055))
            grid_w   = W_img - margin_l - max(4, int(W_img * 0.008))
            grid_h   = H_img - margin_t - max(4, int(H_img * 0.008))
            cell_w   = grid_w / len(COLS)
            cell_h   = grid_h / len(ROWS)

            b64_img    = pil_to_b64(overlay_img, fmt="JPEG")
            coord_json = json.dumps(
                {c: (1 if st.session_state.coord_dict.get(c, True) else 0)
                 for c in ALL_COORDS})
            cols_json  = json.dumps(COLS)
            rows_json  = json.dumps([str(r) for r in ROWS])
            canvas_id  = f"mold_{re.sub(r'[^a-zA-Z0-9]', '_', active_id)}"

            canvas_html = f"""
<style>
#{canvas_id}_wrap {{
  position: relative;
  display: block;
  width: 100%;
  line-height: 0;
}}
#{canvas_id} {{
  width: 100%;
  height: auto;
  display: block;
  border-radius: 6px;
  cursor: crosshair;
  box-shadow: 0 2px 16px rgba(0,0,0,0.45);
}}
#{canvas_id}_tip {{
  position: fixed;
  background: rgba(15,15,15,0.88);
  color: #fff;
  padding: 5px 13px;
  border-radius: 6px;
  font: 700 13px/1.5 Arial, sans-serif;
  pointer-events: none;
  display: none;
  z-index: 9999;
  white-space: nowrap;
  border: 1px solid rgba(255,255,255,0.12);
}}
#{canvas_id}_flash {{
  position: fixed;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  background: rgba(0,0,0,0.78);
  color: #fff;
  padding: 10px 24px;
  border-radius: 10px;
  font: 700 15px Arial, sans-serif;
  pointer-events: none;
  display: none;
  z-index: 9999;
}}
#{canvas_id}_msg {{
  font: 12px Arial, sans-serif;
  color: #7F8C8D;
  text-align: center;
  padding: 5px 0 0;
  min-height: 20px;
}}
</style>

<div id="{canvas_id}_wrap">
  <canvas id="{canvas_id}"></canvas>
  <div id="{canvas_id}_tip"></div>
  <div id="{canvas_id}_flash"></div>
</div>
<div id="{canvas_id}_msg">👆 Click any cavity on the photo to mark it missing or restore it</div>

<script>
(function() {{
  const COLS     = {cols_json};
  const ROWS     = {rows_json};
  const coords   = {coord_json};
  const marginL  = {margin_l};
  const marginT  = {margin_t};
  const cellW    = {cell_w};
  const cellH    = {cell_h};
  const IMG_W    = {W_img};
  const IMG_H    = {H_img};
  const CID      = "{canvas_id}";

  const canvas = document.getElementById(CID);
  const ctx    = canvas.getContext('2d');
  const tip    = document.getElementById(CID + '_tip');
  const flash  = document.getElementById(CID + '_flash');
  const msg    = document.getElementById(CID + '_msg');

  canvas.width  = IMG_W;
  canvas.height = IMG_H;

  const baseImg = new Image();
  baseImg.src = 'data:image/jpeg;base64,{b64_img}';
  baseImg.onload = () => ctx.drawImage(baseImg, 0, 0);

  function getScale() {{
    const r = canvas.getBoundingClientRect();
    return {{ sx: IMG_W / r.width, sy: IMG_H / r.height, r }};
  }}

  function coordFromXY(x, y) {{
    const ci = Math.floor((x - marginL) / cellW);
    const ri = Math.floor((y - marginT) / cellH);
    if (ci < 0 || ci >= COLS.length || ri < 0 || ri >= ROWS.length) return null;
    return COLS[ci] + ROWS[ri];
  }}

  canvas.addEventListener('mousemove', e => {{
    const {{ sx, sy, r }} = getScale();
    const x = (e.clientX - r.left) * sx;
    const y = (e.clientY - r.top)  * sy;
    const coord = coordFromXY(x, y);
    if (coord) {{
      const status = coords[coord] === 1 ? '🟢 Present' : '🔴 Missing';
      tip.textContent = coord + ' — ' + status + '  (click to toggle)';
      tip.style.display = 'block';
      tip.style.left = (e.clientX + 16) + 'px';
      tip.style.top  = (e.clientY - 12) + 'px';
    }} else {{
      tip.style.display = 'none';
    }}
  }});
  canvas.addEventListener('mouseleave', () => tip.style.display = 'none');

  canvas.addEventListener('click', e => {{
    const {{ sx, sy, r }} = getScale();
    const x = (e.clientX - r.left) * sx;
    const y = (e.clientY - r.top)  * sy;
    const coord = coordFromXY(x, y);
    if (!coord) return;

    // Optimistic local toggle
    coords[coord] = coords[coord] === 1 ? 0 : 1;
    const nowMissing = coords[coord] === 0;

    // Repaint just that cell for instant feedback
    const ci = COLS.indexOf(coord[0]);
    const ri = ROWS.indexOf(coord.slice(1));
    const x0 = marginL + ci * cellW;
    const y0 = marginT + ri * cellH;
    ctx.drawImage(baseImg, x0, y0, cellW, cellH, x0, y0, cellW, cellH);
    ctx.fillStyle = nowMissing
      ? 'rgba(231,76,60,0.75)' : 'rgba(46,204,113,0.50)';
    ctx.fillRect(x0 + 1, y0 + 1, cellW - 2, cellH - 2);
    ctx.strokeStyle = 'rgba(255,255,255,0.75)';
    ctx.lineWidth = 1.5;
    ctx.strokeRect(x0 + 0.75, y0 + 0.75, cellW - 1.5, cellH - 1.5);

    // Label
    ctx.fillStyle = nowMissing ? '#fff' : 'rgba(10,10,10,0.85)';
    ctx.font = 'bold ' + Math.max(9, Math.floor(Math.min(cellW, cellH) * 0.30)) + 'px Arial';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(coord, x0 + cellW / 2, y0 + cellH / 2);

    tip.style.display = 'none';
    msg.innerHTML = (nowMissing
      ? '<span style="color:#E74C3C">🔴 Marked MISSING: <strong>' + coord + '</strong></span>'
      : '<span style="color:#2ECC71">🟢 Marked PRESENT: <strong>' + coord + '</strong></span>');

    // Flash confirmation
    flash.textContent = (nowMissing ? '🔴 ' : '🟢 ') + coord;
    flash.style.display = 'block';
    setTimeout(() => flash.style.display = 'none', 700);

    // ── Relay to Streamlit via URL query param ─────────────────────────────
    // We append ?click=<coord> to the parent URL; Streamlit reads it
    // via st.query_params on the next rerun, processes it, then clears it.
    try {{
      const url = new URL(window.parent.location.href);
      url.searchParams.set('click', coord);
      window.parent.history.pushState({{}}, '', url.toString());
      // Trigger Streamlit's rerun by posting a message it recognises
      window.parent.postMessage({{ type: 'streamlit:forceRerender' }}, '*');
    }} catch(err) {{
      // Cross-origin fallback: try direct DOM relay to any visible text input
      const inputs = window.parent.document.querySelectorAll(
        '[data-testid="stTextInput"] input');
      if (inputs.length) {{
        const inp = inputs[inputs.length - 1];
        Object.getOwnPropertyDescriptor(
          window.HTMLInputElement.prototype, 'value')
          .set.call(inp, coord);
        inp.dispatchEvent(new Event('input', {{ bubbles: true }}));
      }}
    }}
  }});
}})();
</script>
"""

            components.html(canvas_html,
                            height=int(H_img * 820 / max(W_img, 1)) + 70,
                            scrolling=False)

            # Download button for the overlay image
            buf = io.BytesIO()
            overlay_img.save(buf, format="PNG")
            st.download_button(
                "⬇️ Download overlay image",
                data=buf.getvalue(),
                file_name=f"{re.sub(r'[^a-zA-Z0-9_-]','_',frame_name)}_overlay.png",
                mime="image/png",
            )

    # ── Tab: Grid Editor ──────────────────────────────────────────────────────
    with tab_grid:
        st.markdown(
            "Click a cell to toggle **Present 🟢** ↔ **Missing 🔴**. "
            "Changes auto-save when photo inspector is active; "
            "use **Save Frame** below otherwise.")

        tb1, tb2, tb3, tb4 = st.columns(4)
        with tb1:
            if st.button("✅ All Present", use_container_width=True, key="g_all_pres"):
                for c in ALL_COORDS:
                    st.session_state.coord_dict[c] = True
                st.session_state.dirty = True; st.rerun()
        with tb2:
            if st.button("❌ All Missing", use_container_width=True, key="g_all_miss"):
                for c in ALL_COORDS:
                    st.session_state.coord_dict[c] = False
                st.session_state.dirty = True; st.rerun()
        with tb3:
            if st.button("🔄 Invert", use_container_width=True, key="g_invert"):
                for c in ALL_COORDS:
                    st.session_state.coord_dict[c] = not st.session_state.coord_dict[c]
                st.session_state.dirty = True; st.rerun()
        with tb4:
            if st.button("💾 Save Frame", use_container_width=True,
                         type="primary", key="g_save",
                         disabled=not st.session_state.dirty):
                auto_save(active_id, frame_name)
                st.success("Saved!"); st.rerun()

        hcols = st.columns([0.5] + [1]*len(COLS))
        hcols[0].markdown("**↓**")
        for i, lbl in enumerate(COLS):
            hcols[i+1].markdown(
                f"<div style='text-align:center;font-weight:700;"
                f"color:#3498DB;font-size:0.8rem'>{lbl}</div>",
                unsafe_allow_html=True)

        for row_num in ROWS:
            rcols = st.columns([0.5] + [1]*len(COLS))
            rcols[0].markdown(
                f"<div style='text-align:center;font-weight:700;"
                f"color:#3498DB'>{row_num}</div>", unsafe_allow_html=True)
            for ci, col_lbl in enumerate(COLS):
                coord   = f"{col_lbl}{row_num}"
                present = st.session_state.coord_dict.get(coord, True)
                with rcols[ci+1]:
                    if st.button("🟢" if present else "🔴",
                                 key=f"grid_btn_{coord}",
                                 help=f"{coord}: {'Present' if present else 'MISSING'}",
                                 use_container_width=True):
                        st.session_state.coord_dict[coord] = not present
                        st.session_state.dirty = True; st.rerun()

        if st.session_state.dirty:
            st.warning("⚠️ Unsaved changes — click **Save Frame** above.")

    # ── Tab: Quick Entry ──────────────────────────────────────────────────────
    with tab_quick:
        st.markdown("""
        Paste **missing** coordinates as a comma-separated list.
        All others are assumed **present**.
        **Examples:** `A1, C3, O8`  or  `B2 D5 F7`
        """)

        if missing_list:
            st.markdown(
                f"**Currently missing ({len(missing_list)}):** " +
                ", ".join(f"`{c}`" for c in sorted(missing_list)))
        else:
            st.success("All positions currently marked as present.")

        with st.form("quick_entry"):
            raw = st.text_area(
                "Missing coordinates",
                value=", ".join(sorted(missing_list)) if missing_list else "",
                height=90, placeholder="e.g. A1, B3, G5, O8")
            ca, cb = st.columns(2)
            apply_btn = ca.form_submit_button("Apply (replace all missing)",
                                              use_container_width=True, type="primary")
            add_btn   = cb.form_submit_button("Add to existing missing",
                                              use_container_width=True)

        if apply_btn or add_btn:
            tokens = re.split(r"[\s,;]+", raw.strip().upper())
            valid, invalid = [], []
            for t in tokens:
                if not t: continue
                (valid if t in ALL_COORDS else invalid).append(t)
            if invalid:
                st.error(f"Invalid coordinate(s): {', '.join(invalid)}")
            else:
                if apply_btn:
                    for c in ALL_COORDS:
                        st.session_state.coord_dict[c] = (c not in valid)
                else:
                    for c in valid:
                        st.session_state.coord_dict[c] = False
                auto_save(active_id, frame_name)
                n = len([c for c, v in st.session_state.coord_dict.items() if not v])
                st.success(f"Saved! {n} position(s) missing.")
                st.rerun()

        st.subheader("Row-by-row status")
        status_rows = []
        for row_num in ROWS:
            rm = [f"{cl}{row_num}" for cl in COLS
                  if not st.session_state.coord_dict.get(f"{cl}{row_num}", True)]
            status_rows.append({
                "Row": row_num,
                "Missing": len(rm),
                "Missing Coords": ", ".join(rm) if rm else "—",
            })
        st.dataframe(pd.DataFrame(status_rows), use_container_width=True, hide_index=True)

    # ── Tab: Grid / Overlay Preview ───────────────────────────────────────────
    with tab_vis:
        if photo:
            st.markdown("Overlay rendered on actual mold photo. 🟢 Present &nbsp; 🔴 Missing")
            opa = st.slider("Overlay opacity", 60, 220, 60, 10, key="vis_opacity")
            ov  = render_overlay_on_photo(photo, st.session_state.coord_dict, opacity=opa)
            st.image(ov, use_container_width=True, caption=frame_name)
            buf = io.BytesIO(); ov.save(buf, format="PNG")
            st.download_button(
                "⬇️ Download overlay (.png)", data=buf.getvalue(),
                file_name=f"{re.sub(r'[^a-zA-Z0-9_-]','_',frame_name)}_overlay.png",
                mime="image/png")
        else:
            st.markdown("Synthetic grid (upload a photo to see the real mold). "
                        "🟢 Present &nbsp; 🔴 Missing")
            grid_img = render_plain_grid(st.session_state.coord_dict)
            st.image(grid_img, use_container_width=True, caption=frame_name)
            buf = io.BytesIO(); grid_img.save(buf, format="PNG")
            st.download_button(
                "⬇️ Download grid (.png)", data=buf.getvalue(),
                file_name=f"{re.sub(r'[^a-zA-Z0-9_-]','_',frame_name)}_grid.png",
                mime="image/png")

        if missing_list:
            st.subheader(f"🔴 Missing positions ({len(missing_list)})")
            rows_of_missing = {}
            for c in sorted(missing_list):
                rows_of_missing.setdefault(c[1:], []).append(c)
            for rn, coords in sorted(rows_of_missing.items(), key=lambda x: int(x[0])):
                st.markdown(
                    f"**Row {rn}:** " +
                    " ".join(
                        f"<span style='background:#E74C3C;color:#FFF;"
                        f"padding:2px 6px;border-radius:4px;font-size:0.85rem'>{c}</span>"
                        for c in coords),
                    unsafe_allow_html=True)
        else:
            st.success("🎉 All 120 positions are present!")

    # ── Bottom navigation (Next Frame button) ──────────────────────────────────
    st.divider()
    bot_left, bot_mid, bot_right = st.columns([1, 2, 1])
    with bot_left:
        if st.button("◀ Previous Frame", use_container_width=True,
                     disabled=(total_frames <= 1)):
            navigate_to_adjacent_frame(-1)
            st.rerun()
    with bot_mid:
        st.markdown(
            f"<div style='text-align:center;color:#95A5A6;padding-top:8px'>"
            f"Frame <strong>{current_idx + 1}</strong> of <strong>{total_frames}</strong></div>",
            unsafe_allow_html=True)
    with bot_right:
        if st.button("Next Frame ▶", use_container_width=True,
                     disabled=(total_frames <= 1),
                     type="primary"):
            navigate_to_adjacent_frame(+1)
            st.rerun()


if __name__ == "__main__":
    main()
