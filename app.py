"""
Chocolate Mold Inspector — Streamlit App
15 columns (A–O) × 8 rows (1–8) = 120 coordinates per mold frame
Upload frame photos, click coordinates to mark them missing,
and export a cumulative Excel/PDF report across all frames.
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

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ── Constants ──────────────────────────────────────────────────────────────────
COLS       = list("ABCDEFGHIJKLMNO")
ROWS       = list(range(1, 9))
ALL_COORDS = [f"{c}{r}" for r in ROWS for c in COLS]

DATA_FILE        = "mold_data.csv"
IMAGES_DIR       = "frame_images"
MAX_UPLOAD_FILES = 50

PALETTE = {"present": "#2ECC71", "empty": "#E74C3C", "selected": "#F39C12"}

os.makedirs(IMAGES_DIR, exist_ok=True)


# ── Data persistence ───────────────────────────────────────────────────────────

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
        columns=["frame_id", "frame_name", "timestamp", "image_path"] + ALL_COORDS)


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
    ext  = os.path.splitext(uploaded_file.name)[-1].lower() or ".png"
    path = os.path.join(IMAGES_DIR, f"{frame_id}{ext}")
    with open(path, "wb") as f:
        f.write(uploaded_file.getvalue())
    return path


def load_frame_image(image_path: str):
    if image_path and os.path.exists(image_path):
        return Image.open(image_path).convert("RGB")
    return None


# ── Auto-save: immediately persist coord_dict to CSV ──────────────────────────
# This is the ONLY save path used in the Photo Inspector tab.
# It writes directly to disk so switching tabs never loses data.

def auto_save(frame_id: str, frame_name: str):
    """Persist current coord_dict to the DataFrame and CSV immediately."""
    row = st.session_state.df[st.session_state.df["frame_id"] == frame_id]
    image_path = str(row.iloc[0].get("image_path", "")) if not row.empty else ""
    st.session_state.df = upsert_frame(
        st.session_state.df, frame_id, frame_name,
        st.session_state.coord_dict, image_path)
    save_data(st.session_state.df)
    st.session_state.dirty = False


# ── Image helpers ──────────────────────────────────────────────────────────────

def pil_to_b64(img: Image.Image, fmt="JPEG") -> str:
    buf = io.BytesIO()
    img.save(buf, format=fmt)
    return base64.b64encode(buf.getvalue()).decode()


def render_overlay_on_photo(photo: Image.Image, coord_dict: dict,
                             opacity: int = 40) -> Image.Image:
    base = photo.convert("RGBA")
    W, H = base.size
    margin_l = max(24, int(W * 0.038))
    margin_t = max(20, int(H * 0.055))
    grid_w   = W - margin_l - max(4, int(W * 0.008))
    grid_h   = H - margin_t - max(4, int(H * 0.008))
    cell_w   = grid_w / len(COLS)
    cell_h   = grid_h / len(ROWS)

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
            fill = (46, 204, 113, opacity) if present else (231, 76, 60, opacity)
            draw.rectangle([x0 + 1, y0 + 1, x1 - 1, y1 - 1], fill=fill)
            draw.rectangle([x0, y0, x1, y1], outline=(255, 255, 255, 180), width=1)
            txt_col = (10, 10, 10, 255) if present else (255, 255, 255, 255)
            draw.text(((x0 + x1) / 2, (y0 + y1) / 2),
                      coord, fill=txt_col, font=fnt, anchor="mm")

    for ci, col in enumerate(COLS):
        x = margin_l + (ci + 0.5) * cell_w
        draw.text((x, margin_t / 2), col,
                  fill=(255, 255, 255, 230), font=fnt_hdr, anchor="mm")
    for ri, row_num in enumerate(ROWS):
        y = margin_t + (ri + 0.5) * cell_h
        draw.text((margin_l / 2, y), str(row_num),
                  fill=(255, 255, 255, 230), font=fnt_hdr, anchor="mm")

    return Image.alpha_composite(base, overlay).convert("RGB")


def render_plain_grid(coord_dict: dict,
                      cell_px: int = 68, label_px: int = 34) -> Image.Image:
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


# ── Excel export ───────────────────────────────────────────────────────────────

def _hex_to_argb(hex_color: str) -> str:
    return "FF" + hex_color.lstrip("#").upper()


def _interpolate_color(count: int, max_count: int) -> str:
    t = min(count / max_count, 1.0) if max_count > 0 else 0.0
    r = int(46  + (231 - 46)  * t)
    g = int(204 + (76  - 204) * t)
    b = int(113 + (60  - 113) * t)
    return f"FF{r:02X}{g:02X}{b:02X}"


def _compute_missing_counts(df: pd.DataFrame) -> dict:
    counts = {}
    for coord in ALL_COORDS:
        counts[coord] = int((df[coord].astype(str) == "0").sum()) \
            if coord in df.columns else 0
    return counts


def build_cumulative_heatmap_workbook(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Cumulative Heatmap"

    thin         = Side(style="thin", color="999999")
    border       = Border(left=thin, right=thin, top=thin, bottom=thin)
    center       = Alignment(horizontal="center", vertical="center")
    header_fill  = PatternFill("solid", fgColor="FF2C3E50")
    summary_fill = PatternFill("solid", fgColor="FF34495E")
    white_bold   = Font(bold=True, color="FFFFFFFF", name="Arial", size=11)
    cell_font    = Font(name="Arial", size=10)
    cell_font_wh = Font(name="Arial", size=10, color="FFFFFFFF")

    total_frames   = len(df)
    missing_counts = _compute_missing_counts(df)
    max_missing    = max(missing_counts.values()) if missing_counts else 1

    # Sheet 1: Cumulative Heatmap
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
        er = ri + 3
        ws.row_dimensions[er].height = 24
        rh = ws.cell(er, 1, str(row_num))
        rh.font = white_bold; rh.fill = header_fill
        rh.alignment = center; rh.border = border

        row_total = 0
        for ci, col in enumerate(COLS, start=2):
            coord = f"{col}{row_num}"
            count = missing_counts.get(coord, 0)
            row_total += count
            argb  = _interpolate_color(count, max_missing)
            t     = count / max_missing if max_missing > 0 else 0
            cell  = ws.cell(er, ci, count)
            cell.fill      = PatternFill("solid", fgColor=argb)
            cell.font      = cell_font_wh if t > 0.45 else cell_font
            cell.alignment = center; cell.border = border

        rtc = ws.cell(er, tot_col, row_total)
        rtc.font = Font(bold=True, name="Arial", size=10, color="FFFFFFFF")
        rtc.fill = summary_fill; rtc.alignment = center; rtc.border = border

    totals_row = len(ROWS) + 3
    ws.row_dimensions[totals_row].height = 24
    cl = ws.cell(totals_row, 1, "Total")
    cl.font = Font(bold=True, color="FFFFFFFF", name="Arial", size=10)
    cl.fill = summary_fill; cl.alignment = center; cl.border = border

    grand_total = 0
    for ci, col_letter in enumerate(COLS, start=2):
        col_total = sum(missing_counts.get(f"{COLS[ci-2]}{r}", 0) for r in ROWS)
        grand_total += col_total
        ct = ws.cell(totals_row, ci, col_total)
        ct.font = Font(bold=True, name="Arial", size=10, color="FFFFFFFF")
        ct.fill = summary_fill; ct.alignment = center; ct.border = border

    gt = ws.cell(totals_row, tot_col, grand_total)
    gt.font = Font(bold=True, name="Arial", size=11, color="FFFFFFFF")
    gt.fill = PatternFill("solid", fgColor="FF1A252F")
    gt.alignment = center; gt.border = border

    legend_row = totals_row + 2
    ws.merge_cells(start_row=legend_row, start_column=1,
                   end_row=legend_row, end_column=len(COLS) + 2)
    leg = ws.cell(legend_row, 1,
        f"Each cell = number of frames where that cavity was missing  |  "
        f"Green = never missing  to  Red = missing in all {total_frames} frame(s)")
    leg.font = Font(italic=True, name="Arial", size=9, color="FF555555")
    leg.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[legend_row].height = 18

    # Sheet 2: Frame Summary
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
        for ci, val in enumerate([fname, missing, present, total,
                                   f"=C{ri}/D{ri}", ts], 1):
            c = ws_sum.cell(ri, ci, val)
            c.alignment = center; c.border = border
            c.font = Font(name="Arial", size=10)
            if ci == 2 and total > 0:
                inten = min(missing / total, 1.0)
                rv = int(231 * inten + 46  * (1 - inten))
                gv = int(76  * inten + 204 * (1 - inten))
                bv = int(60  * inten + 113 * (1 - inten))
                c.fill = PatternFill("solid", fgColor=f"FF{rv:02X}{gv:02X}{bv:02X}")
            if ci == 5:
                c.number_format = "0.0%"
        ws_sum.row_dimensions[ri].height = 20

    tdr = len(df) + 3
    ws_sum.cell(tdr, 1, "TOTALS").font = Font(bold=True, name="Arial", size=10)
    if not df.empty:
        ld = tdr - 1
        for ci2, formula in enumerate(
            [None, f"=SUM(B3:B{ld})", f"=SUM(C3:C{ld})",
             f"=SUM(D3:D{ld})", f"=C{tdr}/D{tdr}", None], 1):
            if formula:
                c2 = ws_sum.cell(tdr, ci2, formula)
                c2.font = Font(bold=True, name="Arial")
                if ci2 == 5:
                    c2.number_format = "0.0%"
    for ci in range(1, 7):
        ws_sum.cell(tdr, ci).border = border
        ws_sum.cell(tdr, ci).alignment = center
    ws_sum.row_dimensions[tdr].height = 22

    # Sheet 3: Coordinate Frequency
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
        er2   = rank + 2
        count = missing_counts.get(coord, 0)
        pct   = count / total_frames if total_frames > 0 else 0.0
        argb2 = _interpolate_color(count, max_missing if max_missing > 0 else 1)
        t_val = count / max_missing if max_missing > 0 else 0
        flag_font = Font(name="Arial", size=10,
                         color="FFFFFFFF" if t_val > 0.45 else "FF000000",
                         bold=(count > 0))
        for ci, val in enumerate([coord, count, total_frames, pct, rank], 1):
            cell = ws_freq.cell(er2, ci, val)
            cell.alignment = center; cell.border = border
            cell.font = Font(name="Arial", size=10)
            if ci == 4:
                cell.number_format = "0.0%"
                cell.fill = PatternFill("solid", fgColor=argb2)
                cell.font = flag_font
            elif ci == 2 and count > 0:
                inten = min(count / total_frames, 1.0)
                rv = int(231 * inten + 236 * (1 - inten))
                gv = int(76  * inten + 240 * (1 - inten))
                bv = int(60  * inten + 241 * (1 - inten))
                cell.fill = PatternFill("solid", fgColor=f"FF{rv:02X}{gv:02X}{bv:02X}")
        ws_freq.row_dimensions[er2].height = 18

    footer_row = len(ALL_COORDS) + 3
    ws_freq.merge_cells(start_row=footer_row, start_column=1,
                        end_row=footer_row, end_column=len(freq_headers))
    total_events = sum(missing_counts.values())
    avg_miss = total_events / total_frames if total_frames > 0 else 0
    foot = ws_freq.cell(footer_row, 1,
        f"Total missing events: {total_events}   |   "
        f"Avg missing per frame: {avg_miss:.1f} / {len(ALL_COORDS)}   |   "
        f"Coordinates never flagged: {sum(1 for v in missing_counts.values() if v == 0)}")
    foot.font = Font(italic=True, name="Arial", size=9, color="FF555555")
    foot.alignment = Alignment(horizontal="left", vertical="center")
    ws_freq.row_dimensions[footer_row].height = 18

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── PDF export ─────────────────────────────────────────────────────────────────

def _rl_color(hex6: str):
    h = hex6.lstrip("#")
    return colors.Color(int(h[0:2], 16) / 255,
                        int(h[2:4], 16) / 255,
                        int(h[4:6], 16) / 255)


def _interp_rl_color(count: int, max_count: int):
    t = min(count / max_count, 1.0) if max_count > 0 else 0.0
    return colors.Color(
        (46  + (231 - 46)  * t) / 255,
        (204 + (76  - 204) * t) / 255,
        (113 + (60  - 113) * t) / 255)


def build_cumulative_heatmap_pdf(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=15*mm, bottomMargin=15*mm,
        title="Chocolate Mold Inspection Report")

    COL_DARK  = _rl_color("#2C3E50")
    COL_MID   = _rl_color("#34495E")
    WHITE     = colors.white
    BLACK     = colors.black

    title_style = ParagraphStyle(
        "title", fontName="Helvetica-Bold", fontSize=14,
        leading=18, textColor=WHITE, alignment=TA_CENTER)
    caption_style = ParagraphStyle(
        "caption", fontName="Helvetica-Oblique", fontSize=7,
        leading=10, textColor=colors.HexColor("#555555"))

    total_frames   = len(df)
    missing_counts = _compute_missing_counts(df)
    max_missing    = max(missing_counts.values()) if missing_counts else 1
    now_str        = datetime.now().strftime("%Y-%m-%d %H:%M")

    story = []

    def _hdr_para(txt, size=9):
        return Paragraph(f"<b>{txt}</b>", ParagraphStyle(
            "h", fontName="Helvetica-Bold", fontSize=size,
            textColor=WHITE, alignment=TA_CENTER))

    # ── Page 1: Cumulative Heatmap ─────────────────────────────────────────────
    title_tbl = Table(
        [[Paragraph(
            f"Chocolate Mold — Cumulative Missing Count "
            f"({total_frames} frame(s) inspected)", title_style)]],
        colWidths=[doc.width])
    title_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), COL_DARK),
        ("TOPPADDING",    (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    story.append(title_tbl)
    story.append(Spacer(1, 6*mm))

    row_hdr_w   = 14*mm
    tot_col_w   = 18*mm
    cell_w_each = (doc.width - row_hdr_w - tot_col_w) / len(COLS)
    col_widths  = [row_hdr_w] + [cell_w_each] * len(COLS) + [tot_col_w]

    hdr_row = [""] + [_hdr_para(c) for c in COLS] + [_hdr_para("Row\nTotal", 7)]
    grid_rows = [hdr_row]

    for row_num in ROWS:
        row = [_hdr_para(str(row_num))]
        row_total = 0
        for col in COLS:
            count = missing_counts.get(f"{col}{row_num}", 0)
            row_total += count
            row.append(str(count) if count > 0 else "")
        row.append(_hdr_para(str(row_total)))
        grid_rows.append(row)

    tot_row = [_hdr_para("Total")]
    grand_total = 0
    for col in COLS:
        ct = sum(missing_counts.get(f"{col}{r}", 0) for r in ROWS)
        grand_total += ct
        tot_row.append(_hdr_para(str(ct)))
    tot_row.append(_hdr_para(str(grand_total), 10))
    grid_rows.append(tot_row)

    ts_cmds = [
        ("FONTNAME",      (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE",      (0, 0), (-1, -1), 8),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",          (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("BACKGROUND",    (0, 0), (-1, 0),  COL_DARK),
        ("BACKGROUND",    (0, 1), (0, -2),  COL_DARK),
        ("BACKGROUND",    (-1, 1), (-1, -2), COL_MID),
        ("BACKGROUND",    (0, -1), (-1, -1), COL_MID),
        ("BACKGROUND",    (-1, -1), (-1, -1), _rl_color("#1A252F")),
    ]
    for ri, row_num in enumerate(ROWS, start=1):
        for ci, col in enumerate(COLS, start=1):
            count = missing_counts.get(f"{col}{row_num}", 0)
            bg    = _interp_rl_color(count, max_missing)
            ts_cmds.append(("BACKGROUND", (ci, ri), (ci, ri), bg))
            t = count / max_missing if max_missing > 0 else 0
            ts_cmds.append(("TEXTCOLOR", (ci, ri), (ci, ri),
                             WHITE if t > 0.45 else BLACK))

    Table(grid_rows, colWidths=col_widths).setStyle(TableStyle(ts_cmds))
    story.append(Table(grid_rows, colWidths=col_widths,
                       style=TableStyle(ts_cmds)))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph(
        f"Green = never missing  |  Red = missing in all {total_frames} frames  |  "
        f"Generated: {now_str}", caption_style))

    # ── Page 2: Frame Summary ──────────────────────────────────────────────────
    story.append(PageBreak())
    story.append(Table(
        [[Paragraph("Chocolate Mold Inspection — Frame Summary", title_style)]],
        colWidths=[doc.width],
        style=TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), COL_DARK),
            ("TOPPADDING",    (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ])))
    story.append(Spacer(1, 6*mm))

    sum_cw = [doc.width*0.28, doc.width*0.10, doc.width*0.10,
              doc.width*0.10, doc.width*0.12, doc.width*0.30]
    sum_hdrs = ["Frame", "Missing", "Present", "Total", "% Present", "Timestamp"]
    sum_rows = [[_hdr_para(h) for h in sum_hdrs]]

    for _, rec in df.iterrows():
        fname   = str(rec.get("frame_name", rec["frame_id"]))
        missing = sum(1 for c in ALL_COORDS if str(rec.get(c, "1")) == "0")
        present = len(ALL_COORDS) - missing
        total   = len(ALL_COORDS)
        pct_str = f"{present / total * 100:.1f}%"
        ts      = str(rec.get("timestamp", ""))
        sum_rows.append([
            Paragraph(fname, ParagraphStyle("fn", fontName="Helvetica", fontSize=8)),
            missing, present, total, pct_str,
            Paragraph(ts, ParagraphStyle("ts2", fontName="Helvetica", fontSize=7,
                                          textColor=colors.HexColor("#555555"),
                                          alignment=TA_CENTER)),
        ])

    total_missing_all = sum(
        sum(1 for c in ALL_COORDS if str(rec.get(c, "1")) == "0")
        for _, rec in df.iterrows())
    total_present_all = len(ALL_COORDS) * total_frames - total_missing_all
    total_total_all   = len(ALL_COORDS) * total_frames
    pct_all           = f"{total_present_all / total_total_all * 100:.1f}%" \
                        if total_total_all > 0 else "—"
    sum_rows.append([
        _hdr_para("TOTALS"), _hdr_para(str(total_missing_all)),
        _hdr_para(str(total_present_all)), _hdr_para(str(total_total_all)),
        _hdr_para(pct_all), ""])

    sum_ts = [
        ("FONTSIZE",      (0, 0), (-1, -1), 8),
        ("ALIGN",         (1, 0), (-1, -1), "CENTER"),
        ("ALIGN",         (0, 0), (0, -1),  "LEFT"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",          (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("BACKGROUND",    (0, 0), (-1, 0),  COL_DARK),
        ("BACKGROUND",    (0, -1), (-1, -1), COL_MID),
        ("LEFTPADDING",   (0, 0), (0, -1),   6),
    ]
    for ri in range(1, len(sum_rows) - 1):
        rec     = df.iloc[ri - 1]
        miss    = sum(1 for c in ALL_COORDS if str(rec.get(c, "1")) == "0")
        inten   = min(miss / len(ALL_COORDS), 1.0)
        bg      = _interp_rl_color(int(inten * max_missing), max_missing)
        sum_ts.append(("BACKGROUND", (1, ri), (1, ri), bg))
        sum_ts.append(("TEXTCOLOR",  (1, ri), (1, ri),
                        WHITE if inten > 0.45 else BLACK))

    story.append(Table(sum_rows, colWidths=sum_cw, style=TableStyle(sum_ts)))

    # ── Page 3: Coordinate Frequency ──────────────────────────────────────────
    story.append(PageBreak())
    story.append(Table(
        [[Paragraph(
            f"Coordinate Flag Frequency  —  {total_frames} frame(s) analysed",
            title_style)]],
        colWidths=[doc.width],
        style=TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), COL_DARK),
            ("TOPPADDING",    (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ])))
    story.append(Spacer(1, 6*mm))

    freq_cw   = [doc.width*0.15, doc.width*0.18, doc.width*0.15,
                 doc.width*0.15, doc.width*0.10]
    freq_hdrs = ["Coordinate", "Times Missing", "Total Frames", "% Flagged", "Rank"]
    freq_rows = [[_hdr_para(h) for h in freq_hdrs]]

    sorted_coords = sorted(ALL_COORDS,
                           key=lambda coord: (-missing_counts.get(coord, 0), coord))
    for rank, coord in enumerate(sorted_coords, start=1):
        count = missing_counts.get(coord, 0)
        pct   = f"{count / total_frames * 100:.1f}%" if total_frames > 0 else "0.0%"
        freq_rows.append([coord, count, total_frames, pct, rank])

    total_events  = sum(missing_counts.values())
    avg_miss      = total_events / total_frames if total_frames > 0 else 0
    never_flagged = sum(1 for v in missing_counts.values() if v == 0)
    freq_rows.append([
        Paragraph(
            f"<b>Total events: {total_events}  |  "
            f"Avg per frame: {avg_miss:.1f}  |  "
            f"Never flagged: {never_flagged}</b>",
            ParagraphStyle("ff", fontName="Helvetica-BoldOblique",
                           fontSize=7, textColor=WHITE)),
        "", "", "", ""])

    freq_ts = [
        ("FONTSIZE",      (0, 0), (-1, -1), 8),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",          (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
        ("TOPPADDING",    (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("BACKGROUND",    (0, 0), (-1, 0),  COL_DARK),
        ("BACKGROUND",    (0, -1), (-1, -1), COL_MID),
        ("SPAN",          (0, -1), (-1, -1)),
    ]
    for ri, coord in enumerate(sorted_coords, start=1):
        count = missing_counts.get(coord, 0)
        t_val = count / max_missing if max_missing > 0 else 0
        freq_ts.append(("BACKGROUND", (3, ri), (3, ri),
                         _interp_rl_color(count, max_missing)))
        freq_ts.append(("TEXTCOLOR",  (3, ri), (3, ri),
                         WHITE if t_val > 0.45 else BLACK))
        if count > 0:
            inten = min(count / total_frames, 1.0)
            freq_ts.append(("BACKGROUND", (1, ri), (1, ri),
                             colors.Color(
                                 (231*inten + 236*(1-inten))/255,
                                 (76 *inten + 240*(1-inten))/255,
                                 (60 *inten + 241*(1-inten))/255)))

    story.append(Table(freq_rows, colWidths=freq_cw, style=TableStyle(freq_ts)))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph(f"Generated: {now_str}", caption_style))

    doc.build(story)
    return buf.getvalue()


# ── Session-state helpers ──────────────────────────────────────────────────────

def init_state():
    defaults = {
        "df":              load_data(),
        "active_frame_id": None,
        "coord_dict":      {c: True for c in ALL_COORDS},
        "dirty":           False,
        "frame_image":     None,
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


def navigate_to_adjacent_frame(direction: int):
    df        = st.session_state.df
    frame_ids = df["frame_id"].tolist()
    if not frame_ids:
        return
    current = st.session_state.active_frame_id
    idx     = frame_ids.index(current) if current in frame_ids else 0
    load_frame(frame_ids[(idx + direction) % len(frame_ids)])


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
    .stat-box {
        border-radius: 10px; padding: 14px 18px; text-align: center;
        margin-bottom: 4px;
        background: var(--stat-bg, #16213E);
        border: 1px solid var(--stat-border, #2C3E50);
    }
    .stat-number { font-size: 1.9rem; font-weight: 700; }
    .stat-label  { font-size: 0.82rem; color: #95A5A6; margin-top: 3px; }
    .green { color: #27AE60; }
    .red   { color: #E74C3C; }
    .stButton > button { border-radius: 8px; font-weight: 600; }
    .upload-hint {
        border: 1px dashed #3498DB; border-radius: 8px;
        padding: 12px 16px; font-size: 0.88rem; margin-bottom: 8px;
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Sidebar ────────────────────────────────────────────────────────────────
    with st.sidebar:
        st.title("🍫 Mold Inspector")
        st.caption("15 × 8 grid | 120 positions")
        st.divider()

        st.subheader("📸 Upload Frame Photos")
        st.markdown(
            f'<div class="upload-hint">Upload up to {MAX_UPLOAD_FILES} images at once. '
            'A new record is created automatically per photo.</div>',
            unsafe_allow_html=True)

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
            for uf in uploaded_photos:
                base_name = os.path.splitext(uf.name)[0]
                existing  = st.session_state.df["frame_name"].tolist() \
                    if not st.session_state.df.empty else []
                if base_name in existing:
                    continue
                fid   = f"frame_{int(time.time() * 1000)}_{new_count}"
                ipath = save_frame_image(fid, uf)
                st.session_state.df = upsert_frame(
                    st.session_state.df, fid, base_name,
                    {c: True for c in ALL_COORDS}, ipath)
                new_count += 1
            if new_count:
                save_data(st.session_state.df)
                st.success(f"Imported {new_count} new frame(s).")
                load_frame(st.session_state.df["frame_id"].iloc[-1])
                st.rerun()

        st.divider()
        st.subheader("📋 Select Frame")
        df = st.session_state.df

        if st.button("🏠 All Frames Overview", use_container_width=True):
            st.session_state.active_frame_id = None
            st.rerun()

        if df.empty:
            st.info("No frames yet — upload photos above.")
        else:
            frame_options = df["frame_name"].tolist()
            frame_ids     = df["frame_id"].tolist()
            current_idx   = 0
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
                    st.session_state.df["frame_id"] != fid_to_del
                ].reset_index(drop=True)
                save_data(st.session_state.df)
                if st.session_state.active_frame_id == fid_to_del:
                    st.session_state.active_frame_id = None
                    st.session_state.frame_image     = None
                st.rerun()

            st.caption(f"Total frames: **{len(df)}**")

        st.divider()
        st.subheader("📥 Export")
        df = st.session_state.df

        if st.button("⬇️ Download Heatmap (.xlsx)",
                     use_container_width=True, disabled=df.empty):
            xlsx_bytes = build_cumulative_heatmap_workbook(df)
            now        = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                "💾 Save .xlsx", data=xlsx_bytes,
                file_name=f"mold_heatmap_{now}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

        if st.button("⬇️ Download Report (.pdf)",
                     use_container_width=True, disabled=df.empty):
            pdf_bytes = build_cumulative_heatmap_pdf(df)
            now       = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                "💾 Save .pdf", data=pdf_bytes,
                file_name=f"mold_report_{now}.pdf",
                mime="application/pdf",
                use_container_width=True)

        if not df.empty:
            st.download_button(
                "⬇️ Backup CSV",
                data=df.to_csv(index=False).encode(),
                file_name="mold_data_backup.csv",
                mime="text/csv", use_container_width=True)

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
        if df.empty:
            st.info("Upload frame photos in the sidebar to get started.")
            return

        st.subheader("All Frames — Overview")
        summary = []
        for _, rec in df.iterrows():
            missing = sum(1 for c in ALL_COORDS if str(rec.get(c, "1")) != "1")
            summary.append({
                "Frame":        rec["frame_name"],
                "Photo":        "✅" if str(rec.get("image_path", "")) else "—",
                "Missing":      missing,
                "Present":      len(ALL_COORDS) - missing,
                "% Present":    f"{(len(ALL_COORDS)-missing)/len(ALL_COORDS)*100:.1f}%",
                "Last Updated": rec.get("timestamp", ""),
            })
        st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)

        st.markdown("#### Load a frame")
        frame_ids   = df["frame_id"].tolist()
        frame_names = df["frame_name"].tolist()
        btn_cols    = st.columns(min(len(frame_ids), 5))
        for i, (fid, fname) in enumerate(zip(frame_ids, frame_names)):
            with btn_cols[i % 5]:
                if st.button(fname, key=f"ov_{fid}", use_container_width=True):
                    load_frame(fid)
                    st.rerun()
        return

    # ── Active frame ───────────────────────────────────────────────────────────
    active_id    = st.session_state.active_frame_id
    df           = st.session_state.df
    active_row   = df[df["frame_id"] == active_id].iloc[0]
    frame_name   = active_row["frame_name"]
    coord_dict   = st.session_state.coord_dict
    photo        = st.session_state.frame_image

    frame_ids    = df["frame_id"].tolist()
    current_idx  = frame_ids.index(active_id) if active_id in frame_ids else 0
    total_frames = len(frame_ids)

    missing_list  = [c for c, v in coord_dict.items() if not v]
    present_count = len(ALL_COORDS) - len(missing_list)

    # Header + prev/next navigation
    header_l, header_r = st.columns([3, 1])
    with header_l:
        st.markdown(f"## 🍫 {frame_name}")
        st.caption(
            f"Frame {current_idx + 1} of {total_frames}  |  "
            f"Saved: {active_row.get('timestamp', '—')}")
    with header_r:
        nav1, nav2 = st.columns(2)
        with nav1:
            if st.button("◀ Prev", use_container_width=True,
                         disabled=(total_frames <= 1)):
                navigate_to_adjacent_frame(-1); st.rerun()
        with nav2:
            if st.button("Next ▶", use_container_width=True,
                         disabled=(total_frames <= 1), type="primary"):
                navigate_to_adjacent_frame(+1); st.rerun()

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
            f'<div class="stat-label">Total Positions</div></div>',
            unsafe_allow_html=True)

    st.divider()

    # Per-frame photo attachment
    with st.expander(
            "📷 " + ("Replace frame photo" if photo else "Attach a photo to this frame"),
            expanded=(photo is None)):
        per_frame_upload = st.file_uploader(
            "Upload mold photo for this frame",
            type=["png", "jpg", "jpeg", "bmp", "tiff", "webp"],
            key=f"photo_{active_id}", label_visibility="collapsed")
        if per_frame_upload:
            ipath  = save_frame_image(active_id, per_frame_upload)
            photo  = Image.open(per_frame_upload).convert("RGB")
            st.session_state.frame_image = photo
            st.session_state.df = upsert_frame(
                st.session_state.df, active_id, frame_name, coord_dict, ipath)
            save_data(st.session_state.df)
            st.success("Photo saved.")
            st.rerun()

    # Tabs
    if photo:
        tab_photo, tab_grid, tab_quick, tab_vis = st.tabs([
            "🖼️ Photo Inspector", "🔲 Grid Editor",
            "⌨️ Quick Entry",     "📊 Overlay Preview"])
    else:
        tab_photo = None
        tab_grid, tab_quick, tab_vis = st.tabs([
            "🔲 Grid Editor", "⌨️ Quick Entry", "📊 Grid Preview"])

    # ── Tab: Photo Inspector ───────────────────────────────────────────────────
    # KEY FIX: every toggle calls auto_save() immediately — no dirty flag,
    # no Save button needed. The coord_dict in session_state AND the CSV are
    # always in sync, so switching tabs never loses anything.
    if tab_photo and photo:
        with tab_photo:

            # Top toolbar
            tc1, tc2, tc3, tc4, tc5 = st.columns([1, 1, 1, 1, 3])
            with tc1:
                if st.button("✅ All Present", use_container_width=True,
                             key="pi_all_pres"):
                    for c in ALL_COORDS:
                        st.session_state.coord_dict[c] = True
                    auto_save(active_id, frame_name)
                    st.rerun()
            with tc2:
                if st.button("❌ All Missing", use_container_width=True,
                             key="pi_all_miss"):
                    for c in ALL_COORDS:
                        st.session_state.coord_dict[c] = False
                    auto_save(active_id, frame_name)
                    st.rerun()
            with tc3:
                if st.button("🔄 Invert", use_container_width=True, key="pi_inv"):
                    for c in ALL_COORDS:
                        st.session_state.coord_dict[c] = \
                            not st.session_state.coord_dict[c]
                    auto_save(active_id, frame_name)
                    st.rerun()
            with tc4:
                opacity = st.slider("Overlay", 20, 220, 40, 10,
                                    key="pi_opacity",
                                    label_visibility="collapsed")
                st.caption("Opacity")
            with tc5:
                if missing_list:
                    badges = " ".join(
                        f"<span style='background:#E74C3C;color:#FFF;"
                        f"padding:1px 7px;border-radius:4px;font-size:0.78rem;"
                        f"margin:1px;display:inline-block'>{c}</span>"
                        for c in sorted(missing_list))
                    st.markdown(
                        f"<div style='line-height:2;padding-top:2px'>"
                        f"<strong style='color:#E74C3C'>"
                        f"Missing ({len(missing_list)}):</strong> {badges}</div>",
                        unsafe_allow_html=True)
                else:
                    st.success("All 120 positions present!")

            st.markdown("<hr style='margin:4px 0 8px'>", unsafe_allow_html=True)

            # ── Two-column layout: overlay image (left) + coord grid (right) ──
            img_col, grid_col = st.columns([3, 2])

            with img_col:
                # Render and display the overlay image (static, for reference)
                cur_opacity = st.session_state.get("pi_opacity", 40)
                overlay_img = render_overlay_on_photo(
                    photo, st.session_state.coord_dict, opacity=cur_opacity)
                st.image(overlay_img, use_container_width=True, caption=frame_name)

                buf_dl = io.BytesIO()
                overlay_img.save(buf_dl, format="PNG")
                st.download_button(
                    "⬇️ Download overlay image",
                    data=buf_dl.getvalue(),
                    file_name=f"{re.sub(r'[^a-zA-Z0-9_-]','_',frame_name)}_overlay.png",
                    mime="image/png", use_container_width=True)

            with grid_col:
                # ── Clickable coordinate grid ──────────────────────────────────
                # This is the RELIABLE toggle mechanism.
                # Every button click calls auto_save() — persisted immediately.
                st.markdown(
                    "**Click any cell** to toggle it.  "
                    "🟢 = present &nbsp; 🔴 = missing  \n"
                    "_Changes save instantly._")

                # Column headers
                hcols = st.columns([0.4] + [1] * len(COLS))
                hcols[0].markdown(
                    "<div style='text-align:center;font-size:0.7rem'><b>↓</b></div>",
                    unsafe_allow_html=True)
                for i, lbl in enumerate(COLS):
                    hcols[i + 1].markdown(
                        f"<div style='text-align:center;font-weight:700;"
                        f"color:#3498DB;font-size:0.72rem'>{lbl}</div>",
                        unsafe_allow_html=True)

                # Data rows — each button auto_saves on click
                for row_num in ROWS:
                    rcols = st.columns([0.4] + [1] * len(COLS))
                    rcols[0].markdown(
                        f"<div style='text-align:center;font-weight:700;"
                        f"color:#3498DB;font-size:0.72rem'>{row_num}</div>",
                        unsafe_allow_html=True)
                    for ci, col_lbl in enumerate(COLS):
                        coord   = f"{col_lbl}{row_num}"
                        present = st.session_state.coord_dict.get(coord, True)
                        with rcols[ci + 1]:
                            if st.button(
                                "🟢" if present else "🔴",
                                key=f"pi_btn_{coord}",
                                help=f"{coord}: "
                                     f"{'Present — click to mark missing' if present else 'MISSING — click to restore'}",
                                use_container_width=True,
                            ):
                                # Toggle and IMMEDIATELY persist
                                st.session_state.coord_dict[coord] = not present
                                auto_save(active_id, frame_name)
                                st.rerun()

    # ── Tab: Grid Editor ──────────────────────────────────────────────────────
    with tab_grid:
        st.markdown("Click a cell to toggle **Present 🟢** ↔ **Missing 🔴**, "
                    "then hit **Save Frame**.")

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
            if st.button("🔄 Invert", use_container_width=True, key="g_inv"):
                for c in ALL_COORDS:
                    st.session_state.coord_dict[c] = not st.session_state.coord_dict[c]
                st.session_state.dirty = True; st.rerun()
        with tb4:
            if st.button("💾 Save Frame", use_container_width=True,
                         type="primary", key="g_save",
                         disabled=not st.session_state.dirty):
                auto_save(active_id, frame_name)
                st.success("Saved!"); st.rerun()

        hcols = st.columns([0.5] + [1] * len(COLS))
        hcols[0].markdown("**↓**")
        for i, lbl in enumerate(COLS):
            hcols[i + 1].markdown(
                f"<div style='text-align:center;font-weight:700;"
                f"color:#3498DB;font-size:0.8rem'>{lbl}</div>",
                unsafe_allow_html=True)

        for row_num in ROWS:
            rcols = st.columns([0.5] + [1] * len(COLS))
            rcols[0].markdown(
                f"<div style='text-align:center;font-weight:700;"
                f"color:#3498DB'>{row_num}</div>", unsafe_allow_html=True)
            for ci, col_lbl in enumerate(COLS):
                coord   = f"{col_lbl}{row_num}"
                present = st.session_state.coord_dict.get(coord, True)
                with rcols[ci + 1]:
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
            apply_btn = ca.form_submit_button(
                "Apply (replace all missing)", use_container_width=True, type="primary")
            add_btn   = cb.form_submit_button(
                "Add to existing missing", use_container_width=True)

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
        st.dataframe(pd.DataFrame(status_rows), use_container_width=True,
                     hide_index=True)

    # ── Tab: Overlay / Grid Preview ───────────────────────────────────────────
    with tab_vis:
        if photo:
            st.markdown("Overlay on mold photo. 🟢 Present  🔴 Missing")
            opa = st.slider("Overlay opacity", 20, 220, 40, 10, key="vis_opacity")
            ov  = render_overlay_on_photo(photo, st.session_state.coord_dict,
                                          opacity=opa)
            st.image(ov, use_container_width=True, caption=frame_name)
            buf2 = io.BytesIO(); ov.save(buf2, format="PNG")
            st.download_button(
                "⬇️ Download overlay (.png)", data=buf2.getvalue(),
                file_name=f"{re.sub(r'[^a-zA-Z0-9_-]','_',frame_name)}_overlay.png",
                mime="image/png")
        else:
            st.markdown("Synthetic grid. 🟢 Present  🔴 Missing")
            grid_img = render_plain_grid(st.session_state.coord_dict)
            st.image(grid_img, use_container_width=True, caption=frame_name)
            buf2 = io.BytesIO(); grid_img.save(buf2, format="PNG")
            st.download_button(
                "⬇️ Download grid (.png)", data=buf2.getvalue(),
                file_name=f"{re.sub(r'[^a-zA-Z0-9_-]','_',frame_name)}_grid.png",
                mime="image/png")

        if missing_list:
            st.subheader(f"🔴 Missing positions ({len(missing_list)})")
            rows_of_missing: dict = {}
            for c in sorted(missing_list):
                rows_of_missing.setdefault(c[1:], []).append(c)
            for rn, coords_in_row in sorted(rows_of_missing.items(),
                                            key=lambda x: int(x[0])):
                st.markdown(
                    f"**Row {rn}:** " + " ".join(
                        f"<span style='background:#E74C3C;color:#FFF;"
                        f"padding:2px 6px;border-radius:4px;font-size:0.85rem'>{c}</span>"
                        for c in coords_in_row),
                    unsafe_allow_html=True)
        else:
            st.success("🎉 All 120 positions are present!")

    # Bottom navigation
    st.divider()
    bot_l, bot_m, bot_r = st.columns([1, 2, 1])
    with bot_l:
        if st.button("◄ Previous Frame", use_container_width=True,
                     disabled=(total_frames <= 1)):
            navigate_to_adjacent_frame(-1); st.rerun()
    with bot_m:
        st.markdown(
            f"<div style='text-align:center;padding-top:8px'>"
            f"Frame <strong>{current_idx + 1}</strong> "
            f"of <strong>{total_frames}</strong></div>",
            unsafe_allow_html=True)
    with bot_r:
        if st.button("Next Frame ►", use_container_width=True,
                     disabled=(total_frames <= 1), type="primary"):
            navigate_to_adjacent_frame(+1); st.rerun()


if __name__ == "__main__":
    main()
