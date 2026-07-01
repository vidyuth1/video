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

# ReportLab imports for PDF export
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
    HRFlowable,
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ── Constants ──────────────────────────────────────────────────────────────────
COLS       = list("ABCDEFGHIJKLMNO")          # 15 columns  (A–O)
ROWS       = list(range(1, 9))                 # 8 rows       (1–8)
ALL_COORDS = [f"{c}{r}" for r in ROWS for c in COLS]

DATA_FILE        = "mold_data.csv"
IMAGES_DIR       = "frame_images"
MAX_UPLOAD_FILES = 50

PALETTE = {
    "present":  "#2ECC71",
    "empty":    "#E74C3C",
    "selected": "#F39C12",
}

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


# ── Image helpers ──────────────────────────────────────────────────────────────

def pil_to_b64(img: Image.Image, fmt="PNG") -> str:
    buf = io.BytesIO()
    img.save(buf, format=fmt)
    return base64.b64encode(buf.getvalue()).decode()


def render_overlay_on_photo(photo: Image.Image, coord_dict: dict,
                             opacity: int = 40) -> Image.Image:
    """Draw a semi-transparent 15×8 grid overlay on the mold photo."""
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
            # Both present and missing use the same opacity value
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
        draw.text((x, label_px // 2), col, fill="#ECF0F1",
                  font=fnt_lbl, anchor="mm")

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
    if max_count == 0:
        t = 0.0
    else:
        t = min(count / max_count, 1.0)
    r = int(46  + (231 - 46)  * t)
    g = int(204 + (76  - 204) * t)
    b = int(113 + (60  - 113) * t)
    return f"FF{r:02X}{g:02X}{b:02X}"


def _compute_missing_counts(df: pd.DataFrame) -> dict:
    counts = {}
    for coord in ALL_COORDS:
        if coord in df.columns:
            counts[coord] = int((df[coord].astype(str) == "0").sum())
        else:
            counts[coord] = 0
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

    total_frames  = len(df)
    missing_counts = _compute_missing_counts(df)
    max_missing   = max(missing_counts.values()) if missing_counts else 1

    # ── Sheet 1: Cumulative Heatmap ────────────────────────────────────────────
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
        f"Green = never missing  to  Red = missing in all {total_frames} frame(s)  |  "
        f"Total inspected: {total_frames} frames, {len(ALL_COORDS)} positions each")
    leg.font = Font(italic=True, name="Arial", size=9, color="FF555555")
    leg.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[legend_row].height = 18

    # ── Sheet 2: Frame Summary ─────────────────────────────────────────────────
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

    # ── Sheet 3: Coordinate Frequency ─────────────────────────────────────────
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
        er2 = rank + 2
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
    """Convert '#RRGGBB' to a ReportLab Color."""
    h = hex6.lstrip("#")
    return colors.Color(int(h[0:2], 16) / 255,
                        int(h[2:4], 16) / 255,
                        int(h[4:6], 16) / 255)


def _interp_rl_color(count: int, max_count: int):
    if max_count == 0:
        t = 0.0
    else:
        t = min(count / max_count, 1.0)
    r = (46  + (231 - 46)  * t) / 255
    g = (204 + (76  - 204) * t) / 255
    b = (113 + (60  - 113) * t) / 255
    return colors.Color(r, g, b)


def build_cumulative_heatmap_pdf(df: pd.DataFrame) -> bytes:
    """
    Generate a multi-page PDF report matching the Excel export:
      Page 1  — Cumulative heatmap grid (landscape A4)
      Page 2  — Frame summary table
      Page 3  — Coordinate frequency table
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
        title="Chocolate Mold Inspection Report",
        author="Mold Inspector",
    )

    # ── Shared styles ──────────────────────────────────────────────────────────
    COL_DARK  = _rl_color("#2C3E50")
    COL_MID   = _rl_color("#34495E")
    COL_LIGHT = _rl_color("#ECF0F1")
    COL_GREEN = _rl_color("#2ECC71")
    COL_RED   = _rl_color("#E74C3C")
    WHITE     = colors.white
    BLACK     = colors.black

    base_style = ParagraphStyle(
        "base", fontName="Helvetica", fontSize=9, leading=13, textColor=BLACK)
    title_style = ParagraphStyle(
        "title", fontName="Helvetica-Bold", fontSize=14,
        leading=18, textColor=WHITE, alignment=TA_CENTER)
    h2_style = ParagraphStyle(
        "h2", fontName="Helvetica-Bold", fontSize=11,
        leading=15, textColor=COL_DARK, spaceAfter=4)
    caption_style = ParagraphStyle(
        "caption", fontName="Helvetica-Oblique", fontSize=7,
        leading=10, textColor=colors.HexColor("#555555"))

    total_frames   = len(df)
    missing_counts = _compute_missing_counts(df)
    max_missing    = max(missing_counts.values()) if missing_counts else 1
    now_str        = datetime.now().strftime("%Y-%m-%d %H:%M")

    story = []

    # ════════════════════════════════════════════════════════════════════════════
    # PAGE 1 — Cumulative Heatmap Grid
    # ════════════════════════════════════════════════════════════════════════════

    # Title banner
    title_data = [[Paragraph(
        f"Chocolate Mold — Cumulative Missing Count  "
        f"({total_frames} frame(s) inspected)", title_style)]]
    title_tbl = Table(title_data, colWidths=[doc.width])
    title_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), COL_DARK),
        ("TOPPADDING",    (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("ROUNDEDCORNERS", [4]),
    ]))
    story.append(title_tbl)
    story.append(Spacer(1, 6 * mm))

    # Build heatmap table
    # Row 0: header (blank + col letters + "Total")
    # Rows 1–8: row number + counts + row total
    # Last row: "Total" + col totals + grand total

    page_w = doc.width  # usable width in landscape
    n_data_cols = len(COLS)
    row_hdr_w   = 14 * mm
    tot_col_w   = 18 * mm
    cell_w_each = (page_w - row_hdr_w - tot_col_w) / n_data_cols

    col_widths = [row_hdr_w] + [cell_w_each] * n_data_cols + [tot_col_w]

    # Header row
    hdr_row = [""] + [Paragraph(f"<b>{c}</b>", ParagraphStyle(
        "ch", fontName="Helvetica-Bold", fontSize=8, textColor=WHITE,
        alignment=TA_CENTER)) for c in COLS] + \
        [Paragraph("<b>Row\nTotal</b>", ParagraphStyle(
            "ct", fontName="Helvetica-Bold", fontSize=7, textColor=WHITE,
            alignment=TA_CENTER, leading=9))]

    grid_rows = [hdr_row]

    for row_num in ROWS:
        row = [Paragraph(f"<b>{row_num}</b>", ParagraphStyle(
            "rh", fontName="Helvetica-Bold", fontSize=9, textColor=WHITE,
            alignment=TA_CENTER))]
        row_total = 0
        for col in COLS:
            count = missing_counts.get(f"{col}{row_num}", 0)
            row_total += count
            row.append(str(count) if count > 0 else "")
        row.append(Paragraph(f"<b>{row_total}</b>", ParagraphStyle(
            "rt", fontName="Helvetica-Bold", fontSize=9, textColor=WHITE,
            alignment=TA_CENTER)))
        grid_rows.append(row)

    # Totals row
    tot_row = [Paragraph("<b>Total</b>", ParagraphStyle(
        "tl", fontName="Helvetica-Bold", fontSize=9, textColor=WHITE,
        alignment=TA_CENTER))]
    grand_total = 0
    for col in COLS:
        ct = sum(missing_counts.get(f"{col}{r}", 0) for r in ROWS)
        grand_total += ct
        tot_row.append(Paragraph(f"<b>{ct}</b>", ParagraphStyle(
            "tv", fontName="Helvetica-Bold", fontSize=9, textColor=WHITE,
            alignment=TA_CENTER)))
    tot_row.append(Paragraph(f"<b>{grand_total}</b>", ParagraphStyle(
        "gv", fontName="Helvetica-Bold", fontSize=10, textColor=WHITE,
        alignment=TA_CENTER)))
    grid_rows.append(tot_row)

    grid_tbl = Table(grid_rows, colWidths=col_widths, rowHeights=None)

    # Base style commands
    ts_cmds = [
        ("FONTNAME",    (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE",    (0, 0), (-1, -1), 8),
        ("ALIGN",       (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",        (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
        ("TOPPADDING",  (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        # Header row background
        ("BACKGROUND",  (0, 0), (-1, 0), COL_DARK),
        # Row header column
        ("BACKGROUND",  (0, 1), (0, -2), COL_DARK),
        # Row total column
        ("BACKGROUND",  (-1, 1), (-1, -2), COL_MID),
        # Totals row
        ("BACKGROUND",  (0, -1), (-1, -1), COL_MID),
        # Grand total corner
        ("BACKGROUND",  (-1, -1), (-1, -1), _rl_color("#1A252F")),
    ]

    # Per-cell background colors based on missing count
    for ri, row_num in enumerate(ROWS, start=1):
        for ci, col in enumerate(COLS, start=1):
            coord = f"{col}{row_num}"
            count = missing_counts.get(coord, 0)
            bg    = _interp_rl_color(count, max_missing)
            ts_cmds.append(("BACKGROUND", (ci, ri), (ci, ri), bg))
            # Text color: white for darker (more missing) cells
            t = count / max_missing if max_missing > 0 else 0
            txt_c = WHITE if t > 0.45 else BLACK
            ts_cmds.append(("TEXTCOLOR", (ci, ri), (ci, ri), txt_c))

    grid_tbl.setStyle(TableStyle(ts_cmds))
    story.append(grid_tbl)
    story.append(Spacer(1, 4 * mm))

    # Legend
    story.append(Paragraph(
        f"Each cell = number of frames where that cavity was missing.  "
        f"Green = never missing  |  Red = missing in all {total_frames} frame(s).  "
        f"Generated: {now_str}",
        caption_style))

    # ════════════════════════════════════════════════════════════════════════════
    # PAGE 2 — Frame Summary
    # ════════════════════════════════════════════════════════════════════════════
    story.append(PageBreak())

    story.append(Table(
        [[Paragraph("Chocolate Mold Inspection — Frame Summary", title_style)]],
        colWidths=[doc.width],
    ))
    story.append(Spacer(1, 6 * mm))

    sum_headers_pdf = ["Frame", "Missing", "Present", "Total", "% Present", "Timestamp"]
    sum_col_widths  = [
        doc.width * 0.28,
        doc.width * 0.10,
        doc.width * 0.10,
        doc.width * 0.10,
        doc.width * 0.12,
        doc.width * 0.30,
    ]

    def _hdr_para(txt):
        return Paragraph(f"<b>{txt}</b>", ParagraphStyle(
            "sh", fontName="Helvetica-Bold", fontSize=9,
            textColor=WHITE, alignment=TA_CENTER))

    sum_rows = [[_hdr_para(h) for h in sum_headers_pdf]]

    for _, rec in df.iterrows():
        fname   = str(rec.get("frame_name", rec["frame_id"]))
        missing = sum(1 for c in ALL_COORDS if str(rec.get(c, "1")) == "0")
        present = len(ALL_COORDS) - missing
        total   = len(ALL_COORDS)
        pct_str = f"{present / total * 100:.1f}%"
        ts      = str(rec.get("timestamp", ""))
        # Color for missing cell
        inten  = min(missing / total, 1.0)
        miss_c = _interp_rl_color(int(inten * max_missing), max_missing)
        sum_rows.append([
            Paragraph(fname, ParagraphStyle("fn", fontName="Helvetica", fontSize=8,
                                             textColor=BLACK)),
            missing, present, total, pct_str,
            Paragraph(ts, ParagraphStyle("ts2", fontName="Helvetica", fontSize=7,
                                          textColor=colors.HexColor("#555555"),
                                          alignment=TA_CENTER)),
        ])

    # Totals row
    total_missing_all = sum(
        sum(1 for c in ALL_COORDS if str(rec.get(c, "1")) == "0")
        for _, rec in df.iterrows())
    total_present_all = len(ALL_COORDS) * total_frames - total_missing_all
    total_total_all   = len(ALL_COORDS) * total_frames
    pct_all           = f"{total_present_all / total_total_all * 100:.1f}%" \
                        if total_total_all > 0 else "—"
    sum_rows.append([
        Paragraph("<b>TOTALS</b>", ParagraphStyle(
            "tot", fontName="Helvetica-Bold", fontSize=9, textColor=WHITE)),
        Paragraph(f"<b>{total_missing_all}</b>", ParagraphStyle(
            "tm", fontName="Helvetica-Bold", fontSize=9, textColor=WHITE,
            alignment=TA_CENTER)),
        Paragraph(f"<b>{total_present_all}</b>", ParagraphStyle(
            "tp", fontName="Helvetica-Bold", fontSize=9, textColor=WHITE,
            alignment=TA_CENTER)),
        Paragraph(f"<b>{total_total_all}</b>", ParagraphStyle(
            "tt", fontName="Helvetica-Bold", fontSize=9, textColor=WHITE,
            alignment=TA_CENTER)),
        Paragraph(f"<b>{pct_all}</b>", ParagraphStyle(
            "tpct", fontName="Helvetica-Bold", fontSize=9, textColor=WHITE,
            alignment=TA_CENTER)),
        "",
    ])

    sum_tbl = Table(sum_rows, colWidths=sum_col_widths)
    sum_ts = [
        ("FONTSIZE",    (0, 0), (-1, -1), 8),
        ("ALIGN",       (1, 0), (-1, -1), "CENTER"),
        ("ALIGN",       (0, 0), (0, -1), "LEFT"),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",        (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
        ("TOPPADDING",  (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("BACKGROUND",  (0, 0), (-1, 0), COL_DARK),
        ("BACKGROUND",  (0, -1), (-1, -1), COL_MID),
        ("LEFTPADDING", (0, 0), (0, -1), 6),
    ]
    # Color missing column per row
    for ri in range(1, len(sum_rows) - 1):
        rec_idx = ri - 1
        rec     = df.iloc[rec_idx]
        miss    = sum(1 for c in ALL_COORDS if str(rec.get(c, "1")) == "0")
        inten   = min(miss / len(ALL_COORDS), 1.0)
        bg      = _interp_rl_color(int(inten * max_missing), max_missing)
        sum_ts.append(("BACKGROUND", (1, ri), (1, ri), bg))
        txt_c = WHITE if inten > 0.45 else BLACK
        sum_ts.append(("TEXTCOLOR", (1, ri), (1, ri), txt_c))

    sum_tbl.setStyle(TableStyle(sum_ts))
    story.append(sum_tbl)

    # ════════════════════════════════════════════════════════════════════════════
    # PAGE 3 — Coordinate Frequency
    # ════════════════════════════════════════════════════════════════════════════
    story.append(PageBreak())

    story.append(Table(
        [[Paragraph(
            f"Coordinate Flag Frequency  —  {total_frames} frame(s) analysed",
            title_style)]],
        colWidths=[doc.width],
    ))
    story.append(Spacer(1, 6 * mm))

    sorted_coords = sorted(ALL_COORDS,
                           key=lambda coord: (-missing_counts.get(coord, 0), coord))

    freq_col_widths = [
        doc.width * 0.15,
        doc.width * 0.18,
        doc.width * 0.15,
        doc.width * 0.15,
        doc.width * 0.10,
    ]
    freq_hdr = ["Coordinate", "Times Missing", "Total Frames", "% Flagged", "Rank"]
    freq_rows = [[_hdr_para(h) for h in freq_hdr]]

    for rank, coord in enumerate(sorted_coords, start=1):
        count = missing_counts.get(coord, 0)
        pct   = f"{count / total_frames * 100:.1f}%" if total_frames > 0 else "0.0%"
        freq_rows.append([coord, count, total_frames, pct, rank])

    # Footer summary row
    total_events = sum(missing_counts.values())
    avg_miss     = total_events / total_frames if total_frames > 0 else 0
    never_flagged = sum(1 for v in missing_counts.values() if v == 0)
    freq_rows.append([
        Paragraph(
            f"<b>Total events: {total_events}  |  "
            f"Avg per frame: {avg_miss:.1f}  |  "
            f"Never flagged: {never_flagged}</b>",
            ParagraphStyle("ffoot", fontName="Helvetica-BoldOblique",
                           fontSize=7, textColor=WHITE)),
        "", "", "", "",
    ])

    freq_tbl = Table(freq_rows, colWidths=freq_col_widths)
    freq_ts = [
        ("FONTSIZE",    (0, 0), (-1, -1), 8),
        ("ALIGN",       (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",        (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
        ("TOPPADDING",  (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("BACKGROUND",  (0, 0), (-1, 0), COL_DARK),
        ("BACKGROUND",  (0, -1), (-1, -1), COL_MID),
        ("SPAN",        (0, -1), (-1, -1)),
    ]
    # Color the % Flagged column per coord row
    for ri, coord in enumerate(sorted_coords, start=1):
        count = missing_counts.get(coord, 0)
        t_val = count / max_missing if max_missing > 0 else 0
        bg    = _interp_rl_color(count, max_missing)
        freq_ts.append(("BACKGROUND", (3, ri), (3, ri), bg))
        txt_c = WHITE if t_val > 0.45 else BLACK
        freq_ts.append(("TEXTCOLOR", (3, ri), (3, ri), txt_c))
        # Subtle tint on Times Missing when non-zero
        if count > 0:
            inten = min(count / total_frames, 1.0)
            r2 = (231 * inten + 236 * (1 - inten)) / 255
            g2 = (76  * inten + 240 * (1 - inten)) / 255
            b2 = (60  * inten + 241 * (1 - inten)) / 255
            freq_ts.append(("BACKGROUND", (1, ri), (1, ri),
                             colors.Color(r2, g2, b2)))

    freq_tbl.setStyle(TableStyle(freq_ts))
    story.append(freq_tbl)
    story.append(Spacer(1, 4 * mm))
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
        "show_overview":   False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def load_frame(frame_id: str):
    st.session_state.active_frame_id = frame_id
    st.session_state.coord_dict      = get_frame_dict(st.session_state.df, frame_id)
    st.session_state.dirty           = False
    st.session_state.show_overview   = False
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
    df        = st.session_state.df
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

    # ── CSS — light/dark mode adaptive ────────────────────────────────────────
    st.markdown("""
    <style>
    /* ── Stat boxes ── */
    .stat-box {
        border-radius: 10px;
        padding: 14px 18px;
        text-align: center;
        margin-bottom: 4px;
        background: var(--stat-bg, #16213E);
        border: 1px solid var(--stat-border, #2C3E50);
    }
    .stat-number { font-size: 1.9rem; font-weight: 700; }
    .stat-label  { font-size: 0.82rem; color: var(--muted, #95A5A6); margin-top: 3px; }
    .green { color: #27AE60; }
    .red   { color: #E74C3C; }

    /* ── Buttons ── */
    .stButton > button { border-radius: 8px; font-weight: 600; }

    /* ── Upload hint ── */
    .upload-hint {
        border: 1px dashed #3498DB;
        border-radius: 8px;
        padding: 12px 16px;
        font-size: 0.88rem;
        margin-bottom: 8px;
        background: var(--hint-bg, transparent);
        color: var(--hint-text, inherit);
    }

    /* ── Dark mode overrides ── */
    @media (prefers-color-scheme: dark) {
        .stat-box   { --stat-bg: #16213E; --stat-border: #2C3E50; }
        .stat-label { color: #95A5A6; }
        .upload-hint { --hint-bg: #16213E; --hint-text: #BDC3C7; }
    }

    /* ── Light mode overrides ── */
    @media (prefers-color-scheme: light) {
        .stat-box   { --stat-bg: #F4F6F8; --stat-border: #D5D8DC; }
        .stat-label { color: #6C757D; }
        .upload-hint { --hint-bg: #EAF4FB; --hint-text: #2C3E50; }
    }

    /* Streamlit's own theme vars already handle most text/bg;
       the above targets only our custom HTML blocks. */
    </style>
    """, unsafe_allow_html=True)

    # ── Sidebar ────────────────────────────────────────────────────────────────
    with st.sidebar:
        st.title("🍫 Mold Inspector")
        st.caption("15 × 8 grid | 120 positions")
        st.divider()

        # Upload
        st.subheader("📸 Upload Frame Photos")
        st.markdown(
            f'<div class="upload-hint">Upload up to {MAX_UPLOAD_FILES} frame images '
            'at once. A new record is created automatically per photo.</div>',
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

        # Frame selector
        st.subheader("📋 Select Frame")
        df = st.session_state.df

        # Homepage / overview button
        if st.button("🏠 All Frames Overview", use_container_width=True):
            st.session_state.show_overview   = True
            st.session_state.active_frame_id = None
            st.rerun()

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
                    st.session_state.df["frame_id"] != fid_to_del
                ].reset_index(drop=True)
                save_data(st.session_state.df)
                if st.session_state.active_frame_id == fid_to_del:
                    st.session_state.active_frame_id = None
                    st.session_state.frame_image     = None
                st.rerun()

            st.caption(f"Total frames: **{len(df)}**")

        st.divider()

        # Export
        st.subheader("📥 Export")
        df = st.session_state.df

        # XLSX
        if st.button("⬇️ Download Heatmap (.xlsx)",
                     use_container_width=True, disabled=df.empty):
            xlsx_bytes = build_cumulative_heatmap_workbook(df)
            now        = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                "💾 Save .xlsx", data=xlsx_bytes,
                file_name=f"mold_heatmap_{now}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

        # PDF
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

    # ── Homepage / overview ────────────────────────────────────────────────────
    if st.session_state.active_frame_id is None or st.session_state.show_overview:
        st.markdown("## 🍫 Chocolate Mold Inspector")

        if df.empty:
            st.info("Upload frame photos in the sidebar to get started.")
            return

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

        # Quick-load buttons for each frame
        st.markdown("#### Load a frame to begin inspection")
        frame_ids   = df["frame_id"].tolist()
        frame_names = df["frame_name"].tolist()
        btn_cols    = st.columns(min(len(frame_ids), 5))
        for i, (fid, fname) in enumerate(zip(frame_ids, frame_names)):
            with btn_cols[i % 5]:
                if st.button(fname, key=f"overview_load_{fid}",
                             use_container_width=True):
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

    # ── Header + navigation ────────────────────────────────────────────────────
    header_left, header_right = st.columns([3, 1])
    with header_left:
        st.markdown(f"## 🍫 {frame_name}")
        st.caption(
            f"Frame {current_idx + 1} of {total_frames}  |  "
            f"ID: `{active_id}`  |  Saved: {active_row.get('timestamp', '—')}"
        )
    with header_right:
        nav_c1, nav_c2 = st.columns(2)
        with nav_c1:
            if st.button("◀ Prev", use_container_width=True,
                         disabled=(total_frames <= 1), help="Previous frame"):
                navigate_to_adjacent_frame(-1)
                st.rerun()
        with nav_c2:
            if st.button("Next ▶", use_container_width=True,
                         disabled=(total_frames <= 1), help="Next frame",
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
            f'<div class="stat-label">Total Positions</div></div>',
            unsafe_allow_html=True)

    st.divider()

    # Photo upload expander
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
            st.success("Photo saved to this frame.")
            st.rerun()

    # Tabs
    if photo:
        tab_photo, tab_grid, tab_quick, tab_vis = st.tabs([
            "🖼️ Photo Inspector", "🔲 Grid Editor", "⌨️ Quick Entry",
            "📊 Overlay Preview"])
    else:
        tab_grid, tab_quick, tab_vis = st.tabs([
            "🔲 Grid Editor", "⌨️ Quick Entry", "📊 Grid Preview"])
        tab_photo = None

    # ── Tab: Photo Inspector ───────────────────────────────────────────────────
    if tab_photo and photo:
        with tab_photo:
            # Process canvas click from query params (single-click relay)
            qp = st.query_params
            pending_click = qp.get("click", "")
            if pending_click and pending_click in ALL_COORDS:
                # Toggle in session state, mark dirty — do NOT auto-save
                st.session_state.coord_dict[pending_click] = \
                    not st.session_state.coord_dict.get(pending_click, True)
                st.session_state.dirty = True
                st.query_params.clear()
                st.rerun()

            # Top control strip
            ctrl_c1, ctrl_c2, ctrl_c3, ctrl_c4, ctrl_c5 = st.columns([1, 1, 1, 1, 3])
            with ctrl_c1:
                if st.button("✅ All Present", use_container_width=True,
                             key="pi_all_pres"):
                    for c in ALL_COORDS:
                        st.session_state.coord_dict[c] = True
                    st.session_state.dirty = True
                    st.rerun()
            with ctrl_c2:
                if st.button("❌ All Missing", use_container_width=True,
                             key="pi_all_miss"):
                    for c in ALL_COORDS:
                        st.session_state.coord_dict[c] = False
                    st.session_state.dirty = True
                    st.rerun()
            with ctrl_c3:
                if st.button("🔄 Invert", use_container_width=True, key="pi_invert"):
                    for c in ALL_COORDS:
                        st.session_state.coord_dict[c] = \
                            not st.session_state.coord_dict[c]
                    st.session_state.dirty = True
                    st.rerun()
            with ctrl_c4:
                opacity = st.slider("Overlay opacity", 20, 220, 40, 10,
                                    key="overlay_opacity",
                                    label_visibility="collapsed")
                st.caption("Opacity")
            with ctrl_c5:
                if missing_list:
                    badges = " ".join(
                        f"<span style='background:#E74C3C;color:#FFF;"
                        f"padding:1px 7px;border-radius:4px;font-size:0.78rem;"
                        f"margin:1px;display:inline-block'>{c}</span>"
                        for c in sorted(missing_list))
                    st.markdown(
                        f"<div style='line-height:2.0;padding-top:2px'>"
                        f"<strong style='color:#E74C3C'>"
                        f"Missing ({len(missing_list)}):</strong> {badges}</div>",
                        unsafe_allow_html=True)
                else:
                    st.success("All 120 positions present!")

            st.markdown("<hr style='margin:4px 0 8px'>", unsafe_allow_html=True)

            # Build overlay
            cur_opacity = st.session_state.get("overlay_opacity", 40)
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
  position: relative; display: block; width: 100%; line-height: 0;
}}
#{canvas_id} {{
  width: 100%; height: auto; display: block;
  border-radius: 6px; cursor: crosshair;
  box-shadow: 0 2px 16px rgba(0,0,0,0.35);
}}
#{canvas_id}_tip {{
  position: fixed; background: rgba(15,15,15,0.88); color: #fff;
  padding: 5px 13px; border-radius: 6px;
  font: 700 13px/1.5 Arial,sans-serif; pointer-events: none;
  display: none; z-index: 9999; white-space: nowrap;
  border: 1px solid rgba(255,255,255,0.12);
}}
#{canvas_id}_flash {{
  position: fixed; top: 50%; left: 50%;
  transform: translate(-50%, -50%);
  background: rgba(0,0,0,0.78); color: #fff;
  padding: 10px 24px; border-radius: 10px;
  font: 700 15px Arial,sans-serif; pointer-events: none;
  display: none; z-index: 9999;
}}
#{canvas_id}_msg {{
  font: 12px Arial,sans-serif; color: #7F8C8D;
  text-align: center; padding: 5px 0 0; min-height: 20px;
}}
</style>
<div id="{canvas_id}_wrap">
  <canvas id="{canvas_id}"></canvas>
  <div id="{canvas_id}_tip"></div>
  <div id="{canvas_id}_flash"></div>
</div>
<div id="{canvas_id}_msg">👆 Click any cavity to mark it missing or restore it — then hit Save Frame</div>
<script>
(function() {{
  const COLS={cols_json}, ROWS={rows_json}, coords={coord_json};
  const marginL={margin_l}, marginT={margin_t};
  const cellW={cell_w}, cellH={cell_h};
  const IMG_W={W_img}, IMG_H={H_img}, CID="{canvas_id}";

  const canvas=document.getElementById(CID);
  const ctx=canvas.getContext('2d');
  const tip=document.getElementById(CID+'_tip');
  const flash=document.getElementById(CID+'_flash');
  const msg=document.getElementById(CID+'_msg');

  canvas.width=IMG_W; canvas.height=IMG_H;

  const baseImg=new Image();
  baseImg.src='data:image/jpeg;base64,{b64_img}';
  baseImg.onload=()=>ctx.drawImage(baseImg,0,0);

  function getScale(){{
    const r=canvas.getBoundingClientRect();
    return {{sx:IMG_W/r.width, sy:IMG_H/r.height, r}};
  }}
  function coordFromXY(x,y){{
    const ci=Math.floor((x-marginL)/cellW);
    const ri=Math.floor((y-marginT)/cellH);
    if(ci<0||ci>=COLS.length||ri<0||ri>=ROWS.length) return null;
    return COLS[ci]+ROWS[ri];
  }}

  canvas.addEventListener('mousemove', e=>{{
    const {{sx,sy,r}}=getScale();
    const x=(e.clientX-r.left)*sx, y=(e.clientY-r.top)*sy;
    const coord=coordFromXY(x,y);
    if(coord){{
      const st=coords[coord]===1?'🟢 Present':'🔴 Missing';
      tip.textContent=coord+' — '+st+'  (click to toggle)';
      tip.style.display='block';
      tip.style.left=(e.clientX+16)+'px';
      tip.style.top=(e.clientY-12)+'px';
    }} else tip.style.display='none';
  }});
  canvas.addEventListener('mouseleave',()=>tip.style.display='none');

  canvas.addEventListener('click', e=>{{
    const {{sx,sy,r}}=getScale();
    const x=(e.clientX-r.left)*sx, y=(e.clientY-r.top)*sy;
    const coord=coordFromXY(x,y);
    if(!coord) return;

    coords[coord]=coords[coord]===1?0:1;
    const nowMissing=coords[coord]===0;

    const ci=COLS.indexOf(coord[0]);
    const ri=ROWS.indexOf(coord.slice(1));
    const x0=marginL+ci*cellW, y0=marginT+ri*cellH;
    ctx.drawImage(baseImg,x0,y0,cellW,cellH,x0,y0,cellW,cellH);
    ctx.fillStyle=nowMissing?'rgba(231,76,60,0.40)':'rgba(46,204,113,0.40)';
    ctx.fillRect(x0+1,y0+1,cellW-2,cellH-2);
    ctx.strokeStyle='rgba(255,255,255,0.75)';
    ctx.lineWidth=1.5;
    ctx.strokeRect(x0+0.75,y0+0.75,cellW-1.5,cellH-1.5);
    ctx.fillStyle=nowMissing?'#fff':'rgba(10,10,10,0.85)';
    ctx.font='bold '+Math.max(9,Math.floor(Math.min(cellW,cellH)*0.30))+'px Arial';
    ctx.textAlign='center'; ctx.textBaseline='middle';
    ctx.fillText(coord,x0+cellW/2,y0+cellH/2);

    tip.style.display='none';
    msg.innerHTML=nowMissing
      ?'<span style="color:#E74C3C">🔴 Marked MISSING: <strong>'+coord+'</strong> — remember to Save Frame</span>'
      :'<span style="color:#2ECC71">🟢 Marked PRESENT: <strong>'+coord+'</strong> — remember to Save Frame</span>';

    flash.textContent=(nowMissing?'🔴 ':'🟢 ')+coord;
    flash.style.display='block';
    setTimeout(()=>flash.style.display='none',700);

    try {{
      const url=new URL(window.parent.location.href);
      url.searchParams.set('click',coord);
      window.parent.history.pushState({{}},'',url.toString());
      window.parent.postMessage({{type:'streamlit:forceRerender'}},'*');
    }} catch(err) {{
      const inputs=window.parent.document.querySelectorAll(
        '[data-testid="stTextInput"] input');
      if(inputs.length){{
        const inp=inputs[inputs.length-1];
        Object.getOwnPropertyDescriptor(
          window.HTMLInputElement.prototype,'value')
          .set.call(inp,coord);
        inp.dispatchEvent(new Event('input',{{bubbles:true}}));
      }}
    }}
  }});
}})();
</script>
"""
            components.html(canvas_html,
                            height=int(H_img * 820 / max(W_img, 1)) + 70,
                            scrolling=False)

            # ── Save Frame button + unsaved warning ───────────────────────────
            save_col, dl_col = st.columns([1, 1])
            with save_col:
                if st.button("💾 Save Frame", use_container_width=True,
                             type="primary", key="pi_save",
                             disabled=not st.session_state.dirty):
                    auto_save(active_id, frame_name)
                    st.success("Frame saved!")
                    st.rerun()
            with dl_col:
                buf_dl = io.BytesIO()
                overlay_img.save(buf_dl, format="PNG")
                st.download_button(
                    "⬇️ Download overlay image",
                    data=buf_dl.getvalue(),
                    file_name=f"{re.sub(r'[^a-zA-Z0-9_-]','_',frame_name)}_overlay.png",
                    mime="image/png",
                    use_container_width=True,
                )
            if st.session_state.dirty:
                st.warning("⚠️ You have unsaved changes — click **Save Frame** above.")

    # ── Tab: Grid Editor ──────────────────────────────────────────────────────
    with tab_grid:
        st.markdown(
            "Click a cell to toggle **Present 🟢** ↔ **Missing 🔴**, "
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
                f"color:#3498DB'>{row_num}</div>",
                unsafe_allow_html=True)
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
        st.dataframe(pd.DataFrame(status_rows), use_container_width=True,
                     hide_index=True)

    # ── Tab: Overlay Preview ────────────────────────────────────────────
    with tab_vis:
        if photo:
            st.markdown("Overlay rendered on actual mold photo. 🟢 Present  🔴 Missing")
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
            st.markdown("Synthetic grid (upload a photo to see the real mold). "
                        "🟢 Present  🔴 Missing")
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
                        f"padding:2px 6px;border-radius:4px;"
                        f"font-size:0.85rem'>{c}</span>"
                        for c in coords_in_row),
                    unsafe_allow_html=True)
        else:
            st.success("🎉 All 120 positions are present!")

    # ── Bottom navigation ──────────────────────────────────────────────────────
    st.divider()
    bot_l, bot_m, bot_r = st.columns([1, 2, 1])
    with bot_l:
        if st.button("◄ Previous Frame", use_container_width=True,
                     disabled=(total_frames <= 1)):
            navigate_to_adjacent_frame(-1)
            st.rerun()
    with bot_m:
        st.markdown(
            f"<div style='text-align:center;padding-top:8px'>"
            f"Frame <strong>{current_idx + 1}</strong> "
            f"of <strong>{total_frames}</strong></div>",
            unsafe_allow_html=True)
    with bot_r:
        if st.button("Next Frame ►", use_container_width=True,
                     disabled=(total_frames <= 1), type="primary"):
            navigate_to_adjacent_frame(+1)
            st.rerun()


if __name__ == "__main__":
    main()
