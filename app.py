"""
Chocolate Mold Inspector — Streamlit App
15 columns (A–O) × 8 rows (1–8) = 120 coordinates per mold frame

Workflow:
  • Upload frame photos in the sidebar
  • Click directly on the photo to toggle cavities present ↔ missing
  • All changes auto-save immediately to CSV
  • Export a cumulative Excel/PDF report across all frames
"""

import base64
import io
import json
import os
import re
import time
from datetime import datetime

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
N_COORDS   = len(ALL_COORDS)  # 120

DATA_FILE        = "mold_data.csv"
IMAGES_DIR       = "frame_images"
MAX_UPLOAD_FILES = 50

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


def get_coord_dict(df: pd.DataFrame, frame_id: str) -> dict:
    row = df[df["frame_id"] == frame_id]
    if row.empty:
        return {c: True for c in ALL_COORDS}
    r = row.iloc[0]
    return {c: (str(r.get(c, "1")) == "1") for c in ALL_COORDS}


def upsert_frame(df: pd.DataFrame, frame_id: str, frame_name: str,
                 coord_dict: dict, image_path: str = "") -> pd.DataFrame:
    ts       = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row_data = {
        "frame_id": frame_id, "frame_name": frame_name,
        "timestamp": ts, "image_path": image_path,
        **{c: ("1" if v else "0") for c, v in coord_dict.items()},
    }
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


# ── Immediate persist ──────────────────────────────────────────────────────────

def auto_save(frame_id: str, frame_name: str):
    """Write coord_dict → DataFrame → CSV. Single source of truth."""
    row = st.session_state.df[st.session_state.df["frame_id"] == frame_id]
    image_path = str(row.iloc[0].get("image_path", "")) if not row.empty else ""
    st.session_state.df = upsert_frame(
        st.session_state.df, frame_id, frame_name,
        st.session_state.coord_dict, image_path)
    save_data(st.session_state.df)


# ── Image helpers ──────────────────────────────────────────────────────────────

def pil_to_b64(img: Image.Image, fmt: str = "JPEG") -> str:
    buf = io.BytesIO()
    img.save(buf, format=fmt)
    return base64.b64encode(buf.getvalue()).decode()


def render_overlay(photo: Image.Image, coord_dict: dict, opacity: int = 55) -> Image.Image:
    """Render the mold photo with a semi-transparent colour overlay per cell."""
    base = photo.convert("RGBA")
    W, H = base.size
    margin_l = max(24, int(W * 0.038))
    margin_t = max(20, int(H * 0.055))
    grid_w   = W - margin_l - max(4, int(W * 0.008))
    grid_h   = H - margin_t - max(4, int(H * 0.008))
    cell_w   = grid_w / len(COLS)
    cell_h   = grid_h / len(ROWS)

    overlay  = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw     = ImageDraw.Draw(overlay)

    font_size = max(9, int(min(cell_w, cell_h) * 0.30))
    try:
        fnt     = ImageFont.truetype(
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
            x1, y1 = x0 + cell_w, y0 + cell_h
            fill = (46, 204, 113, opacity) if present else (231, 76, 60, opacity)
            draw.rectangle([x0 + 1, y0 + 1, x1 - 1, y1 - 1], fill=fill)
            draw.rectangle([x0, y0, x1, y1], outline=(255, 255, 255, 180), width=1)
            txt_col = (10, 10, 10, 255) if present else (255, 255, 255, 255)
            draw.text(((x0 + x1) / 2, (y0 + y1) / 2),
                      coord, fill=txt_col, font=fnt, anchor="mm")

    for ci, col in enumerate(COLS):
        draw.text((margin_l + (ci + 0.5) * cell_w, margin_t / 2),
                  col, fill=(255, 255, 255, 230), font=fnt_hdr, anchor="mm")
    for ri, row_num in enumerate(ROWS):
        draw.text((margin_l / 2, margin_t + (ri + 0.5) * cell_h),
                  str(row_num), fill=(255, 255, 255, 230), font=fnt_hdr, anchor="mm")

    return Image.alpha_composite(base, overlay).convert("RGB")


def build_clickable_photo_html(photo: Image.Image, coord_dict: dict,
                                opacity: int = 55,
                                display_width: int = 900) -> str:
    """
    Return an HTML snippet with a <canvas> drawn over the photo.
    Clicking a cell sends a postMessage with the coordinate.
    The overlay is rendered by JS on the canvas so no round-trip is needed.
    """
    W, H = photo.size
    aspect = H / W
    display_height = int(display_width * aspect)

    margin_l_frac = max(0.024, 0.038)
    margin_t_frac = max(0.025, 0.055)

    # Encode photo as base64 JPEG
    buf = io.BytesIO()
    photo.save(buf, format="JPEG", quality=85)
    img_b64 = base64.b64encode(buf.getvalue()).decode()

    # Encode coord_dict as JSON (1/0)
    coord_json = json.dumps({c: (1 if v else 0) for c, v in coord_dict.items()})
    cols_json  = json.dumps(COLS)
    rows_json  = json.dumps(ROWS)

    html = f"""
<style>
  #mold-wrap {{
    position: relative;
    width: {display_width}px;
    margin: 0 auto;
    cursor: crosshair;
    border-radius: 6px;
    overflow: hidden;
    box-shadow: 0 4px 24px rgba(0,0,0,0.4);
  }}
  #mold-canvas {{
    display: block;
    width: {display_width}px;
    height: {display_height}px;
  }}
  #coord-tooltip {{
    position: fixed;
    background: rgba(0,0,0,0.75);
    color: #fff;
    padding: 4px 10px;
    border-radius: 6px;
    font: 600 13px/1.4 monospace;
    pointer-events: none;
    display: none;
    z-index: 9999;
  }}
</style>
<div id="mold-wrap">
  <canvas id="mold-canvas" width="{display_width}" height="{display_height}"></canvas>
</div>
<div id="coord-tooltip"></div>

<script>
(function() {{
  const COLS  = {cols_json};
  const ROWS  = {rows_json};
  const STATE = {coord_json};

  const canvas  = document.getElementById('mold-canvas');
  const ctx     = canvas.getContext('2d');
  const tooltip = document.getElementById('coord-tooltip');

  const W = {display_width};
  const H = {display_height};

  // Grid geometry (mirrors Python render_overlay)
  const marginL = Math.max(24, W * 0.038);
  const marginT = Math.max(20, H * 0.055);
  const gridW   = W - marginL - Math.max(4, W * 0.008);
  const gridH   = H - marginT - Math.max(4, H * 0.008);
  const cellW   = gridW / COLS.length;
  const cellH   = gridH / ROWS.length;
  const OPACITY = {opacity / 255:.3f};

  // Load the photo
  const img = new Image();
  img.onload = () => draw();
  img.src = 'data:image/jpeg;base64,{img_b64}';

  function draw() {{
    ctx.drawImage(img, 0, 0, W, H);

    const fontSize = Math.max(9, Math.min(cellW, cellH) * 0.30);
    ctx.textAlign    = 'center';
    ctx.textBaseline = 'middle';
    ctx.font         = `bold ${{fontSize}}px sans-serif`;

    ROWS.forEach((rowNum, ri) => {{
      COLS.forEach((col, ci) => {{
        const coord   = col + rowNum;
        const present = STATE[coord] === 1;
        const x0 = marginL + ci * cellW;
        const y0 = marginT + ri * cellH;

        // Fill
        ctx.fillStyle = present
          ? `rgba(46,204,113,${{OPACITY}})`
          : `rgba(231,76,60,${{OPACITY}})`;
        ctx.fillRect(x0 + 1, y0 + 1, cellW - 2, cellH - 2);

        // Border
        ctx.strokeStyle = 'rgba(255,255,255,0.7)';
        ctx.lineWidth   = 1;
        ctx.strokeRect(x0, y0, cellW, cellH);

        // Label
        ctx.fillStyle = present ? 'rgba(10,10,10,0.9)' : 'rgba(255,255,255,0.95)';
        ctx.fillText(coord, x0 + cellW / 2, y0 + cellH / 2);
      }});
    }});

    // Column headers
    const hdrSize = Math.max(10, Math.min(marginL, marginT) * 0.55);
    ctx.font      = `bold ${{hdrSize}}px sans-serif`;
    ctx.fillStyle = 'rgba(255,255,255,0.9)';
    COLS.forEach((col, ci) => {{
      ctx.fillText(col, marginL + (ci + 0.5) * cellW, marginT / 2);
    }});
    ROWS.forEach((rowNum, ri) => {{
      ctx.fillText(String(rowNum), marginL / 2, marginT + (ri + 0.5) * cellH);
    }});
  }}

  function hitTest(px, py) {{
    if (px < marginL || py < marginT) return null;
    const ci = Math.floor((px - marginL) / cellW);
    const ri = Math.floor((py - marginT) / cellH);
    if (ci < 0 || ci >= COLS.length || ri < 0 || ri >= ROWS.length) return null;
    return COLS[ci] + ROWS[ri];
  }}

  canvas.addEventListener('mousemove', e => {{
    const rect  = canvas.getBoundingClientRect();
    const scaleX = W / rect.width;
    const scaleY = H / rect.height;
    const px = (e.clientX - rect.left) * scaleX;
    const py = (e.clientY - rect.top)  * scaleY;
    const coord = hitTest(px, py);
    if (coord) {{
      tooltip.style.display = 'block';
      tooltip.style.left    = (e.clientX + 14) + 'px';
      tooltip.style.top     = (e.clientY - 28) + 'px';
      const status = STATE[coord] === 1 ? '✅ Present' : '❌ Missing';
      tooltip.textContent   = `${{coord}}  ${{status}}`;
    }} else {{
      tooltip.style.display = 'none';
    }}
  }});
  canvas.addEventListener('mouseleave', () => {{ tooltip.style.display = 'none'; }});

  canvas.addEventListener('click', e => {{
    const rect  = canvas.getBoundingClientRect();
    const scaleX = W / rect.width;
    const scaleY = H / rect.height;
    const px = (e.clientX - rect.left) * scaleX;
    const py = (e.clientY - rect.top)  * scaleY;
    const coord = hitTest(px, py);
    if (!coord) return;

    // Optimistic toggle on canvas
    STATE[coord] = STATE[coord] === 1 ? 0 : 1;
    draw();

    // Notify Streamlit via query-param trick
    window.parent.postMessage({{type: 'mold_toggle', coord: coord}}, '*');
  }});
}})();
</script>
"""
    return html


# ── Excel export ───────────────────────────────────────────────────────────────

def _interpolate_argb(count: int, max_count: int) -> str:
    t = min(count / max_count, 1.0) if max_count > 0 else 0.0
    r = int(46  + (231 - 46)  * t)
    g = int(204 + (76  - 204) * t)
    b = int(113 + (60  - 113) * t)
    return f"FF{r:02X}{g:02X}{b:02X}"


def _missing_counts(df: pd.DataFrame) -> dict:
    return {
        coord: int((df[coord].astype(str) == "0").sum())
        if coord in df.columns else 0
        for coord in ALL_COORDS
    }


def build_excel(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Cumulative Heatmap"

    thin        = Side(style="thin", color="999999")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)
    center      = Alignment(horizontal="center", vertical="center")
    hdr_fill    = PatternFill("solid", fgColor="FF2C3E50")
    sum_fill    = PatternFill("solid", fgColor="FF34495E")
    white_bold  = Font(bold=True, color="FFFFFFFF", name="Arial", size=11)
    cell_font   = Font(name="Arial", size=10)
    cell_font_w = Font(name="Arial", size=10, color="FFFFFFFF")

    n_frames  = len(df)
    mc        = _missing_counts(df)
    max_miss  = max(mc.values()) if mc else 1

    # ── Sheet 1: Heatmap ───────────────────────────────────────────────────────
    n_data_cols = len(COLS) + 2   # row label + cols + row-total
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_data_cols)
    tc = ws.cell(1, 1,
        f"Chocolate Mold — Cumulative Missing  ({n_frames} frame(s) inspected)")
    tc.font = Font(bold=True, color="FFFFFFFF", name="Arial", size=13)
    tc.fill = hdr_fill; tc.alignment = center
    ws.row_dimensions[1].height = 30

    ws.cell(2, 1, "").fill = hdr_fill
    for ci, col in enumerate(COLS, 2):
        c = ws.cell(2, ci, col)
        c.font = white_bold; c.fill = hdr_fill
        c.alignment = center; c.border = border
    tot_col = n_data_cols
    c2 = ws.cell(2, tot_col, "Row\nMissing")
    c2.font = white_bold; c2.fill = sum_fill
    c2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c2.border = border
    ws.row_dimensions[2].height = 30

    ws.column_dimensions["A"].width = 6
    for ci in range(2, n_data_cols):
        ws.column_dimensions[get_column_letter(ci)].width = 10
    ws.column_dimensions[get_column_letter(tot_col)].width = 11

    for ri, row_num in enumerate(ROWS):
        er = ri + 3
        ws.row_dimensions[er].height = 24
        rh = ws.cell(er, 1, str(row_num))
        rh.font = white_bold; rh.fill = hdr_fill
        rh.alignment = center; rh.border = border

        row_total = 0
        for ci, col in enumerate(COLS, 2):
            coord = f"{col}{row_num}"
            count = mc.get(coord, 0)
            row_total += count
            t     = count / max_miss if max_miss > 0 else 0
            cell  = ws.cell(er, ci, count)
            cell.fill      = PatternFill("solid", fgColor=_interpolate_argb(count, max_miss))
            cell.font      = cell_font_w if t > 0.45 else cell_font
            cell.alignment = center; cell.border = border

        rt = ws.cell(er, tot_col, row_total)
        rt.font = Font(bold=True, name="Arial", size=10, color="FFFFFFFF")
        rt.fill = sum_fill; rt.alignment = center; rt.border = border

    tot_row = len(ROWS) + 3
    ws.row_dimensions[tot_row].height = 24
    ws.cell(tot_row, 1, "Total").font = Font(bold=True, color="FFFFFFFF", name="Arial", size=10)
    ws.cell(tot_row, 1).fill = sum_fill; ws.cell(tot_row, 1).alignment = center
    ws.cell(tot_row, 1).border = border

    grand = 0
    for ci, col in enumerate(COLS, 2):
        ct = sum(mc.get(f"{col}{r}", 0) for r in ROWS)
        grand += ct
        c = ws.cell(tot_row, ci, ct)
        c.font = Font(bold=True, name="Arial", size=10, color="FFFFFFFF")
        c.fill = sum_fill; c.alignment = center; c.border = border
    gt = ws.cell(tot_row, tot_col, grand)
    gt.font = Font(bold=True, name="Arial", size=11, color="FFFFFFFF")
    gt.fill = PatternFill("solid", fgColor="FF1A252F")
    gt.alignment = center; gt.border = border

    leg_row = tot_row + 2
    ws.merge_cells(start_row=leg_row, start_column=1, end_row=leg_row, end_column=n_data_cols)
    leg = ws.cell(leg_row, 1,
        f"Green = never missing  →  Red = missing in all {n_frames} frame(s)  |  "
        f"Each cell = count of frames where that cavity was absent")
    leg.font = Font(italic=True, name="Arial", size=9, color="FF555555")
    leg.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[leg_row].height = 18

    # ── Sheet 2: Frame Summary ─────────────────────────────────────────────────
    ws2 = wb.create_sheet("Frame Summary", 1)
    hdrs = ["Frame", "Missing", "Present", "Total", "% Present", "Timestamp"]
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(hdrs))
    t2 = ws2.cell(1, 1, "Chocolate Mold Inspection — Frame Summary")
    t2.font = Font(bold=True, color="FFFFFFFF", name="Arial", size=14)
    t2.fill = hdr_fill; t2.alignment = center
    ws2.row_dimensions[1].height = 30

    for ci, h in enumerate(hdrs, 1):
        c = ws2.cell(2, ci, h)
        c.font = white_bold; c.fill = sum_fill
        c.alignment = center; c.border = border
    ws2.row_dimensions[2].height = 22
    for ci, w in enumerate([30, 10, 10, 10, 12, 22], 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w

    for ri, (_, rec) in enumerate(df.iterrows(), 3):
        missing = sum(1 for c in ALL_COORDS if str(rec.get(c, "1")) == "0")
        present = N_COORDS - missing
        for ci, val in enumerate(
            [str(rec.get("frame_name", rec["frame_id"])),
             missing, present, N_COORDS, f"=C{ri}/D{ri}",
             str(rec.get("timestamp", ""))], 1):
            cell = ws2.cell(ri, ci, val)
            cell.alignment = center; cell.border = border
            cell.font = Font(name="Arial", size=10)
            if ci == 2 and N_COORDS > 0:
                inten = min(missing / N_COORDS, 1.0)
                rv = int(231 * inten + 46  * (1 - inten))
                gv = int(76  * inten + 204 * (1 - inten))
                bv = int(60  * inten + 113 * (1 - inten))
                cell.fill = PatternFill("solid", fgColor=f"FF{rv:02X}{gv:02X}{bv:02X}")
            if ci == 5:
                cell.number_format = "0.0%"
        ws2.row_dimensions[ri].height = 20

    last  = len(df) + 2
    tot_r = last + 1
    ws2.cell(tot_r, 1, "TOTALS").font = Font(bold=True, name="Arial", size=10)
    if not df.empty:
        for ci, formula in enumerate(
            [None, f"=SUM(B3:B{last})", f"=SUM(C3:C{last})",
             f"=SUM(D3:D{last})", f"=C{tot_r}/D{tot_r}", None], 1):
            if formula:
                c = ws2.cell(tot_r, ci, formula)
                c.font = Font(bold=True, name="Arial")
                if ci == 5:
                    c.number_format = "0.0%"
    for ci in range(1, 7):
        ws2.cell(tot_r, ci).border = border
        ws2.cell(tot_r, ci).alignment = center
    ws2.row_dimensions[tot_r].height = 22

    # ── Sheet 3: Coordinate Frequency ─────────────────────────────────────────
    ws3 = wb.create_sheet("Coordinate Frequency", 2)
    f_hdrs = ["Coordinate", "Times Missing", "Total Frames", "% Flagged", "Rank"]
    ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(f_hdrs))
    ft = ws3.cell(1, 1, f"Coordinate Flag Frequency  —  {n_frames} frame(s) analysed")
    ft.font = Font(bold=True, color="FFFFFFFF", name="Arial", size=13)
    ft.fill = hdr_fill; ft.alignment = center
    ws3.row_dimensions[1].height = 30

    for ci, h in enumerate(f_hdrs, 1):
        c = ws3.cell(2, ci, h)
        c.font = white_bold; c.fill = sum_fill
        c.alignment = center; c.border = border
    ws3.row_dimensions[2].height = 22
    for ci, w in enumerate([14, 16, 14, 12, 8], 1):
        ws3.column_dimensions[get_column_letter(ci)].width = w

    sorted_coords = sorted(ALL_COORDS, key=lambda c: (-mc.get(c, 0), c))
    for rank, coord in enumerate(sorted_coords, 1):
        er  = rank + 2
        cnt = mc.get(coord, 0)
        pct = cnt / n_frames if n_frames > 0 else 0.0
        t   = cnt / max_miss if max_miss > 0 else 0
        argb = _interpolate_argb(cnt, max_miss if max_miss > 0 else 1)
        flag_font = Font(name="Arial", size=10,
                         color="FFFFFFFF" if t > 0.45 else "FF000000",
                         bold=(cnt > 0))
        for ci, val in enumerate([coord, cnt, n_frames, pct, rank], 1):
            cell = ws3.cell(er, ci, val)
            cell.alignment = center; cell.border = border
            cell.font = Font(name="Arial", size=10)
            if ci == 4:
                cell.number_format = "0.0%"
                cell.fill = PatternFill("solid", fgColor=argb)
                cell.font = flag_font
            elif ci == 2 and cnt > 0:
                inten = min(cnt / n_frames, 1.0)
                rv = int(231 * inten + 236 * (1 - inten))
                gv = int(76  * inten + 240 * (1 - inten))
                bv = int(60  * inten + 241 * (1 - inten))
                cell.fill = PatternFill("solid", fgColor=f"FF{rv:02X}{gv:02X}{bv:02X}")
        ws3.row_dimensions[er].height = 18

    footer = len(ALL_COORDS) + 3
    ws3.merge_cells(start_row=footer, start_column=1, end_row=footer, end_column=len(f_hdrs))
    total_ev = sum(mc.values())
    avg_miss = total_ev / n_frames if n_frames > 0 else 0
    ws3.cell(footer, 1,
        f"Total missing events: {total_ev}   |   "
        f"Avg missing per frame: {avg_miss:.1f} / {N_COORDS}   |   "
        f"Never flagged: {sum(1 for v in mc.values() if v == 0)}")
    ws3.cell(footer, 1).font = Font(italic=True, name="Arial", size=9, color="FF555555")
    ws3.cell(footer, 1).alignment = Alignment(horizontal="left", vertical="center")
    ws3.row_dimensions[footer].height = 18

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── PDF export ─────────────────────────────────────────────────────────────────

def _rl(hex6: str) -> colors.Color:
    h = hex6.lstrip("#")
    return colors.Color(int(h[:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:], 16) / 255)


def _rl_interp(count: int, max_count: int) -> colors.Color:
    t = min(count / max_count, 1.0) if max_count > 0 else 0.0
    return colors.Color(
        (46  + (231 - 46)  * t) / 255,
        (204 + (76  - 204) * t) / 255,
        (113 + (60  - 113) * t) / 255)


def build_pdf(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=15*mm, bottomMargin=15*mm,
        title="Chocolate Mold Inspection Report")

    COL_DARK = _rl("#2C3E50")
    COL_MID  = _rl("#34495E")
    WHITE    = colors.white
    BLACK    = colors.black

    title_style   = ParagraphStyle("t",  fontName="Helvetica-Bold",    fontSize=14,
                                   leading=18, textColor=WHITE, alignment=TA_CENTER)
    caption_style = ParagraphStyle("ca", fontName="Helvetica-Oblique", fontSize=7,
                                   leading=10, textColor=colors.HexColor("#555555"))

    n_frames = len(df)
    mc       = _missing_counts(df)
    max_miss = max(mc.values()) if mc else 1
    now_str  = datetime.now().strftime("%Y-%m-%d %H:%M")

    def hdr(txt, size=9):
        return Paragraph(f"<b>{txt}</b>", ParagraphStyle(
            "h", fontName="Helvetica-Bold", fontSize=size,
            textColor=WHITE, alignment=TA_CENTER))

    def title_tbl(text):
        t = Table([[Paragraph(text, title_style)]], colWidths=[doc.width])
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), COL_DARK),
            ("TOPPADDING",    (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ]))
        return t

    story = []

    # Page 1 — Heatmap
    story.append(title_tbl(
        f"Chocolate Mold — Cumulative Missing Count  ({n_frames} frame(s) inspected)"))
    story.append(Spacer(1, 6*mm))

    row_hdr_w  = 14*mm
    tot_col_w  = 18*mm
    cell_each  = (doc.width - row_hdr_w - tot_col_w) / len(COLS)
    col_widths = [row_hdr_w] + [cell_each] * len(COLS) + [tot_col_w]

    grid_rows = [[""] + [hdr(c) for c in COLS] + [hdr("Row\nTotal", 7)]]
    for row_num in ROWS:
        row_total = 0
        row = [hdr(str(row_num))]
        for col in COLS:
            count = mc.get(f"{col}{row_num}", 0)
            row_total += count
            row.append(str(count) if count > 0 else "")
        row.append(hdr(str(row_total)))
        grid_rows.append(row)

    tot_row = [hdr("Total")]
    grand   = 0
    for col in COLS:
        ct = sum(mc.get(f"{col}{r}", 0) for r in ROWS)
        grand += ct
        tot_row.append(hdr(str(ct)))
    tot_row.append(hdr(str(grand), 10))
    grid_rows.append(tot_row)

    ts = [
        ("FONTNAME",      (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE",      (0, 0), (-1, -1), 8),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",          (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("BACKGROUND",    (0, 0), (-1, 0),   COL_DARK),
        ("BACKGROUND",    (0, 1), (0, -2),   COL_DARK),
        ("BACKGROUND",    (-1, 1), (-1, -2), COL_MID),
        ("BACKGROUND",    (0, -1), (-1, -1), COL_MID),
        ("BACKGROUND",    (-1, -1), (-1, -1), _rl("#1A252F")),
    ]
    for ri, row_num in enumerate(ROWS, 1):
        for ci, col in enumerate(COLS, 1):
            count = mc.get(f"{col}{row_num}", 0)
            bg    = _rl_interp(count, max_miss)
            t_val = count / max_miss if max_miss > 0 else 0
            ts += [
                ("BACKGROUND", (ci, ri), (ci, ri), bg),
                ("TEXTCOLOR",  (ci, ri), (ci, ri), WHITE if t_val > 0.45 else BLACK),
            ]

    story.append(Table(grid_rows, colWidths=col_widths, style=TableStyle(ts)))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph(
        f"Green = never missing  |  Red = missing in all {n_frames} frames  |  "
        f"Generated: {now_str}", caption_style))

    # Page 2 — Frame Summary
    story.append(PageBreak())
    story.append(title_tbl("Chocolate Mold Inspection — Frame Summary"))
    story.append(Spacer(1, 6*mm))

    sum_cw   = [doc.width*0.28, doc.width*0.10, doc.width*0.10,
                doc.width*0.10, doc.width*0.12, doc.width*0.30]
    sum_hdrs = ["Frame", "Missing", "Present", "Total", "% Present", "Timestamp"]
    sum_rows = [[hdr(h) for h in sum_hdrs]]

    for _, rec in df.iterrows():
        miss = sum(1 for c in ALL_COORDS if str(rec.get(c, "1")) == "0")
        pres = N_COORDS - miss
        sum_rows.append([
            Paragraph(str(rec.get("frame_name", rec["frame_id"])),
                      ParagraphStyle("fn", fontName="Helvetica", fontSize=8)),
            miss, pres, N_COORDS, f"{pres / N_COORDS * 100:.1f}%",
            Paragraph(str(rec.get("timestamp", "")),
                      ParagraphStyle("ts", fontName="Helvetica", fontSize=7,
                                     textColor=colors.HexColor("#555555"),
                                     alignment=TA_CENTER)),
        ])

    total_miss_all = sum(
        sum(1 for c in ALL_COORDS if str(rec.get(c, "1")) == "0")
        for _, rec in df.iterrows())
    total_pres_all = N_COORDS * n_frames - total_miss_all
    total_tot_all  = N_COORDS * n_frames
    pct_all = f"{total_pres_all / total_tot_all * 100:.1f}%" if total_tot_all > 0 else "—"
    sum_rows.append([hdr("TOTALS"), hdr(str(total_miss_all)),
                     hdr(str(total_pres_all)), hdr(str(total_tot_all)), hdr(pct_all), ""])

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
        ("LEFTPADDING",   (0, 0), (0, -1),  6),
    ]
    for ri, (_, rec) in enumerate(df.iterrows(), 1):
        miss  = sum(1 for c in ALL_COORDS if str(rec.get(c, "1")) == "0")
        inten = min(miss / N_COORDS, 1.0)
        bg    = _rl_interp(int(inten * max_miss), max_miss)
        sum_ts += [
            ("BACKGROUND", (1, ri), (1, ri), bg),
            ("TEXTCOLOR",  (1, ri), (1, ri), WHITE if inten > 0.45 else BLACK),
        ]

    story.append(Table(sum_rows, colWidths=sum_cw, style=TableStyle(sum_ts)))

    # Page 3 — Coordinate Frequency
    story.append(PageBreak())
    story.append(title_tbl(
        f"Coordinate Flag Frequency  —  {n_frames} frame(s) analysed"))
    story.append(Spacer(1, 6*mm))

    freq_cw   = [doc.width*0.15, doc.width*0.18, doc.width*0.15,
                 doc.width*0.15, doc.width*0.10]
    freq_hdrs = ["Coordinate", "Times Missing", "Total Frames", "% Flagged", "Rank"]
    freq_rows = [[hdr(h) for h in freq_hdrs]]

    sorted_coords = sorted(ALL_COORDS, key=lambda c: (-mc.get(c, 0), c))
    for rank, coord in enumerate(sorted_coords, 1):
        count = mc.get(coord, 0)
        pct   = f"{count / n_frames * 100:.1f}%" if n_frames > 0 else "0.0%"
        freq_rows.append([coord, count, n_frames, pct, rank])

    total_ev  = sum(mc.values())
    avg_miss  = total_ev / n_frames if n_frames > 0 else 0
    never_flg = sum(1 for v in mc.values() if v == 0)
    freq_rows.append([
        Paragraph(
            f"<b>Total: {total_ev}  |  Avg/frame: {avg_miss:.1f}  |  "
            f"Never flagged: {never_flg}</b>",
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
    for ri, coord in enumerate(sorted_coords, 1):
        count = mc.get(coord, 0)
        t_val = count / max_miss if max_miss > 0 else 0
        freq_ts.append(("BACKGROUND", (3, ri), (3, ri), _rl_interp(count, max_miss)))
        freq_ts.append(("TEXTCOLOR",  (3, ri), (3, ri), WHITE if t_val > 0.45 else BLACK))
        if count > 0:
            inten = min(count / n_frames, 1.0)
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


# ── Session state ──────────────────────────────────────────────────────────────

def init_state():
    defaults = {
        "df":              load_data(),
        "active_frame_id": None,
        "coord_dict":      {c: True for c in ALL_COORDS},
        "frame_image":     None,
        "pending_toggle":  None,   # coord toggled via canvas postMessage
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def load_frame(frame_id: str):
    st.session_state.active_frame_id = frame_id
    st.session_state.coord_dict      = get_coord_dict(st.session_state.df, frame_id)
    row = st.session_state.df[st.session_state.df["frame_id"] == frame_id]
    img_path = str(row.iloc[0].get("image_path", "")) if not row.empty else ""
    st.session_state.frame_image = load_frame_image(img_path)


def navigate(direction: int):
    frame_ids = st.session_state.df["frame_id"].tolist()
    if not frame_ids:
        return
    cur = st.session_state.active_frame_id
    idx = frame_ids.index(cur) if cur in frame_ids else 0
    load_frame(frame_ids[(idx + direction) % len(frame_ids)])


# ── Main ───────────────────────────────────────────────────────────────────────

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
        background: #16213E; border: 1px solid #2C3E50; margin-bottom: 4px;
    }
    .stat-number { font-size: 1.9rem; font-weight: 700; }
    .stat-label  { font-size: 0.82rem; color: #95A5A6; margin-top: 3px; }
    .green { color: #27AE60; } .red { color: #E74C3C; }
    .stButton > button { border-radius: 8px; font-weight: 600; }
    .upload-hint {
        border: 1px dashed #3498DB; border-radius: 8px;
        padding: 12px 16px; font-size: 0.88rem; margin-bottom: 8px;
    }
    .missing-badge {
        background:#E74C3C; color:#FFF; padding:1px 7px;
        border-radius:4px; font-size:0.78rem;
        margin:1px; display:inline-block;
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Handle postMessage toggle from canvas ──────────────────────────────────
    # We use a URL query-param trick: the JS canvas posts to the parent;
    # a tiny receiver script writes a hidden Streamlit text_input.
    # Simpler approach: use st.query_params to pass the toggled coord.
    toggle_coord = st.query_params.get("toggle")
    if toggle_coord and toggle_coord in ALL_COORDS:
        active_id = st.session_state.active_frame_id
        if active_id:
            st.session_state.coord_dict[toggle_coord] = \
                not st.session_state.coord_dict.get(toggle_coord, True)
            auto_save(active_id,
                      st.session_state.df.loc[
                          st.session_state.df["frame_id"] == active_id,
                          "frame_name"].iloc[0])
        # Clear the param
        st.query_params.clear()
        st.rerun()

    # ── Sidebar ────────────────────────────────────────────────────────────────
    with st.sidebar:
        st.title("🍫 Mold Inspector")
        st.caption("15 × 8 grid | 120 positions")
        st.divider()

        st.subheader("📸 Upload Frame Photos")
        st.markdown(
            f'<div class="upload-hint">Upload up to {MAX_UPLOAD_FILES} images. '
            'A new record is created per photo.</div>', unsafe_allow_html=True)

        uploaded_photos = st.file_uploader(
            "Drop frame images here",
            type=["png", "jpg", "jpeg", "bmp", "tiff", "webp"],
            accept_multiple_files=True,
            label_visibility="collapsed",
            key="batch_uploader",
        )

        if uploaded_photos:
            if len(uploaded_photos) > MAX_UPLOAD_FILES:
                st.warning(f"Only the first {MAX_UPLOAD_FILES} will be imported.")
                uploaded_photos = uploaded_photos[:MAX_UPLOAD_FILES]
            existing_names = set(st.session_state.df["frame_name"].tolist()) \
                if not st.session_state.df.empty else set()
            new_count = 0
            last_fid  = None
            for uf in uploaded_photos:
                base_name = os.path.splitext(uf.name)[0]
                if base_name in existing_names:
                    continue
                fid   = f"frame_{int(time.time() * 1000)}_{new_count}"
                ipath = save_frame_image(fid, uf)
                st.session_state.df = upsert_frame(
                    st.session_state.df, fid, base_name,
                    {c: True for c in ALL_COORDS}, ipath)
                existing_names.add(base_name)
                last_fid = fid
                new_count += 1
            if new_count:
                save_data(st.session_state.df)
                st.success(f"Imported {new_count} new frame(s).")
                if last_fid:
                    load_frame(last_fid)
                st.rerun()

        st.divider()
        st.subheader("📋 Frames")
        df = st.session_state.df

        if st.button("🏠 Overview", use_container_width=True):
            st.session_state.active_frame_id = None
            st.rerun()

        if df.empty:
            st.info("No frames yet — upload photos above.")
        else:
            frame_names = df["frame_name"].tolist()
            frame_ids   = df["frame_id"].tolist()
            current_idx = 0
            if st.session_state.active_frame_id in frame_ids:
                current_idx = frame_ids.index(st.session_state.active_frame_id)

            sel_idx = st.selectbox(
                "Frame", range(len(frame_names)),
                format_func=lambda i: frame_names[i],
                index=current_idx, label_visibility="collapsed")

            col_load, col_del = st.columns(2)
            if col_load.button("Load", use_container_width=True):
                load_frame(frame_ids[sel_idx]); st.rerun()
            if col_del.button("🗑 Delete", use_container_width=True):
                fid = frame_ids[sel_idx]
                row = st.session_state.df[st.session_state.df["frame_id"] == fid]
                if not row.empty:
                    ip = str(row.iloc[0].get("image_path", ""))
                    if ip and os.path.exists(ip):
                        os.remove(ip)
                st.session_state.df = st.session_state.df[
                    st.session_state.df["frame_id"] != fid].reset_index(drop=True)
                save_data(st.session_state.df)
                if st.session_state.active_frame_id == fid:
                    st.session_state.active_frame_id = None
                    st.session_state.frame_image     = None
                st.rerun()

            st.caption(f"Total frames: **{len(df)}**")

        st.divider()
        st.subheader("📥 Export")
        df = st.session_state.df

        if st.button("⬇️ Download Heatmap (.xlsx)",
                     use_container_width=True, disabled=df.empty):
            xlsx_bytes = build_excel(df)
            st.download_button(
                "💾 Save .xlsx", data=xlsx_bytes,
                file_name=f"mold_heatmap_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

        if st.button("⬇️ Download Report (.pdf)",
                     use_container_width=True, disabled=df.empty):
            pdf_bytes = build_pdf(df)
            st.download_button(
                "💾 Save .pdf", data=pdf_bytes,
                file_name=f"mold_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf", use_container_width=True)

        if not df.empty:
            st.download_button(
                "⬇️ Backup CSV",
                data=df.to_csv(index=False).encode(),
                file_name="mold_data_backup.csv",
                mime="text/csv", use_container_width=True)

        st.divider()
        st.subheader("⬆️ Restore CSV")
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
            st.success("Restored!"); st.rerun()

    # ── Main panel ─────────────────────────────────────────────────────────────
    df = st.session_state.df

    # Overview page
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
                "Frame":      rec["frame_name"],
                "Photo":      "✅" if str(rec.get("image_path", "")) else "—",
                "Missing":    missing,
                "Present":    N_COORDS - missing,
                "% Present":  f"{(N_COORDS - missing) / N_COORDS * 100:.1f}%",
                "Updated":    rec.get("timestamp", ""),
            })
        st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)

        st.markdown("#### Load a frame")
        frame_ids   = df["frame_id"].tolist()
        frame_names = df["frame_name"].tolist()
        cols        = st.columns(min(len(frame_ids), 5))
        for i, (fid, fname) in enumerate(zip(frame_ids, frame_names)):
            with cols[i % 5]:
                if st.button(fname, key=f"ov_{fid}", use_container_width=True):
                    load_frame(fid); st.rerun()
        return

    # ── Active frame ───────────────────────────────────────────────────────────
    active_id   = st.session_state.active_frame_id
    active_row  = df[df["frame_id"] == active_id].iloc[0]
    frame_name  = active_row["frame_name"]
    coord_dict  = st.session_state.coord_dict
    photo       = st.session_state.frame_image

    frame_ids   = df["frame_id"].tolist()
    cur_idx     = frame_ids.index(active_id) if active_id in frame_ids else 0
    total_fr    = len(frame_ids)

    missing_list  = sorted(c for c, v in coord_dict.items() if not v)
    present_count = N_COORDS - len(missing_list)

    # Header + navigation
    h_left, h_right = st.columns([3, 1])
    with h_left:
        st.markdown(f"## 🍫 {frame_name}")
        st.caption(
            f"Frame {cur_idx + 1} of {total_fr}  |  "
            f"Last saved: {active_row.get('timestamp', '—')}")
    with h_right:
        n1, n2 = st.columns(2)
        if n1.button("◀ Prev", use_container_width=True, disabled=(total_fr <= 1)):
            navigate(-1); st.rerun()
        if n2.button("Next ▶", use_container_width=True, disabled=(total_fr <= 1),
                     type="primary"):
            navigate(+1); st.rerun()

    # Stats row
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            f'<div class="stat-box"><div class="stat-number green">{present_count}</div>'
            f'<div class="stat-label">Present</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(
            f'<div class="stat-box"><div class="stat-number red">{len(missing_list)}</div>'
            f'<div class="stat-label">Missing</div></div>', unsafe_allow_html=True)
    with c3:
        pct = present_count / N_COORDS * 100
        st.markdown(
            f'<div class="stat-box"><div class="stat-number">{pct:.1f}%</div>'
            f'<div class="stat-label">Fill Rate</div></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(
            f'<div class="stat-box"><div class="stat-number">{N_COORDS}</div>'
            f'<div class="stat-label">Total Positions</div></div>', unsafe_allow_html=True)

    st.divider()

    # Photo attachment expander
    with st.expander(
            "📷 " + ("Replace frame photo" if photo else "Attach a photo to this frame"),
            expanded=(photo is None)):
        per_frame_upload = st.file_uploader(
            "Upload mold photo",
            type=["png", "jpg", "jpeg", "bmp", "tiff", "webp"],
            key=f"photo_{active_id}", label_visibility="collapsed")
        if per_frame_upload:
            ipath  = save_frame_image(active_id, per_frame_upload)
            photo  = Image.open(per_frame_upload).convert("RGB")
            st.session_state.frame_image = photo
            st.session_state.df = upsert_frame(
                st.session_state.df, active_id, frame_name, coord_dict, ipath)
            save_data(st.session_state.df)
            st.success("Photo saved."); st.rerun()

    if not photo:
        st.info("📷 Attach a photo to this frame to enable click-to-toggle inspection.")
        st.markdown(
            "**No photo?** Use the expander above to upload a mold image. "
            "Once uploaded, click any cavity cell directly on the photo to mark it present or missing.")
        return

    # ── Photo Inspector (click-to-toggle, no separate grid tab) ───────────────

    # Toolbar
    t1, t2, t3, t4 = st.columns([1, 1, 1, 3])
    with t1:
        if st.button("✅ All Present", use_container_width=True, key="all_pres"):
            for c in ALL_COORDS:
                st.session_state.coord_dict[c] = True
            auto_save(active_id, frame_name); st.rerun()
    with t2:
        if st.button("❌ All Missing", use_container_width=True, key="all_miss"):
            for c in ALL_COORDS:
                st.session_state.coord_dict[c] = False
            auto_save(active_id, frame_name); st.rerun()
    with t3:
        if st.button("🔄 Invert", use_container_width=True, key="inv"):
            for c in ALL_COORDS:
                st.session_state.coord_dict[c] = not st.session_state.coord_dict[c]
            auto_save(active_id, frame_name); st.rerun()
    with t4:
        if missing_list:
            badges = " ".join(
                f'<span class="missing-badge">{c}</span>'
                for c in missing_list)
            st.markdown(
                f'<div style="line-height:2;padding-top:2px">'
                f'<strong style="color:#E74C3C">Missing ({len(missing_list)}):</strong> '
                f'{badges}</div>', unsafe_allow_html=True)
        else:
            st.success("All 120 positions present!")

    opacity = st.slider("Overlay opacity", 20, 220, 20, 5, key="opacity")

    # ── Clickable photo via HTML canvas ───────────────────────────────────────
    # Strategy: render a canvas-based overlay in an iframe.
    # User clicks → JS resolves grid coord → updates URL query param → Streamlit reruns.
    # This avoids any round-trip image encode on every click.
    st.markdown(
        "**Click any cell** on the photo to toggle present ↔ missing. "
        "Changes save instantly. Hover to see coordinate.",
        help="The coloured overlay aligns with the mold grid. "
             "Green = present, Red = missing.")

    # Build the HTML with a JS bridge that sets window.location search param
    W_photo, H_photo = photo.size
    display_w = 860

    buf_photo = io.BytesIO()
    photo.save(buf_photo, format="JPEG", quality=82)
    img_b64 = base64.b64encode(buf_photo.getvalue()).decode()

    coord_json = json.dumps({c: (1 if v else 0) for c, v in coord_dict.items()})

    # We use postMessage → parent window URL manipulation.
    # The trick: the iframe posts the coord; a small <script> in the main page
    # would normally handle it. Since Streamlit doesn't expose that,
    # we set location.href of the top frame to trigger the query param.
    html_canvas = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{ background:transparent; overflow:hidden; }}
  canvas {{ display:block; cursor:crosshair; border-radius:6px;
            box-shadow:0 4px 24px rgba(0,0,0,0.35); }}
  #tip {{
    position:fixed; background:rgba(0,0,0,0.78); color:#fff;
    padding:4px 10px; border-radius:6px; font:600 13px/1.4 monospace;
    pointer-events:none; display:none; z-index:9999;
  }}
</style>
</head>
<body>
<canvas id="c"></canvas>
<div id="tip"></div>
<script>
const COLS  = {json.dumps(COLS)};
const ROWS  = {json.dumps(ROWS)};
const STATE = {coord_json};
const OPACITY = {opacity / 255:.3f};

const canvas = document.getElementById('c');
const ctx    = canvas.getContext('2d');
const tip    = document.getElementById('tip');

const DW = {display_w};
const img = new Image();
img.src   = 'data:image/jpeg;base64,{img_b64}';
img.onload = () => {{
  const DH = Math.round(DW * img.naturalHeight / img.naturalWidth);
  canvas.width  = DW;
  canvas.height = DH;
  document.body.style.height = DH + 'px';
  draw(DW, DH);
}};

function gridGeom(W, H) {{
  const mL = Math.max(24, W * 0.038);
  const mT = Math.max(20, H * 0.055);
  const gW = W - mL - Math.max(4, W * 0.008);
  const gH = H - mT - Math.max(4, H * 0.008);
  return {{ mL, mT, cW: gW / COLS.length, cH: gH / ROWS.length }};
}}

function draw(W, H) {{
  ctx.clearRect(0, 0, W, H);
  ctx.drawImage(img, 0, 0, W, H);
  const {{ mL, mT, cW, cH }} = gridGeom(W, H);
  const fs = Math.max(9, Math.min(cW, cH) * 0.30);
  ctx.font         = `bold ${{fs}}px sans-serif`;
  ctx.textAlign    = 'center';
  ctx.textBaseline = 'middle';

  ROWS.forEach((rn, ri) => {{
    COLS.forEach((col, ci) => {{
      const coord   = col + rn;
      const present = STATE[coord] === 1;
      const x0 = mL + ci * cW, y0 = mT + ri * cH;
      ctx.fillStyle = present
        ? `rgba(46,204,113,${{OPACITY}})`
        : `rgba(231,76,60,${{OPACITY}})`;
      ctx.fillRect(x0+1, y0+1, cW-2, cH-2);
      ctx.strokeStyle = 'rgba(255,255,255,0.7)';
      ctx.lineWidth   = 1;
      ctx.strokeRect(x0, y0, cW, cH);
      ctx.fillStyle = present ? 'rgba(10,10,10,0.9)' : 'rgba(255,255,255,0.95)';
      ctx.fillText(coord, x0 + cW/2, y0 + cH/2);
    }});
  }});

  const hs = Math.max(10, Math.min(mL, mT) * 0.55);
  ctx.font      = `bold ${{hs}}px sans-serif`;
  ctx.fillStyle = 'rgba(255,255,255,0.9)';
  COLS.forEach((col, ci) => ctx.fillText(col, mL+(ci+0.5)*cW, mT/2));
  ROWS.forEach((rn, ri)  => ctx.fillText(String(rn), mL/2, mT+(ri+0.5)*cH));
}}

function hitTest(px, py, W, H) {{
  const {{ mL, mT, cW, cH }} = gridGeom(W, H);
  if (px < mL || py < mT) return null;
  const ci = Math.floor((px - mL) / cW);
  const ri = Math.floor((py - mT) / cH);
  if (ci < 0 || ci >= COLS.length || ri < 0 || ri >= ROWS.length) return null;
  return COLS[ci] + ROWS[ri];
}}

canvas.addEventListener('mousemove', e => {{
  const r     = canvas.getBoundingClientRect();
  const scX   = canvas.width  / r.width;
  const scY   = canvas.height / r.height;
  const coord = hitTest((e.clientX-r.left)*scX, (e.clientY-r.top)*scY,
                        canvas.width, canvas.height);
  if (coord) {{
    tip.style.display = 'block';
    tip.style.left    = (e.clientX + 14) + 'px';
    tip.style.top     = (e.clientY - 28) + 'px';
    tip.textContent   = coord + '  ' + (STATE[coord]===1 ? '✅ Present' : '❌ Missing');
  }} else {{
    tip.style.display = 'none';
  }}
}});
canvas.addEventListener('mouseleave', () => tip.style.display = 'none');

canvas.addEventListener('click', e => {{
  const r     = canvas.getBoundingClientRect();
  const scX   = canvas.width  / r.width;
  const scY   = canvas.height / r.height;
  const coord = hitTest((e.clientX-r.left)*scX, (e.clientY-r.top)*scY,
                        canvas.width, canvas.height);
  if (!coord) return;
  // Optimistic toggle on canvas
  STATE[coord] = STATE[coord] === 1 ? 0 : 1;
  draw(canvas.width, canvas.height);
  // Tell Streamlit via top-level URL (query param)
  window.top.location.href =
    window.top.location.pathname + '?toggle=' + coord;
}});
</script>
</body>
</html>
"""

    components.html(html_canvas, height=int(display_w * H_photo / W_photo) + 20,
                    scrolling=False)

    # Download overlay image (server-rendered, for export)
    st.markdown("---")
    exp_col1, exp_col2 = st.columns([2, 1])
    with exp_col1:
        with st.expander("⌨️ Quick-entry: paste missing coordinates"):
            st.markdown(
                "Comma/space-separated list of missing coords. "
                "**Replaces** current missing set.")
            with st.form("quick_entry"):
                raw = st.text_area(
                    "Missing coordinates",
                    value=", ".join(missing_list),
                    height=70, placeholder="e.g. A1, B3, G5, O8",
                    label_visibility="collapsed")
                c_apply, c_add = st.columns(2)
                apply_btn = c_apply.form_submit_button(
                    "Replace missing", use_container_width=True, type="primary")
                add_btn   = c_add.form_submit_button(
                    "Add to missing", use_container_width=True)
            if apply_btn or add_btn:
                tokens = re.split(r"[\s,;]+", raw.strip().upper())
                valid, invalid = [], []
                for t in tokens:
                    if not t: continue
                    (valid if t in ALL_COORDS else invalid).append(t)
                if invalid:
                    st.error(f"Invalid coordinates: {', '.join(invalid)}")
                else:
                    if apply_btn:
                        for c in ALL_COORDS:
                            st.session_state.coord_dict[c] = (c not in valid)
                    else:
                        for c in valid:
                            st.session_state.coord_dict[c] = False
                    auto_save(active_id, frame_name)
                    st.rerun()

    with exp_col2:
        ov_img = render_overlay(photo, coord_dict, opacity=opacity)
        buf_dl = io.BytesIO()
        ov_img.save(buf_dl, format="PNG")
        st.download_button(
            "⬇️ Download overlay image (.png)",
            data=buf_dl.getvalue(),
            file_name=f"{re.sub(r'[^a-zA-Z0-9_-]', '_', frame_name)}_overlay.png",
            mime="image/png", use_container_width=True)

        if missing_list:
            st.subheader(f"🔴 Missing ({len(missing_list)})")
            rows_missing: dict = {}
            for c in missing_list:
                rows_missing.setdefault(c[1:], []).append(c)
            for rn, coords_in_row in sorted(rows_missing.items(), key=lambda x: int(x[0])):
                st.markdown(
                    f"**Row {rn}:** " + " ".join(
                        f'<span class="missing-badge">{c}</span>'
                        for c in coords_in_row),
                    unsafe_allow_html=True)
        else:
            st.success("🎉 All 120 positions present!")

    # Bottom navigation
    st.divider()
    bot_l, bot_m, bot_r = st.columns([1, 2, 1])
    with bot_l:
        if st.button("◄ Previous", use_container_width=True, disabled=(total_fr <= 1)):
            navigate(-1); st.rerun()
    with bot_m:
        st.markdown(
            f"<div style='text-align:center;padding-top:8px'>"
            f"Frame <b>{cur_idx + 1}</b> of <b>{total_fr}</b></div>",
            unsafe_allow_html=True)
    with bot_r:
        if st.button("Next ►", use_container_width=True,
                     disabled=(total_fr <= 1), type="primary"):
            navigate(+1); st.rerun()


if __name__ == "__main__":
    main()
