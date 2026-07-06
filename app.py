"""
Chocolate Mold Inspector — Streamlit App
15 columns (A–O) × 8 rows (1–8) = 120 coordinates per mold frame

Save model
──────────
• Bulk actions (All Present / All Missing / Invert) → auto-save immediately.
• Quick-entry text form                             → auto-save on submit.
• Sidebar operations (upload, delete, restore)      → auto-save immediately.
• Navigation (Prev / Next)                          → prompts save if unsaved.
• Photo canvas clicks                               → accumulate in JS; a
  "💾 Save Changes" button on the right panel
  commits the full pending state in one write.
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
from reportlab.lib.enums import TA_CENTER

# ── Constants ──────────────────────────────────────────────────────────────────
COLS       = list("ABCDEFGHIJKLMNO")
ROWS       = list(range(1, 9))
ALL_COORDS = [f"{c}{r}" for r in ROWS for c in COLS]
N_COORDS   = len(ALL_COORDS)          # 120
COORD_SET  = set(ALL_COORDS)

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
    row_data = {
        "frame_id":   frame_id,
        "frame_name": frame_name,
        "timestamp":  datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "image_path": image_path,
        **{c: ("1" if v else "0") for c, v in coord_dict.items()},
    }
    if frame_id in df["frame_id"].values:
        for k, v in row_data.items():
            df.loc[df["frame_id"] == frame_id, k] = v
    else:
        df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
    return df


def _image_path_for(frame_id: str) -> str:
    row = st.session_state.df[st.session_state.df["frame_id"] == frame_id]
    return str(row.iloc[0].get("image_path", "")) if not row.empty else ""


# ── Save helpers ───────────────────────────────────────────────────────────────

def persist(frame_id: str, frame_name: str, coord_dict: dict | None = None) -> None:
    """Write coord_dict (or session_state.coord_dict) → DataFrame → CSV."""
    if coord_dict is None:
        coord_dict = st.session_state.coord_dict
    st.session_state.df = upsert_frame(
        st.session_state.df, frame_id, frame_name,
        coord_dict, _image_path_for(frame_id))
    save_data(st.session_state.df)
    st.session_state.unsaved = False


def auto_save(frame_id: str, frame_name: str, coord_dict: dict | None = None) -> None:
    """Convenience alias for bulk actions that always persist immediately."""
    persist(frame_id, frame_name, coord_dict)


# ── Image / file helpers ───────────────────────────────────────────────────────

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


def render_overlay(photo: Image.Image, coord_dict: dict, opacity: int = 20) -> Image.Image:
    """Server-side PIL render of the mold overlay (used for PNG export only)."""
    base   = photo.convert("RGBA")
    W, H   = base.size
    mL     = max(24, int(W * 0.038))
    mT     = max(20, int(H * 0.055))
    cell_w = (W - mL - max(4, int(W * 0.008))) / len(COLS)
    cell_h = (H - mT - max(4, int(H * 0.008))) / len(ROWS)

    overlay = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw    = ImageDraw.Draw(overlay)
    fs      = max(9, int(min(cell_w, cell_h) * 0.30))
    try:
        fnt  = ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", fs)
        fhdr = ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            max(10, int(min(mL, mT) * 0.55)))
    except Exception:
        fnt = fhdr = ImageFont.load_default()

    for ri, row_num in enumerate(ROWS):
        for ci, col in enumerate(COLS):
            coord   = f"{col}{row_num}"
            present = coord_dict.get(coord, True)
            x0 = mL + ci * cell_w;  y0 = mT + ri * cell_h
            x1 = x0 + cell_w;       y1 = y0 + cell_h
            fill = (46, 204, 113, opacity) if present else (231, 76, 60, opacity)
            draw.rectangle([x0+1, y0+1, x1-1, y1-1], fill=fill)
            draw.rectangle([x0, y0, x1, y1], outline=(255, 255, 255, 180), width=1)
            draw.text(((x0+x1)/2, (y0+y1)/2), coord,
                      fill=((10,10,10,255) if present else (255,255,255,255)),
                      font=fnt, anchor="mm")

    for ci, col in enumerate(COLS):
        draw.text((mL+(ci+0.5)*cell_w, mT/2), col,
                  fill=(255,255,255,230), font=fhdr, anchor="mm")
    for ri, row_num in enumerate(ROWS):
        draw.text((mL/2, mT+(ri+0.5)*cell_h), str(row_num),
                  fill=(255,255,255,230), font=fhdr, anchor="mm")

    return Image.alpha_composite(base, overlay).convert("RGB")


# ── Canvas HTML ────────────────────────────────────────────────────────────────

def build_canvas_html(photo: Image.Image, coord_dict: dict,
                      opacity: int = 20, display_w: int = 860) -> str:
    """
    Renders an interactive <canvas> in an iframe.

    Interaction model:
      • Each click optimistically toggles the cell colour on the canvas.
      • Toggled state is accumulated in JS (pendingState).
      • The iframe posts pendingState to the parent via postMessage only when
        the Python "Save Changes" button triggers a JS call (see bridge below).
      • Separately, the iframe posts state on every click so the right-panel
        pending list can refresh — this uses a lightweight postMessage that
        Streamlit's component bridge reads via st.query_params on the next rerun.

    To avoid a full page-navigation rerun on every click the iframe posts a
    message; a tiny <script> injected into the Streamlit page receives it and
    writes the serialised pending state into a hidden URL fragment, which
    Streamlit reads as a query param on the NEXT rerun triggered only by the
    Save button (a normal st.button click).  Because we store the entire
    pending state in the URL fragment, no state is lost between reruns.
    """
    W_img, H_img = photo.size
    display_h    = int(display_w * H_img / W_img)

    buf = io.BytesIO()
    photo.save(buf, format="JPEG", quality=82)
    img_b64 = base64.b64encode(buf.getvalue()).decode()

    coord_json = json.dumps({c: (1 if v else 0) for c, v in coord_dict.items()})

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ background:transparent; overflow:hidden; }}
canvas {{ display:block; cursor:crosshair;
          border-radius:6px; box-shadow:0 4px 24px rgba(0,0,0,0.30); }}
#tip {{
  position:fixed; background:rgba(0,0,0,0.78); color:#fff;
  padding:4px 10px; border-radius:6px; font:600 13px/1.4 monospace;
  pointer-events:none; display:none; z-index:9999; white-space:nowrap;
}}
</style>
</head>
<body>
<canvas id="c"></canvas>
<div id="tip"></div>
<script>
const COLS    = {json.dumps(COLS)};
const ROWS    = {json.dumps(ROWS)};
// saved  = what is currently persisted (never mutated after init)
// pending = working copy; accumulates unsaved toggles
const saved   = {coord_json};
const pending = Object.assign({{}}, saved);

const canvas = document.getElementById('c');
const ctx    = canvas.getContext('2d');
const tip    = document.getElementById('tip');

const DW  = {display_w};
const DH  = {display_h};
canvas.width  = DW;
canvas.height = DH;
document.body.style.height = DH + 'px';

const OPACITY = {opacity / 255:.4f};

const img = new Image();
img.onload = () => draw();
img.src = 'data:image/jpeg;base64,{img_b64}';

function geom() {{
  const mL = Math.max(24, DW * 0.038);
  const mT = Math.max(20, DH * 0.055);
  const cW = (DW - mL - Math.max(4, DW * 0.008)) / COLS.length;
  const cH = (DH - mT - Math.max(4, DH * 0.008)) / ROWS.length;
  return {{ mL, mT, cW, cH }};
}}

function draw() {{
  const {{ mL, mT, cW, cH }} = geom();
  ctx.clearRect(0, 0, DW, DH);
  ctx.drawImage(img, 0, 0, DW, DH);

  const fs = Math.max(9, Math.min(cW, cH) * 0.30);
  ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
  ctx.font = `bold ${{fs}}px sans-serif`;

  ROWS.forEach((rn, ri) => {{
    COLS.forEach((col, ci) => {{
      const coord   = col + rn;
      const present = pending[coord] === 1;
      const dirty   = pending[coord] !== saved[coord];
      const x0 = mL + ci * cW, y0 = mT + ri * cH;

      // Fill: use orange tint for unsaved toggles
      if (dirty) {{
        ctx.fillStyle = present
          ? `rgba(52,152,219,${{OPACITY * 2.2}})`   // blue = restored (unsaved)
          : `rgba(241,196,15,${{OPACITY * 2.2}})`;  // amber = newly missing (unsaved)
      }} else {{
        ctx.fillStyle = present
          ? `rgba(46,204,113,${{OPACITY}})`
          : `rgba(231,76,60,${{OPACITY}})`;
      }}
      ctx.fillRect(x0+1, y0+1, cW-2, cH-2);

      // Border: highlight dirty cells
      ctx.strokeStyle = dirty ? 'rgba(255,220,0,0.9)' : 'rgba(255,255,255,0.55)';
      ctx.lineWidth   = dirty ? 2 : 1;
      ctx.strokeRect(x0, y0, cW, cH);

      ctx.fillStyle = present ? 'rgba(10,10,10,0.85)' : 'rgba(255,255,255,0.95)';
      ctx.fillText(coord, x0 + cW/2, y0 + cH/2);
    }});
  }});

  // Axis labels
  const hs = Math.max(10, Math.min(mL, mT) * 0.55);
  ctx.font = `bold ${{hs}}px sans-serif`;
  ctx.fillStyle = 'rgba(255,255,255,0.9)';
  COLS.forEach((col, ci) => ctx.fillText(col, mL+(ci+0.5)*cW, mT/2));
  ROWS.forEach((rn, ri)  => ctx.fillText(String(rn), mL/2, mT+(ri+0.5)*cH));
}}

function hitTest(px, py) {{
  const {{ mL, mT, cW, cH }} = geom();
  if (px < mL || py < mT) return null;
  const ci = Math.floor((px - mL) / cW);
  const ri = Math.floor((py - mT) / cH);
  if (ci < 0 || ci >= COLS.length || ri < 0 || ri >= ROWS.length) return null;
  return COLS[ci] + ROWS[ri];
}}

function scalePos(e) {{
  const r  = canvas.getBoundingClientRect();
  return [(e.clientX - r.left) * (DW / r.width),
          (e.clientY - r.top)  * (DH / r.height)];
}}

canvas.addEventListener('mousemove', e => {{
  const [px, py] = scalePos(e);
  const coord    = hitTest(px, py);
  if (coord) {{
    tip.style.display = 'block';
    tip.style.left    = (e.clientX + 14) + 'px';
    tip.style.top     = (e.clientY - 30) + 'px';
    const cur   = pending[coord] === 1;
    const dirty = pending[coord] !== saved[coord];
    tip.textContent = coord
      + '  ' + (cur ? '✅ Present' : '❌ Missing')
      + (dirty ? '  ✏️ unsaved' : '');
  }} else {{
    tip.style.display = 'none';
  }}
}});
canvas.addEventListener('mouseleave', () => tip.style.display = 'none');

canvas.addEventListener('click', e => {{
  const [px, py] = scalePos(e);
  const coord    = hitTest(px, py);
  if (!coord) return;
  pending[coord] = pending[coord] === 1 ? 0 : 1;
  draw();
  // Notify parent of the full pending state so it can refresh the panel
  window.parent.postMessage(
    {{ type: 'mold_pending', state: pending }}, '*'
  );
}});

// Listen for "commit" signal from the Save button (parent posts back)
window.addEventListener('message', ev => {{
  if (ev.data && ev.data.type === 'mold_commit') {{
    // Encode pending state into the parent URL so Streamlit picks it up
    const encoded = btoa(JSON.stringify(pending));
    window.top.location.href =
      window.top.location.pathname + '?mold_save=' + encoded;
  }}
}});
</script>
</body>
</html>"""


# ── Excel export ───────────────────────────────────────────────────────────────

def _interp_argb(count: int, max_count: int) -> str:
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
    thin       = Side(style="thin", color="999999")
    border     = Border(left=thin, right=thin, top=thin, bottom=thin)
    center     = Alignment(horizontal="center", vertical="center")
    hdr_fill   = PatternFill("solid", fgColor="FF2C3E50")
    sum_fill   = PatternFill("solid", fgColor="FF34495E")
    white_bold = Font(bold=True, color="FFFFFFFF", name="Arial", size=11)
    norm_font  = Font(name="Arial", size=10)
    white_font = Font(name="Arial", size=10, color="FFFFFFFF")

    n_frames = len(df)
    mc       = _missing_counts(df)
    max_miss = max(mc.values()) if mc else 1

    # Sheet 1 — Cumulative Heatmap
    ws = wb.active
    ws.title = "Cumulative Heatmap"
    nc = len(COLS) + 2
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    h = ws.cell(1, 1, f"Chocolate Mold — Cumulative Missing  ({n_frames} frame(s))")
    h.font = Font(bold=True, color="FFFFFFFF", name="Arial", size=13)
    h.fill = hdr_fill; h.alignment = center
    ws.row_dimensions[1].height = 30

    ws.cell(2, 1, "").fill = hdr_fill
    for ci, col in enumerate(COLS, 2):
        c = ws.cell(2, ci, col)
        c.font = white_bold; c.fill = hdr_fill; c.alignment = center; c.border = border
    tot_col = nc
    c2 = ws.cell(2, tot_col, "Row\nMissing")
    c2.font = white_bold; c2.fill = sum_fill
    c2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c2.border = border
    ws.row_dimensions[2].height = 30
    ws.column_dimensions["A"].width = 6
    for ci in range(2, nc):
        ws.column_dimensions[get_column_letter(ci)].width = 10
    ws.column_dimensions[get_column_letter(tot_col)].width = 11

    for ri, row_num in enumerate(ROWS):
        er = ri + 3
        ws.row_dimensions[er].height = 24
        rh = ws.cell(er, 1, str(row_num))
        rh.font = white_bold; rh.fill = hdr_fill; rh.alignment = center; rh.border = border
        row_total = 0
        for ci, col in enumerate(COLS, 2):
            count = mc.get(f"{col}{row_num}", 0); row_total += count
            t = count / max_miss if max_miss else 0
            cell = ws.cell(er, ci, count)
            cell.fill = PatternFill("solid", fgColor=_interp_argb(count, max_miss))
            cell.font = white_font if t > 0.45 else norm_font
            cell.alignment = center; cell.border = border
        rt = ws.cell(er, tot_col, row_total)
        rt.font = Font(bold=True, name="Arial", size=10, color="FFFFFFFF")
        rt.fill = sum_fill; rt.alignment = center; rt.border = border

    tr = len(ROWS) + 3
    ws.row_dimensions[tr].height = 24
    tl = ws.cell(tr, 1, "Total")
    tl.font = Font(bold=True, color="FFFFFFFF", name="Arial", size=10)
    tl.fill = sum_fill; tl.alignment = center; tl.border = border
    grand = 0
    for ci, col in enumerate(COLS, 2):
        ct = sum(mc.get(f"{col}{r}", 0) for r in ROWS); grand += ct
        c = ws.cell(tr, ci, ct)
        c.font = Font(bold=True, name="Arial", size=10, color="FFFFFFFF")
        c.fill = sum_fill; c.alignment = center; c.border = border
    gt = ws.cell(tr, tot_col, grand)
    gt.font = Font(bold=True, name="Arial", size=11, color="FFFFFFFF")
    gt.fill = PatternFill("solid", fgColor="FF1A252F"); gt.alignment = center; gt.border = border

    lr = tr + 2
    ws.merge_cells(start_row=lr, start_column=1, end_row=lr, end_column=nc)
    leg = ws.cell(lr, 1,
        f"Green = never missing  →  Red = missing in all {n_frames} frame(s)  |  "
        "Each cell = count of frames where that cavity was absent")
    leg.font = Font(italic=True, name="Arial", size=9, color="FF555555")
    leg.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[lr].height = 18

    # Sheet 2 — Frame Summary
    ws2 = wb.create_sheet("Frame Summary", 1)
    hdrs2 = ["Frame", "Missing", "Present", "Total", "% Present", "Timestamp"]
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(hdrs2))
    t2 = ws2.cell(1, 1, "Chocolate Mold Inspection — Frame Summary")
    t2.font = Font(bold=True, color="FFFFFFFF", name="Arial", size=14)
    t2.fill = hdr_fill; t2.alignment = center; ws2.row_dimensions[1].height = 30
    for ci, h2 in enumerate(hdrs2, 1):
        c = ws2.cell(2, ci, h2)
        c.font = white_bold; c.fill = sum_fill; c.alignment = center; c.border = border
    ws2.row_dimensions[2].height = 22
    for ci, w in enumerate([30, 10, 10, 10, 12, 22], 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w

    for ri, (_, rec) in enumerate(df.iterrows(), 3):
        miss = sum(1 for c in ALL_COORDS if str(rec.get(c, "1")) == "0")
        pres = N_COORDS - miss
        for ci, val in enumerate(
            [str(rec.get("frame_name", rec["frame_id"])),
             miss, pres, N_COORDS, f"=C{ri}/D{ri}",
             str(rec.get("timestamp", ""))], 1):
            cell = ws2.cell(ri, ci, val)
            cell.alignment = center; cell.border = border; cell.font = norm_font
            if ci == 2 and N_COORDS:
                inten = min(miss / N_COORDS, 1.0)
                rv = int(231*inten + 46*(1-inten))
                gv = int(76 *inten + 204*(1-inten))
                bv = int(60 *inten + 113*(1-inten))
                cell.fill = PatternFill("solid", fgColor=f"FF{rv:02X}{gv:02X}{bv:02X}")
            if ci == 5:
                cell.number_format = "0.0%"
        ws2.row_dimensions[ri].height = 20

    last = len(df) + 2; tot_r = last + 1
    ws2.cell(tot_r, 1, "TOTALS").font = Font(bold=True, name="Arial", size=10)
    if not df.empty:
        for ci, formula in enumerate(
            [None, f"=SUM(B3:B{last})", f"=SUM(C3:C{last})",
             f"=SUM(D3:D{last})", f"=C{tot_r}/D{tot_r}", None], 1):
            if formula:
                c = ws2.cell(tot_r, ci, formula)
                c.font = Font(bold=True, name="Arial")
                if ci == 5: c.number_format = "0.0%"
    for ci in range(1, 7):
        ws2.cell(tot_r, ci).border = border
        ws2.cell(tot_r, ci).alignment = center
    ws2.row_dimensions[tot_r].height = 22

    # Sheet 3 — Coordinate Frequency
    ws3 = wb.create_sheet("Coordinate Frequency", 2)
    fhdrs = ["Coordinate", "Times Missing", "Total Frames", "% Flagged", "Rank"]
    ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(fhdrs))
    ft = ws3.cell(1, 1, f"Coordinate Flag Frequency  —  {n_frames} frame(s)")
    ft.font = Font(bold=True, color="FFFFFFFF", name="Arial", size=13)
    ft.fill = hdr_fill; ft.alignment = center; ws3.row_dimensions[1].height = 30
    for ci, h3 in enumerate(fhdrs, 1):
        c = ws3.cell(2, ci, h3)
        c.font = white_bold; c.fill = sum_fill; c.alignment = center; c.border = border
    ws3.row_dimensions[2].height = 22
    for ci, w in enumerate([14, 16, 14, 12, 8], 1):
        ws3.column_dimensions[get_column_letter(ci)].width = w

    sorted_coords = sorted(ALL_COORDS, key=lambda c: (-mc.get(c, 0), c))
    for rank, coord in enumerate(sorted_coords, 1):
        er = rank + 2
        cnt = mc.get(coord, 0)
        t   = cnt / max_miss if max_miss else 0
        argb = _interp_argb(cnt, max_miss or 1)
        flag_font = Font(name="Arial", size=10,
                         color="FFFFFFFF" if t > 0.45 else "FF000000", bold=(cnt > 0))
        for ci, val in enumerate([coord, cnt, n_frames, cnt/n_frames if n_frames else 0, rank], 1):
            cell = ws3.cell(er, ci, val)
            cell.alignment = center; cell.border = border; cell.font = norm_font
            if ci == 4:
                cell.number_format = "0.0%"; cell.fill = PatternFill("solid", fgColor=argb)
                cell.font = flag_font
            elif ci == 2 and cnt > 0:
                inten = min(cnt / n_frames, 1.0)
                rv = int(231*inten + 236*(1-inten))
                gv = int(76 *inten + 240*(1-inten))
                bv = int(60 *inten + 241*(1-inten))
                cell.fill = PatternFill("solid", fgColor=f"FF{rv:02X}{gv:02X}{bv:02X}")
        ws3.row_dimensions[er].height = 18

    footer = len(ALL_COORDS) + 3
    ws3.merge_cells(start_row=footer, start_column=1, end_row=footer, end_column=len(fhdrs))
    total_ev = sum(mc.values())
    ws3.cell(footer, 1,
        f"Total events: {total_ev}  |  Avg/frame: {total_ev/n_frames:.1f}  |  "
        f"Never flagged: {sum(1 for v in mc.values() if v == 0)}")
    ws3.cell(footer, 1).font = Font(italic=True, name="Arial", size=9, color="FF555555")
    ws3.cell(footer, 1).alignment = Alignment(horizontal="left", vertical="center")
    ws3.row_dimensions[footer].height = 18

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


# ── PDF export ─────────────────────────────────────────────────────────────────

def _rl(hex6: str) -> colors.Color:
    h = hex6.lstrip("#")
    return colors.Color(int(h[:2],16)/255, int(h[2:4],16)/255, int(h[4:],16)/255)


def _rl_interp(count: int, max_count: int) -> colors.Color:
    t = min(count/max_count, 1.0) if max_count > 0 else 0.0
    return colors.Color(
        (46+(231-46)*t)/255, (204+(76-204)*t)/255, (113+(60-113)*t)/255)


def build_pdf(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=15*mm, bottomMargin=15*mm,
        title="Chocolate Mold Inspection Report")

    DARK = _rl("#2C3E50"); MID = _rl("#34495E")
    W = colors.white;      B = colors.black

    ts = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=14,
                        leading=18, textColor=W, alignment=TA_CENTER)
    cs = ParagraphStyle("c", fontName="Helvetica-Oblique", fontSize=7,
                        leading=10, textColor=colors.HexColor("#555555"))

    n_frames = len(df)
    mc       = _missing_counts(df)
    max_miss = max(mc.values()) if mc else 1
    now_str  = datetime.now().strftime("%Y-%m-%d %H:%M")

    def hdr(txt, sz=9):
        return Paragraph(f"<b>{txt}</b>", ParagraphStyle(
            "h", fontName="Helvetica-Bold", fontSize=sz, textColor=W, alignment=TA_CENTER))

    def ttbl(text):
        tbl = Table([[Paragraph(text, ts)]], colWidths=[doc.width])
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0),(-1,-1), DARK),
            ("TOPPADDING", (0,0),(-1,-1), 7),
            ("BOTTOMPADDING",(0,0),(-1,-1), 7)]))
        return tbl

    story = []

    # Page 1 — Heatmap
    story += [ttbl(f"Chocolate Mold — Cumulative Missing  ({n_frames} frame(s))"),
              Spacer(1, 6*mm)]
    rw = 14*mm; tw = 18*mm
    cw = [(doc.width-rw-tw)/len(COLS)]
    col_ws = [rw] + cw*len(COLS) + [tw]

    grid = [[""] + [hdr(c) for c in COLS] + [hdr("Row\nTotal", 7)]]
    for rn in ROWS:
        row_t = 0; row = [hdr(str(rn))]
        for col in COLS:
            cnt = mc.get(f"{col}{rn}", 0); row_t += cnt
            row.append(str(cnt) if cnt else "")
        row.append(hdr(str(row_t))); grid.append(row)
    tot_r = [hdr("Total")]; grand = 0
    for col in COLS:
        ct = sum(mc.get(f"{col}{r}", 0) for r in ROWS); grand += ct; tot_r.append(hdr(str(ct)))
    tot_r.append(hdr(str(grand), 10)); grid.append(tot_r)

    gcmds = [
        ("FONTNAME",(0,0),(-1,-1),"Helvetica"), ("FONTSIZE",(0,0),(-1,-1),8),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),       ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#CCCCCC")),
        ("TOPPADDING",(0,0),(-1,-1),3),          ("BOTTOMPADDING",(0,0),(-1,-1),3),
        ("BACKGROUND",(0,0),(-1,0),DARK),        ("BACKGROUND",(0,1),(0,-2),DARK),
        ("BACKGROUND",(-1,1),(-1,-2),MID),       ("BACKGROUND",(0,-1),(-1,-1),MID),
        ("BACKGROUND",(-1,-1),(-1,-1),_rl("#1A252F")),
    ]
    for ri, rn in enumerate(ROWS, 1):
        for ci, col in enumerate(COLS, 1):
            cnt = mc.get(f"{col}{rn}", 0); tv = cnt/max_miss if max_miss else 0
            gcmds += [("BACKGROUND",(ci,ri),(ci,ri),_rl_interp(cnt,max_miss)),
                      ("TEXTCOLOR",(ci,ri),(ci,ri), W if tv>0.45 else B)]
    story.append(Table(grid, colWidths=col_ws, style=TableStyle(gcmds)))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph(
        f"Green = never missing  |  Red = missing in all {n_frames} frames  |  "
        f"Generated: {now_str}", cs))

    # Page 2 — Frame Summary
    story += [PageBreak(), ttbl("Chocolate Mold — Frame Summary"), Spacer(1, 6*mm)]
    scw = [doc.width*f for f in [0.28,0.10,0.10,0.10,0.12,0.30]]
    shdrs = ["Frame","Missing","Present","Total","% Present","Timestamp"]
    srows = [[hdr(h) for h in shdrs]]
    for _, rec in df.iterrows():
        miss = sum(1 for c in ALL_COORDS if str(rec.get(c,"1"))=="0")
        pres = N_COORDS - miss
        srows.append([
            Paragraph(str(rec.get("frame_name", rec["frame_id"])),
                      ParagraphStyle("fn", fontName="Helvetica", fontSize=8)),
            miss, pres, N_COORDS, f"{pres/N_COORDS*100:.1f}%",
            Paragraph(str(rec.get("timestamp","")),
                      ParagraphStyle("ts2", fontName="Helvetica", fontSize=7,
                                     textColor=colors.HexColor("#555555"),
                                     alignment=TA_CENTER)),
        ])
    total_m = sum(sum(1 for c in ALL_COORDS if str(rec.get(c,"1"))=="0")
                  for _,rec in df.iterrows())
    total_p = N_COORDS*n_frames - total_m; total_t = N_COORDS*n_frames
    srows.append([hdr("TOTALS"), hdr(str(total_m)), hdr(str(total_p)),
                  hdr(str(total_t)),
                  hdr(f"{total_p/total_t*100:.1f}%" if total_t else "—"), ""])
    scmds = [
        ("FONTSIZE",(0,0),(-1,-1),8), ("ALIGN",(1,0),(-1,-1),"CENTER"),
        ("ALIGN",(0,0),(0,-1),"LEFT"), ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#CCCCCC")),
        ("TOPPADDING",(0,0),(-1,-1),3), ("BOTTOMPADDING",(0,0),(-1,-1),3),
        ("BACKGROUND",(0,0),(-1,0),DARK), ("BACKGROUND",(0,-1),(-1,-1),MID),
        ("LEFTPADDING",(0,0),(0,-1),6),
    ]
    for ri, (_,rec) in enumerate(df.iterrows(), 1):
        miss = sum(1 for c in ALL_COORDS if str(rec.get(c,"1"))=="0")
        inten = min(miss/N_COORDS, 1.0)
        scmds += [("BACKGROUND",(1,ri),(1,ri),_rl_interp(int(inten*max_miss),max_miss)),
                  ("TEXTCOLOR",(1,ri),(1,ri), W if inten>0.45 else B)]
    story.append(Table(srows, colWidths=scw, style=TableStyle(scmds)))

    # Page 3 — Coordinate Frequency
    story += [PageBreak(),
              ttbl(f"Coordinate Flag Frequency  —  {n_frames} frame(s)"),
              Spacer(1, 6*mm)]
    fcw = [doc.width*f for f in [0.15,0.18,0.15,0.15,0.10]]
    fhdrs_pdf = ["Coordinate","Times Missing","Total Frames","% Flagged","Rank"]
    frows = [[hdr(h) for h in fhdrs_pdf]]
    scoords = sorted(ALL_COORDS, key=lambda c: (-mc.get(c,0), c))
    for rank, coord in enumerate(scoords, 1):
        cnt = mc.get(coord, 0)
        frows.append([coord, cnt, n_frames,
                      f"{cnt/n_frames*100:.1f}%" if n_frames else "0.0%", rank])
    total_ev = sum(mc.values()); avg = total_ev/n_frames if n_frames else 0
    frows.append([
        Paragraph(f"<b>Total: {total_ev}  |  Avg/frame: {avg:.1f}  |  "
                  f"Never flagged: {sum(1 for v in mc.values() if v==0)}</b>",
                  ParagraphStyle("ff", fontName="Helvetica-BoldOblique",
                                 fontSize=7, textColor=W)),
        "","","",""])
    fcmds = [
        ("FONTSIZE",(0,0),(-1,-1),8), ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#CCCCCC")),
        ("TOPPADDING",(0,0),(-1,-1),2), ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ("BACKGROUND",(0,0),(-1,0),DARK), ("BACKGROUND",(0,-1),(-1,-1),MID),
        ("SPAN",(0,-1),(-1,-1)),
    ]
    for ri, coord in enumerate(scoords, 1):
        cnt = mc.get(coord, 0); tv = cnt/max_miss if max_miss else 0
        fcmds += [("BACKGROUND",(3,ri),(3,ri),_rl_interp(cnt,max_miss)),
                  ("TEXTCOLOR",(3,ri),(3,ri), W if tv>0.45 else B)]
        if cnt > 0:
            inten = min(cnt/n_frames, 1.0)
            fcmds.append(("BACKGROUND",(1,ri),(1,ri),
                colors.Color((231*inten+236*(1-inten))/255,
                             (76 *inten+240*(1-inten))/255,
                             (60 *inten+241*(1-inten))/255)))
    story += [Table(frows, colWidths=fcw, style=TableStyle(fcmds)),
              Spacer(1,4*mm), Paragraph(f"Generated: {now_str}", cs)]
    doc.build(story)
    return buf.getvalue()


# ── Session state ──────────────────────────────────────────────────────────────

def init_state():
    defaults = {
        "df":              load_data(),
        "active_frame_id": None,
        "coord_dict":      {c: True for c in ALL_COORDS},
        "frame_image":     None,
        "unsaved":         False,   # True when canvas has pending (not yet saved) changes
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def load_frame(frame_id: str):
    st.session_state.active_frame_id = frame_id
    st.session_state.coord_dict      = get_coord_dict(st.session_state.df, frame_id)
    st.session_state.unsaved         = False
    row      = st.session_state.df[st.session_state.df["frame_id"] == frame_id]
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
        border-radius:10px; padding:14px 18px; text-align:center;
        background:#16213E; border:1px solid #2C3E50; margin-bottom:4px;
    }
    .stat-number { font-size:1.9rem; font-weight:700; }
    .stat-label  { font-size:0.82rem; color:#95A5A6; margin-top:3px; }
    .green { color:#27AE60; } .red { color:#E74C3C; }
    .amber { color:#F39C12; }
    .stButton > button { border-radius:8px; font-weight:600; }
    .upload-hint {
        border:1px dashed #3498DB; border-radius:8px;
        padding:12px 16px; font-size:0.88rem; margin-bottom:8px;
    }
    .badge {
        display:inline-block; padding:1px 7px; border-radius:4px;
        font-size:0.78rem; margin:1px;
    }
    .badge-red    { background:#E74C3C; color:#FFF; }
    .badge-amber  { background:#F39C12; color:#000; }
    .save-panel {
        background:#1A2A1A; border:2px solid #27AE60;
        border-radius:10px; padding:16px; position:sticky; top:80px;
    }
    .save-panel-dirty {
        background:#2A1A08; border:2px solid #F39C12;
        border-radius:10px; padding:16px; position:sticky; top:80px;
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Handle canvas save commit (query param written by JS) ─────────────────
    # When the user clicks "Save Changes", Python posts a message to the iframe,
    # which encodes its full pending state into ?mold_save=<base64-json> and
    # navigates the top frame → Streamlit reruns and we pick it up here.
    mold_save_param = st.query_params.get("mold_save")
    if mold_save_param and st.session_state.active_frame_id:
        try:
            raw_state = json.loads(base64.b64decode(mold_save_param).decode())
            # Validate and apply
            new_dict = {}
            for coord in ALL_COORDS:
                new_dict[coord] = (raw_state.get(coord, 1) == 1)
            st.session_state.coord_dict = new_dict
            active_id   = st.session_state.active_frame_id
            frame_name  = st.session_state.df.loc[
                st.session_state.df["frame_id"] == active_id, "frame_name"].iloc[0]
            persist(active_id, frame_name, new_dict)
            st.session_state.unsaved = False
        except Exception:
            st.warning("Could not parse save payload — please try again.")
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
            type=["png","jpg","jpeg","bmp","tiff","webp"],
            accept_multiple_files=True,
            label_visibility="collapsed",
            key="batch_uploader",
        )
        if uploaded_photos:
            if len(uploaded_photos) > MAX_UPLOAD_FILES:
                st.warning(f"Only the first {MAX_UPLOAD_FILES} will be imported.")
                uploaded_photos = uploaded_photos[:MAX_UPLOAD_FILES]
            existing = set(st.session_state.df["frame_name"].tolist()) \
                if not st.session_state.df.empty else set()
            new_count = 0; last_fid = None
            for uf in uploaded_photos:
                base_name = os.path.splitext(uf.name)[0]
                if base_name in existing: continue
                fid   = f"frame_{int(time.time()*1000)}_{new_count}"
                ipath = save_frame_image(fid, uf)
                st.session_state.df = upsert_frame(
                    st.session_state.df, fid, base_name,
                    {c: True for c in ALL_COORDS}, ipath)
                existing.add(base_name); last_fid = fid; new_count += 1
            if new_count:
                save_data(st.session_state.df)
                st.success(f"Imported {new_count} new frame(s).")
                if last_fid: load_frame(last_fid)
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
            cur_idx     = (frame_ids.index(st.session_state.active_frame_id)
                           if st.session_state.active_frame_id in frame_ids else 0)
            sel_idx = st.selectbox(
                "Frame", range(len(frame_names)),
                format_func=lambda i: frame_names[i],
                index=cur_idx, label_visibility="collapsed")

            col_load, col_del = st.columns(2)
            if col_load.button("Load", use_container_width=True):
                load_frame(frame_ids[sel_idx]); st.rerun()
            if col_del.button("🗑 Delete", use_container_width=True):
                fid = frame_ids[sel_idx]
                row = st.session_state.df[st.session_state.df["frame_id"] == fid]
                if not row.empty:
                    ip = str(row.iloc[0].get("image_path", ""))
                    if ip and os.path.exists(ip): os.remove(ip)
                st.session_state.df = st.session_state.df[
                    st.session_state.df["frame_id"] != fid].reset_index(drop=True)
                save_data(st.session_state.df)
                if st.session_state.active_frame_id == fid:
                    st.session_state.active_frame_id = None
                    st.session_state.frame_image     = None
                    st.session_state.unsaved         = False
                st.rerun()
            st.caption(f"Total frames: **{len(df)}**")

        st.divider()
        st.subheader("📥 Export")
        df = st.session_state.df
        if st.button("⬇️ Download Heatmap (.xlsx)", use_container_width=True, disabled=df.empty):
            xlsx_bytes = build_excel(df)
            st.download_button(
                "💾 Save .xlsx", data=xlsx_bytes,
                file_name=f"mold_heatmap_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

        if st.button("⬇️ Download Report (.pdf)", use_container_width=True, disabled=df.empty):
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
                if coord not in restored.columns: restored[coord] = "1"
            if "image_path" not in restored.columns: restored["image_path"] = ""
            st.session_state.df = restored
            save_data(restored)
            st.success("Restored!"); st.rerun()

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
            miss = sum(1 for c in ALL_COORDS if str(rec.get(c,"1")) != "1")
            summary.append({
                "Frame":     rec["frame_name"],
                "Photo":     "✅" if str(rec.get("image_path","")) else "—",
                "Missing":   miss,
                "Present":   N_COORDS - miss,
                "% Present": f"{(N_COORDS-miss)/N_COORDS*100:.1f}%",
                "Updated":   rec.get("timestamp",""),
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
    active_id  = st.session_state.active_frame_id
    active_row = df[df["frame_id"] == active_id].iloc[0]
    frame_name = active_row["frame_name"]
    coord_dict = st.session_state.coord_dict
    photo      = st.session_state.frame_image

    frame_ids  = df["frame_id"].tolist()
    cur_idx    = frame_ids.index(active_id) if active_id in frame_ids else 0
    total_fr   = len(frame_ids)

    missing_list  = sorted(c for c, v in coord_dict.items() if not v)
    present_count = N_COORDS - len(missing_list)

    # ── Header ─────────────────────────────────────────────────────────────────
    h_left, h_right = st.columns([3, 1])
    with h_left:
        st.markdown(f"## 🍫 {frame_name}")
        saved_ts = active_row.get("timestamp", "—")
        unsaved_indicator = " · **⚠️ unsaved canvas changes**" if st.session_state.unsaved else ""
        st.caption(f"Frame {cur_idx+1} of {total_fr}  |  Last saved: {saved_ts}{unsaved_indicator}")
    with h_right:
        n1, n2 = st.columns(2)
        if n1.button("◀ Prev", use_container_width=True, disabled=(total_fr<=1)):
            navigate(-1); st.rerun()
        if n2.button("Next ▶", use_container_width=True,
                     disabled=(total_fr<=1), type="primary"):
            navigate(+1); st.rerun()

    # Stats
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

    # Photo upload expander
    with st.expander(
            "📷 " + ("Replace frame photo" if photo else "Attach a photo to this frame"),
            expanded=(photo is None)):
        per_frame_upload = st.file_uploader(
            "Upload mold photo",
            type=["png","jpg","jpeg","bmp","tiff","webp"],
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
        st.info("📷 Attach a photo above to enable click-to-toggle inspection.")
        return

    # ── Bulk-action toolbar (auto-save immediately) ────────────────────────────
    t1, t2, t3, t4 = st.columns([1, 1, 1, 3])
    with t1:
        if st.button("✅ All Present", use_container_width=True):
            for c in ALL_COORDS: st.session_state.coord_dict[c] = True
            auto_save(active_id, frame_name)
            st.session_state.unsaved = False
            st.rerun()
    with t2:
        if st.button("❌ All Missing", use_container_width=True):
            for c in ALL_COORDS: st.session_state.coord_dict[c] = False
            auto_save(active_id, frame_name)
            st.session_state.unsaved = False
            st.rerun()
    with t3:
        if st.button("🔄 Invert", use_container_width=True):
            for c in ALL_COORDS:
                st.session_state.coord_dict[c] = not st.session_state.coord_dict[c]
            auto_save(active_id, frame_name)
            st.session_state.unsaved = False
            st.rerun()
    with t4:
        if missing_list:
            badges = " ".join(
                f'<span class="badge badge-red">{c}</span>' for c in missing_list)
            st.markdown(
                f'<div style="line-height:2;padding-top:2px">'
                f'<strong style="color:#E74C3C">Missing ({len(missing_list)}):</strong> '
                f'{badges}</div>', unsafe_allow_html=True)
        else:
            st.success("All 120 positions present!")

    opacity = st.slider("Overlay opacity", 0, 220, 20, 5, key="opacity")

    # ── Two-column layout: canvas left, save panel right ──────────────────────
    canvas_col, panel_col = st.columns([3, 1])

    # We need the iframe's component key to stay stable across reruns so the
    # canvas doesn't reload the image unnecessarily.  We only include the
    # coord_dict in the HTML so the saved state is always shown on first load.
    with canvas_col:
        st.markdown(
            "**Click cells** to toggle present ↔ missing.  "
            "🟡 = unsaved toggle.  "
            "Hit **💾 Save Changes** on the right when done.",
            help="Green=present, Red=missing, Yellow border=pending unsaved change.")

        W_photo, H_photo = photo.size
        display_w = 820
        canvas_html = build_canvas_html(photo, coord_dict, opacity=opacity,
                                        display_w=display_w)
        iframe_h = int(display_w * H_photo / W_photo) + 24
        components.html(canvas_html, height=iframe_h, scrolling=False)

    # ── Right panel ────────────────────────────────────────────────────────────
    with panel_col:
        panel_cls = "save-panel-dirty" if st.session_state.unsaved else "save-panel"

        # The Save Changes button posts a message to the iframe which then
        # navigates the top frame with ?mold_save=<encoded-state>.
        # We render a button that, when clicked in the same Streamlit rerun,
        # injects a postMessage script into the page.
        st.markdown(f'<div class="{panel_cls}">', unsafe_allow_html=True)

        if st.session_state.unsaved:
            st.markdown("### ⚠️ Unsaved Changes")
            st.markdown("You have pending toggles on the canvas not yet written to disk.")
        else:
            st.markdown("### ✅ All Saved")
            st.markdown("Canvas state matches the last saved record.")

        # The save button triggers a JS postMessage to the iframe which then
        # causes the iframe to write ?mold_save=... and navigate the top frame.
        save_clicked = st.button(
            "💾 Save Changes",
            use_container_width=True,
            type="primary",
            disabled=not st.session_state.unsaved,
            key="canvas_save_btn",
        )
        if save_clicked:
            # Inject a script that posts 'mold_commit' to the iframe.
            # The iframe receives it, encodes pending state, and navigates top.
            components.html("""
<script>
(function() {
  // Find the mold canvas iframe and post commit message
  const frames = window.parent.document.querySelectorAll('iframe');
  frames.forEach(f => {
    try { f.contentWindow.postMessage({type:'mold_commit'}, '*'); } catch(e) {}
  });
})();
</script>""", height=0)
            # Mark unsaved so the caption shows correctly until the rerun completes
            st.session_state.unsaved = True

        st.markdown("</div>", unsafe_allow_html=True)

        st.divider()

        # Quick-entry (auto-saves immediately, no canvas involved)
        with st.expander("⌨️ Quick-entry", expanded=False):
            st.caption("Comma-separated missing coords. Saves immediately.")
            with st.form("quick_entry"):
                raw = st.text_area(
                    "Missing coords",
                    value=", ".join(missing_list),
                    height=80, placeholder="e.g. A1, B3, G5, O8",
                    label_visibility="collapsed")
                fa, fb = st.columns(2)
                apply_btn = fa.form_submit_button("Replace", type="primary",
                                                   use_container_width=True)
                add_btn   = fb.form_submit_button("Add", use_container_width=True)
            if apply_btn or add_btn:
                tokens = re.split(r"[\s,;]+", raw.strip().upper())
                valid, invalid = [], []
                for t in tokens:
                    if not t: continue
                    (valid if t in COORD_SET else invalid).append(t)
                if invalid:
                    st.error(f"Invalid: {', '.join(invalid)}")
                else:
                    if apply_btn:
                        for c in ALL_COORDS:
                            st.session_state.coord_dict[c] = (c not in valid)
                    else:
                        for c in valid:
                            st.session_state.coord_dict[c] = False
                    auto_save(active_id, frame_name)
                    st.session_state.unsaved = False
                    st.rerun()

        st.divider()

        # Download overlay PNG (server-rendered from last saved state)
        ov_img = render_overlay(photo, coord_dict, opacity=max(opacity, 40))
        buf_dl = io.BytesIO(); ov_img.save(buf_dl, format="PNG")
        st.download_button(
            "⬇️ Download overlay (.png)",
            data=buf_dl.getvalue(),
            file_name=f"{re.sub(r'[^a-zA-Z0-9_-]','_',frame_name)}_overlay.png",
            mime="image/png", use_container_width=True)

        st.divider()

        # Per-row breakdown of saved missing coords
        if missing_list:
            st.markdown(f"**🔴 Missing ({len(missing_list)})**")
            rows_miss: dict = {}
            for c in missing_list:
                rows_miss.setdefault(c[1:], []).append(c)
            for rn, cs in sorted(rows_miss.items(), key=lambda x: int(x[0])):
                st.markdown(
                    f"**Row {rn}:** " +
                    " ".join(f'<span class="badge badge-red">{c}</span>' for c in cs),
                    unsafe_allow_html=True)
        else:
            st.success("🎉 All 120 present!")

    # ── Bottom navigation ──────────────────────────────────────────────────────
    st.divider()
    bot_l, bot_m, bot_r = st.columns([1, 2, 1])
    with bot_l:
        if st.button("◄ Previous", use_container_width=True, disabled=(total_fr<=1)):
            navigate(-1); st.rerun()
    with bot_m:
        st.markdown(
            f"<div style='text-align:center;padding-top:8px'>"
            f"Frame <b>{cur_idx+1}</b> of <b>{total_fr}</b></div>",
            unsafe_allow_html=True)
    with bot_r:
        if st.button("Next ►", use_container_width=True,
                     disabled=(total_fr<=1), type="primary"):
            navigate(+1); st.rerun()


if __name__ == "__main__":
    main()
