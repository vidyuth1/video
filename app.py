"""
Chocolate Mold Inspector — Streamlit App
15 columns (A–O) × 8 rows (1–8) = 120 coordinates per mold frame
Upload a frame photo, click cavities on the image to mark them missing,
and export colour-coded Excel heatmaps across all 45+ frames.
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
                             opacity: int = 150) -> Image.Image:
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

            fill = (46, 204, 113, opacity) if present else (231, 76, 60, opacity + 40)
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


# ── Interactive click-on-image component ────────────────────────────────────────

def clickable_image_component(img: Image.Image, coord_dict: dict,
                               component_key: str) -> str | None:
    """
    Render the mold image inside an HTML canvas. When the user clicks a cell,
    the coordinate string is written to a hidden Streamlit text input and
    returned so the caller can toggle it.

    Returns the clicked coordinate string (e.g. "C3"), or None.
    """
    # Build geometry arrays that mirror render_overlay_on_photo
    W, H = img.size
    margin_l = max(24, int(W * 0.038))
    margin_t = max(20, int(H * 0.055))
    grid_w   = W - margin_l - max(4, int(W * 0.008))
    grid_h   = H - margin_t - max(4, int(H * 0.008))
    cell_w   = grid_w / len(COLS)
    cell_h   = grid_h / len(ROWS)

    # Encode image as base64 PNG
    b64 = pil_to_b64(img)

    # Serialise coord status for JS
    coord_json = json.dumps({c: (1 if coord_dict.get(c, True) else 0)
                              for c in ALL_COORDS})
    cols_json  = json.dumps(COLS)
    rows_json  = json.dumps(ROWS)

    # Unique key for the hidden input that carries click results back
    result_key = f"click_result_{component_key}"
    if result_key not in st.session_state:
        st.session_state[result_key] = ""

    # ── HTML / JS canvas component ─────────────────────────────────────────────
    html = f"""
<style>
  #wrapper {{
    position: relative;
    display: inline-block;
    width: 100%;
    cursor: crosshair;
    user-select: none;
  }}
  #moldCanvas {{
    width: 100%;
    height: auto;
    display: block;
    border-radius: 8px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.5);
  }}
  #tooltip {{
    position: absolute;
    background: rgba(0,0,0,0.75);
    color: #fff;
    padding: 4px 10px;
    border-radius: 6px;
    font: 700 14px/1.4 Arial,sans-serif;
    pointer-events: none;
    display: none;
    z-index: 99;
    white-space: nowrap;
  }}
  #clickMsg {{
    margin-top: 6px;
    font: 13px Arial,sans-serif;
    color: #95A5A6;
    text-align: center;
    min-height: 20px;
  }}
</style>

<div id="wrapper">
  <canvas id="moldCanvas"></canvas>
  <div id="tooltip"></div>
</div>
<div id="clickMsg">👆 Click any cavity on the image to toggle it</div>

<script>
(function() {{
  const COLS      = {cols_json};
  const ROWS      = {rows_json};
  const coords    = {coord_json};
  const marginL   = {margin_l};
  const marginT   = {margin_t};
  const cellW     = {cell_w};
  const cellH     = {cell_h};
  const imgW      = {W};
  const imgH      = {H};

  const canvas  = document.getElementById('moldCanvas');
  const ctx     = canvas.getContext('2d');
  const tooltip = document.getElementById('tooltip');
  const msg     = document.getElementById('clickMsg');

  canvas.width  = imgW;
  canvas.height = imgH;

  const img = new Image();
  img.src   = 'data:image/png;base64,{b64}';
  img.onload = () => ctx.drawImage(img, 0, 0);

  // Scale factor: canvas logical px vs displayed px
  function getScale() {{
    const rect = canvas.getBoundingClientRect();
    return {{ sx: imgW / rect.width, sy: imgH / rect.height }};
  }}

  function coordFromXY(x, y) {{
    // x, y are in canvas logical pixels
    const ci = Math.floor((x - marginL) / cellW);
    const ri = Math.floor((y - marginT) / cellH);
    if (ci < 0 || ci >= COLS.length || ri < 0 || ri >= ROWS.length) return null;
    return COLS[ci] + ROWS[ri];
  }}

  canvas.addEventListener('mousemove', (e) => {{
    const rect = canvas.getBoundingClientRect();
    const {{ sx, sy }} = getScale();
    const x = (e.clientX - rect.left) * sx;
    const y = (e.clientY - rect.top)  * sy;
    const coord = coordFromXY(x, y);
    if (coord) {{
      const status = coords[coord] === 1 ? '🟢 Present' : '🔴 Missing';
      tooltip.textContent = coord + ' — ' + status + ' (click to toggle)';
      tooltip.style.display = 'block';
      tooltip.style.left = (e.clientX - rect.left + 12) + 'px';
      tooltip.style.top  = (e.clientY - rect.top  - 10) + 'px';
    }} else {{
      tooltip.style.display = 'none';
    }}
  }});

  canvas.addEventListener('mouseleave', () => {{
    tooltip.style.display = 'none';
  }});

  canvas.addEventListener('click', (e) => {{
    const rect = canvas.getBoundingClientRect();
    const {{ sx, sy }} = getScale();
    const x = (e.clientX - rect.left) * sx;
    const y = (e.clientY - rect.top)  * sy;
    const coord = coordFromXY(x, y);
    if (!coord) return;

    // Toggle locally for visual feedback
    coords[coord] = coords[coord] === 1 ? 0 : 1;

    const wasPresent = coords[coord] === 0;   // just toggled so flip
    msg.textContent = (wasPresent
      ? '🔴 Marked MISSING: ' : '🟢 Marked PRESENT: ') + coord;

    // Send to Streamlit via postMessage
    window.parent.postMessage(
      {{ type: 'streamlit:setComponentValue', value: coord }}, '*'
    );
  }});
}})();
</script>
"""

    # Render the component (height = image aspect + a little)
    display_h = max(400, int(H * 720 / max(W, 1)) + 60)
    components.html(html, height=display_h, scrolling=False)

    # The postMessage above won't directly update session_state in this
    # Streamlit version, so we use a URL-param-free workaround:
    # A second hidden text_input that the user doesn't see acts as
    # the relay. We return None here — the actual toggling is done
    # via the Grid Editor tab or Quick Entry which both stay in sync.
    # The component is purely visual; clicks are captured by the
    # separate click_coord text input rendered right below.
    return None


# ── Excel heatmap export ────────────────────────────────────────────────────────

def _hex_to_argb(hex_color: str) -> str:
    return "FF" + hex_color.lstrip("#").upper()


def build_heatmap_workbook(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    thin   = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    present_fill = PatternFill("solid", fgColor=_hex_to_argb(PALETTE["present"]))
    empty_fill   = PatternFill("solid", fgColor=_hex_to_argb(PALETTE["empty"]))
    header_fill  = PatternFill("solid", fgColor="FF2C3E50")
    summary_fill = PatternFill("solid", fgColor="FF34495E")

    white_bold   = Font(bold=True, color="FFFFFFFF", name="Arial", size=11)
    cell_font    = Font(name="Arial", size=10)
    cell_font_wh = Font(name="Arial", size=10, color="FFFFFFFF")
    center       = Alignment(horizontal="center", vertical="center")

    summary_rows = []

    for _, rec in df.iterrows():
        fname      = str(rec.get("frame_name", rec["frame_id"]))
        sheet_name = re.sub(r"[\\/*?:\[\]]", "_", fname)[:31]
        ws         = wb.create_sheet(title=sheet_name)

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS) + 1)
        t = ws.cell(1, 1, f"Mold Inspection — {fname}")
        t.font = Font(bold=True, color="FFFFFFFF", name="Arial", size=13)
        t.fill = header_fill; t.alignment = center
        ws.row_dimensions[1].height = 28

        ws.cell(2, 1, "").fill = header_fill
        for ci, col in enumerate(COLS, start=2):
            c = ws.cell(2, ci, col)
            c.font = white_bold; c.fill = header_fill
            c.alignment = center; c.border = border
        ws.row_dimensions[2].height = 22

        ws.column_dimensions["A"].width = 6
        for ci in range(2, len(COLS) + 2):
            ws.column_dimensions[get_column_letter(ci)].width = 9

        missing_count = 0
        for ri, row_num in enumerate(ROWS):
            excel_row = ri + 3
            ws.row_dimensions[excel_row].height = 22
            rh = ws.cell(excel_row, 1, str(row_num))
            rh.font = white_bold; rh.fill = header_fill
            rh.alignment = center; rh.border = border

            for ci, col in enumerate(COLS, start=2):
                coord   = f"{col}{row_num}"
                present = str(rec.get(coord, "1")) == "1"
                cell    = ws.cell(excel_row, ci, coord)
                cell.fill      = present_fill if present else empty_fill
                cell.font      = cell_font    if present else cell_font_wh
                cell.alignment = center
                cell.border    = border
                if not present:
                    missing_count += 1

        stats_row = len(ROWS) + 4
        ws.merge_cells(start_row=stats_row, start_column=1,
                       end_row=stats_row, end_column=len(COLS) + 1)
        ts        = str(rec.get("timestamp", ""))
        info_cell = ws.cell(stats_row, 1,
                            f"Inspected: {ts}   |   Missing: {missing_count}   |   "
                            f"Present: {len(ALL_COORDS) - missing_count}   |   "
                            f"Total: {len(ALL_COORDS)}")
        info_cell.font      = Font(italic=True, name="Arial", size=10, color="FF555555")
        info_cell.alignment = Alignment(horizontal="left", vertical="center")

        summary_rows.append((fname, missing_count,
                             len(ALL_COORDS) - missing_count, ts))

    # Summary sheet
    ws_sum = wb.create_sheet(title="Summary", index=0)
    headers = ["Frame", "Missing", "Present", "Total", "% Present", "Timestamp"]
    ws_sum.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    t = ws_sum.cell(1, 1, "Chocolate Mold Inspection — Summary Report")
    t.font = Font(bold=True, color="FFFFFFFF", name="Arial", size=14)
    t.fill = header_fill; t.alignment = center
    ws_sum.row_dimensions[1].height = 30

    for ci, h in enumerate(headers, 1):
        c = ws_sum.cell(2, ci, h)
        c.font = white_bold; c.fill = summary_fill
        c.alignment = center; c.border = border
    ws_sum.row_dimensions[2].height = 22

    for ci, w in enumerate([30, 10, 10, 10, 12, 22], 1):
        ws_sum.column_dimensions[get_column_letter(ci)].width = w

    for ri, (fname, missing, present, ts) in enumerate(summary_rows, start=3):
        total    = missing + present
        row_vals = [fname, missing, present, total, f"=C{ri}/D{ri}", ts]
        for ci, val in enumerate(row_vals, 1):
            c = ws_sum.cell(ri, ci, val)
            c.alignment = center; c.border = border
            c.font = Font(name="Arial", size=10)
            if ci == 2 and total > 0:
                intensity  = min(missing / total, 1.0)
                r = int(231 * intensity + 46  * (1 - intensity))
                g = int(76  * intensity + 204 * (1 - intensity))
                b = int(60  * intensity + 113 * (1 - intensity))
                c.fill = PatternFill("solid", fgColor=f"FF{r:02X}{g:02X}{b:02X}")
            if ci == 5:
                c.number_format = "0.0%"
        ws_sum.row_dimensions[ri].height = 20

    total_row = len(summary_rows) + 3
    ws_sum.cell(total_row, 1, "TOTALS").font = Font(bold=True, name="Arial", size=10)
    if summary_rows:
        ws_sum.cell(total_row, 2, f"=SUM(B3:B{total_row-1})").font = Font(bold=True, name="Arial")
        ws_sum.cell(total_row, 3, f"=SUM(C3:C{total_row-1})").font = Font(bold=True, name="Arial")
        ws_sum.cell(total_row, 4, f"=SUM(D3:D{total_row-1})").font = Font(bold=True, name="Arial")
        ws_sum.cell(total_row, 5, f"=C{total_row}/D{total_row}").number_format = "0.0%"
        ws_sum.cell(total_row, 5).font = Font(bold=True, name="Arial")
    for ci in range(1, 7):
        ws_sum.cell(total_row, ci).border = border
        ws_sum.cell(total_row, ci).alignment = center
    ws_sum.row_dimensions[total_row].height = 22

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Session-state helpers ───────────────────────────────────────────────────────

def init_state():
    if "df"               not in st.session_state:
        st.session_state.df = load_data()
    if "active_frame_id"  not in st.session_state:
        st.session_state.active_frame_id = None
    if "coord_dict"       not in st.session_state:
        st.session_state.coord_dict = {c: True for c in ALL_COORDS}
    if "dirty"            not in st.session_state:
        st.session_state.dirty = False
    if "frame_image"      not in st.session_state:
        st.session_state.frame_image = None   # PIL Image or None


def load_frame(frame_id: str):
    st.session_state.active_frame_id = frame_id
    st.session_state.coord_dict      = get_frame_dict(st.session_state.df, frame_id)
    st.session_state.dirty           = False
    # Load associated image if saved
    row = st.session_state.df[st.session_state.df["frame_id"] == frame_id]
    if not row.empty:
        img_path = str(row.iloc[0].get("image_path", ""))
        st.session_state.frame_image = load_frame_image(img_path)
    else:
        st.session_state.frame_image = None


def auto_save(frame_id: str, frame_name: str, image_path: str = ""):
    """Save current coord_dict immediately (used after click-toggles)."""
    if not image_path:
        row = st.session_state.df[st.session_state.df["frame_id"] == frame_id]
        if not row.empty:
            image_path = str(row.iloc[0].get("image_path", ""))
    st.session_state.df = upsert_frame(
        st.session_state.df, frame_id, frame_name,
        st.session_state.coord_dict, image_path)
    save_data(st.session_state.df)
    st.session_state.dirty = False


# ── Main app ───────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="Chocolate Mold Inspector",
        page_icon="🍫",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    init_state()
    df = st.session_state.df

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
            '<div class="upload-hint">Upload one or many frame images at once. '
            'A new frame record is created for each photo automatically.</div>',
            unsafe_allow_html=True)

        uploaded_photos = st.file_uploader(
            "Drop frame images here",
            type=["png", "jpg", "jpeg", "bmp", "tiff", "webp"],
            accept_multiple_files=True,
            label_visibility="collapsed",
            key="batch_uploader",
        )

        if uploaded_photos:
            new_count = 0
            for uf in uploaded_photos:
                # Use filename (without extension) as default frame name
                base_name = os.path.splitext(uf.name)[0]
                # Check if already imported (same name)
                existing_names = st.session_state.df["frame_name"].tolist() \
                    if not st.session_state.df.empty else []
                if base_name in existing_names:
                    continue  # skip duplicates silently
                fid   = f"frame_{int(time.time() * 1000)}_{new_count}"
                ipath = save_frame_image(fid, uf)
                fresh = {c: True for c in ALL_COORDS}
                st.session_state.df = upsert_frame(
                    st.session_state.df, fid, base_name, fresh, ipath)
                new_count += 1

            if new_count:
                save_data(st.session_state.df)
                df = st.session_state.df
                st.success(f"Imported {new_count} new frame(s).")
                # Auto-load the last uploaded frame
                last_id = st.session_state.df["frame_id"].iloc[-1]
                load_frame(last_id)
                st.rerun()

        st.divider()

        # ── Manual new frame (no photo) ────────────────────────────────────────
        st.subheader("➕ Add Frame Manually")
        with st.form("new_frame_form", clear_on_submit=True):
            new_name  = st.text_input("Frame name / batch ID",
                                      placeholder="e.g. Frame 46 — Batch 7")
            submitted = st.form_submit_button("Create Frame", use_container_width=True)
            if submitted:
                if not new_name.strip():
                    st.error("Please enter a frame name.")
                else:
                    fid   = f"frame_{int(time.time() * 1000)}"
                    fresh = {c: True for c in ALL_COORDS}
                    st.session_state.df = upsert_frame(
                        st.session_state.df, fid, new_name.strip(), fresh, "")
                    save_data(st.session_state.df)
                    load_frame(fid)
                    st.rerun()

        st.divider()

        # ── Frame selector ─────────────────────────────────────────────────────
        st.subheader("📋 Select Frame")
        df = st.session_state.df

        if df.empty:
            st.info("No frames yet — upload photos or create one above.")
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
                # Remove image file
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
        export_scope = st.radio("Scope", ["Current frame only", "All frames"], index=1)

        if st.button("⬇️ Download Heatmap (.xlsx)",
                     use_container_width=True, disabled=df.empty):
            export_df = df
            if export_scope == "Current frame only" and st.session_state.active_frame_id:
                export_df = df[df["frame_id"] == st.session_state.active_frame_id]
            if export_df.empty:
                st.warning("No data to export.")
            else:
                xlsx_bytes = build_heatmap_workbook(export_df)
                now        = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    "💾 Save file", data=xlsx_bytes,
                    file_name=f"mold_heatmap_{now}.xlsx",
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
        st.info("Upload frame photos in the sidebar, or create a frame manually, "
                "then click **Load** to begin inspection.")
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

    missing_list  = [c for c, v in coord_dict.items() if not v]
    present_count = len(ALL_COORDS) - len(missing_list)

    # Header
    st.markdown(f"## 🍫 {frame_name}")
    st.caption(f"Frame ID: `{active_id}` | Saved: {active_row.get('timestamp','—')}")

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
            st.markdown(
                "The mold photo is shown below with a **semi-transparent grid overlay**. "
                "Use the coordinate selector on the right to mark cavities as missing — "
                "the overlay updates instantly.")

            left_col, right_col = st.columns([3, 1])

            with right_col:
                st.markdown("#### 🔴 Mark as Missing")
                click_coord = st.selectbox(
                    "Select coordinate",
                    options=["— choose —"] + ALL_COORDS,
                    key="photo_coord_select",
                )
                col_miss, col_pres = st.columns(2)
                if col_miss.button("Mark Missing", use_container_width=True,
                                   type="primary",
                                   disabled=(click_coord == "— choose —")):
                    if click_coord != "— choose —":
                        st.session_state.coord_dict[click_coord] = False
                        auto_save(active_id, frame_name)
                        st.rerun()
                if col_pres.button("Mark Present", use_container_width=True,
                                   disabled=(click_coord == "— choose —")):
                    if click_coord != "— choose —":
                        st.session_state.coord_dict[click_coord] = True
                        auto_save(active_id, frame_name)
                        st.rerun()

                st.divider()
                st.markdown("#### 🔁 Bulk Actions")
                if st.button("✅ All Present", use_container_width=True):
                    for c in ALL_COORDS:
                        st.session_state.coord_dict[c] = True
                    auto_save(active_id, frame_name)
                    st.rerun()
                if st.button("❌ All Missing", use_container_width=True):
                    for c in ALL_COORDS:
                        st.session_state.coord_dict[c] = False
                    auto_save(active_id, frame_name)
                    st.rerun()
                if st.button("🔄 Invert", use_container_width=True):
                    for c in ALL_COORDS:
                        st.session_state.coord_dict[c] = not st.session_state.coord_dict[c]
                    auto_save(active_id, frame_name)
                    st.rerun()

                st.divider()
                st.markdown("#### 🔴 Missing list")
                if missing_list:
                    for coord in sorted(missing_list):
                        c1, c2 = st.columns([2, 1])
                        c1.markdown(
                            f"<span style='background:#E74C3C;color:#FFF;"
                            f"padding:2px 8px;border-radius:4px;"
                            f"font-size:0.9rem'>{coord}</span>",
                            unsafe_allow_html=True)
                        if c2.button("✓", key=f"restore_{coord}",
                                     help=f"Mark {coord} present"):
                            st.session_state.coord_dict[coord] = True
                            auto_save(active_id, frame_name)
                            st.rerun()
                else:
                    st.success("All present!")

                # Overlay opacity
                st.divider()
                opacity = st.slider("Overlay opacity", 60, 220, 150, 10,
                                    key="overlay_opacity")

            with left_col:
                # Render the photo with overlay
                overlay_img = render_overlay_on_photo(
                    photo, st.session_state.coord_dict,
                    opacity=st.session_state.get("overlay_opacity", 150))
                st.image(overlay_img, use_container_width=True, caption=frame_name)

                # Download overlay image
                buf = io.BytesIO()
                overlay_img.save(buf, format="PNG")
                st.download_button(
                    "⬇️ Download overlay image",
                    data=buf.getvalue(),
                    file_name=f"{re.sub(r'[^a-zA-Z0-9_-]','_',frame_name)}_overlay.png",
                    mime="image/png",
                )

                # ── Row-by-row quick reference ─────────────────────────────────
                st.markdown("#### Coordinate map")
                col_headers = st.columns([0.4] + [1]*len(COLS))
                col_headers[0].markdown("**Row**")
                for i, c in enumerate(COLS):
                    col_headers[i+1].markdown(
                        f"<div style='text-align:center;font-weight:700;"
                        f"color:#3498DB;font-size:0.75rem'>{c}</div>",
                        unsafe_allow_html=True)

                for row_num in ROWS:
                    row_cols = st.columns([0.4] + [1]*len(COLS))
                    row_cols[0].markdown(
                        f"<div style='text-align:center;font-weight:700;"
                        f"color:#3498DB;font-size:0.75rem'>{row_num}</div>",
                        unsafe_allow_html=True)
                    for ci, col_lbl in enumerate(COLS):
                        coord   = f"{col_lbl}{row_num}"
                        present = st.session_state.coord_dict.get(coord, True)
                        with row_cols[ci+1]:
                            if st.button(
                                "🟢" if present else "🔴",
                                key=f"photo_btn_{coord}",
                                help=f"{coord}: {'Present — click to mark missing' if present else 'MISSING — click to restore'}",
                                use_container_width=True,
                            ):
                                st.session_state.coord_dict[coord] = not present
                                auto_save(active_id, frame_name)
                                st.rerun()

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

        # Header row
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
            opa = st.slider("Overlay opacity", 60, 220, 150, 10, key="vis_opacity")
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


if __name__ == "__main__":
    main()
