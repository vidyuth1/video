"""
Chocolate Mold Inspector — Streamlit App
15 columns (A–O) × 8 rows (1–8) = 120 coordinates per mold frame
Supports 45+ frames with CSV-backed persistence and XLSX heatmap export.
"""

import io
import json
import os
import re
import time
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    GradientFill,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont

# ── Constants ──────────────────────────────────────────────────────────────────
COLS = list("ABCDEFGHIJKLMNO")          # 15 columns  (A–O)
ROWS = list(range(1, 9))                 # 8 rows       (1–8)
ALL_COORDS = [f"{c}{r}" for r in ROWS for c in COLS]  # A1…O8, row-major

DATA_FILE = "mold_data.csv"             # persistent store

PALETTE = {
    "present":  "#2ECC71",             # green  — chocolate present
    "empty":    "#E74C3C",             # red    — missing chocolate
    "selected": "#F39C12",             # amber  — cursor hover / selection
}


# ── Data persistence ────────────────────────────────────────────────────────────

def load_data() -> pd.DataFrame:
    """Load or initialise the master DataFrame (one row per frame)."""
    if os.path.exists(DATA_FILE):
        df = pd.read_csv(DATA_FILE, dtype=str)
        # Ensure all coordinate columns exist (handles schema upgrades)
        for coord in ALL_COORDS:
            if coord not in df.columns:
                df[coord] = "1"
        return df
    # Bootstrap empty dataframe
    df = pd.DataFrame(columns=["frame_id", "frame_name", "timestamp"] + ALL_COORDS)
    return df


def save_data(df: pd.DataFrame) -> None:
    df.to_csv(DATA_FILE, index=False)


def get_frame_dict(df: pd.DataFrame, frame_id: str) -> dict:
    """Return coord→bool dict for a frame (True = present)."""
    row = df[df["frame_id"] == frame_id]
    if row.empty:
        return {c: True for c in ALL_COORDS}
    r = row.iloc[0]
    return {c: (str(r.get(c, "1")) == "1") for c in ALL_COORDS}


def upsert_frame(df: pd.DataFrame, frame_id: str, frame_name: str,
                 coord_dict: dict) -> pd.DataFrame:
    """Insert or update a frame record."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row_data = {"frame_id": frame_id, "frame_name": frame_name, "timestamp": ts}
    row_data.update({c: ("1" if v else "0") for c, v in coord_dict.items()})
    if frame_id in df["frame_id"].values:
        for k, v in row_data.items():
            df.loc[df["frame_id"] == frame_id, k] = v
    else:
        df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
    return df


# ── Mold visualisation ──────────────────────────────────────────────────────────

def render_mold_image(coord_dict: dict,
                      cell_px: int = 72,
                      label_px: int = 36) -> Image.Image:
    """
    Draw the 15×8 mold grid as a PIL image.
    Green cells = chocolate present, Red = missing.
    """
    n_cols, n_rows = len(COLS), len(ROWS)
    img_w = label_px + n_cols * cell_px + 2
    img_h = label_px + n_rows * cell_px + 2
    img = Image.new("RGB", (img_w, img_h), "#1A1A2E")
    draw = ImageDraw.Draw(img)

    # Try to load a small font; fall back to default
    try:
        fnt_label = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 16)
        fnt_cell  = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 13)
    except Exception:
        fnt_label = ImageFont.load_default()
        fnt_cell  = fnt_label

    # Column headers (A–O)
    for ci, col in enumerate(COLS):
        x = label_px + ci * cell_px + cell_px // 2
        y = label_px // 2
        draw.text((x, y), col, fill="#ECF0F1", font=fnt_label, anchor="mm")

    # Row headers + cells
    for ri, row in enumerate(ROWS):
        y_top = label_px + ri * cell_px

        # Row number label
        draw.text((label_px // 2, y_top + cell_px // 2),
                  str(row), fill="#ECF0F1", font=fnt_label, anchor="mm")

        for ci, col in enumerate(COLS):
            coord = f"{col}{row}"
            present = coord_dict.get(coord, True)
            x_left = label_px + ci * cell_px

            fill = PALETTE["present"] if present else PALETTE["empty"]
            # Cell background
            draw.rounded_rectangle(
                [x_left + 2, y_top + 2,
                 x_left + cell_px - 2, y_top + cell_px - 2],
                radius=6,
                fill=fill,
                outline="#1A1A2E",
                width=2,
            )
            # Coord label inside cell
            label_col = "#1A1A2E" if present else "#FFFFFF"
            draw.text((x_left + cell_px // 2, y_top + cell_px // 2),
                      coord, fill=label_col, font=fnt_cell, anchor="mm")

    return img


# ── Excel heatmap export ────────────────────────────────────────────────────────

def _hex_to_argb(hex_color: str) -> str:
    """Convert #RRGGBB → FFRRGGBB (openpyxl ARGB)."""
    h = hex_color.lstrip("#")
    return "FF" + h.upper()


def build_heatmap_workbook(df: pd.DataFrame) -> bytes:
    """
    Build an xlsx workbook with:
    - One sheet per frame (named by frame_name)
    - A 'Summary' sheet counting missing chocolates per frame
    - Each mold sheet has a 15×8 heatmap (green=present, red=missing)
      plus counts for present / missing.
    """
    wb = Workbook()
    wb.remove(wb.active)  # remove default blank sheet

    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    present_fill  = PatternFill("solid", fgColor=_hex_to_argb(PALETTE["present"]))
    empty_fill    = PatternFill("solid", fgColor=_hex_to_argb(PALETTE["empty"]))
    header_fill   = PatternFill("solid", fgColor="FF2C3E50")
    summary_fill  = PatternFill("solid", fgColor="FF34495E")

    white_bold   = Font(bold=True, color="FFFFFFFF", name="Arial", size=11)
    cell_font    = Font(name="Arial", size=10)
    cell_font_wh = Font(name="Arial", size=10, color="FFFFFFFF")
    center       = Alignment(horizontal="center", vertical="center")

    # ── Per-frame sheets ──────────────────────────────────────────────────────
    summary_rows = []   # (frame_name, missing_count, present_count)

    for _, rec in df.iterrows():
        fname = str(rec.get("frame_name", rec["frame_id"]))
        # Sanitise sheet name (Excel limit: 31 chars, no special chars)
        sheet_name = re.sub(r"[\\/*?:\[\]]", "_", fname)[:31]
        ws = wb.create_sheet(title=sheet_name)

        # Title row
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS) + 1)
        title_cell = ws.cell(1, 1, f"Mold Inspection — {fname}")
        title_cell.font  = Font(bold=True, color="FFFFFFFF", name="Arial", size=13)
        title_cell.fill  = header_fill
        title_cell.alignment = center
        ws.row_dimensions[1].height = 28

        # Column header row (row 2): blank + A…O
        ws.cell(2, 1, "").fill = header_fill
        for ci, col in enumerate(COLS, start=2):
            c = ws.cell(2, ci, col)
            c.font = white_bold; c.fill = header_fill
            c.alignment = center; c.border = border
        ws.row_dimensions[2].height = 22

        # Column widths
        ws.column_dimensions["A"].width = 6
        for ci in range(2, len(COLS) + 2):
            ws.column_dimensions[get_column_letter(ci)].width = 9

        # Data rows
        missing_count = 0
        for ri, row_num in enumerate(ROWS):
            excel_row = ri + 3   # rows 3–10
            ws.row_dimensions[excel_row].height = 22

            # Row number header
            rh = ws.cell(excel_row, 1, str(row_num))
            rh.font = white_bold; rh.fill = header_fill
            rh.alignment = center; rh.border = border

            for ci, col in enumerate(COLS, start=2):
                coord = f"{col}{row_num}"
                present = str(rec.get(coord, "1")) == "1"
                cell = ws.cell(excel_row, ci, coord)
                cell.fill      = present_fill if present else empty_fill
                cell.font      = cell_font if present else cell_font_wh
                cell.alignment = center
                cell.border    = border
                if not present:
                    missing_count += 1

        # Stats rows (below the grid)
        stats_row = len(ROWS) + 4
        ws.merge_cells(start_row=stats_row, start_column=1,
                       end_row=stats_row, end_column=len(COLS) + 1)
        ts = str(rec.get("timestamp", ""))
        info_cell = ws.cell(stats_row, 1,
                            f"Inspected: {ts}   |   Missing: {missing_count}   |   "
                            f"Present: {len(ALL_COORDS) - missing_count}   |   "
                            f"Total: {len(ALL_COORDS)}")
        info_cell.font = Font(italic=True, name="Arial", size=10, color="FF555555")
        info_cell.alignment = Alignment(horizontal="left", vertical="center")

        summary_rows.append((fname, missing_count, len(ALL_COORDS) - missing_count,
                             str(rec.get("timestamp", ""))))

    # ── Summary sheet ─────────────────────────────────────────────────────────
    ws_sum = wb.create_sheet(title="Summary", index=0)

    # Header
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

    col_widths = [30, 10, 10, 10, 12, 22]
    for ci, w in enumerate(col_widths, 1):
        ws_sum.column_dimensions[get_column_letter(ci)].width = w

    for ri, (fname, missing, present, ts) in enumerate(summary_rows, start=3):
        total = missing + present
        pct   = f"=C{ri}/D{ri}"   # Excel formula for % present
        row_vals = [fname, missing, present, total, pct, ts]
        for ci, val in enumerate(row_vals, 1):
            c = ws_sum.cell(ri, ci, val)
            c.alignment = center; c.border = border
            c.font = Font(name="Arial", size=10)
            # Colour-code the missing column
            if ci == 2:
                intensity = min(missing / max(len(ALL_COORDS), 1), 1.0)
                r = int(231 * intensity + 46 * (1 - intensity))
                g = int(76  * intensity + 204 * (1 - intensity))
                b = int(60  * intensity + 113 * (1 - intensity))
                cell_color = f"FF{r:02X}{g:02X}{b:02X}"
                c.fill = PatternFill("solid", fgColor=cell_color)
            if ci == 5:   # percentage column
                c.number_format = "0.0%"
        ws_sum.row_dimensions[ri].height = 20

    # Grand totals row
    total_row = len(summary_rows) + 3
    ws_sum.cell(total_row, 1, "TOTALS").font = Font(bold=True, name="Arial", size=10)
    n = len(summary_rows)
    if n:
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
    if "df" not in st.session_state:
        st.session_state.df = load_data()
    if "active_frame_id" not in st.session_state:
        st.session_state.active_frame_id = None
    if "coord_dict" not in st.session_state:
        st.session_state.coord_dict = {c: True for c in ALL_COORDS}
    if "dirty" not in st.session_state:
        st.session_state.dirty = False


def load_frame(frame_id: str):
    st.session_state.active_frame_id = frame_id
    st.session_state.coord_dict = get_frame_dict(st.session_state.df, frame_id)
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

    # ── Custom CSS ─────────────────────────────────────────────────────────────
    st.markdown("""
    <style>
    .main { background: #0F0F1A; }
    [data-testid="stSidebar"] { background: #1A1A2E; }
    h1, h2, h3 { color: #ECF0F1; }
    .stat-box {
        background: #16213E;
        border-radius: 10px;
        padding: 16px 20px;
        text-align: center;
        border: 1px solid #2C3E50;
    }
    .stat-number { font-size: 2rem; font-weight: 700; }
    .stat-label  { font-size: 0.85rem; color: #95A5A6; margin-top: 4px; }
    .green { color: #2ECC71; }
    .red   { color: #E74C3C; }
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.2s;
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Sidebar ────────────────────────────────────────────────────────────────
    with st.sidebar:
        st.title("🍫 Mold Inspector")
        st.caption("15 × 8 grid | 120 positions")
        st.divider()

        # ── New frame form ─────────────────────────────────────────────────────
        st.subheader("➕ Add New Frame")
        with st.form("new_frame_form", clear_on_submit=True):
            new_name = st.text_input("Frame name / batch ID",
                                     placeholder="e.g. Frame 46 — Batch 7")
            submitted = st.form_submit_button("Create Frame", use_container_width=True)
            if submitted:
                if not new_name.strip():
                    st.error("Please enter a frame name.")
                else:
                    fid = f"frame_{int(time.time() * 1000)}"
                    fresh = {c: True for c in ALL_COORDS}
                    st.session_state.df = upsert_frame(
                        st.session_state.df, fid, new_name.strip(), fresh)
                    save_data(st.session_state.df)
                    load_frame(fid)
                    st.success(f"Created '{new_name.strip()}'")
                    st.rerun()

        st.divider()

        # ── Frame selector ─────────────────────────────────────────────────────
        st.subheader("📋 Select Frame")
        df = st.session_state.df

        if df.empty:
            st.info("No frames yet. Create one above.")
        else:
            frame_options = df["frame_name"].tolist()
            frame_ids     = df["frame_id"].tolist()

            current_idx = 0
            if st.session_state.active_frame_id in frame_ids:
                current_idx = frame_ids.index(st.session_state.active_frame_id)

            selected_idx = st.selectbox(
                "Frame",
                options=range(len(frame_options)),
                format_func=lambda i: frame_options[i],
                index=current_idx,
                label_visibility="collapsed",
            )

            if st.button("Load Frame", use_container_width=True):
                load_frame(frame_ids[selected_idx])
                st.rerun()

            st.caption(f"Total frames: **{len(df)}**")

        st.divider()

        # ── Export ─────────────────────────────────────────────────────────────
        st.subheader("📥 Export")
        export_scope = st.radio(
            "Export scope",
            ["Current frame only", "All frames"],
            index=1,
        )

        if st.button("⬇️ Download Heatmap (.xlsx)", use_container_width=True,
                     disabled=df.empty):
            export_df = df
            if export_scope == "Current frame only" and st.session_state.active_frame_id:
                export_df = df[df["frame_id"] == st.session_state.active_frame_id]

            if export_df.empty:
                st.warning("No data to export.")
            else:
                xlsx_bytes = build_heatmap_workbook(export_df)
                now = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    "💾 Save file",
                    data=xlsx_bytes,
                    file_name=f"mold_heatmap_{now}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        # CSV backup
        if not df.empty:
            st.download_button(
                "⬇️ Backup CSV",
                data=df.to_csv(index=False).encode(),
                file_name="mold_data_backup.csv",
                mime="text/csv",
                use_container_width=True,
            )

        # CSV restore
        st.divider()
        st.subheader("⬆️ Restore from CSV")
        uploaded = st.file_uploader("Upload backup CSV", type="csv",
                                    label_visibility="collapsed")
        if uploaded:
            restored = pd.read_csv(uploaded, dtype=str)
            for coord in ALL_COORDS:
                if coord not in restored.columns:
                    restored[coord] = "1"
            st.session_state.df = restored
            save_data(restored)
            st.success("Data restored!")
            st.rerun()

    # ── Main panel ─────────────────────────────────────────────────────────────
    if st.session_state.active_frame_id is None:
        st.markdown("## 🍫 Chocolate Mold Inspector")
        st.info("Create or select a frame in the sidebar to begin inspection.")

        if not df.empty:
            st.subheader("Overview — All Frames")
            summary = []
            for _, rec in df.iterrows():
                missing = sum(1 for c in ALL_COORDS if str(rec.get(c, "1")) != "1")
                summary.append({
                    "Frame": rec["frame_name"],
                    "Missing": missing,
                    "Present": len(ALL_COORDS) - missing,
                    "% Present": f"{(len(ALL_COORDS)-missing)/len(ALL_COORDS)*100:.1f}%",
                    "Last Updated": rec.get("timestamp", ""),
                })
            st.dataframe(pd.DataFrame(summary), use_container_width=True)
        return

    # ── Active frame editing ───────────────────────────────────────────────────
    active_id   = st.session_state.active_frame_id
    active_row  = df[df["frame_id"] == active_id].iloc[0]
    frame_name  = active_row["frame_name"]
    coord_dict  = st.session_state.coord_dict

    missing_list  = [c for c, v in coord_dict.items() if not v]
    present_count = len(ALL_COORDS) - len(missing_list)

    st.markdown(f"## 🍫 {frame_name}")
    st.caption(f"Frame ID: `{active_id}` | Saved: {active_row.get('timestamp','—')}")

    # Stats row
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

    # ── Tabs: Grid Editor  |  Quick Entry  |  Visual ──────────────────────────
    tab_grid, tab_quick, tab_vis = st.tabs([
        "🔲 Grid Editor", "⌨️ Quick Entry", "🖼️ Visual Preview"
    ])

    # ── Tab 1: Grid Editor ─────────────────────────────────────────────────────
    with tab_grid:
        st.markdown(
            "**Click a coordinate** to toggle it between "
            "<span style='color:#2ECC71'>●&nbsp;Present</span> and "
            "<span style='color:#E74C3C'>●&nbsp;Missing</span>.",
            unsafe_allow_html=True,
        )

        toolbar_c1, toolbar_c2, toolbar_c3, toolbar_c4 = st.columns([2, 2, 2, 2])
        with toolbar_c1:
            if st.button("✅ Mark All Present", use_container_width=True):
                for c in ALL_COORDS:
                    st.session_state.coord_dict[c] = True
                st.session_state.dirty = True
                st.rerun()
        with toolbar_c2:
            if st.button("❌ Mark All Missing", use_container_width=True):
                for c in ALL_COORDS:
                    st.session_state.coord_dict[c] = False
                st.session_state.dirty = True
                st.rerun()
        with toolbar_c3:
            if st.button("🔄 Invert Selection", use_container_width=True):
                for c in ALL_COORDS:
                    st.session_state.coord_dict[c] = not st.session_state.coord_dict[c]
                st.session_state.dirty = True
                st.rerun()
        with toolbar_c4:
            if st.button("💾 Save Frame", use_container_width=True,
                         type="primary", disabled=not st.session_state.dirty):
                st.session_state.df = upsert_frame(
                    st.session_state.df, active_id, frame_name,
                    st.session_state.coord_dict)
                save_data(st.session_state.df)
                st.session_state.dirty = False
                st.success("Saved!")
                st.rerun()

        # Column header row
        header_cols = st.columns([0.5] + [1] * len(COLS))
        header_cols[0].markdown("**↓ Row**")
        for i, col_lbl in enumerate(COLS):
            header_cols[i + 1].markdown(
                f"<div style='text-align:center;font-weight:700;color:#3498DB;'>{col_lbl}</div>",
                unsafe_allow_html=True)

        # Data rows
        for row_num in ROWS:
            row_cols = st.columns([0.5] + [1] * len(COLS))
            row_cols[0].markdown(
                f"<div style='text-align:center;font-weight:700;color:#3498DB;'>{row_num}</div>",
                unsafe_allow_html=True)

            for ci, col_lbl in enumerate(COLS):
                coord   = f"{col_lbl}{row_num}"
                present = st.session_state.coord_dict.get(coord, True)
                emoji   = "🟢" if present else "🔴"
                btn_key = f"btn_{coord}"

                with row_cols[ci + 1]:
                    if st.button(f"{emoji}", key=btn_key,
                                 help=f"{coord}: {'Present' if present else 'MISSING'}",
                                 use_container_width=True):
                        st.session_state.coord_dict[coord] = not present
                        st.session_state.dirty = True
                        st.rerun()

        if st.session_state.dirty:
            st.warning("⚠️ Unsaved changes — click **Save Frame** above.")

    # ── Tab 2: Quick Entry ─────────────────────────────────────────────────────
    with tab_quick:
        st.markdown("""
        Enter **missing** coordinate(s) as a comma-separated list.
        All other positions are assumed **present**.

        **Format examples:** `A1, C3, O8` or `B2 D5 F7`
        """)

        # Show currently missing
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
                height=100,
                placeholder="e.g.  A1, B3, G5, O8",
            )
            col_apply, col_add = st.columns(2)
            apply_btn = col_apply.form_submit_button(
                "Apply (replace all missing)", use_container_width=True, type="primary")
            add_btn   = col_add.form_submit_button(
                "Add to existing missing", use_container_width=True)

        if apply_btn or add_btn:
            tokens = re.split(r"[\s,;]+", raw.strip().upper())
            valid, invalid = [], []
            for t in tokens:
                if not t:
                    continue
                if t in ALL_COORDS:
                    valid.append(t)
                else:
                    invalid.append(t)

            if invalid:
                st.error(f"Invalid coordinate(s): {', '.join(invalid)}")
            else:
                if apply_btn:
                    for c in ALL_COORDS:
                        st.session_state.coord_dict[c] = (c not in valid)
                else:   # add_btn
                    for c in valid:
                        st.session_state.coord_dict[c] = False

                st.session_state.df = upsert_frame(
                    st.session_state.df, active_id, frame_name,
                    st.session_state.coord_dict)
                save_data(st.session_state.df)
                st.session_state.dirty = False
                n = len([c for c, v in st.session_state.coord_dict.items() if not v])
                st.success(f"Saved! {n} position(s) marked missing.")
                st.rerun()

        # Per-row overview table
        st.subheader("Row-by-row status")
        status_rows = []
        for row_num in ROWS:
            row_missing = [f"{col_lbl}{row_num}"
                           for col_lbl in COLS
                           if not st.session_state.coord_dict.get(f"{col_lbl}{row_num}", True)]
            status_rows.append({
                "Row": row_num,
                "Missing Count": len(row_missing),
                "Missing Coords": ", ".join(row_missing) if row_missing else "—",
            })
        st.dataframe(pd.DataFrame(status_rows), use_container_width=True, hide_index=True)

    # ── Tab 3: Visual Preview ──────────────────────────────────────────────────
    with tab_vis:
        st.markdown("Rendered mold grid. 🟢 = Present &nbsp; 🔴 = Missing")
        img = render_mold_image(st.session_state.coord_dict, cell_px=68, label_px=34)
        st.image(img, caption=frame_name, use_container_width=True)

        # Download the image
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        st.download_button(
            "⬇️ Download Grid Image (.png)",
            data=buf.getvalue(),
            file_name=f"{re.sub(r'[^a-zA-Z0-9_-]','_',frame_name)}_grid.png",
            mime="image/png",
        )

        # Missing positions panel
        if missing_list:
            st.subheader(f"🔴 Missing positions ({len(missing_list)})")
            # Display in a compact grid
            rows_of_missing = {}
            for c in sorted(missing_list):
                row_n = c[1:]
                rows_of_missing.setdefault(row_n, []).append(c)
            for row_n, coords in sorted(rows_of_missing.items(), key=lambda x: int(x[0])):
                st.markdown(
                    f"**Row {row_n}:** " +
                    " ".join(f"<span style='background:#E74C3C;color:#FFF;"
                             f"padding:2px 6px;border-radius:4px;font-size:0.85rem'>{c}</span>"
                             for c in coords),
                    unsafe_allow_html=True,
                )
        else:
            st.success("🎉 All 120 positions are present!")


if __name__ == "__main__":
    main()
