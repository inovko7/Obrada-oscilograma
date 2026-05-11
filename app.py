import streamlit as st
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import tempfile
import os
import io
import zipfile
from pathlib import Path

# Import core logic (all the heavy lifting stays exactly as-is)
from oscilogram_core import (
    read_measurement_txt,
    load_excel_metadata,
    make_oscillogram_png,
    insert_images_into_template_strict,
    export_strict_template_multipart,
    parse_float_any,
    OSCGEN_PRO_PASSWORD,
)

# ─────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Oscilogram Generator",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# Custom CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

h1, h2, h3 {
    font-family: 'IBM Plex Mono', monospace !important;
    letter-spacing: -0.03em;
}

.stApp {
    background: #0f0f0f;
    color: #e8e8e8;
}

section[data-testid="stSidebar"] {
    background: #161616 !important;
    border-right: 1px solid #2a2a2a;
}

.stButton > button {
    background: #e8ff00 !important;
    color: #0f0f0f !important;
    border: none !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-weight: 600 !important;
    letter-spacing: 0.05em !important;
    border-radius: 2px !important;
    padding: 0.5rem 1.5rem !important;
    transition: all 0.15s ease !important;
}

.stButton > button:hover {
    background: #ffffff !important;
    transform: translateY(-1px);
}

.stDownloadButton > button {
    background: #00ff88 !important;
    color: #0f0f0f !important;
    border: none !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-weight: 600 !important;
    border-radius: 2px !important;
}

div[data-testid="stExpander"] {
    border: 1px solid #2a2a2a !important;
    border-radius: 4px !important;
    background: #161616 !important;
}

.stSelectbox label, .stNumberInput label, .stTextInput label, .stSlider label {
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.8rem !important;
    color: #888 !important;
    letter-spacing: 0.05em !important;
    text-transform: uppercase !important;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 0px;
    background: #161616;
    border-bottom: 1px solid #2a2a2a;
}

.stTabs [data-baseweb="tab"] {
    font-family: 'IBM Plex Mono', monospace;
    letter-spacing: 0.05em;
    font-size: 0.85rem;
    color: #666;
    padding: 0.75rem 1.5rem;
}

.stTabs [aria-selected="true"] {
    color: #e8ff00 !important;
    border-bottom: 2px solid #e8ff00 !important;
    background: transparent !important;
}

.stFileUploader {
    border: 1px dashed #2a2a2a !important;
    border-radius: 4px !important;
    background: #161616 !important;
}

.stAlert {
    border-radius: 2px !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.85rem !important;
}

.metric-card {
    background: #161616;
    border: 1px solid #2a2a2a;
    border-radius: 4px;
    padding: 1rem 1.25rem;
    font-family: 'IBM Plex Mono', monospace;
}

.metric-value {
    font-size: 1.8rem;
    font-weight: 600;
    color: #e8ff00;
}

.metric-label {
    font-size: 0.7rem;
    color: #666;
    text-transform: uppercase;
    letter-spacing: 0.1em;
}

hr {
    border-color: #2a2a2a !important;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Session state init
# ─────────────────────────────────────────────
def init_state():
    defaults = {
        "pro_unlocked": False,
        "meta_df": None,
        "serial_number": "",
        "items": [],
        "settings": [],
        "preview_index": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────
TEXTBOX_MODES = [
    "Auto (from Excel)",
    "Default",
    "LI full (Upk,T1,T2)",
    "LI full (Ipk)",
    "LI full (Ipk max/min)",
    "LI tailchopped (Upk min/max, T1, T2--, Tc)",
    "LI tailchopped (Ipk)",
    "SI IEC 60060 (Upk,Tp,T2)",
    "SI IEC 60060 (Ipk,Tp,T2)",
    "SI IEC 60076 (Upk,T1,Td,T0)",
    "SI IEC 60076 (Ipk max/min)",
]

def match_txt_to_meta(meta_df, txt_files_map):
    """Match uploaded TXT files to Excel rows by name pattern v<No>.txt / c<No>.txt."""
    items = []
    for i, row in meta_df.iterrows():
        no = str(row.get("No", "")).strip()
        if not no:
            continue
        no_clean = re.sub(r"\.0$", "", no) if "." in str(no) else no
        meta = row.to_dict()

        for prefix in ["v", "c"]:
            expected = f"{prefix}{no_clean}.txt"
            matched_key = None
            for fname in txt_files_map:
                if fname.lower() == expected.lower():
                    matched_key = fname
                    break
            if matched_key:
                items.append({
                    "excel_index": i,
                    "r": i,
                    "No": no_clean,
                    "prefix": prefix,
                    "expected_name": expected,
                    "txt_key": matched_key,
                    "meta": meta,
                })
    return items

import re

def render_preview(item, settings, sidebar_params):
    """Render a preview figure inline (lower DPI for speed)."""
    txt_bytes = item["_txt_bytes"]
    with tempfile.NamedTemporaryFile(suffix=".txt", delete=False, mode="wb") as tf:
        tf.write(txt_bytes)
        tmp_txt = tf.name
    tmp_png = tmp_txt.replace(".txt", ".png")
    try:
        make_oscillogram_png(
            txt_path=tmp_txt,
            out_png=tmp_png,
            serial_number=sidebar_params["serial_number"],
            divider_str=sidebar_params["divider1"] if item.get("prefix") == "v" else sidebar_params["divider2"],
            meta=item["meta"],
            file_prefix=item.get("prefix"),
            divider1_str=sidebar_params["divider1"],
            divider2_str=sidebar_params["divider2"],
            show_matlab_header=True,
            textbox_mode_c=sidebar_params["textbox_mode_c"],
            xmax_us=float(settings["xmax_us"]),
            y_mode=settings["y_mode"],
            y_step_kv=settings.get("y_step_kv"),
            grid_scale=sidebar_params["grid_scale"],
            grid_major_lw=sidebar_params["grid_major_lw"],
            grid_minor_lw=sidebar_params["grid_minor_lw"],
            plot_frame_lw=sidebar_params["plot_frame_lw"],
            textbox_frame_lw=sidebar_params["textbox_frame_lw"],
            out_width_cm=sidebar_params["width_cm"],
            out_height_cm=sidebar_params["height_cm"],
            out_dpi=150,  # low DPI for fast preview
            font_label=sidebar_params["font_label"],
            font_tick=sidebar_params["font_tick"],
            font_header=sidebar_params["font_header"],
            font_textbox=sidebar_params["font_textbox"],
            tick_major_lw=sidebar_params["tick_major_lw"],
            tick_minor_lw=sidebar_params["tick_minor_lw"],
            textbox_pos_mode=sidebar_params["textbox_pos_mode"],
            margin_left=sidebar_params["margin_left"],
            margin_right=sidebar_params["margin_right"],
            margin_bottom=sidebar_params["margin_bottom"],
            margin_top=sidebar_params["margin_top"],
        )
        with open(tmp_png, "rb") as f:
            return f.read()
    finally:
        try:
            os.unlink(tmp_txt)
        except Exception:
            pass
        try:
            os.unlink(tmp_png)
        except Exception:
            pass


# ─────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────
st.markdown("""
<div style="padding: 2rem 0 1rem 0; border-bottom: 1px solid #2a2a2a; margin-bottom: 2rem;">
    <span style="font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem; color: #e8ff00; letter-spacing: 0.2em; text-transform: uppercase;">⚡ High Voltage Lab</span>
    <h1 style="margin: 0.25rem 0 0 0; font-size: 2rem; color: #ffffff;">Oscilogram Generator</h1>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Sidebar – settings
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Settings")

    # PRO unlock
    if not st.session_state.pro_unlocked:
        st.markdown('<span style="font-family:IBM Plex Mono;font-size:0.75rem;color:#666;text-transform:uppercase;letter-spacing:0.1em">Edition: BASIC</span>', unsafe_allow_html=True)
        with st.expander("🔒 Unlock PRO"):
            pw = st.text_input("Password", type="password", key="pro_pw")
            if st.button("Unlock"):
                if pw == OSCGEN_PRO_PASSWORD:
                    st.session_state.pro_unlocked = True
                    st.rerun()
                else:
                    st.error("Pogrešna lozinka.")
    else:
        st.markdown('<span style="font-family:IBM Plex Mono;font-size:0.75rem;color:#e8ff00;text-transform:uppercase;letter-spacing:0.1em">⭐ Edition: PRO</span>', unsafe_allow_html=True)

    pro = st.session_state.pro_unlocked
    st.divider()

    st.markdown("**Dividers**")
    divider1 = st.text_input("Divider 1 (v channel)", value="1000000", disabled=not pro)
    divider2 = st.text_input("Divider 2 (c channel)", value="1000000", disabled=not pro)

    st.divider()
    st.markdown("**X axis**")
    xmax_v = st.number_input("Default X max [v] (µs)", value=100.0, min_value=1.0, disabled=not pro)
    xmax_c = st.number_input("Default X max [c] (µs)", value=100.0, min_value=1.0, disabled=not pro)

    st.divider()
    st.markdown("**Textbox template [c]**")
    textbox_mode_c = st.selectbox("Mode", TEXTBOX_MODES, disabled=not pro)

    st.divider()
    st.markdown("**Image size**")
    width_cm = st.number_input("Width (cm)", value=9.61, disabled=not pro)
    height_cm = st.number_input("Height (cm)", value=5.37, disabled=not pro)
    output_dpi = st.number_input("Output DPI", value=600, min_value=72, max_value=1200, disabled=not pro)

    st.divider()
    with st.expander("Advanced (PRO)", expanded=False):
        grid_scale = st.selectbox("Grid scale", ["Normal", "Fine", "Coarse"], disabled=not pro)
        grid_major_lw = st.number_input("Grid major lw", value=0.85, disabled=not pro)
        grid_minor_lw = st.number_input("Grid minor lw", value=0.0, disabled=not pro)
        plot_frame_lw = st.number_input("Plot frame lw", value=1.2, disabled=not pro)
        textbox_frame_lw = st.number_input("Textbox frame lw", value=1.0, disabled=not pro)
        tick_major_lw = st.number_input("Tick major lw", value=1.0, disabled=not pro)
        tick_minor_lw = st.number_input("Tick minor lw", value=0.8, disabled=not pro)
        font_label = st.number_input("Font label", value=6, disabled=not pro)
        font_tick = st.number_input("Font tick", value=6, disabled=not pro)
        font_header = st.number_input("Font header", value=6, disabled=not pro)
        font_textbox = st.number_input("Font textbox", value=6, disabled=not pro)
        margin_left = st.number_input("Margin left", value=0.12, disabled=not pro)
        margin_right = st.number_input("Margin right", value=0.93, disabled=not pro)
        margin_bottom = st.number_input("Margin bottom", value=0.22, disabled=not pro)
        margin_top = st.number_input("Margin top", value=0.88, disabled=not pro)
        textbox_pos_mode = st.selectbox("Textbox position", ["auto", "top-right", "bottom-right"], disabled=not pro)
    
    if not pro:
        grid_scale = "Normal"
        grid_major_lw = 0.85
        grid_minor_lw = 0.0
        plot_frame_lw = 1.2
        textbox_frame_lw = 1.0
        tick_major_lw = 1.0
        tick_minor_lw = 0.8
        font_label = 6
        font_tick = 6
        font_header = 6
        font_textbox = 6
        margin_left = 0.12
        margin_right = 0.93
        margin_bottom = 0.22
        margin_top = 0.88
        textbox_pos_mode = "auto"

sidebar_params = dict(
    serial_number=st.session_state.serial_number,
    divider1=divider1,
    divider2=divider2,
    textbox_mode_c=textbox_mode_c,
    grid_scale=grid_scale,
    grid_major_lw=grid_major_lw,
    grid_minor_lw=grid_minor_lw,
    plot_frame_lw=plot_frame_lw,
    textbox_frame_lw=textbox_frame_lw,
    width_cm=width_cm,
    height_cm=height_cm,
    output_dpi=int(output_dpi),
    font_label=int(font_label),
    font_tick=int(font_tick),
    font_header=int(font_header),
    font_textbox=int(font_textbox),
    tick_major_lw=tick_major_lw,
    tick_minor_lw=tick_minor_lw,
    textbox_pos_mode=textbox_pos_mode,
    margin_left=margin_left,
    margin_right=margin_right,
    margin_bottom=margin_bottom,
    margin_top=margin_top,
)


# ─────────────────────────────────────────────
# Main tabs
# ─────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["① INPUT", "② PREVIEW", "③ EXPORT"])


# ══════════════════════════════════════════════
# TAB 1 — Input
# ══════════════════════════════════════════════
with tab1:
    col_left, col_right = st.columns([1, 1], gap="large")

    with col_left:
        st.markdown("#### Metadata")
        excel_file = st.file_uploader(
            "Excel or CSV file (.xlsx / .csv)",
            type=["xlsx", "csv"],
            help="File with columns: No, Shape, Type, PkMax, Min, T1, Tp, T2, Td, Tc, T0"
        )

        serial_input = st.text_input(
            "Serial number (optional — defaults to filename)",
            value=st.session_state.serial_number
        )

        st.markdown("#### Word Template")
        template_file = st.file_uploader(
            "Template .docx",
            type=["docx"],
            help="Word document with a table — one cell per oscillogram"
        )

    with col_right:
        st.markdown("#### Measurement files")
        txt_files = st.file_uploader(
            "TXT files (v1.txt, c1.txt, v2.txt …)",
            type=["txt"],
            accept_multiple_files=True,
            help="Files must be named v<No>.txt or c<No>.txt"
        )

        if txt_files:
            st.markdown(f'<div class="metric-card"><div class="metric-value">{len(txt_files)}</div><div class="metric-label">TXT files uploaded</div></div>', unsafe_allow_html=True)

    st.divider()

    if st.button("▶ Load inputs & prepare preview"):
        if not excel_file:
            st.error("Nedostaje metadata file (Excel/CSV).")
        elif not txt_files:
            st.error("Nedostaje TXT fajlovi.")
        else:
            with st.spinner("Učitavam…"):
                try:
                    # Save excel to temp
                    with tempfile.NamedTemporaryFile(
                        suffix=Path(excel_file.name).suffix, delete=False
                    ) as ef:
                        ef.write(excel_file.read())
                        tmp_excel = ef.name

                    meta_df = load_excel_metadata(tmp_excel)
                    os.unlink(tmp_excel)

                    sn = serial_input.strip() or Path(excel_file.name).stem
                    st.session_state.serial_number = sn
                    sidebar_params["serial_number"] = sn

                    # Map TXT files
                    txt_map = {f.name: f.read() for f in txt_files}

                    items = match_txt_to_meta(meta_df, txt_map)
                    if not items:
                        st.error("Nema podudaranja između Excel redova i TXT fajlova.\nProvjeri imenovanje: v1.txt, c1.txt, v2.txt …")
                    else:
                        # Attach bytes to items
                        for it in items:
                            it["_txt_bytes"] = txt_map[it["txt_key"]]

                        # Init per-item settings
                        settings = []
                        for it in items:
                            pfx = it.get("prefix", "v")
                            settings.append({
                                "xmax_us": xmax_c if pfx == "c" else xmax_v,
                                "y_mode": "auto",
                                "y_step_kv": None,
                                "scale_factor": 1.0,
                            })

                        # Store template bytes
                        if template_file:
                            st.session_state["template_bytes"] = template_file.read()
                            st.session_state["template_name"] = template_file.name
                        else:
                            st.session_state.pop("template_bytes", None)

                        st.session_state.meta_df = meta_df
                        st.session_state.items = items
                        st.session_state.settings = settings
                        st.session_state.preview_index = 0

                        st.success(f"✓ Učitano {len(items)} oscilogram(a) — serial: {sn}")

                        # Metrics
                        v_count = sum(1 for it in items if it.get("prefix") == "v")
                        c_count = sum(1 for it in items if it.get("prefix") == "c")
                        m1, m2, m3 = st.columns(3)
                        with m1:
                            st.markdown(f'<div class="metric-card"><div class="metric-value">{len(items)}</div><div class="metric-label">Ukupno</div></div>', unsafe_allow_html=True)
                        with m2:
                            st.markdown(f'<div class="metric-card"><div class="metric-value">{v_count}</div><div class="metric-label">Naponski (v)</div></div>', unsafe_allow_html=True)
                        with m3:
                            st.markdown(f'<div class="metric-card"><div class="metric-value">{c_count}</div><div class="metric-label">Strujni (c)</div></div>', unsafe_allow_html=True)

                except Exception as e:
                    st.error(f"Greška pri učitavanju: {e}")


# ══════════════════════════════════════════════
# TAB 2 — Preview
# ══════════════════════════════════════════════
with tab2:
    items = st.session_state.items
    settings = st.session_state.settings

    if not items:
        st.info("Učitaj fajlove na tabu ① pa se vrati ovdje.")
    else:
        col_list, col_main = st.columns([1, 3], gap="large")

        with col_list:
            st.markdown("#### Oscilogrami")
            labels = []
            for it in items:
                meta = it.get("meta", {}) or {}
                no = meta.get("No", it.get("No", ""))
                shape = str(meta.get("Shape", ""))
                pfx = it.get("prefix", "?")
                labels.append(f"{pfx}{no} — {shape[:20]}")

            sel_idx = st.radio(
                "Odabir",
                options=range(len(labels)),
                format_func=lambda i: labels[i],
                index=st.session_state.preview_index,
                label_visibility="collapsed",
            )
            st.session_state.preview_index = sel_idx

        with col_main:
            it = items[sel_idx]
            cur_settings = settings[sel_idx]
            meta = it.get("meta", {}) or {}

            st.markdown(f"#### {labels[sel_idx]}")

            # Per-oscillogram controls
            cc1, cc2, cc3, cc4 = col_main.columns(4)
            with cc1:
                xmax_val = st.number_input(
                    "X max (µs)",
                    value=float(cur_settings["xmax_us"]),
                    min_value=1.0,
                    key=f"xmax_{sel_idx}"
                )
            with cc2:
                y_mode = st.selectbox(
                    "Y osa",
                    ["auto", "manual"],
                    index=0 if cur_settings["y_mode"] == "auto" else 1,
                    key=f"ymode_{sel_idx}"
                )
            with cc3:
                y_step = st.number_input(
                    "Y step (kV)",
                    value=float(cur_settings.get("y_step_kv") or 50.0),
                    min_value=0.1,
                    key=f"ystep_{sel_idx}",
                    disabled=(y_mode == "auto")
                )
            with cc4:
                scale = st.number_input(
                    "Scale",
                    value=float(cur_settings.get("scale_factor", 1.0)),
                    key=f"scale_{sel_idx}"
                )

            # Update settings
            settings[sel_idx]["xmax_us"] = xmax_val
            settings[sel_idx]["y_mode"] = y_mode
            settings[sel_idx]["y_step_kv"] = y_step if y_mode == "manual" else None
            settings[sel_idx]["scale_factor"] = scale
            st.session_state.settings = settings

            sidebar_params["serial_number"] = st.session_state.serial_number

            # Apply to all button
            if st.button("Apply X max to ALL"):
                for s in st.session_state.settings:
                    s["xmax_us"] = xmax_val
                st.rerun()

            # Render preview
            with st.spinner("Renderiranje…"):
                try:
                    png_bytes = render_preview(it, settings[sel_idx], sidebar_params)
                    st.image(png_bytes, use_container_width=True)
                except Exception as e:
                    st.error(f"Preview greška: {e}")


# ══════════════════════════════════════════════
# TAB 3 — Export
# ══════════════════════════════════════════════
with tab3:
    items = st.session_state.items
    settings = st.session_state.settings

    if not items:
        st.info("Učitaj fajlove na tabu ① pa se vrati ovdje.")
    else:
        serial = st.session_state.serial_number
        template_bytes = st.session_state.get("template_bytes")

        col_exp, col_info = st.columns([2, 1], gap="large")

        with col_info:
            st.markdown("#### Summary")
            st.markdown(f"""
<div class="metric-card" style="margin-bottom:1rem">
  <div class="metric-value">{len(items)}</div>
  <div class="metric-label">Oscilogrami</div>
</div>
<div class="metric-card" style="margin-bottom:1rem">
  <div class="metric-value">{serial or '—'}</div>
  <div class="metric-label">Serial number</div>
</div>
<div class="metric-card">
  <div class="metric-value">{int(sidebar_params['output_dpi'])}</div>
  <div class="metric-label">DPI</div>
</div>
""", unsafe_allow_html=True)

        with col_exp:
            st.markdown("#### Generiraj izvještaj")

            if not template_bytes:
                st.warning("⚠️ Nije uploadan Word template. Generira se ZIP sa slikama.")

            if st.button("🔨 Generiraj Word dokument / ZIP"):
                sidebar_params["serial_number"] = serial

                with st.spinner("Generiranje oscilogram slika…"):
                    tmpdir = tempfile.mkdtemp(prefix="osc_web_")
                    images = []
                    errors = []

                    progress = st.progress(0)
                    for i, it in enumerate(items):
                        try:
                            # Write TXT to temp
                            with tempfile.NamedTemporaryFile(
                                suffix=".txt", delete=False, dir=tmpdir, mode="wb"
                            ) as tf:
                                tf.write(it["_txt_bytes"])
                                tmp_txt = tf.name

                            out_png = os.path.join(tmpdir, f"osc_{i+1:03d}.png")
                            st_i = settings[i]
                            make_oscillogram_png(
                                txt_path=tmp_txt,
                                out_png=out_png,
                                serial_number=serial,
                                divider_str=sidebar_params["divider1"] if it.get("prefix") == "v" else sidebar_params["divider2"],
                                meta=it["meta"],
                                file_prefix=it.get("prefix"),
                                divider1_str=sidebar_params["divider1"],
                                divider2_str=sidebar_params["divider2"],
                                show_matlab_header=True,
                                textbox_mode_c=sidebar_params["textbox_mode_c"],
                                xmax_us=float(st_i["xmax_us"]),
                                y_mode=st_i["y_mode"],
                                y_step_kv=st_i.get("y_step_kv"),
                                grid_scale=sidebar_params["grid_scale"],
                                grid_major_lw=sidebar_params["grid_major_lw"],
                                grid_minor_lw=sidebar_params["grid_minor_lw"],
                                plot_frame_lw=sidebar_params["plot_frame_lw"],
                                textbox_frame_lw=sidebar_params["textbox_frame_lw"],
                                out_width_cm=sidebar_params["width_cm"],
                                out_height_cm=sidebar_params["height_cm"],
                                out_dpi=sidebar_params["output_dpi"],
                                font_label=sidebar_params["font_label"],
                                font_tick=sidebar_params["font_tick"],
                                font_header=sidebar_params["font_header"],
                                font_textbox=sidebar_params["font_textbox"],
                                tick_major_lw=sidebar_params["tick_major_lw"],
                                tick_minor_lw=sidebar_params["tick_minor_lw"],
                                textbox_pos_mode=sidebar_params["textbox_pos_mode"],
                                margin_left=sidebar_params["margin_left"],
                                margin_right=sidebar_params["margin_right"],
                                margin_bottom=sidebar_params["margin_bottom"],
                                margin_top=sidebar_params["margin_top"],
                                y_scale=float(st_i.get("scale_factor", 1.0)),
                            )
                            images.append(out_png)
                        except Exception as e:
                            errors.append(f"Oscilogram {i+1}: {e}")
                        finally:
                            try:
                                os.unlink(tmp_txt)
                            except Exception:
                                pass

                        progress.progress((i + 1) / len(items))

                    if errors:
                        for err in errors:
                            st.warning(err)

                    if template_bytes:
                        # Write template to temp
                        with tempfile.NamedTemporaryFile(
                            suffix=".docx", delete=False, dir=tmpdir, mode="wb"
                        ) as tpl:
                            tpl.write(template_bytes)
                            tmp_template = tpl.name

                        out_docx = os.path.join(tmpdir, f"OUTPUT_{serial}.docx")
                        try:
                            insert_images_into_template_strict(
                                tmp_template, images, out_docx,
                                width_cm=sidebar_params["width_cm"],
                                height_cm=sidebar_params["height_cm"],
                                lock_aspect=True,
                            )
                            with open(out_docx, "rb") as f:
                                docx_bytes = f.read()

                            st.success(f"✓ Generirano {len(images)} oscilogram(a)!")
                            st.download_button(
                                label=f"⬇ Preuzmi OUTPUT_{serial}.docx",
                                data=docx_bytes,
                                file_name=f"OUTPUT_{serial}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            )

                        except ValueError as e:
                            # Template too small — multipart
                            parts, capacity = export_strict_template_multipart(
                                tmp_template, images,
                                os.path.join(tmpdir, f"OUTPUT_{serial}.docx"),
                                width_cm=sidebar_params["width_cm"],
                                height_cm=sidebar_params["height_cm"],
                                lock_aspect=True,
                            )
                            st.warning(f"Template kapacitet: {capacity}. Generirano {len(parts)} dijelova.")
                            # Bundle into ZIP
                            zip_buf = io.BytesIO()
                            with zipfile.ZipFile(zip_buf, "w") as zf:
                                for p in parts:
                                    zf.write(p, arcname=Path(p).name)
                            zip_buf.seek(0)
                            st.download_button(
                                label=f"⬇ Preuzmi OUTPUT_{serial}_parts.zip",
                                data=zip_buf.read(),
                                file_name=f"OUTPUT_{serial}_parts.zip",
                                mime="application/zip",
                            )
                    else:
                        # No template — ZIP with PNGs
                        zip_buf = io.BytesIO()
                        with zipfile.ZipFile(zip_buf, "w") as zf:
                            for img_path in images:
                                zf.write(img_path, arcname=Path(img_path).name)
                        zip_buf.seek(0)
                        st.success(f"✓ Generirano {len(images)} slika!")
                        st.download_button(
                            label=f"⬇ Preuzmi oscilogrami_{serial}.zip",
                            data=zip_buf.read(),
                            file_name=f"oscilogrami_{serial}.zip",
                            mime="application/zip",
                        )
