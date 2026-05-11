import os
import re
import math
import tempfile
import io
import json
import hashlib
from pathlib import Path
# import tkinter as tk
# from tkinter import filedialog, messagebox, ttk, simpledialog
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, FuncFormatter
# from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_ALIGN_VERTICAL  # optional
# =========================
# JPG / Clipboard helpers
# =========================

def get_base_dir() -> Path:
    """Folder where the script is located (or cwd fallback)."""
    try:
        return Path(__file__).resolve().parent
    except Exception:
        return Path.cwd()

def auto_find_template_docx() -> str:
    """Find template in the same folder as the .py (prefer TEMPLATE.docx)."""
    base = get_base_dir()
    preferred = base / "template.docx"
    if preferred.exists():
        return str(preferred)

    # fallback: first .docx found in the same folder
    docxs = sorted(base.glob("*.docx"))
    if docxs:
        return str(docxs[0])

    return ""

def _ensure_parent_dir(path: str) -> None:
    """Ensure the folder for a file path exists."""
    Path(path).parent.mkdir(parents=True, exist_ok=True)


def save_figure_as_jpg(fig, out_path: str, dpi: int = 300, quality: int = 95):
    """
    Save matplotlib figure as JPG.
    Requires Pillow installed.
    """

    from PIL import Image  # pip install pillow

    # Save figure temporarily as PNG into memory
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, facecolor="white")
    buf.seek(0)

    # Convert PNG -> JPG
    img = Image.open(buf).convert("RGB")

    _ensure_parent_dir(out_path)
    img.save(out_path, "JPEG", quality=int(quality))

    buf.close()
def copy_figure_to_clipboard_windows(fig, dpi: int = 300):
    """
    Copy matplotlib figure to Windows clipboard.
    Requires Pillow + pywin32.
    """

    from PIL import Image
    import win32clipboard  # pip install pywin32

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, facecolor="white")
    buf.seek(0)

    img = Image.open(buf).convert("RGB")

    # Convert to BMP (clipboard format)
    output = io.BytesIO()
    img.save(output, "BMP")
    data = output.getvalue()[14:]  # remove BMP header

    output.close()
    buf.close()

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    finally:
        win32clipboard.CloseClipboard()
        
# =========================
# Basic / Pro unlock config
# =========================
# Default password can be overridden via environment variable OSCGEN_PRO_PASSWORD
OSCGEN_PRO_PASSWORD = os.environ.get("OSCGEN_PRO_PASSWORD", "pro123")

# =========================
# Unit helpers
# =========================
def cm_to_inches(cm: float) -> float:
    return float(cm) / 2.54
def parse_float_any(s: str, default: float) -> float:
    try:
        return float(str(s).replace(",", "."))
    except Exception:
        return default
# =========================
# Data reading / utilities
# =========================
def read_measurement_txt(txt_path: str) -> pd.DataFrame:
    """Reads TXT with header line containing 'Time [s]' and 'Value'."""
    with open(txt_path, "r", errors="ignore") as f:
        lines = f.readlines()
    start_idx = None
    for i, line in enumerate(lines):
        if "Time [s]" in line and "Value" in line:
            start_idx = i + 1
            break
    if start_idx is None:
        raise ValueError(f"Cannot find data table header 'Time [s] ... Value ...' in: {txt_path}")
    data = []
    for line in lines[start_idx:]:
        line = line.strip()
        if not line:
            continue
        parts = re.split(r"\s+", line)
        if len(parts) < 2:
            continue
        try:
            t = float(parts[0])
            v = float(parts[1])
            data.append((t, v))
        except ValueError:
            continue
    if not data:
        raise ValueError(f"No numeric data found in: {txt_path}")
    return pd.DataFrame(data, columns=["time_s", "value_v"])
def convert_to_kv(value_v: pd.Series) -> pd.Series:
    return value_v / 1000.0
def fmt_3dec_decimal_comma(x) -> str:
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return ""
        return f"{float(str(x).replace(',', '.')):.3f}".replace(".", ",")
    except Exception:
        return ""
# =========================
# Missing formatting helpers (hotfix)
# =========================
def _to_float_maybe(x):
    """Parse number from Excel/string allowing decimal comma."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    try:
        if isinstance(x, str):
            xs = x.replace(",", ".")
            m = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", xs)
            if not m:
                return None
            return float(m.group(0))
        return float(x)
    except Exception:
        return None

def fmt_voltage_to_kv(x) -> str:
    """Format a voltage-like value to kV string with decimal comma (3 dec)."""
    v = _to_float_maybe(x)
    if v is None:
        return ""
    return f"{v:.3f}".replace(".", ",")

def fmt_current_with_unit(x):
    """Return (value_str, unit) for current. Uses A / kA with decimal comma."""
    v = _to_float_maybe(x)
    if v is None:
        return ("", "")
    unit = "A"
    if abs(v) >= 1000.0:
        v = v / 1000.0
        unit = "kA"
    return (f"{v:.3f}".replace(".", ","), unit)

def fmt_time_to_us(x) -> str:
    """Format a time-like value to microseconds with decimal comma.

    Supports strings with units: ns, us (µs), ms, s.
    If no unit is present, falls back to the old heuristic:
      - If |x| < 0.01 -> treat as seconds and convert to us (x*1e6)
      - Otherwise -> treat as already in microseconds
    """
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""

    xs = str(x).strip()
    low = xs.lower().replace("µ", "u")

    v = _to_float_maybe(xs)
    if v is None:
        return ""

    # Unit-aware parsing
    if "ns" in low:
        v = v / 1000.0
    elif "us" in low or re.search(r"\bus\b", low):
        v = v
    elif "ms" in low:
        v = v * 1000.0
    elif re.search(r"\bs\b", low) and ("ms" not in low and "us" not in low and "ns" not in low):
        v = v * 1e6
    else:
        # fallback heuristic (your current behavior)
        if abs(v) < 0.01 and v != 0:
            v = v * 1e6

    return f"{v:.3f}".replace(".", ",")

def format_kv_or_mv(x) -> str:
    """Auto-format voltage: kV or MV depending on magnitude."""
    v = _to_float_maybe(x)
    if v is None:
        return ""
    if abs(v) >= 1000.0:
        return f"{v/1000.0:.3f}".replace(".", ",") + " MV"
    return f"{v:.3f}".replace(".", ",") + " kV"

def format_time_us(x) -> str:
    """Format time value as microseconds with unit 'us'."""
    s = fmt_time_to_us(x)
    return (s + " us").strip() if s else ""

def format_voltage_from_excel(x) -> str:
    """Format voltage preserving the unit implied by the Excel cell.

    Accepts either numeric or strings like '535.858 kV', '394.640 V', '1.20 MV'.
    Outputs with decimal comma.

    Rules:
    - If cell text contains 'kV' -> treat number as kV; show kV (or MV if >=1000 kV).
    - If contains 'MV' -> treat as MV; show MV.
    - If contains 'V' (but not kV/MV) -> treat as V; show V (<1000 V) else kV/MV.
    - If no unit -> fall back to kV/MV formatter (legacy behavior).
    """
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    xs = str(x)
    low = xs.lower()
    v = _to_float_maybe(x)
    if v is None:
        return ""
    # Detect units in text (order matters)
    if "mv" in low:
        return f"{v:.3f}".replace(".", ",") + " MV"
    if "kv" in low:
        # value is in kV
        if abs(v) >= 1000.0:
            return f"{(v/1000.0):.3f}".replace(".", ",") + " MV"
        return f"{v:.3f}".replace(".", ",") + " kV"
    # plain volts (avoid matching kV)
    if " v" in low or low.endswith("v"):
        return format_voltage_v_kv_mv(v)
    # unknown -> assume kV legacy
    return format_kv_or_mv(v)


def format_voltage_from_excel_preserve_unit(x, blank_if_missing_unit: bool = False) -> str:
    """Format voltage exactly in the unit written in the Excel cell.

    This is used for channel 2 (c*.txt) textbox, where the unit must match Excel.

    Accepted units in the cell text: V, kV, MV (case-insensitive). The numeric part is NOT converted.
    If no unit is present:
      - if blank_if_missing_unit=True -> return only the numeric value (no unit)
      - else -> fall back to legacy auto formatter (kV/MV)
    """
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    xs = str(x).strip()
    low = xs.lower()

    v = _to_float_maybe(x)
    if v is None:
        return ""

    # Detect unit tokens (order matters: mv, kv, then plain v)
    unit = None
    if re.search(r"\bmv\b", low):
        unit = "MV"
    elif re.search(r"\bkv\b", low):
        unit = "kV"
    elif re.search(r"(?:\bv\b|v$)", low) and not re.search(r"\bkv\b|\bmv\b", low):
        unit = "V"

    num = f"{v:.3f}".replace(".", ",")

    if unit:
        return f"{num} {unit}"
    if blank_if_missing_unit:
        return num
    # fallback (legacy)
    return format_kv_or_mv(v)


def format_voltage_v_kv_mv(x) -> str:
    """Format voltage where the *input is in volts*.

    - < 1000 V   -> V
    - < 1e6 V    -> kV
    - otherwise  -> MV

    Uses decimal comma and 3 decimals (except V: 1 decimal is usually enough, but keep 1 to stay readable).
    """
    v = _to_float_maybe(x)
    if v is None:
        return ""
    av = abs(v)
    if av < 1000.0:
        return f"{v:.1f}".replace(".", ",") + " V"
    if av < 1_000_000.0:
        return f"{(v/1000.0):.3f}".replace(".", ",") + " kV"
    return f"{(v/1_000_000.0):.3f}".replace(".", ",") + " MV"

    s = fmt_time_to_us(x)
    return (s + " us").strip() if s else ""

    try:
        if isinstance(x, str):
            x = x.replace(",", ".")
            m = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", x)
            if not m:
                return ""
            x = m.group(0)
        s = f"{float(x):.3f}"
        return s.replace(".", ",")
    except Exception:
        return ""

def fmt_tick_decimal_comma(x, _pos=None) -> str:
    """Matplotlib tick formatter using decimal comma."""
    try:
        if x is None or (isinstance(x, float) and (math.isnan(x) or math.isinf(x))):
            return ""
        # Prefer integer formatting when very close to an integer
        if abs(x - round(x)) < 1e-9:
            s = str(int(round(x)))
        else:
            s = f"{x:.3f}".rstrip("0").rstrip(".")
        return s.replace(".", ",")
    except Exception:
        try:
            return str(x).replace(".", ",")
        except Exception:
            return ""

def normalize_shape(s: str) -> str:
    """Normalize Shape strings (trim and collapse multiple spaces)."""
    s = (s or "").strip()
    s = re.sub(r"\s+", " ", s)
    return s
def is_current_type(type_str: str) -> bool:
    return normalize_shape(type_str).lower() == "current"


# =========================
# Excel metadata loading
# =========================
def load_excel_metadata(excel_path: str) -> pd.DataFrame:
    """Load required metadata columns from Excel.

    NOTE (hardcoded positional mode):
    This version supports loading an Excel sheet even when it has NO header row.
    In that case, columns are taken by fixed position (1-based Excel columns):

      1  -> No
      2  -> Shape
      3  -> Type
      4  -> PkMax
      5  -> Min
      6  -> T1
      7  -> Tp
      8  -> T2
      9  -> Td
      10 -> Tc
      11 -> T0

    The normal mode (with headers) remains supported for convenience.
    """

    # -------------------------
    # CSV SUPPORT (minimal add)
    # -------------------------
    ext = Path(excel_path).suffix.lower()

    if ext == ".csv":
        # 1) Read CSV with delimiter sniffing
        df = pd.read_csv(excel_path, sep=None, engine="python", encoding="utf-8-sig")

        # 2) If it ended up as ONE column, split it manually
        if df.shape[1] == 1:
            col0 = df.columns[0]
            s = df[col0].astype(str)

            split = s.str.split(",", expand=True)

            if split.shape[1] == 1:
                split = s.str.split(";", expand=True)

            split.columns = split.iloc[0].astype(str).tolist()
            df = split.iloc[1:].reset_index(drop=True)

        cols = {str(c).strip(): c for c in df.columns}

        def find_col(candidates: list[str]):
            for cand in candidates:
                for k in cols.keys():
                    if cand.lower() in k.lower():
                        return cols[k]
            return None

        col_no = find_col(["No"])
        col_shape = find_col(["Shape"])
        col_type = find_col(["Type"])
        col_pkmax = find_col(["Pk;Max", "PkMax", "Max"])
        col_min = find_col(["Min"])
        col_t1 = find_col(["T1"])
        col_tp = find_col(["Tp"])
        col_t2 = find_col(["T2"])
        col_td = find_col(["Td"])
        col_tc = find_col(["Tc"])
        col_t0 = find_col(["To", "T0"])

        missing = [n for n, c in [
            ("No", col_no), ("Shape", col_shape), ("Type", col_type),
            ("PkMax", col_pkmax), ("Min", col_min),
            ("T1", col_t1), ("Tp", col_tp), ("T2", col_t2),
            ("Td", col_td), ("Tc", col_tc), ("T0", col_t0)
        ] if c is None]

        if missing:
            raise ValueError("CSV loaded but missing required columns: " + ", ".join(missing))

        meta = df[[col_no, col_shape, col_type,
                   col_pkmax, col_min,
                   col_t1, col_tp, col_t2,
                   col_td, col_tc, col_t0]].copy()

        meta.columns = ["No", "Shape", "Type", "PkMax", "Min", "T1", "Tp", "T2", "Td", "Tc", "T0"]
        meta = meta.dropna(subset=["No"]).reset_index(drop=True)
        return meta

    # -------------------------
    # Attempt A: normal workbook with header row (existing behavior)
    # -------------------------
    try:
        df = pd.read_excel(excel_path, engine="openpyxl")
        cols = {str(c).strip(): c for c in df.columns}

        def find_col(candidates: list[str]):
            for cand in candidates:
                for k in cols.keys():
                    if cand.lower() in k.lower():
                        return cols[k]
            return None

        col_no = find_col(["No"])  # "No." or "No"
        col_shape = find_col(["Shape"])
        col_type = find_col(["Type"])

        # Amplitudes
        col_pkmax = find_col(["Pk;Max", "Pk; Max", "PkMax", "Max"])
        col_min = find_col(["Min"])

        # Times
        col_t1 = find_col(["T1"])
        col_tp = find_col(["Tp"])
        col_t2 = find_col(["T2"])
        col_td = find_col(["Td"])
        col_tc = find_col(["Tc"])
        col_t0 = find_col(["To", "T0"])

        missing = [
            name
            for name, col in [
                ("No", col_no),
                ("Shape", col_shape),
                ("Type", col_type),
                ("PkMax", col_pkmax),
                ("Min", col_min),
                ("T1", col_t1),
                ("Tp", col_tp),
                ("T2", col_t2),
                ("Td", col_td),
                ("Tc", col_tc),
                ("T0", col_t0),
            ]
            if col is None
        ]

        if not missing:
            meta = df[
                [col_no, col_shape, col_type, col_pkmax, col_min, col_t1, col_tp, col_t2, col_td, col_tc, col_t0]
            ].copy()
            meta.columns = ["No", "Shape", "Type", "PkMax", "Min", "T1", "Tp", "T2", "Td", "Tc", "T0"]
            meta = meta.dropna(subset=["No"]).reset_index(drop=True)
            return meta

        # If headers exist but required columns are not found, fall back to positional.
    except Exception:
        # If normal read fails, fall back to positional.
        pass

    # -------------------------
    # Attempt B: NO header row -> fixed positions
    # -------------------------
    df2 = pd.read_excel(excel_path, engine="openpyxl", header=None)

    if df2.shape[1] < 11:
        raise ValueError(
            f"Excel without header detected, but it has only {df2.shape[1]} columns. "
            "Expected at least 11 columns in fixed order: No, Shape, Type, PkMax, Min, T1, Tp, T2, Td, Tc, T0."
        )

    meta = df2.iloc[:, 0:11].copy()
    meta.columns = ["No", "Shape", "Type", "PkMax", "Min", "T1", "Tp", "T2", "Td", "Tc", "T0"]
    meta = meta.dropna(subset=["No"]).reset_index(drop=True)
    return meta

def fmt_kv_from_v(x) -> str:
    """Back-compat wrapper: format voltage-like values to kV."""
    return fmt_voltage_to_kv(x)


def fmt_a(x) -> str:
    """Back-compat wrapper: format current-like values (numeric part only)."""
    v, _u = fmt_current_with_unit(x)
    return v


    try:
        return fmt_3dec_decimal_comma(x)
    except Exception:
        return ""
def fmt_us_from_s(x) -> str:
    """Back-compat wrapper: format time-like values to microseconds."""
    return fmt_time_to_us(x)

def format_current_auto(value_a: float | None) -> str:
    """Format current value with automatic A/kA unit and decimal comma."""
    if value_a is None or (isinstance(value_a, float) and pd.isna(value_a)):
        return ""
    try:
        v = float(value_a)
    except Exception:
        return ""
    unit = "A"
    if abs(v) >= 1000.0:
        v = v / 1000.0
        unit = "kA"
    s = f"{v:.3f}".replace(".", ",")
    return f"{s} {unit}"


def build_textbox_lines(meta: dict, file_prefix: str | None = None, textbox_mode_c: str = "Auto (from Excel)") -> list[str]:
    """Build textbox content.

    - For v-files: always use Excel-driven rules (Shape + Type).
    - For c-files: if textbox_mode_c != 'Auto (from Excel)', use the selected template.
    """
    no = meta.get("No", "")
    shape_raw = str(meta.get("Shape", ""))
    shape_norm = normalize_shape(shape_raw)
    typ = str(meta.get("Type", ""))
    cur = is_current_type(typ)

    pkmax = meta.get("PkMax")
    mn = meta.get("Min")
    t1 = meta.get("T1")
    tp = meta.get("Tp")
    t2 = meta.get("T2")
    td = meta.get("Td")
    tc = meta.get("Tc")
    t0 = meta.get("T0")


    # Voltage formatter:
    # - v-files: keep existing Excel-driven formatting (with units if present)
    # - c-files: show the SAME unit as written in Excel; if Excel has no unit -> show only the number (no unit)
    _fp = (file_prefix or "").lower().strip()
    if _fp == "c":
        fmt_volt = lambda x: format_voltage_from_excel_preserve_unit(x, blank_if_missing_unit=True)
    else:
        fmt_volt = format_voltage_from_excel

    def add_if(out: list[str], label: str, value_str: str):
        value_str = (value_str or "").strip()
        if value_str:
            out.append(f"{label}:  {value_str}")

    # Apply user-selected textbox template for c-files (channel 2)
    mode = (textbox_mode_c or "Auto (from Excel)").strip()
    if (file_prefix or "").lower() == "c" and mode != "Auto (from Excel)":
        # Minimal MATLAB-style template:
        # TEXT_BOX={["No.",num2str(headertext{1})];["Upk:  ",headertext{6}];};
        if mode == "Default":
            return [
                f"No.{no}",
                f"Upk:  {(fmt_volt(pkmax) if (file_prefix or '').lower() == 'c' else fmt_volt(pkmax))}",
            ]
        if mode.startswith("LI full") and "Ipk max/min" in mode:
            return [
                f"No.{no}",
                "LI full",
                f"Ipk max:  {format_current_auto(pkmax)}",
                f"Ipk min:  {format_current_auto(mn)}",
            ]
        if mode.startswith("LI full") and "(Ipk)" in mode:
            return [
                f"No.{no}",
                "LI full",
                f"Ipk:  {format_current_auto(pkmax)}",
            ]
        if mode.startswith("LI full") and "Upk" in mode:
            return [
                f"No.{no}",
                "LI full",
                f"Upk:  {fmt_volt(pkmax)}",
                f"T1:  {format_time_us(t1)}",
                f"T2:  {format_time_us(t2)}",
            ]

        if mode.startswith("LI tailchopped") and "(Ipk)" in mode:
            return [
                f"No.{no}",
                "LI tailchopped",
                f"Ipk:  {format_current_auto(pkmax)}",
            ]
        if mode.startswith("LI tailchopped") and "Upk" in mode:
            # MATLAB style: T2 is shown with '--'
            t2s = format_time_us(t2)
            t2line = f"T2:-- {t2s}".rstrip() if t2s else "T2:--"
            return [
                f"No.{no}",
                "LI tailchopped",
                f"Upk min:  {fmt_volt(mn)}",
                f"Upk max:  {fmt_volt(pkmax)}",
                f"T1:  {format_time_us(t1)}",
                t2line,
                f"Tc:  {format_time_us(tc)}",
            ]

        if mode.startswith("SI IEC 60060") and "(Ipk" in mode:
            return [
                f"No.{no}",
                "SI IEC 60060",
                f"Ipk:  {format_current_auto(pkmax)}",
                f"Tp:  {format_time_us(tp)}",
                f"T2:  {format_time_us(t2)}",
            ]
        if mode.startswith("SI IEC 60060") and "Upk" in mode:
            return [
                f"No.{no}",
                "SI IEC 60060",
                f"Upk:  {fmt_volt(pkmax)}",
                f"Tp:  {format_time_us(tp)}",
                f"T2:  {format_time_us(t2)}",
            ]

        if mode.startswith("SI IEC 60076") and "Ipk max/min" in mode:
            return [
                f"No.{no}",
                "SI IEC 60076",
                f"Ipk max:  {format_current_auto(pkmax)}",
                f"Ipk min:  {format_current_auto(mn)}",
            ]
        if mode.startswith("SI IEC 60076") and "Upk" in mode:
            return [
                f"No.{no}",
                "SI IEC 60076",
                f"Upk:  {fmt_volt(pkmax)}",
                f"T1:  {format_time_us(t1)}",
                f"Td:  {format_time_us(td)}",
                f"T0:  {format_time_us(t0)}",
            ]

        # Fallback: show a minimal box
        return [f"No.{no}", f"Upk:  {fmt_volt(pkmax)}"]

    # Default behavior (Excel-driven rules, for v-files and auto mode for c)
    # Match MATLAB behavior using normalized shape string:
    # Default behavior (Excel-driven rules, for v-files and auto mode for c)
    # Match MATLAB behavior using normalized shape string:
    if shape_norm == "LI full":
        out = [f"No.{no}", "LI full"]

        if cur:
            add_if(out, "Ipk", format_current_auto(pkmax))
            return out

        add_if(out, "Upk", fmt_volt(pkmax))
        add_if(out, "T1", format_time_us(t1))
        add_if(out, "T2", format_time_us(t2))
        return out
    if shape_norm == "LI frontchopped":
        out = [f"No.{no}", "LI frontchopped"]

        if cur:
            add_if(out, "Ipk", format_current_auto(pkmax))
            add_if(out, "T1", format_time_us(t1))
            add_if(out, "Tc", format_time_us(tc))
            return out

        # Voltage variant
        add_if(out, "Upk min", fmt_volt(mn))
        add_if(out, "Upk max", fmt_volt(pkmax))
        add_if(out, "T1", format_time_us(t1))
        add_if(out, "Tc", format_time_us(tc))
        return out
    if shape_norm == "LI tailchopped":
        if cur:
            return [f"No.{no}", "LI tailchopped", f"Ipk:  {format_current_auto(pkmax)}"]
        t2s = format_time_us(t2)
        t2line = f"T2:-- {t2s}".rstrip() if t2s else "T2:--"
        out = [f"No.{no}", "LI tailchopped"]
        add_if(out, "Upk min", fmt_volt(mn))
        add_if(out, "Upk max", fmt_volt(pkmax))
        add_if(out, "T1", format_time_us(t1))

        t2s = (format_time_us(t2) or "").strip()
        if t2s:
            out.append(f"T2:-- {t2s}")

        add_if(out, "Tc", format_time_us(tc))
        return out
    if shape_norm == "SI IEC 60060":
        if cur:
            return [f"No.{no}", "SI IEC 60060", f"Ipk:  {format_current_auto(pkmax)}", f"Tp:  {format_time_us(tp)}", f"T2:  {format_time_us(t2)}"]
        out = [f"No.{no}", "SI IEC 60060"]
        add_if(out, "Upk", fmt_volt(pkmax))
        add_if(out, "Tp", format_time_us(tp))
        add_if(out, "T2", format_time_us(t2))
        return out
    if shape_norm == "SI IEC 60076":
        if cur:
            # Best effort: if it's current and this shape is used, show max/min like MATLAB's spaced variant.
            return [f"No.{no}", "SI IEC 60076", f"Ipk max:  {format_current_auto(pkmax)}", f"Ipk min:  {format_current_auto(mn)}"]
        out = [f"No.{no}", "SI IEC 60076"]
        add_if(out, "Upk", fmt_volt(pkmax))
        add_if(out, "T1", format_time_us(t1))
        add_if(out, "Td", format_time_us(td))
        add_if(out, "T0", format_time_us(t0))
        return out

    # Fallback
    if cur:
        return [f"No.{no}", f"Ipk:  {format_current_auto(pkmax)}"]
    return [f"No.{no}", f"Upk:  {fmt_volt(pkmax)}"]




def choose_nice_step(data_min: float, data_max: float, target_ticks: int = 8) -> float:
    """Pick a 'nice' major tick step for Y axis.

    Uses 1/2/5 * 10^n stepping similar to MATLAB 'nice' ticks.
    """
    span = abs(float(data_max) - float(data_min))
    if span == 0:
        return 1.0
    raw = span / float(target_ticks)
    exp = 10 ** (math.floor(math.log10(raw)))
    frac = raw / exp
    if frac <= 1:
        nice = 1
    elif frac <= 2:
        nice = 2
    elif frac <= 5:
        nice = 5
    else:
        nice = 10
    return float(nice * exp)

def grid_params(grid_scale: str):
    """
    Returns:
    x_minor_step_us, y_minor_div, major_lw, minor_lw
    """
    gs = (grid_scale or "Normal").lower()
    if gs == "fine":
        # keep denser minor ticks but don't draw minor grid by default
        return 2.0, 4, 0.85, 0.0
    if gs == "coarse":
        return 5.0, 2, 0.85, 0.0  # no minor grid
    # "Normal": keep minor ticks (for easier reading) but DO NOT draw minor grid.
    # This avoids the extra grid lines on unlabeled minor ticks.
    return 5.0, 2, 0.85, 0.0
def draw_oscillogram_on_axes(
    ax,
    df: pd.DataFrame,
    serial_number: str,
    divider_str: str,
    meta: dict,
    xmax_us: float,
    y_mode: str,
    y_step_kv: float | None,
    grid_scale: str,
    file_prefix: str | None = None,
    divider1_str: str = "",
    divider2_str: str = "",
    show_matlab_header: bool = True,
    textbox_mode_c: str = "Auto (from Excel)",
    grid_major_lw: float | None = None,
    grid_minor_lw: float | None = None,
    plot_frame_lw: float = 1.2,
    textbox_frame_lw: float = 1.0,
    tick_major_lw: float = 1.0,
    tick_minor_lw: float = 0.8,
    font_label: int = 6,
    font_tick: int = 6,
    font_header: int = 6,
    font_textbox: int = 6,
    y_scale: float = 1.0,
    textbox_pos_mode: str = "auto",
):
    t_us = df["time_s"] * 1e6

    # df["value_v"] is in VOLTS. Apply optional scaling, then choose display unit automatically: V / kV / MV.
    try:
        _ys = float(y_scale)
    except Exception:
        _ys = 1.0
    y_v = df["value_v"] * _ys
    peak_v = float(max(abs(float(y_v.min())), abs(float(y_v.max()))))

    if peak_v < 1000.0:
        y_plot = y_v
        y_unit_label = "V"
    elif peak_v < 1_000_000.0:
        y_plot = y_v / 1000.0
        y_unit_label = "kV"
    else:
        y_plot = y_v / 1_000_000.0
        y_unit_label = "MV"

    use_mv = (y_unit_label == "MV")

    ax.clear()
    # Curve (thin)
    ax.plot(t_us, y_plot, linewidth=1.0, color="green")
    # Font sizes (a bit smaller to better match typical Word report look)
    ax.set_xlabel("TIME (us)", fontsize=font_label)
    ax.set_ylabel(f"VOLTAGE ({y_unit_label})", fontsize=font_label)
    ax.tick_params(labelsize=font_tick)
    ax.tick_params(which="major", width=float(tick_major_lw))
    ax.tick_params(which="minor", width=float(tick_minor_lw))
    # X
    # Special rule: tailchopped impulses should use a fixed window:
    #   Xmax = 10 us, with 10% of that in front of zero.
    # i.e. [-1 us .. 10 us]
    shape_norm = normalize_shape(str(meta.get("Shape", ""))).lower()

    if "tailchopped" in shape_norm:
        xmin_us = -1.0
        xmax_plot = 10.0
    else:
        xmax_plot = float(xmax_us)
        # For 'c' files, use 10% pretrigger (as requested)
        if (file_prefix or "").lower() == "c":
            xmin_us = -0.10 * xmax_plot
        else:
            xmin_us = -5.0

    ax.set_xlim(xmin_us, xmax_plot)

    # Grid params (shared)
    x_minor_step_us, y_minor_div, major_lw, minor_lw = grid_params(grid_scale)
    if grid_major_lw is not None:
        major_lw = float(grid_major_lw)
    if grid_minor_lw is not None:
        minor_lw = float(grid_minor_lw)

    # X ticks:
    # - tailchopped: fixed [-1..10] us with 1 us major ticks
    # - otherwise (v & c): 10.5 major divisions total
    #   -> 0.5 division left of zero, 10 divisions right of zero
    if "tailchopped" in shape_norm:
        ax.xaxis.set_major_locator(MultipleLocator(1))
        ax.xaxis.set_minor_locator(MultipleLocator(0.2))
    else:
        # Major step so that right side has exactly 10 divisions up to xmax_plot
        major_step_x = float(xmax_plot) / 10.0

        # Shift so that zero is not on the border:
        # 0.5 major division to the left, 10 to the right
        xmin_us = -0.5 * major_step_x
        xmax_plot = 10.0 * major_step_x
        ax.set_xlim(xmin_us, xmax_plot)

        ax.xaxis.set_major_locator(MultipleLocator(major_step_x))
        ax.xaxis.set_minor_locator(MultipleLocator(major_step_x / 5.0))

    # Decimal comma on ticks (both axes)
    ax.xaxis.set_major_formatter(FuncFormatter(fmt_tick_decimal_comma))
    ax.yaxis.set_major_formatter(FuncFormatter(fmt_tick_decimal_comma))
    # Y
    y_min = float(y_plot.min())
    y_max = float(y_plot.max())
    if y_mode == "manual":
        major_step = float(y_step_kv) if (y_step_kv is not None and y_step_kv > 0) else 20.0
        if use_mv:
            major_step = major_step / 1000.0
    else:
        major_step = choose_nice_step(y_min, y_max, target_ticks=8)
    y0 = major_step * math.floor(y_min / major_step)
    y1 = major_step * math.ceil(y_max / major_step)
    if y1 == y0:
        y1 += major_step
    ax.set_ylim(y0, y1)
    ax.yaxis.set_major_locator(MultipleLocator(major_step))
    if y_minor_div and y_minor_div > 1:
        ax.yaxis.set_minor_locator(MultipleLocator(major_step / y_minor_div))
    # Grid (make it crisp like your example)
    ax.grid(True, which="major", linestyle="--", linewidth=major_lw, color="black")
    if minor_lw > 0:
        ax.grid(True, which="minor", linestyle="--", linewidth=minor_lw, color="black")
        # --- Solid axes at x=0 and y=0 (over the dashed grid) ---
    x0, x1 = ax.get_xlim()
    y0, y1 = ax.get_ylim()
    # draw only if 0 is inside the visible range
    if x0 <= 0 <= x1:
        ax.axvline(0, color="black", linewidth=major_lw, linestyle="-", zorder=3)
    if y0 <= 0 <= y1:
        ax.axhline(0, color="black", linewidth=major_lw, linestyle="-", zorder=3)
    # Frame & ticks a bit stronger
    for spine in ax.spines.values():
        spine.set_linewidth(float(plot_frame_lw))
        spine.set_color("black")
    ax.tick_params(axis="both", which="major", width=float(tick_major_lw), length=3, colors="black")
    ax.tick_params(axis="both", which="minor", width=float(tick_minor_lw), length=2, colors="black")
    # Headers
    # MATLAB-style header: TWO anchored texts (as requested):
    #  - LEFT: ' Serial number: ...' (starts at left edge with a single leading space)
    #  - RIGHT: 'CHANNEL X (VOLTAGE)   Divider: ...' anchored to right edge
    if show_matlab_header and file_prefix in ("v", "c"):
        if file_prefix == "v":
            right_text = f"CHANNEL 1 (VOLTAGE)   Divider: {divider1_str}"
        else:
            right_text = f"CHANNEL 2 (VOLTAGE)   Divider: {divider2_str}"

        ax.text(
            -0.02, 1.04,
            f"Serial number: {serial_number}",
            transform=ax.transAxes,
            ha="left", va="bottom",
            fontsize=font_header,
            color="black",
            clip_on=False,
        )
        ax.text(
            1.0, 1.04,
            right_text,
            transform=ax.transAxes,
            ha="right", va="bottom",
            fontsize=font_header,
            color="black",
            clip_on=False,
        )
    else:
        # Fallback (should rarely happen): keep a simple serial line.
        ax.text(
            0.0, 1.08,
            f"Serial number: {serial_number}",
            transform=ax.transAxes,
            ha="left", va="bottom",
            fontsize=font_header,
            color="black",
            clip_on=False,
        )

        ax.text(
            1.0, 1.02,
            f"Divider: {divider_str} V/V",
            transform=ax.transAxes,
            ha="right", va="bottom",
            fontsize=font_header,
            color="black"
        )
    # Info box
    box_lines = build_textbox_lines(meta, file_prefix=file_prefix, textbox_mode_c=textbox_mode_c)
    box_text = "\n".join(box_lines)
    # Info box (position can be auto based on waveform polarity)
    mode = (textbox_pos_mode or "auto").strip().lower()
    if mode in ("auto", "a"):
        # Heuristic (INVERTED by request):
        # - If waveform is mostly positive -> put textbox near TOP-RIGHT
        # - If waveform is mostly negative -> put textbox near BOTTOM-RIGHT
        # - If it crosses zero -> use the dominant absolute peak
        if y_max <= 0:
            corner = "bottom-right"
        elif y_min >= 0:
            corner = "top-right"
        else:
            corner = "top-right" if abs(y_max) >= abs(y_min) else "bottom-right"
    elif mode in ("tr", "top", "top-right", "topright"):
        corner = "top-right"
    else:
        corner = "bottom-right"
    if corner == "top-right":
        tx, ty, ha, va = 0.98, 0.965, "right", "top"
    else:
        tx, ty, ha, va = 0.98, 0.04, "right", "bottom"
    ax.text(
        tx, ty, box_text,
        transform=ax.transAxes,
        ha=ha, va=va,
        multialignment="left",
        bbox=dict(
            boxstyle="square,pad=0.35",
            facecolor="white",
            edgecolor="black",
            linewidth=float(textbox_frame_lw),
        ),
        fontsize=font_textbox,
        color="black",
        clip_on=True,
    )
def make_oscillogram_png(
    txt_path: str,
    out_png: str,
    serial_number: str,
    divider_str: str,
    meta: dict,
    xmax_us: float,
    y_mode: str,
    y_step_kv: float | None,
    grid_scale: str,
    out_width_cm: float,
    out_height_cm: float,
    out_dpi: int,
    grid_major_lw: float | None = None,
    grid_minor_lw: float | None = None,
    plot_frame_lw: float = 1.2,
    textbox_frame_lw: float = 1.0,
    tick_major_lw: float = 1.0,
    tick_minor_lw: float = 0.8,
    font_label: int = 6,
    font_tick: int = 6,
    font_header: int = 6,
    font_textbox: int = 6,
    textbox_mode_c: str = "Auto (from Excel)",
    textbox_pos_mode: str = "auto",
    y_scale: float = 1.0,
    file_prefix: str | None = None,
    divider1_str: str = "",
    divider2_str: str = "",
    show_matlab_header: bool = True,
    margin_left: float = 0.12,
    margin_right: float = 0.93,
    margin_bottom: float = 0.22,
    margin_top: float = 0.88,
    use_tight_bbox: bool = False,
    pad_inches: float = 0.02,
):
    """
    Key for sharpness:
    - figure size matches final Word size (cm)
    - DPI high (300–600)
    """
    df = read_measurement_txt(txt_path)
    fig_w_in = cm_to_inches(out_width_cm)
    fig_h_in = cm_to_inches(out_height_cm)
    fig = plt.figure(figsize=(fig_w_in, fig_h_in), dpi=out_dpi)
    ax = fig.add_subplot(111)
    draw_oscillogram_on_axes(
        ax, df, serial_number, divider_str, meta,
        file_prefix=file_prefix,
        divider1_str=divider1_str,
        divider2_str=divider2_str,
        show_matlab_header=show_matlab_header,
        textbox_mode_c=textbox_mode_c,
        xmax_us=xmax_us, y_mode=y_mode, y_step_kv=y_step_kv, grid_scale=grid_scale,
        grid_major_lw=grid_major_lw,
        grid_minor_lw=grid_minor_lw,
        plot_frame_lw=plot_frame_lw,
        textbox_frame_lw=textbox_frame_lw,
        tick_major_lw=tick_major_lw,
        tick_minor_lw=tick_minor_lw,
        font_label=font_label,
        font_tick=font_tick,
        font_header=font_header,
        font_textbox=font_textbox,
        textbox_pos_mode=textbox_pos_mode,
        y_scale=y_scale,
    )
    # Margins around the plot area (MATLAB-like defaults)
    # Values are relative in [0..1] with respect to the figure.
    fig.subplots_adjust(
        left=float(margin_left),
        right=float(margin_right),
        bottom=float(margin_bottom),
        top=float(margin_top),
    )
    # Save exact size. Optional tight bbox if user wants automatic cropping.
    save_kwargs = dict(dpi=out_dpi, facecolor="white")
    if use_tight_bbox:
        save_kwargs.update({"bbox_inches": "tight", "pad_inches": float(pad_inches)})
    fig.savefig(out_png, **save_kwargs)
    plt.close(fig)
# =========================
# Word insertion (STRICT TEMPLATE)
# =========================
def clear_cell_keep_format(cell):
    for p in cell.paragraphs:
        for r in p.runs:
            r.text = ""
    if len(cell.paragraphs) > 1:
        for p in cell.paragraphs[1:]:
            p._element.getparent().remove(p._element)
    if not cell.paragraphs:
        cell.add_paragraph("")

def _parse_textbox_lines_to_row(lines: list[str]) -> dict:
    """Parse textbox lines (as rendered on the oscillogram) into a flat dict.

    Rules:
    - Lines like 'Upk:  535,858 kV' become {'Upk': '535,858 kV'}
    - 'No.5' becomes {'No.': '5'}
    - The first non-empty line without ':' after 'No.' is treated as 'Shape'
    - Any other non-empty no-colon lines are stored as 'Info' (concatenated)
    """
    row: dict[str, str] = {}
    info_extra: list[str] = []
    for raw in (lines or []):
        s = (raw or "").strip()
        if not s:
            continue
        if ":" in s:
            k, v = s.split(":", 1)
            k = k.strip()
            v = v.strip()
            if k.lower().startswith("upk"):
                v = v.replace("kV", "").replace("MV", "").strip()
            if k:
                row[k] = v
            continue
        low = s.lower()
        if low.startswith("no."):
            row["No."] = s[3:].strip()
            continue
        if "Shape" not in row:
            row["Shape"] = s
        else:
            info_extra.append(s)
    if info_extra:
        row["Info"] = " / ".join(info_extra)
    return row

def _preferred_table_columns(all_keys: list[str]) -> list[str]:
    preferred = [
        "No.",
        "Shape",
        "Upk",
        "Upk min",
        "Upk max",
        "Ipk",
        "Ipk max",
        "Ipk min",
        "T1",
        "Tp",
        "T2",
        "Td",
        "Tc",
        "T0",
        "Info",
    ]
    out = []
    for k in preferred:
        if k in all_keys and k not in out:
            out.append(k)
    # add any remaining keys (stable order)
    for k in all_keys:
        if k not in out:
            out.append(k)
    return out

def _append_textbox_values_table(doc: Document, table_rows: list[dict]) -> None:
    """Append a Word table with textbox values for v* oscillograms."""
    if not table_rows:
        return
    # Collect all keys
    keys = []
    for r in table_rows:
        for k in r.keys():
            if k not in keys:
                keys.append(k)
    cols = _preferred_table_columns(keys)
    if not cols:
        return

    doc.add_paragraph("")
    doc.add_paragraph("Textbox values – voltage channel (v*)")
    t = doc.add_table(rows=1, cols=len(cols))
    try:
        t.style = "Table Grid"
    except Exception:
        pass
    hdr = t.rows[0].cells
    for j, k in enumerate(cols):
        hdr[j].text = k

    for r in table_rows:
        cells = t.add_row().cells
        for j, k in enumerate(cols):
            cells[j].text = str(r.get(k, ""))

def insert_images_into_template_strict(
    template_docx: str,
    images: list[str],
    output_docx: str,
    width_cm: float,
    height_cm: float,
    lock_aspect: bool,
    peaks_table_rows: list[dict] | None = None,
):
    doc = Document(template_docx)
    if not doc.tables:
        raise ValueError("The Word template contains no tables.")
    tbl = doc.tables[0]
    rows = len(tbl.rows)
    cols = len(tbl.columns)
    capacity = rows * cols
    if len(images) > capacity:
        raise ValueError(f"Template capacity is {capacity} images, but you generated {len(images)}.")
    w_in = cm_to_inches(width_cm)
    h_in = cm_to_inches(height_cm)
    idx = 0
    for r in range(rows):
        for c in range(cols):
            if idx >= len(images):
                break
            cell = tbl.cell(r, c)
            clear_cell_keep_format(cell)

            # optional vertical centering
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            tcPr = cell._tc.get_or_add_tcPr()

            # remove existing tcMar if present
            for el in tcPr.findall(qn("w:tcMar")):
                tcPr.remove(el)

            tcMar = OxmlElement("w:tcMar")
            for side in ("top", "left", "bottom", "right"):
                node = OxmlElement(f"w:{side}")
                node.set(qn("w:w"), "0")
                node.set(qn("w:type"), "dxa")
                tcMar.append(node)
            tcPr.append(tcMar)

            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = 0
            p.paragraph_format.space_after = 0
            p.paragraph_format.line_spacing = 1

            run = p.add_run()
            if lock_aspect:
                run.add_picture(images[idx], width=Inches(w_in))
            else:
                run.add_picture(images[idx], width=Inches(w_in), height=Inches(h_in))
            idx += 1
        if idx >= len(images):
            break
    # Optional: append a summary table with textbox values (for v-oscillograms)
    if peaks_table_rows:
        try:
            _append_textbox_values_table(doc, peaks_table_rows)
        except Exception:
            # Don't fail the whole export if table fails
            pass
    doc.save(output_docx)
def export_strict_template_multipart(
    template_docx: str,
    images: list[str],
    output_base_docx: str,
    width_cm: float,
    height_cm: float,
    lock_aspect: bool,
):
    probe = Document(template_docx)
    if not probe.tables:
        raise ValueError("The Word template contains no tables.")
    tbl = probe.tables[0]
    capacity = len(tbl.rows) * len(tbl.columns)
    base = Path(output_base_docx)
    stem = base.stem
    suffix = base.suffix
    parts = []
    for start in range(0, len(images), capacity):
        part_images = images[start:start + capacity]
        part_idx = (start // capacity) + 1
        out_part = str(base.with_name(f"{stem}_part{part_idx:02d}{suffix}"))
        insert_images_into_template_strict(
            template_docx, part_images, out_part,
            width_cm=width_cm, height_cm=height_cm, lock_aspect=lock_aspect
        )
        parts.append(out_part)
    return parts, capacity
