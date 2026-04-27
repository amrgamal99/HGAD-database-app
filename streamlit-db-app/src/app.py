import os
import re
import requests
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Optional, Tuple, Dict, List
import urllib.parse

import pandas as pd
import streamlit as st
import base64

# ReportLab (PDF) + Arabic shaping
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    Image as RLImage,
    PageBreak,
)
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import arabic_reshaper
from bidi.algorithm import get_display

try:
    from PIL import Image as PILImage
except Exception:
    PILImage = None

from db.connection import (
    get_db_connection,
    fetch_data,
    fetch_financial_flow_view,
    fetch_contract_summary_view,
)
from components.filters import (
    create_company_dropdown,
    create_project_dropdown,
    create_type_dropdown,
    create_column_search,
)

# =========================================================
# Paths / Assets
# =========================================================
BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"

LOGO_CANDIDATES = [ASSETS_DIR / "logo.png"]
WIDE_LOGO_CANDIDATES = [
    ASSETS_DIR / "logo_wide.png",
    ASSETS_DIR / "logo_wide.jpg",
    ASSETS_DIR / "logo_wide.jpeg",
    ASSETS_DIR / "logo_wide.webp",
]

AR_FONT_CANDIDATES = [
    ASSETS_DIR / "Cairo-Regular.ttf",
    ASSETS_DIR / "Amiri-Regular.ttf",
    ASSETS_DIR / "NotoNaskhArabic-Regular.ttf",
    Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
]

_AR_RE = re.compile(r"[\u0600-\u06FF]")

_AMIRI_URL = "https://github.com/aliftype/amiri/raw/main/Amiri-Regular.ttf"
_NOTO_URL  = "https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoNaskhArabic/NotoNaskhArabic-Regular.ttf"


def _format_date_arabic(val) -> str:
    """
    Convert various date representations to a clean date string: YYYY/MM/DD.
    Handles: datetime objects, YYYYMMDD integers/strings, ISO strings, and common formats.
    """
    if val is None:
        return ""
    # Handle pandas/numpy NaT and NaN
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    sval = str(val).strip()
    if not sval or sval.lower() in ("nan", "none", "nat", ""):
        return ""
    # Remove decimal part if stored as float e.g. "20250305.0"
    if re.match(r"^\d{8}\.0$", sval):
        sval = sval.split(".")[0]
    # Handle compact YYYYMMDD format (e.g. "20250305")
    if re.match(r"^\d{8}$", sval):
        try:
            from datetime import datetime
            dt = datetime.strptime(sval, "%Y%m%d")
            return dt.strftime("%Y/%m/%d")
        except Exception:
            pass
    # Try pandas parsing
    try:
        dt = pd.to_datetime(sval, dayfirst=False, errors="coerce")
        if pd.isna(dt):
            dt = pd.to_datetime(sval, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%Y/%m/%d")
    except Exception:
        pass
    return sval


# =========================================================
# Small utils
# =========================================================
def _ensure_arabic_font() -> Optional[Path]:
    for p in AR_FONT_CANDIDATES:
        if Path(p).exists() and Path(p).stat().st_size > 10_000:
            return Path(p)
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    for url, fname in [(_AMIRI_URL, "Amiri-Regular.ttf"), (_NOTO_URL, "NotoNaskhArabic-Regular.ttf")]:
        dest = ASSETS_DIR / fname
        try:
            r = requests.get(url, timeout=30)
            if r.status_code == 200 and len(r.content) > 10_000:
                dest.write_bytes(r.content)
                return dest
        except Exception:
            pass
    return None


def _first_existing(paths) -> Optional[Path]:
    for p in paths:
        pth = Path(p)
        if pth.exists() and pth.is_file() and pth.stat().st_size > 0:
            return pth
    return None

def _image_size(path: Path) -> Tuple[int, int]:
    if PILImage:
        try:
            with PILImage.open(path) as im:
                return im.size
        except Exception:
            pass
    return (600, 120)

def _site_logo_path() -> Optional[Path]:
    return _first_existing(LOGO_CANDIDATES)

def _wide_logo_path() -> Optional[Path]:
    return _first_existing(WIDE_LOGO_CANDIDATES)


@st.cache_resource
def register_arabic_font() -> Tuple[str, bool]:
    font_path = _ensure_arabic_font()
    if font_path:
        name = font_path.stem
        try:
            pdfmetrics.registerFont(TTFont(name, str(font_path)))
            return name, True
        except Exception:
            pass
    for fallback_path in [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/TTF/DejaVuSans.ttf",
    ]:
        if Path(fallback_path).exists():
            try:
                pdfmetrics.registerFont(TTFont("DejaVuSans", fallback_path))
                return "DejaVuSans", False
            except Exception:
                pass
    return "Helvetica", False


def looks_arabic(s: str) -> bool:
    return bool(_AR_RE.search(str(s or "")))

def shape_arabic(s: str) -> str:
    try:
        return get_display(arabic_reshaper.reshape(str(s)))
    except Exception:
        return str(s)

def _safe_filename(s: str) -> str:
    return (
        (s or "")
        .replace("/", "-").replace("\\", "-").replace(":", "-")
        .replace("*", "-").replace("?", "-").replace('"', "'")
        .replace("<", "(").replace(">", ")").replace("|", "-")
    )


# =========================================================
# Streamlit Config + CSS
# =========================================================
st.set_page_config(
    page_title="قاعدة البيانات والتقارير المالية | HGAD",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
:root{
  --bg:#0a0f1a; --panel:#0f172a; --panel-2:#0b1220; --muted:#9fb2d9;
  --text:#e5e7eb; --accent:#1E3A8A; --accent-2:#2563eb; --line:#23324d;
}

html, body{
  direction: rtl !important; text-align: right !important;
  font-family: "Cairo","Noto Kufi Arabic","Segoe UI",Tahoma,sans-serif !important;
  color:var(--text) !important; background:var(--bg) !important;
}

[data-testid="stSidebar"]{
  background: linear-gradient(180deg, #0b1220, #0a1020) !important;
  border-inline-start: 1px solid var(--line) !important;
  min-width: 280px !important;
  transition: all 0.3s ease !important;
}

[data-testid="stSidebar"][aria-expanded="false"]{
  min-width: 0px !important;
  width: 0px !important;
  overflow: hidden !important;
  border: none !important;
}

[data-testid="stSidebarCollapseButton"]{
  display: flex !important;
  position: fixed !important;
  z-index: 9999 !important;
  top: 12px !important;
  right: 12px !important;
  background: var(--accent) !important;
  border-radius: 50% !important;
  width: 40px !important;
  height: 40px !important;
  align-items: center !important;
  justify-content: center !important;
  box-shadow: 0 4px 14px rgba(0,0,0,0.5) !important;
  border: 1px solid rgba(255,255,255,0.15) !important;
}
[data-testid="stSidebarCollapseButton"]:hover{
  background: var(--accent-2) !important;
}
[data-testid="stSidebarCollapseButton"] svg{
  fill: white !important;
  color: white !important;
}

@media (max-width: 768px) {
  [data-testid="stSidebar"][aria-expanded="true"]{
    position: fixed !important;
    top: 0 !important;
    right: 0 !important;
    height: 100vh !important;
    z-index: 9998 !important;
    width: 85vw !important;
    min-width: unset !important;
    max-width: 340px !important;
    box-shadow: -8px 0 32px rgba(0,0,0,0.6) !important;
  }
  [data-testid="stSidebarCollapseButton"]{
    top: 8px !important;
    right: 8px !important;
  }
}

.hr-accent{ height:2px; border:0; background:linear-gradient(90deg, transparent, var(--accent), transparent); margin: 8px 0 14px; }

.card{ background:var(--panel); border:1px solid var(--line); border-radius:16px; padding:14px; box-shadow:0 6px 24px rgba(3,10,30,.25); }
.card.soft{ background:var(--panel-2); }

.fin-head{
  display:flex; justify-content:space-between; align-items:center;
  border: 1px dashed rgba(37,99,235,.35); border-radius:16px;
  padding: 16px 18px; margin:8px 0 14px; background:linear-gradient(180deg,#0b1220,#0e1424);
}
.fin-head .line{ font-size:22px; font-weight:900; color:var(--text); }
.badge{ background:var(--accent); color:#fff; padding:6px 12px; border-radius:999px; font-weight:700; }

.date-box{ border:1px solid var(--line); border-radius:16px; padding:12px; background:var(--panel-2); margin-bottom:12px; }
.date-row{ display:flex; gap:12px; flex-wrap:wrap; align-items:center; }
[data-testid="stDateInput"] input{
  background:#0f172a !important; color:var(--text) !important;
  border:1px solid var(--line) !important; border-radius:10px !important;
  text-align:center !important; height:44px !important; min-width:190px !important;
}
[data-testid="stDateInput"] label{ color:var(--muted) !important; font-weight:700; }

[data-testid="stDataFrame"] thead tr th{
  position: sticky; top: 0; z-index: 2;
  background: #132036; color: #e7eefc; font-weight:800; font-size:15px;
  border-bottom: 1px solid var(--line);
}
[data-testid="stDataFrame"] div[role="row"]{ font-size:14.5px; }
[data-testid="stDataFrame"] div[role="row"]:nth-child(even){ background: rgba(255,255,255,.03); }

.hsec{ color:#e7eefc; font-weight:900; margin:6px 0 10px; font-size: 22px; }

.fin-panel{ display:grid; grid-template-columns: 1fr 1fr; gap:20px; margin-top:10px; }
.fin-table{ width:100%; border-collapse:collapse; table-layout:fixed; border-radius:14px; overflow:hidden; }
.fin-table th, .fin-table td{ border:1px solid var(--line); padding:12px; font-size:14.5px; white-space:normal; word-wrap:break-word; }
.fin-table tr:hover td{ background:#111a2d; transition: background .2s ease; }
.fin-table td.value{ background:#0f1a30; font-weight:800; text-align:center; width:34%; }
.fin-table td.label{ background:#0d1628; font-weight:700; text-align:right; width:66%; }

.hsec, .fin-head, h1, h3 {
  text-align: right !important;
  direction: rtl !important;
}
</style>
""",
    unsafe_allow_html=True,
)

# =========================================================
# Header
# =========================================================
def _logo_html() -> str:
    p = _site_logo_path()
    if not p:
        return ""
    ext = p.suffix.lower().lstrip(".") or "png"
    mime = f"image/{'jpeg' if ext in ('jpg','jpeg') else ext}"
    b64 = base64.b64encode(p.read_bytes()).decode("ascii")
    return f'<img src="data:{mime};base64,{b64}" width="64" />'

c_logo, c_title = st.columns([1, 6], gap="small")
with c_logo:
    st.markdown(_logo_html(), unsafe_allow_html=True)
with c_title:
    st.markdown(
        """
<h1 style="color:#e7eefc; font-weight:900; margin:0;">
  قاعدة البيانات والتقارير المالية
  <span style="font-size:18px; color:#9fb2d9; font-weight:600;">| HGAD Company</span>
</h1>
""",
        unsafe_allow_html=True,
    )
st.markdown('<hr class="hr-accent"/>', unsafe_allow_html=True)


# =========================================================
# PDF pre-processing helpers
# =========================================================
_DATE_KEYWORDS = ["تاريخ", "date", " التعاقد", "إصدار", "اصدار"]

def _is_date_col(col_name: str) -> bool:
    cn = str(col_name).lower()
    return any(k in cn for k in _DATE_KEYWORDS)

def _is_percentage_col(col_name: str) -> bool:
    return "نسبة الاعمال المنفذة" in str(col_name)

def _preprocess_df_for_pdf(df: pd.DataFrame) -> pd.DataFrame:
    """
    Apply PDF-specific fixes to a dataframe before rendering:
      1. Date columns → YYYY/MM/DD string
      2. نسبة الاعمال المنفذة → append % if missing
    All other values (bank names, project names, etc.) are shown exactly as stored in the database.
    """
    out = df.copy()
    for col in out.columns:
        if _is_date_col(col):
            out[col] = out[col].map(_format_date_arabic)
        elif _is_percentage_col(col):
            def _add_pct(v):
                s = str(v) if pd.notna(v) else ""
                if not s or s.lower() in ("nan", "none"): return ""
                return s if s.strip().endswith("%") else f"{s}%"
            out[col] = out[col].map(_add_pct)
    return out


# =========================================================
# Excel helpers
# =========================================================
def _pick_excel_engine() -> Optional[str]:
    try:
        import xlsxwriter; return "xlsxwriter"
    except Exception:
        pass
    try:
        import openpyxl; return "openpyxl"
    except Exception:
        return None

def _estimate_col_widths_chars(df: pd.DataFrame) -> List[float]:
    widths = []
    for col in df.columns:
        max_len = max([len(str(col))] + [len(str(v)) for v in df[col].values])
        widths.append(min(max_len + 4, 60))
    return widths

def _chars_to_pixels(chars: float) -> float:
    return chars * 7.2

def _compose_title(company: str, project: str, data_type: str, dfrom, dto) -> str:
    parts = []
    if company:   parts.append(f"الشركة: {company}")
    if project:   parts.append(f"المشروع: {project}")
    if data_type: parts.append(f"النوع: {data_type}")
    if dfrom or dto:
        d_from_str = _format_date_arabic(dfrom) if dfrom else "—"
        d_to_str   = _format_date_arabic(dto)   if dto   else "—"
        parts.append(f"الفترة: {d_from_str} ← {d_to_str}")
    return " | ".join(parts)

def _insert_wide_logo(ws, df: pd.DataFrame, start_row: int, start_col: int = 0) -> int:
    wlp = _wide_logo_path()
    if not wlp:
        return start_row
    widths_chars = _estimate_col_widths_chars(df)
    total_width_px = _chars_to_pixels(sum(widths_chars))
    try:
        img_w_px, img_h_px = _image_size(wlp)
        if img_w_px <= 0: img_w_px = 1000
        x_scale = max(0.1, total_width_px / float(img_w_px))
        y_scale = x_scale
        ws.insert_image(start_row, start_col, str(wlp),
                        {"x_scale": x_scale, "y_scale": y_scale, "object_position": 2})
        scaled_h_px = img_h_px * y_scale
        ws.set_row(start_row, int(scaled_h_px * 0.75))
        return start_row + 1
    except Exception:
        ws.set_row(start_row, 80)
        ws.insert_image(start_row, start_col, str(wlp), {"x_scale": 0.5, "y_scale": 0.5, "object_position": 2})
        return start_row + 1

def _write_excel_table(ws, workbook, df: pd.DataFrame, start_row: int, start_col: int) -> Tuple[int, int, int, int]:
    hdr_fmt  = workbook.add_format({"align": "right", "bold": True})
    fmt_text = workbook.add_format({"align": "right"})
    fmt_date = workbook.add_format({"align": "right", "num_format": "yyyy-mm-dd"})
    fmt_num  = workbook.add_format({"align": "right", "num_format": "#,##0.00"})
    fmt_link = workbook.add_format({"font_color": "blue", "underline": 1, "align": "right"})

    r0, c0 = start_row, start_col

    for j, col in enumerate(df.columns):
        ws.write(r0, c0 + j, col, hdr_fmt)

    for i in range(len(df)):
        for j, col in enumerate(df.columns):
            val = df.iloc[i, j]
            colname = str(col)
            sval = "" if pd.isna(val) else str(val)
            if sval.startswith(("http://", "https://")) or ("رابط" in colname and sval):
                ws.write_url(r0 + 1 + i, c0 + j, sval, fmt_link, string="فتح الرابط")
            elif pd.api.types.is_datetime64_any_dtype(df[col]):
                if pd.notna(val): ws.write_datetime(r0 + 1 + i, c0 + j, pd.to_datetime(val), fmt_date)
                else: ws.write_blank(r0 + 1 + i, c0 + j, None, fmt_text)
            elif pd.api.types.is_numeric_dtype(df[col]):
                if pd.notna(val): ws.write_number(r0 + 1 + i, c0 + j, float(val), fmt_num)
                else: ws.write_blank(r0 + 1 + i, c0 + j, None, fmt_text)
            else:
                ws.write(r0 + 1 + i, c0 + j, sval, fmt_text)

    exclude_keywords = ['id', 'رقم', 'تاريخ', 'date', 'code', 'كود', 'بنك', 'bank', 'نوع', 'type']
    sum_row_idx = r0 + 1 + len(df)
    if len(df.columns) > 0:
        ws.write(sum_row_idx, c0, "المجموع", hdr_fmt)
    for j, col in enumerate(df.columns):
        if j == 0: continue
        col_lower = str(col).lower()
        if not any(kw in col_lower for kw in exclude_keywords):
            try:
                numeric_col = pd.to_numeric(df[col], errors='coerce')
                if not numeric_col.isna().all():
                    s = numeric_col.sum()
                    if not pd.isna(s) and abs(s) > 0.001:
                        ws.write(sum_row_idx, c0 + j, s, fmt_num)
                    else:
                        ws.write(sum_row_idx, c0 + j, "", fmt_text)
                else:
                    ws.write(sum_row_idx, c0 + j, "", fmt_text)
            except Exception:
                ws.write(sum_row_idx, c0 + j, "", fmt_text)
        else:
            ws.write(sum_row_idx, c0 + j, "", fmt_text)

    count_row_idx = sum_row_idx + 1
    if len(df.columns) > 0:
        ws.write(count_row_idx, c0, "عدد الصفوف", hdr_fmt)
    if len(df.columns) > 1:
        ws.write(count_row_idx, c0 + 1, len(df), fmt_text)
        for j in range(2, len(df.columns)):
            ws.write(count_row_idx, c0 + j, "", fmt_text)

    r1 = count_row_idx
    c1 = c0 + len(df.columns) - 1
    ws.add_table(r0, c0, r1, c1, {
        "style": "Table Style Medium 9",
        "columns": [{"header": str(c)} for c in df.columns]
    })
    ws.freeze_panes(r0 + 1, c0)
    widths_chars = _estimate_col_widths_chars(df)
    for j, w in enumerate(widths_chars):
        ws.set_column(c0 + j, c0 + j, w)
    return r0, c0, r1, c1

def _auto_excel_sheet(writer, df: pd.DataFrame, sheet_name: str, title_line: str, put_logo: bool = True):
    engine = writer.engine
    safe_name = (sheet_name or "Sheet1")[:31]
    df_x = df.copy()
    if engine == "xlsxwriter":
        wb = writer.book
        ws = wb.add_worksheet(safe_name)
        writer.sheets[safe_name] = ws
        cur_row = 0
        if put_logo:
            cur_row = _insert_wide_logo(ws, df_x, start_row=cur_row, start_col=0)
        ncols = max(1, len(df_x.columns))
        title_fmt = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "font_size": 16})
        ws.merge_range(cur_row, 0, cur_row, ncols-1, title_line, title_fmt)
        ws.set_row(cur_row, 28)
        cur_row += 1
        ws.set_row(cur_row, 16)
        cur_row += 1
        _write_excel_table(ws, wb, df_x, start_row=cur_row, start_col=0)
        ws.set_zoom(115)
        ws.set_margins(left=0.3, right=0.3, top=0.5, bottom=0.5)
    else:
        df_x.to_excel(writer, index=False, sheet_name=safe_name)

def make_excel_bytes(df: pd.DataFrame, sheet_name: str, title_line: str, put_logo: bool = True) -> Optional[bytes]:
    engine = _pick_excel_engine()
    if engine is None: return None
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine=engine) as writer:
        _auto_excel_sheet(writer, df, sheet_name, title_line, put_logo=put_logo)
    buf.seek(0)
    return buf.getvalue()

def make_excel_combined_two_sheets(dfs: Dict[str, pd.DataFrame], titles: Dict[str, str], put_logo: bool = True) -> Optional[bytes]:
    engine = _pick_excel_engine()
    if engine is None: return None
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine=engine) as writer:
        for sheet, df in dfs.items():
            _auto_excel_sheet(writer, df, sheet, titles.get(sheet, sheet), put_logo=put_logo)
    buf.seek(0)
    return buf.getvalue()

def make_excel_single_sheet_stacked(dfs: Dict[str, pd.DataFrame], title_line: str, sheet_name="تقرير_موحد", put_logo: bool = True) -> Optional[bytes]:
    engine = _pick_excel_engine()
    if engine is None: return None
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine=engine) as writer:
        if writer.engine == "xlsxwriter":
            wb = writer.book
            ws = wb.add_worksheet(sheet_name[:31])
            writer.sheets[sheet_name[:31]] = ws
            cur_row = 0
            if put_logo:
                widest_df = max(dfs.values(), key=lambda d: len(d.columns))
                cur_row = _insert_wide_logo(ws, widest_df, start_row=cur_row, start_col=0)
            max_cols = max((len(df.columns) for df in dfs.values()), default=1)
            big_title_fmt = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "font_size": 16})
            ws.merge_range(cur_row, 0, cur_row, max_cols-1, title_line, big_title_fmt)
            ws.set_row(cur_row, 28)
            cur_row += 2
            for section_title, df in dfs.items():
                title_fmt = wb.add_format({"bold": True, "align": "right", "font_size": 12})
                ws.merge_range(cur_row, 0, cur_row, len(df.columns)-1, section_title, title_fmt)
                cur_row += 2
                _write_excel_table(ws, wb, df, start_row=cur_row, start_col=0)
                cur_row += len(df) + 3
            ws.set_zoom(115)
        else:
            out = []
            for sec, df in dfs.items():
                title_row = pd.DataFrame([[sec] + [""] * (len(df.columns) - 1)], columns=df.columns)
                out += [title_row, df, pd.DataFrame([[""] * len(df.columns)], columns=df.columns)]
            pd.concat(out, ignore_index=True).to_excel(writer, index=False, sheet_name=sheet_name[:31])
    buf.seek(0)
    return buf.getvalue()

def make_csv_utf8(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


# =========================================================
# Drive helpers
# =========================================================
def _drive_share_to_direct_download(share_url: str) -> Optional[str]:
    if not share_url: return None
    m = re.search(r'/d/([a-zA-Z0-9_-]{10,})', share_url)
    if m:
        return f"https://drive.google.com/uc?export=download&id={m.group(1)}"
    m2 = re.search(r'[?&]id=([a-zA-Z0-9_-]{10,})', share_url)
    if m2:
        return f"https://drive.google.com/uc?export=download&id={m2.group(1)}"
    return share_url

def _create_zip_from_links(df: pd.DataFrame, link_col: str) -> Tuple[Optional[bytes], List[Tuple[int, str]]]:
    if df is None or df.empty or link_col not in df.columns:
        return None, [(-1, "لا توجد روابط")]
    buf = BytesIO()
    errors = []
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for idx, row in df.iterrows():
            share_url = str(row.get(link_col, "") or "")
            direct = _drive_share_to_direct_download(share_url)
            if not direct:
                errors.append((idx, "رابط غير صالح"))
                continue
            try:
                resp = requests.get(direct, stream=True, timeout=30)
                resp.raise_for_status()
                fname = None
                cd = resp.headers.get("content-disposition", "") or ""
                m = re.search(r"filename\*=([^']*)''(.+)", cd)
                if m:
                    try: fname = urllib.parse.unquote(m.group(2))
                    except Exception: fname = m.group(2)
                if not fname:
                    m2 = re.search(r'filename="?([^";]+)"?', cd)
                    if m2:
                        raw2 = m2.group(1)
                        try: fname = raw2.encode("latin-1").decode("utf-8")
                        except Exception:
                            try: fname = urllib.parse.unquote(raw2)
                            except Exception: fname = raw2
                if not fname:
                    tail = direct.split("/")[-1].split("?")[0]
                    if tail:
                        try: fname = urllib.parse.unquote(tail)
                        except Exception: fname = tail
                    else:
                        fid = re.search(r'/d/([a-zA-Z0-9_-]{10,})', share_url)
                        fname = f"file_{idx}_{fid.group(1) if fid else idx}"
                if not isinstance(fname, str): fname = str(fname)
                zf.writestr(fname, resp.content)
            except Exception as e:
                errors.append((idx, str(e)))
    buf.seek(0)
    if buf.getbuffer().nbytes == 0: return None, errors
    return buf.getvalue(), errors


# =========================================================
# PDF helpers
# =========================================================
def _normalize_name(s: str) -> str:
    return re.sub(r'[\s\u0640\u200c\u200d\u200e\u200f]+', '', str(s or ''))

def _fmt_integer(x) -> str:
    """Format any numeric value as integer with comma thousands separator.
    e.g. 1074373.00 → 1,074,373   |   0 → ""   |   None/NaN → ""
    """
    if x is None: return ""
    try:
        if pd.isna(x): return ""
    except Exception:
        pass
    sx = str(x).replace(",", "").strip()
    if not sx or sx.lower() in ("nan", "none", ""): return ""
    try:
        f = float(sx)
        if f == 0: return ""
        return f"{int(round(f)):,}"
    except Exception:
        return sx

def _fmt_integer_no_comma(x) -> str:
    """Format numeric value as integer WITHOUT comma separator — for رقم الشيك."""
    if x is None: return ""
    try:
        if pd.isna(x): return ""
    except Exception:
        pass
    sx = str(x).replace(",", "").strip()
    if not sx or sx.lower() in ("nan", "none", ""): return ""
    try:
        f = float(sx)
        if f == 0: return ""
        return str(int(round(f)))
    except Exception:
        return sx

_NO_COMMA_COLS = {"رقمالشيك"}   # normalized names (spaces/tatweel stripped)

def _format_numbers_for_display(df: pd.DataFrame, no_comma_cols: Optional[List[str]] = None) -> pd.DataFrame:
    """Format numbers in a DataFrame for display; also applies PDF pre-processing.
    Rules:
      - Date columns          → YYYY/MM/DD
      - نسبة الاعمال المنفذة  → kept as-is with % appended
      - رقم الشيك             → integer WITHOUT commas (e.g. 63105438)
      - ALL other numeric     → integer with comma thousands (e.g. 1,074,373)
      - zero / null           → empty string
    """
    out = _preprocess_df_for_pdf(df.copy())

    # Build the set of no-comma column names (caller can extend via no_comma_cols)
    extra = {_normalize_name(c) for c in (no_comma_cols or [])}
    no_comma = _NO_COMMA_COLS | extra

    for c in out.columns:
        if _is_date_col(c) or _is_percentage_col(c):
            continue
        if _normalize_name(c) in no_comma:
            out[c] = out[c].map(_fmt_integer_no_comma)
        else:
            out[c] = out[c].map(_fmt_integer)
    return out

def compose_pdf_title(company: str, project: str, data_type: str, dfrom, dto) -> str:
    return _compose_title(company, project, data_type, dfrom, dto)

def _shape(text: str) -> str:
    """Shape + bidi any string — always apply for best Arabic rendering.
    For mixed content (e.g. English bank name, numbers), passes through as-is
    but always applies Arabic reshaping when Arabic chars are present.
    """
    s = str(text) if text is not None else ""
    if not s or s.lower() in ("nan", "none"):
        return ""
    # If any Arabic chars present, reshape the whole string (handles mixed content too)
    if looks_arabic(s):
        return shape_arabic(s)
    # Pure English/numbers — return unchanged, will render fine with Latin-capable font
    return s

def _pdf_header_elements(title_line: str) -> Tuple[List, float]:
    font_name, arabic_ok = register_arabic_font()
    page = landscape(A4)
    left, right = 14, 14
    avail_w = page[0] - left - right

    title_style = ParagraphStyle(
        name="Title", fontName=font_name, fontSize=14, leading=17,
        alignment=1, textColor=colors.HexColor("#1b1b1b")
    )

    shaped_title = _shape(title_line)

    elements = []
    wlp = _wide_logo_path()
    if wlp and wlp.exists():
        try:
            if PILImage:
                w_px, h_px = _image_size(wlp)
                ratio = h_px / float(w_px) if w_px else 0.2
                img_h = max(22, avail_w * ratio * 0.55)
            else:
                img_h = 36
            logo_img = RLImage(str(wlp), hAlign="CENTER")
            logo_img.drawWidth = avail_w
            logo_img.drawHeight = img_h
            elements.append(logo_img)
            elements.append(Spacer(1, 8))
        except Exception:
            pass

    elements.append(Paragraph(shaped_title, title_style))
    elements.append(Spacer(1, 8))
    return elements, avail_w

def _pdf_table(
    df: pd.DataFrame,
    title: str = "",
    max_col_width: int = 120,
    font_size: float = 8.0,
    avail_width: Optional[float] = None,
) -> list:
    font_name, _ = register_arabic_font()

    hdr_style = ParagraphStyle(
        name="Hdr", fontName=font_name, fontSize=font_size + 0.6,
        textColor=colors.whitesmoke, alignment=1, leading=font_size + 1.8
    )
    cell_rtl = ParagraphStyle(
        name="CellR", fontName=font_name, fontSize=font_size,
        leading=font_size + 1.5, alignment=2, textColor=colors.black
    )
    # For pure English/numeric values: use Helvetica (always available, full Latin)
    # so bank names, codes, numbers stored as English render correctly in the PDF.
    cell_ltr = ParagraphStyle(
        name="CellL", fontName="Helvetica", fontSize=font_size,
        leading=font_size + 1.5, alignment=1, textColor=colors.black
    )
    link_style = ParagraphStyle(
        name="Link", fontName=font_name, fontSize=font_size,
        alignment=2, textColor=colors.HexColor("#1a56db"), underline=True
    )

    blocks = []
    if title:
        tstyle = ParagraphStyle(
            name="Sec", fontName=font_name, fontSize=font_size + 2,
            alignment=2, textColor=colors.HexColor("#1E3A8A")
        )
        blocks += [Paragraph(_shape(title), tstyle), Spacer(1, 4)]

    # Headers
    headers = [Paragraph(_shape(str(c)), hdr_style) for c in df.columns]
    rows = [headers]

    # Body rows
    for _, r in df.iterrows():
        cells = []
        for c in df.columns:
            raw = r[c]
            sval = "" if (raw is None or (isinstance(raw, float) and pd.isna(raw))) else str(raw)
            col_str = str(c)

            # % sign for نسبة الاعمال المنفذة (belt-and-suspenders)
            if _is_percentage_col(col_str) and sval and not sval.strip().endswith("%"):
                sval = f"{sval}%"

            if sval.startswith(("http://", "https://")) or ("رابط" in col_str and sval):
                html = f'<link href="{sval}">{_shape("فتح الرابط")}</link>'
                cells.append(Paragraph(html, link_style))
            elif looks_arabic(sval) or not sval.strip():
                # Arabic text or empty — use Arabic font, right-aligned
                cells.append(Paragraph(_shape(sval), cell_rtl))
            else:
                # Pure English / numbers / mixed Latin — use Helvetica so glyphs render
                cells.append(Paragraph(sval, cell_ltr))
        rows.append(cells)

    # Column widths
    col_widths = []
    for c in df.columns:
        max_len = max(len(str(c)), df[c].astype(str).map(len).max())
        col_widths.append(min(max_len * 6.4, max_col_width))
    if avail_width:
        total = sum(col_widths)
        if total > avail_width:
            factor = avail_width / total
            col_widths = [w * factor for w in col_widths]

    table = Table(rows, repeatRows=1, colWidths=col_widths)
    table.setStyle(TableStyle([
        ("FONTNAME",     (0, 0), (-1, -1), font_name),
        ("FONTSIZE",     (0, 0), (-1, -1), font_size),
        ("BACKGROUND",   (0, 0), (-1,  0), colors.HexColor("#1E3A8A")),
        ("TEXTCOLOR",    (0, 0), (-1,  0), colors.whitesmoke),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING",   (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 2),
        ("GRID",         (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1")),
        ("ROWBACKGROUNDS",(0,1), (-1, -1), [colors.white, colors.HexColor("#f7fafc")]),
    ]))
    blocks.append(table)
    return blocks

def _choose_pdf_font(df: pd.DataFrame) -> Tuple[int, float]:
    n = len(df.columns)
    if n >= 12: return 110, 7.0
    if n >= 9:  return 125, 7.5
    return 150, 8.0

def make_pdf_bytes(df: pd.DataFrame, title_line: str) -> bytes:
    page = landscape(A4)
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=page,
                            rightMargin=14, leftMargin=14, topMargin=18, bottomMargin=14)
    elements, avail_w = _pdf_header_elements(title_line)
    max_col_width, base_font = _choose_pdf_font(df)
    elements += _pdf_table(df, max_col_width=max_col_width, font_size=base_font, avail_width=avail_w)
    doc.build(elements)
    buf.seek(0)
    return buf.getvalue()

def make_pdf_combined(summary_df: pd.DataFrame, flow_df: pd.DataFrame, title_line: str) -> bytes:
    page = landscape(A4)
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=page,
                            rightMargin=14, leftMargin=14, topMargin=18, bottomMargin=14)
    header_elements, avail_w = _pdf_header_elements(title_line)
    elements = list(header_elements)
    max_w_s, f_s = _choose_pdf_font(summary_df)
    elements += _pdf_table(summary_df, title="ملخص المشروع", max_col_width=max_w_s,
                           font_size=f_s, avail_width=avail_w)
    elements.append(PageBreak())
    max_w_f, f_f = _choose_pdf_font(flow_df)
    elements += _pdf_table(flow_df, title="دفتر التدفق", max_col_width=max_w_f,
                           font_size=f_f, avail_width=avail_w)
    doc.build(elements)
    buf.seek(0)
    return buf.getvalue()


# =========================================================
# Summary render helpers
# =========================================================
def fin_panel_two_tables(left_items: List[Tuple[str, str]], right_items: List[Tuple[str, str]]):
    def _table_html(items):
        rows = "".join(
            f'<tr><td class="value">{v}</td><td class="label">{l}</td></tr>'
            for l, v in items
        )
        return f'<table class="fin-table">{rows}</table>'
    html = (
        f'<div class="fin-panel card">'
        f'<div class="soft">{_table_html(right_items)}</div>'
        f'<div class="soft">{_table_html(left_items)}</div>'
        f'</div>'
    )
    st.markdown(html, unsafe_allow_html=True)

def _apply_date_filter(df: pd.DataFrame, dfrom, dto) -> pd.DataFrame:
    if df is None or df.empty or (not dfrom and not dto): return df
    date_cols = [c for c in df.columns if any(k in str(c) for k in ["تاريخ", "إصدار", "date", " التعاقد"])]
    if not date_cols: return df
    out = df.copy()
    for col in date_cols:
        try:
            dseries = pd.to_datetime(out[col], errors="coerce").dt.date
            if dfrom: out = out[dseries >= dfrom]
            if dto:   out = out[dseries <= dto]
        except Exception:
            pass
    return out

def _fmt_value(v) -> str:
    try:
        if isinstance(v, str) and v.strip().endswith("%"): return v
        f = float(str(v).replace(",", ""))
        return f"{f:,.2f}"
    except Exception:
        return "" if (v is None or (isinstance(v, float) and pd.isna(v))) else str(v)

def _row_to_pairs_from_data(row: pd.Series) -> List[Tuple[str, str]]:
    ignore_substrings = {"id", "ID", "companyid", "contractid"}
    pairs = []
    for col, val in row.items():
        if any(k in str(col).lower() for k in ignore_substrings): continue
        sval = _fmt_value(val)
        if sval == "" or sval.lower() == "nan": continue
        pairs.append((str(col), sval))
    return pairs

def _split_pairs_two_columns(pairs: List[Tuple[str, str]]) -> Tuple[List[Tuple[str, str]], List[Tuple[str, str]]]:
    n = len(pairs); mid = (n + 1) // 2
    return pairs[mid:], pairs[:mid]


# =========================================================
# Main App
# =========================================================
def main() -> None:
    conn = get_db_connection()
    if conn is None:
        st.error("فشل الاتصال بقاعدة البيانات. يرجى مراجعة بيانات الاتصال والتأكد من تشغيل الخادم.")
        return

    with st.sidebar:
        st.title("عوامل التصفية")
        company_name = create_company_dropdown(conn)
        project_name = create_project_dropdown(conn, company_name)
        type_label, type_key = create_type_dropdown()

        st.markdown("---")
        search_clicked = st.button("🔍 بحث", key="sidebar_search_btn", use_container_width=True)

        import streamlit.components.v1 as components
        if search_clicked:
            components.html(
                """
                <script>
                (function() {
                    function collapse() {
                        var w = window.parent;
                        var sidebar = w.document.querySelector('[data-testid="stSidebar"]');
                        if (!sidebar) return;
                        if (sidebar.getAttribute('aria-expanded') === 'false') return;
                        var btn = w.document.querySelector('[data-testid="stSidebarCollapseButton"] button');
                        if (!btn) {
                            var wrap = w.document.querySelector('[data-testid="stSidebarCollapseButton"]');
                            if (wrap) btn = wrap.querySelector('button');
                        }
                        if (btn) { btn.click(); }
                    }
                    setTimeout(collapse, 200);
                })();
                </script>
                """,
                height=0,
                width=0,
            )

    if not company_name or not project_name or not type_key:
        st.info("برجاء اختيار الشركة والمشروع ونوع البيانات من الشريط الجانبي لعرض النتائج.")
        return

    # Global date filters
    g_date_from, g_date_to = None, None
    with st.container():
        st.markdown('<div class="date-box"><div class="date-row">', unsafe_allow_html=True)
        c1, c2 = st.columns([1, 1], gap="small")
        with c1: g_date_from = st.date_input("من تاريخ", value=None, key="g_from", format="YYYY-MM-DD")
        with c2: g_date_to   = st.date_input("إلى تاريخ", value=None, key="g_to",   format="YYYY-MM-DD")
        st.markdown('</div></div>', unsafe_allow_html=True)

    # =======================
    # Financial Report Mode
    # =======================
    if type_key == "financial_report":
        df_summary = fetch_contract_summary_view(conn, company_name, project_name)
        if df_summary.empty:
            st.warning("لم يتم العثور على ملخص العقد لهذا المشروع.")
            return
        row = df_summary.iloc[0]

        st.markdown(
            f"""
            <div class="fin-head">
                <div class="line">
                    <strong>الشركة:</strong> {company_name or '—'}
                    &nbsp;&nbsp;|&nbsp;&nbsp;
                    <strong>المشروع:</strong> {project_name or '—'}
                    &nbsp;&nbsp;|&nbsp;&nbsp;
                    <strong>تاريخ التعاقد:</strong> {_format_date_arabic(row.get("تاريخ التعاقد", "—"))}
                </div>
                <span class="badge">تقرير مالي</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        summary_pairs = _row_to_pairs_from_data(row)
        if summary_pairs:
            left_items, right_items = _split_pairs_two_columns(summary_pairs)
            st.markdown('<h3 class="hsec">ملخص المشروع</h3>', unsafe_allow_html=True)
            fin_panel_two_tables(left_items=left_items, right_items=right_items)

        title_summary = compose_pdf_title(company_name, project_name, "ملخص", g_date_from, g_date_to)
        title_flow    = compose_pdf_title(company_name, project_name, "دفتر التدفق", g_date_from, g_date_to)
        title_all     = compose_pdf_title(company_name, project_name, "ملخص + دفتر التدفق", g_date_from, g_date_to)

        xlsx_sum = make_excel_bytes(df_summary, sheet_name="ملخص", title_line=title_summary, put_logo=True)
        if xlsx_sum:
            st.download_button("تنزيل الملخص كـ Excel", xlsx_sum,
                               file_name=_safe_filename(f"ملخص_{company_name}_{project_name}.xlsx"),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        pdf_sum = make_pdf_bytes(_format_numbers_for_display(df_summary), title_line=title_summary)
        st.download_button("تنزيل الملخص كـ PDF", pdf_sum,
                           file_name=_safe_filename(f"ملخص_{company_name}_{project_name}.pdf"),
                           mime="application/pdf")

        st.markdown('<hr class="hr-accent"/>', unsafe_allow_html=True)

        st.markdown('<h3 class="hsec">دفتر التدفق (v_financial_flow)</h3>', unsafe_allow_html=True)
        df_flow = fetch_financial_flow_view(conn, company_name, project_name, g_date_from, g_date_to)
        if df_flow.empty:
            st.info("لا توجد حركات مطابقة ضمن النطاق المحدد.")
            return

        col_search, term = create_column_search(df_flow)
        if col_search and term:
            df_flow = df_flow[df_flow[col_search].astype(str).str.contains(str(term), case=False, na=False)]
            if df_flow.empty:
                st.info("لا توجد نتائج بعد تطبيق البحث.")
                return

        df_flow_display = df_flow.drop(columns=["companyid", "contractid", "delta"], errors="ignore")

        flow_col_config = {}
        for col in df_flow_display.columns:
            if "رابط" in str(col):
                flow_col_config[col] = st.column_config.LinkColumn(
                    label=col,
                    display_text="🔗 فتح الرابط",
                )

        st.markdown('<div class="card soft">', unsafe_allow_html=True)
        st.dataframe(df_flow_display, column_config=flow_col_config, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

        xlsx_flow = make_excel_bytes(df_flow_display, sheet_name="دفتر_التدفق", title_line=title_flow, put_logo=True)
        if xlsx_flow:
            st.download_button("تنزيل الدفتر كـ Excel", xlsx_flow,
                               file_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}.xlsx"),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("تنزيل الدفتر كـ CSV", make_csv_utf8(df_flow_display),
                           file_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}.csv"),
                           mime="text/csv")
        pdf_flow = make_pdf_bytes(
            _format_numbers_for_display(df_flow_display, no_comma_cols=["رقم الشيك"]),
            title_line=title_flow
        )
        st.download_button("تنزيل الدفتر كـ PDF", pdf_flow,
                           file_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}.pdf"),
                           mime="application/pdf")

        link_cols = [c for c in df_flow_display.columns if "رابط" in str(c)]
        if link_cols:
            for lc in link_cols:
                btn_label = (
                    "⬇️ تنزيل كل المستخلصات (ZIP)" if "مستخلص" in str(lc)
                    else "⬇️ تنزيل كل الشيكات (ZIP)" if "شيك" in str(lc)
                    else f"⬇️ تنزيل {lc} (ZIP)"
                )
                with st.spinner(f"جارٍ إنشاء أرشيف {lc} ..."):
                    zip_bytes, errors = _create_zip_from_links(df_flow_display, lc)
                if not zip_bytes:
                    st.error(f"فشل إنشاء ملف ZIP لـ {lc}.")
                    if errors: st.warning(f"أخطاء: {len(errors)} حالات.")
                else:
                    st.download_button(
                        btn_label, zip_bytes,
                        file_name=_safe_filename(f"{lc}_{company_name}_{project_name}.zip"),
                        mime="application/zip",
                        key=f"zip_{lc}_{company_name}_{project_name}"
                    )
                    if errors: st.warning(f"بعض الملفات لم تُحمّل ({len(errors)}).")

        st.markdown("### تنزيل تقرير موحّد")
        excel_two = make_excel_combined_two_sheets(
            {"ملخص": df_summary, "دفتر_التدفق": df_flow_display},
            titles={"ملخص": title_summary, "دفتر_التدفق": title_flow}, put_logo=True
        )
        if excel_two:
            st.download_button("Excel موحّد (ورقتان)", excel_two,
                               file_name=_safe_filename(f"تقرير_مالي_{company_name}_{project_name}_ورقتين.xlsx"),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        excel_one = make_excel_single_sheet_stacked(
            {"ملخص": df_summary, "دفتر_التدفق": df_flow_display},
            title_line=title_all, sheet_name="تقرير_موحد", put_logo=True
        )
        if excel_one:
            st.download_button("Excel موحّد (ورقة واحدة)", excel_one,
                               file_name=_safe_filename(f"تقرير_مالي_{company_name}_{project_name}_ورقة_واحدة.xlsx"),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        pdf_all = make_pdf_combined(
            _format_numbers_for_display(df_summary),
            _format_numbers_for_display(df_flow_display, no_comma_cols=["رقم الشيك"]),
            title_line=title_all
        )
        st.download_button("PDF موحّد (ملخص + دفتر)", pdf_all,
                           file_name=_safe_filename(f"تقرير_مالي_{company_name}_{project_name}.pdf"),
                           mime="application/pdf")
        return

    # =======================
    # Other data types
    # =======================
    df = fetch_data(conn, company_name, project_name, type_key)
    if df.empty:
        st.warning("لا توجد بيانات مطابقة للاختيارات المحددة.")
        return

    df = _apply_date_filter(df, g_date_from, g_date_to)
    search_column, search_term = create_column_search(df)
    if search_column and search_term:
        df = df[df[search_column].astype(str).str.contains(str(search_term), case=False, na=False)]
        if df.empty:
            st.info("لا توجد نتائج بعد تطبيق معيار البحث.")
            return

    column_config = {
        col: st.column_config.LinkColumn(label=col, display_text="فتح الرابط")
        for col in df.columns if "رابط" in str(col)
    }

    st.markdown('<h3 class="hsec">البيانات</h3>', unsafe_allow_html=True)
    st.markdown('<div class="card soft">', unsafe_allow_html=True)
    st.dataframe(df, column_config=column_config, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    title_generic = compose_pdf_title(company_name, project_name, type_label, g_date_from, g_date_to)

    xlsx_bytes = make_excel_bytes(df, sheet_name="البيانات", title_line=title_generic, put_logo=True)
    if xlsx_bytes:
        st.download_button("تنزيل كـ Excel (XLSX)", xlsx_bytes,
                           file_name=_safe_filename(f"{type_key}_{company_name}_{project_name}.xlsx"),
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("تنزيل كـ CSV (UTF-8)", make_csv_utf8(df),
                       file_name=_safe_filename(f"{type_key}_{company_name}_{project_name}.csv"),
                       mime="text/csv")
    st.download_button("تنزيل كـ PDF", make_pdf_bytes(_format_numbers_for_display(df), title_line=title_generic),
                       file_name=_safe_filename(f"{type_key}_{company_name}_{project_name}.pdf"),
                       mime="application/pdf")

    link_cols = [c for c in df.columns if "رابط" in str(c)]
    if link_cols:
        with st.spinner("جارٍ إنشاء أرشيف ZIP ..."):
            zip_bytes, errors = _create_zip_from_links(df, link_cols[0])
        if not zip_bytes:
            st.error("فشل إنشاء ملف ZIP.")
        else:
            lbl = "⬇️ تنزيل كل المستخلصات (ZIP)" if type_label and "مستخلص" in str(type_label) else "⬇️ تنزيل الكل (ZIP)"
            st.download_button(lbl, zip_bytes,
                               file_name=_safe_filename(f"{type_label or 'الملفات'}_{company_name}_{project_name}.zip"),
                               mime="application/zip")
            if errors: st.warning(f"بعض الملفات لم تُحمّل ({len(errors)}).")


if __name__ == "__main__":
    main()