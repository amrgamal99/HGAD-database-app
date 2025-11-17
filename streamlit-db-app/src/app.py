# app.py
import os
import re
import base64
from io import BytesIO
from pathlib import Path
from typing import Optional, Tuple, Dict, List

import pandas as pd
import streamlit as st

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

# === DB / filters (your existing modules) ===
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
    Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
]

_AR_RE = re.compile(r"[\u0600-\u06FF]")  # Arabic unicode range


# =========================================================
# Small utils
# =========================================================
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
                return im.size  # (w, h)
        except Exception:
            pass
    return (600, 120)


def _site_logo_path() -> Optional[Path]:
    return _first_existing(LOGO_CANDIDATES)


def _wide_logo_path() -> Optional[Path]:
    return _first_existing(WIDE_LOGO_CANDIDATES)


def _first_existing_font_path() -> Optional[str]:
    p = _first_existing(AR_FONT_CANDIDATES)
    return str(p) if p else None


def register_arabic_font() -> Tuple[str, bool]:
    p = _first_existing_font_path()
    if p:
        name = os.path.splitext(os.path.basename(p))[0]
        try:
            pdfmetrics.registerFont(TTFont(name, p))
            return name, True
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
# Streamlit Config + Polished CSS
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

/* RTL base */
html, body{
  direction: rtl !important; text-align: right !important;
  font-family: "Cairo","Noto Kufi Arabic","Segoe UI",Tahoma,sans-serif !important;
  color:var(--text) !important; background:var(--bg) !important;
}

/* Sidebar always open + style */
[data-testid="stSidebar"]{
  transform:none !important; visibility:visible !important;
  width:340px !important; min-width:340px !important;
  background: linear-gradient(180deg, #0b1220, #0a1020);
  border-inline-start: 1px solid var(--line);
}
[data-testid="collapsedControl"],button[kind="header"],
button[title="Expand sidebar"],button[title="Collapse sidebar"],
[data-testid="stSidebarCollapseButton"]{ display:none !important; }

/* Fancy separator */
.hr-accent{ height:2px; border:0; background:linear-gradient(90deg, transparent, var(--accent), transparent); margin: 8px 0 14px; }

/* Cards */
.card{ background:var(--panel); border:1px solid var(--line); border-radius:16px; padding:14px; box-shadow:0 6px 24px rgba(3,10,30,.25); }
.card.soft{ background:var(--panel-2); }

/* Header banner */
.fin-head{
  display:flex; justify-content:space-between; align-items:center;
  border: 1px dashed rgba(37,99,235,.35); border-radius:16px;
  padding: 16px 18px; margin:8px 0 14px; background:linear-gradient(180deg,#0b1220,#0e1424);
}
.fin-head .line{ font-size:22px; font-weight:900; color:var(--text); }
.badge{ background:var(--accent); color:#fff; padding:6px 12px; border-radius:999px; font-weight:700; }

/* Date area */
.date-box{ border:1px solid var(--line); border-radius:16px; padding:12px; background:var(--panel-2); margin-bottom:12px; }
.date-row{ display:flex; gap:12px; flex-wrap:wrap; align-items:center; }
[data-testid="stDateInput"] input{
  background:#0f172a !important; color:var(--text) !important;
  border:1px solid var(--line) !important; border-radius:10px !important;
  text-align:center !important; height:44px !important; min-width:190px !important;
}
[data-testid="stDateInput"] label{ color:var(--muted) !important; font-weight:700; }

/* DataFrame look */
[data-testid="stDataFrame"] thead tr th{
  position: sticky; top: 0; z-index: 2;
  background: #132036; color: #e7eefc; font-weight:800; font-size:15px;
  border-bottom: 1px solid var(--line);
}
[data-testid="stDataFrame"] div[role="row"]{ font-size:14.5px; }
[data-testid="stDataFrame"] div[role="row"]:nth-child(even){ background: rgba(255,255,255,.03); }

/* Section title */
.hsec{ color:#e7eefc; font-weight:900; margin:6px 0 10px; font-size: 22px; }

/* Summary two-column panel */
.fin-panel{ display:grid; grid-template-columns: 1fr 1fr; gap:20px; margin-top:10px; }
.fin-table{ width:100%; border-collapse:collapse; table-layout:fixed; border-radius:14px; overflow:hidden; }
.fin-table th, .fin-table td{ border:1px solid var(--line); padding:12px; font-size:14.5px; white-space:normal; word-wrap:break-word; }
.fin-table tr:hover td{ background:#111a2d; transition: background .2s ease; }
.fin-table td.value{ background:#0f1a30; font-weight:800; text-align:center; width:34%; }
.fin-table td.label{ background:#0d1628; font-weight:700; text-align:right; width:66%; }
</style>
""",
    unsafe_allow_html=True,
)

# =========================================================
# Header (inline base64 small logo)
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
# Excel helpers (logo spans full width + tight spacing + Excel table)
# =========================================================
def _pick_excel_engine() -> Optional[str]:
    try:
        import xlsxwriter  # noqa: F401
        return "xlsxwriter"
    except Exception:
        pass
    try:
        import openpyxl  # noqa: F401
        return "openpyxl"
    except Exception:
        return None


def _estimate_col_widths_chars(df: pd.DataFrame) -> List[float]:
    """Estimate chars width per column (used for logo scaling & set_column)."""
    widths = []
    for col in df.columns:
        series = df[col]
        max_len = max([len(str(col))] + [len(str(v)) for v in series.values])
        widths.append(min(max_len + 4, 60))
    return widths


def _chars_to_pixels(chars: float) -> float:
    """Approx Excel mapping (≈7.2 px per char)."""
    return chars * 7.2


def _compose_title(company: str, project: str, data_type: str, dfrom, dto) -> str:
    # RTL-friendly reversed arrow
    parts = []
    if company: parts.append(f"الشركة: {company}")
    if project: parts.append(f"المشروع: {project}")
    if data_type: parts.append(f"النوع: {data_type}")
    if dfrom or dto:
        parts.append(f"الفترة: {dfrom or '—'} ← {dto or '—'}")
    return " | ".join(parts)


def _insert_wide_logo(ws, df: pd.DataFrame, start_row: int, start_col: int = 0) -> int:
    """
    Insert the wide logo scaled to span the full table width (first→last column).
    Return the next row index for the title. Only one tight title row after.
    """
    wlp = _wide_logo_path()
    if not wlp:
        return start_row

    widths_chars = _estimate_col_widths_chars(df)
    total_width_px = _chars_to_pixels(sum(widths_chars))

    try:
        img_w_px, img_h_px = _image_size(wlp)
        if img_w_px <= 0: img_w_px = 1000
        x_scale = max(0.1, total_width_px / float(img_w_px))
        y_scale = x_scale  # keep aspect ratio
        ws.insert_image(
            start_row, start_col, str(wlp),
            {"x_scale": x_scale, "y_scale": y_scale, "object_position": 2}
        )
        scaled_h_px = img_h_px * y_scale
        ws.set_row(start_row, int(scaled_h_px * 0.75))  # px→pt approx
        return start_row + 1
    except Exception:
        ws.set_row(start_row, 80)
        ws.insert_image(start_row, start_col, str(wlp), {"x_scale": 0.5, "y_scale": 0.5, "object_position": 2})
        return start_row + 1


def _write_excel_table(ws, workbook, df: pd.DataFrame, start_row: int, start_col: int) -> Tuple[int, int, int, int]:
    """Write df as formatted Excel Table with links labeled 'فتح الرابط'."""
    hdr_fmt = workbook.add_format({"align": "right", "bold": True})
    fmt_text = workbook.add_format({"align": "right"})
    fmt_date = workbook.add_format({"align": "right", "num_format": "yyyy-mm-dd"})
    fmt_num  = workbook.add_format({"align": "right", "num_format": "#,##0.00"})
    fmt_link = workbook.add_format({"font_color": "blue", "underline": 1, "align": "right"})

    r0, c0 = start_row, start_col

    # headers
    for j, col in enumerate(df.columns):
        ws.write(r0, c0 + j, col, hdr_fmt)

    # body
    for i in range(len(df)):
        for j, col in enumerate(df.columns):
            val = df.iloc[i, j]
            colname = str(col)
            sval = "" if pd.isna(val) else str(val)
            if sval.startswith(("http://", "https://")) or ("رابط" in colname and sval):
                ws.write_url(r0 + 1 + i, c0 + j, sval, fmt_link, string="فتح الرابط")
            else:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    if pd.notna(val): ws.write_datetime(r0 + 1 + i, c0 + j, pd.to_datetime(val), fmt_date)
                    else: ws.write_blank(r0 + 1 + i, c0 + j, None, fmt_text)
                elif pd.api.types.is_numeric_dtype(df[col]):
                    if pd.notna(val): ws.write_number(r0 + 1 + i, c0 + j, float(val), fmt_num)
                    else: ws.write_blank(r0 + 1 + i, c0 + j, None, fmt_text)
                else:
                    ws.write(r0 + 1 + i, c0 + j, sval, fmt_text)

    r1 = r0 + len(df)
    c1 = c0 + len(df.columns) - 1

    ws.add_table(r0, c0, r1, c1, {
        "style": "Table Style Medium 9",
        "columns": [{"header": str(c)} for c in df.columns]
    })
    ws.freeze_panes(r0 + 1, c0)

    # column widths
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
            # Logo full width, then title next row
            cur_row = _insert_wide_logo(ws, df_x, start_row=cur_row, start_col=0)

        # Title merged across all columns (right after logo)
        ncols = max(1, len(df_x.columns))
        title_fmt = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "font_size": 16})
        ws.merge_range(cur_row, 0, cur_row, ncols-1, title_line, title_fmt)
        ws.set_row(cur_row, 28)
        cur_row += 1

        # One blank row only
        ws.set_row(cur_row, 16)
        cur_row += 1

        # Table
        _write_excel_table(ws, wb, df_x, start_row=cur_row, start_col=0)
        ws.set_zoom(115)
        ws.set_margins(left=0.3, right=0.3, top=0.5, bottom=0.5)
    else:
        df_x.to_excel(writer, index=False, sheet_name=safe_name)


def make_excel_bytes(df: pd.DataFrame, sheet_name: str, title_line: str, put_logo: bool = True) -> Optional[bytes]:
    engine = _pick_excel_engine()
    if engine is None:
        return None
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine=engine) as writer:
        _auto_excel_sheet(writer, df, sheet_name, title_line, put_logo=put_logo)
    buf.seek(0)
    return buf.getvalue()


def make_excel_combined_two_sheets(dfs: Dict[str, pd.DataFrame], titles: Dict[str, str], put_logo: bool = True) -> Optional[bytes]:
    engine = _pick_excel_engine()
    if engine is None:
        return None
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine=engine) as writer:
        for sheet, df in dfs.items():
            _auto_excel_sheet(writer, df, sheet, titles.get(sheet, sheet), put_logo=put_logo)
    buf.seek(0)
    return buf.getvalue()


def make_excel_single_sheet_stacked(dfs: Dict[str, pd.DataFrame], title_line: str, sheet_name="تقرير_موحد", put_logo: bool = True) -> Optional[bytes]:
    engine = _pick_excel_engine()
    if engine is None:
        return None
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
            cur_row += 2  # one blank row

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
            big = pd.concat(out, ignore_index=True)
            big.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    buf.seek(0)
    return buf.getvalue()


def make_csv_utf8(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


# =========================================================
# PDF helpers (clean + zebra + anchors + dynamic title)
# ******* FINAL FIX FOR 'رقم الشيك' (no commas) **********
# =========================================================
def _normalize_name(s: str) -> str:
    """Normalize Arabic column names: remove spaces, tatweel and zero-width marks."""
    return re.sub(r'[\s\u0640\u200c\u200d\u200e\u200f]+', '', str(s or ''))

def _plain_number_no_commas(x) -> str:
    """Render number as plain string without thousands separators; trim trailing .00."""
    if pd.isna(x):
        return ""
    sx = str(x).replace(",", "").strip()
    try:
        f = float(sx)
        if float(int(f)) == f:
            return str(int(f))
        s = f"{f}"
        if "." in s:
            s = s.rstrip("0").rstrip(".")
        return s
    except Exception:
        return str(x)

def _format_numbers_for_display(df: pd.DataFrame, no_comma_cols: Optional[List[str]] = None) -> pd.DataFrame:
    """
    Format numbers for PDF display.
    * Any column explicitly listed in no_comma_cols
    * OR any column whose normalized name contains 'شيك'
      is rendered as plain text (no thousands separators).
    """
    out = df.copy()
    requested = {_normalize_name(c) for c in (no_comma_cols or [])}

    for c in out.columns:
        c_norm = _normalize_name(c)
        force_plain = (c_norm in requested) or ("شيك" in c_norm)

        if force_plain:
            out[c] = out[c].map(_plain_number_no_commas)
            continue

        if pd.api.types.is_numeric_dtype(out[c]):
            out[c] = out[c].map(lambda x: "" if pd.isna(x) else f"{float(x):,.2f}")
        else:
            def _fmt_cell(v):
                s = str(v)
                try:
                    if s.strip().endswith("%"):
                        return s
                    fv = float(s.replace(",", ""))
                    return f"{fv:,.2f}"
                except Exception:
                    return s
            out[c] = out[c].map(_fmt_cell)
    return out


def compose_pdf_title(company: str, project: str, data_type: str, dfrom, dto) -> str:
    return _compose_title(company, project, data_type, dfrom, dto)


def _pdf_header_elements(title_line: str) -> Tuple[List, float]:
    font_name, arabic_ok = register_arabic_font()
    page = landscape(A4)
    left, right, top, bottom = 14, 14, 18, 14
    avail_w = page[0] - left - right

    title_style = ParagraphStyle(
        name="Title", fontName=font_name, fontSize=14, leading=17,
        alignment=1, textColor=colors.HexColor("#1b1b1b")
    )

    if arabic_ok:
        title_line = shape_arabic(title_line)

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

    elements.append(Paragraph(title_line, title_style))
    elements.append(Spacer(1, 8))
    return elements, avail_w


def _pdf_table(df: pd.DataFrame, title: str = "", max_col_width: int = 120, font_size: float = 8.0, avail_width: Optional[float] = None) -> list:
    font_name, _ = register_arabic_font()
    hdr_style = ParagraphStyle(name="Hdr", fontName=font_name, fontSize=font_size+0.6, textColor=colors.whitesmoke, alignment=1, leading=font_size+1.8)
    cell_rtl  = ParagraphStyle(name="CellR", fontName=font_name, fontSize=font_size, leading=font_size+1.5, alignment=2, textColor=colors.black)
    cell_ltr  = ParagraphStyle(name="CellL", fontName=font_name, fontSize=font_size, leading=font_size+1.5, alignment=0, textColor=colors.black)
    link_style = ParagraphStyle(name="Link", fontName=font_name, fontSize=font_size, alignment=2, textColor=colors.HexColor("#1a56db"), underline=True)

    blocks = []
    if title:
        tstyle = ParagraphStyle(name="Sec", fontName=font_name, fontSize=font_size+2, alignment=2, textColor=colors.HexColor("#1E3A8A"))
        blocks += [Paragraph(shape_arabic(title), tstyle), Spacer(1, 4)]

    headers = [Paragraph(shape_arabic(c) if looks_arabic(c) else str(c), hdr_style) for c in df.columns]
    rows = [headers]
    for _, r in df.iterrows():
        cells = []
        for c in df.columns:
            sval = "" if pd.isna(r[c]) else str(r[c])
            if sval.startswith(("http://", "https://")) or ("رابط" in str(c) and sval):
                html = f'<link href="{sval}">{shape_arabic("فتح الرابط")}</link>'
                cells.append(Paragraph(html, link_style))
            else:
                is_ar = looks_arabic(sval)
                cells.append(Paragraph(shape_arabic(sval) if is_ar else sval, cell_rtl if is_ar else cell_ltr))
        rows.append(cells)

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
        ("FONTNAME", (0,0), (-1,-1), font_name),
        ("FONTSIZE", (0,0), (-1,-1), font_size),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1E3A8A")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.whitesmoke),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ("GRID", (0,0), (-1,-1), 0.35, colors.HexColor("#cbd5e1")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f7fafc")]),
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
    left, right, top, bottom = 14, 14, 18, 14
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=page, rightMargin=right, leftMargin=left, topMargin=top, bottomMargin=bottom)

    elements, avail_w = _pdf_header_elements(title_line)
    max_col_width, base_font = _choose_pdf_font(df)
    elements += _pdf_table(df, max_col_width=max_col_width, font_size=base_font, avail_width=avail_w)
    doc.build(elements)
    buf.seek(0)
    return buf.getvalue()


def make_pdf_combined(summary_df: pd.DataFrame, flow_df: pd.DataFrame, title_line: str) -> bytes:
    page = landscape(A4)
    left, right, top, bottom = 14, 14, 18, 14
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=page, rightMargin=right, leftMargin=left, topMargin=top, bottomMargin=bottom)

    header_elements, avail_w = _pdf_header_elements(title_line)
    elements = list(header_elements)

    max_w_s, f_s = _choose_pdf_font(summary_df)
    elements += _pdf_table(summary_df, title="ملخص المشروع", max_col_width=max_w_s, font_size=f_s, avail_width=avail_w)
    elements.append(PageBreak())
    max_w_f, f_f = _choose_pdf_font(flow_df)
    elements += _pdf_table(flow_df, title="دفتر التدفق", max_col_width=max_w_f, font_size=f_f, avail_width=avail_w)

    doc.build(elements)
    buf.seek(0)
    return buf.getvalue()


# =========================================================
# Summary render helpers
# =========================================================
def fin_panel_two_tables(left_items: List[Tuple[str, str]], right_items: List[Tuple[str, str]]):
    def _table_html(items):
        rows = []
        for label, value in items:
            rows.append(f'<tr><td class="value">{value}</td><td class="label">{label}</td></tr>')
        return f'<table class="fin-table">{"".join(rows)}</table>'
    html = f'<div class="fin-panel card"><div class="soft">{_table_html(right_items)}</div><div class="soft">{_table_html(left_items)}</div></div>'
    st.markdown(html, unsafe_allow_html=True)


def _apply_date_filter(df: pd.DataFrame, dfrom, dto) -> pd.DataFrame:
    if df is None or df.empty or (not dfrom and not dto): return df
    date_cols = [c for c in df.columns if any(k in str(c) for k in ["تاريخ", "إصدار", "date", "تعاقد"])]
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
        f = float(str(v).replace(",", "")); return f"{f:,.2f}"
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
    n = len(pairs); mid = (n + 1)//2
    right = pairs[:mid]; left = pairs[mid:]
    return left, right


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

    if not company_name or not project_name or not type_key:
        st.info("برجاء اختيار الشركة والمشروع ونوع البيانات من الشريط الجانبي لعرض النتائج.")
        return

    # Global date filters (main area)
    g_date_from, g_date_to = None, None
    with st.container():
        st.markdown('<div class="date-box"><div class="date-row">', unsafe_allow_html=True)
        c1, c2 = st.columns([1, 1], gap="small")
        with c1: g_date_from = st.date_input("من تاريخ", value=None, key="g_from", format="YYYY-MM-DD")
        with c2: g_date_to   = st.date_input("إلى تاريخ", value=None, key="g_to", format="YYYY-MM-DD")
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

        header_company = company_name or "—"
        header_project = project_name or "—"
        header_date = str(row.get("تاريخ التعاقد", "—"))
        st.markdown(
            f"""
            <div class="fin-head">
                <div class="line">
                    <strong>الشركة:</strong> {header_company}
                    &nbsp;&nbsp;|&nbsp;&nbsp;
                    <strong>المشروع:</strong> {header_project}
                    &nbsp;&nbsp;|&nbsp;&nbsp;
                    <strong>تاريخ التعاقد:</strong> {header_date}
                </div>
                <span class="badge">تقرير مالي</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # Summary panel from data only (name right, value left)
        summary_pairs = _row_to_pairs_from_data(row)
        if summary_pairs:
            left_items, right_items = _split_pairs_two_columns(summary_pairs)
            st.markdown('<h3 class="hsec">ملخص المشروع</h3>', unsafe_allow_html=True)
            fin_panel_two_tables(left_items=left_items, right_items=right_items)

        # Titles for exports (RTL arrow)
        title_summary = compose_pdf_title(company_name, project_name, "ملخص", g_date_from, g_date_to)
        title_flow    = compose_pdf_title(company_name, project_name, "دفتر التدفق", g_date_from, g_date_to)
        title_all     = compose_pdf_title(company_name, project_name, "ملخص + دفتر التدفق", g_date_from, g_date_to)

        # ---- Downloads (summary) ----
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

        # ---- Ledger (v_financial_flow) ----
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

        df_flow_display = df_flow.drop(columns=["companyid", "contractid"], errors="ignore")
        st.markdown('<div class="card soft">', unsafe_allow_html=True)
        st.dataframe(df_flow_display, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # Individual downloads
        xlsx_flow = make_excel_bytes(df_flow_display, sheet_name="دفتر_التدفق", title_line=title_flow, put_logo=True)
        if xlsx_flow:
            st.download_button("تنزيل الدفتر كـ Excel", xlsx_flow,
                               file_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}.xlsx"),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        csv_flow = make_csv_utf8(df_flow_display)
        st.download_button("تنزيل الدفتر كـ CSV", csv_flow,
                           file_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}.csv"),
                           mime="text/csv")

        # PDF: keep رقم الشيك without commas (and any column containing 'شيك')
        pdf_flow_df = _format_numbers_for_display(df_flow_display, no_comma_cols=["رقم الشيك"])
        pdf_flow = make_pdf_bytes(pdf_flow_df, title_line=title_flow)
        st.download_button("تنزيل الدفتر كـ PDF", pdf_flow,
                           file_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}.pdf"),
                           mime="application/pdf")

        # Combined
        st.markdown("### تنزيل تقرير موحّد")
        excel_two = make_excel_combined_two_sheets(
            {"ملخص": df_summary, "دفتر_التدفق": df_flow_display},
            titles={"ملخص": title_summary, "دفتر_التدفق": title_flow},
            put_logo=True
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
    # Other data types (classic table modes)
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

    column_config = {}
    for col in df.columns:
        if "رابط" in str(col):
            column_config[col] = st.column_config.LinkColumn(label=col, display_text="فتح الرابط")

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

    csv_bytes = make_csv_utf8(df)
    st.download_button("تنزيل كـ CSV (UTF-8)", csv_bytes,
                       file_name=_safe_filename(f"{type_key}_{company_name}_{project_name}.csv"),
                       mime="text/csv")

    pdf_bytes = make_pdf_bytes(_format_numbers_for_display(df), title_line=title_generic)
    st.download_button("تنزيل كـ PDF", pdf_bytes,
                       file_name=_safe_filename(f"{type_key}_{company_name}_{project_name}.pdf"),
                       mime="application/pdf")


if __name__ == "__main__":
    main()