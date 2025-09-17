# app.py
import os
import re
import base64
from io import BytesIO
from pathlib import Path
from typing import Optional, Tuple, Dict

import pandas as pd
import streamlit as st

# ReportLab (PDF) + Arabic
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
    create_date_range,
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

_AR_RE = re.compile(r"[\u0600-\u06FF]")  # Arabic block


def _first_existing(paths) -> Optional[Path]:
    for p in paths:
        pth = Path(p)
        if pth.exists() and pth.is_file() and (pth.stat().st_size > 0):
            return pth
    return None


def _image_size(path: Path) -> Tuple[int, int]:
    if PILImage:
        try:
            with PILImage.open(path) as im:
                return im.size  # (w,h) px
        except Exception:
            pass
    return (600, 120)


def _img_to_data_uri(path: Path) -> Optional[str]:
    try:
        ext = path.suffix.lower().lstrip(".") or "png"
        mime = f"image/{'jpeg' if ext in ('jpg','jpeg') else ext}"
        b64 = base64.b64encode(path.read_bytes()).decode("ascii")
        return f"data:{mime};base64,{b64}"
    except Exception:
        return None


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
# Streamlit Page Config
# =========================================================
st.set_page_config(
    page_title="قاعدة البيانات والتقارير المالية | HGAD",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# Global Styles (CSS)
# =========================================================
st.markdown(
    """
<style>
/* Sidebar always open */
[data-testid="stSidebar"] { transform:none !important; visibility:visible !important; width:340px !important; min-width:340px !important; }
[data-testid="stSidebar"][aria-expanded="false"] { transform:none !important; visibility:visible !important; }
[data-testid="collapsedControl"], button[kind="header"],
button[title="Expand sidebar"], button[title="Collapse sidebar"],
[data-testid="stSidebarCollapseButton"] { display:none !important; }

/* RTL root */
html, body {
    direction: rtl !important;
    text-align: right !important;
    font-family: "Cairo","Noto Kufi Arabic","Segoe UI",Tahoma,sans-serif !important;
    white-space: normal !important;
    word-wrap: break-word !important;
    overflow-x: hidden !important;
}

/* DataFrame readability */
[data-testid="stDataFrame"] thead tr th {
    position: sticky; top: 0; background: #1f2937; color: #f9fafb; z-index: 2;
    font-weight: 700; font-size: 16px;
}
[data-testid="stDataFrame"] div[role="row"] { font-size: 15px; }
[data-testid="stDataFrame"] div[role="row"]:nth-child(even) { background-color: rgba(255,255,255,0.04); }

/* Date inputs: fix popover & contrast */
[data-testid="stDateInput"] input {
    background:#0f172a !important; color:#e5e7eb !important;
    border:1px solid #334155 !important; border-radius:10px !important;
    text-align:center !important; height:42px !important;
}
[data-testid="stDateInput"] label { color:#cbd5e1 !important; font-weight:700; }
div[role="dialog"], .stDateInput, .stDateInput > div[aria-modal="true"] {
    z-index: 9999 !important;
}
.css-1o7jrs8, .css-1n76uvr { z-index: 9999 !important; } /* (best-effort on some Streamlit themes) */

/* Financial header */
.fin-head {
    display:flex; justify-content: space-between; align-items:center;
    border: 1px dashed #1e3a8a55; border-radius: 14px;
    padding: 14px 18px; margin-bottom: 10px; background: #0b1220;
}
.fin-head .line {
    font-size: 24px; font-weight: 900; color: #e5e7eb;
}
.badge { display:inline-block; background:#1e3a8a; color:white; font-weight:700; padding:6px 12px; border-radius:999px; }

/* KPI grid: 2 per row */
.kpi-grid { display:grid; grid-template-columns: repeat(2, minmax(260px, 1fr)); gap: 16px; margin-top: 8px; }
.kpi {
    background: #0b1220; border: 1px solid #1e3a8a33; border-radius: 14px; padding: 16px; box-shadow: 0 4px 20px rgba(0,0,0,0.08);
}
.kpi h4 { margin: 0 0 8px 0; font-size: 14px; color: #93c5fd; font-weight: 700; }
.kpi .val { font-size: 22px; font-weight: 800; color: #e5e7eb; }

/* Summary panel: two-column financial table (like your screenshot) */
.fin-panel { display:grid; grid-template-columns: 1fr 1fr; gap: 18px; margin-top: 10px; }
.fin-table { width:100%; border-collapse: collapse; }
.fin-table th, .fin-table td {
    border: 1px solid #334155; padding: 10px 12px; font-size: 15px;
}
.fin-table th { background:#0f1530; color:#e5e7eb; font-weight:900; }
.fin-table td:first-child { background:#111827; color:#e5e7eb; font-weight:800; text-align:center; width: 40%; }
.fin-table td:last-child  { background:#0b1220; color:#e5e7eb; font-weight:700; text-align:right; }

/* Section titles */
.hsec { color:#1E3A8A; font-weight:800; margin:0.2rem 0 0.6rem 0; font-size: 20px; }
</style>
""",
    unsafe_allow_html=True,
)

# =========================================================
# Header (inline base64 logo)
# =========================================================
logo_path = _site_logo_path()
logo_data_uri = _img_to_data_uri(logo_path) if logo_path else None

c_logo, c_title = st.columns([1, 6], gap="small")
with c_logo:
    if logo_data_uri:
        st.markdown(f'<img src="{logo_data_uri}" width="64" />', unsafe_allow_html=True)
with c_title:
    st.markdown(
        """
<h1 style="color:#1E3A8A; font-weight:800; margin:0;">
    قاعدة البيانات والتقارير المالية
    <span style="font-size:20px; color:#4b5563;">| HGAD Company</span>
</h1>
""",
        unsafe_allow_html=True,
    )
st.markdown('<hr style="border:0; height:2px; background:linear-gradient(to left, transparent, #1E3A8A, transparent);"/>', unsafe_allow_html=True)

# =========================================================
# Excel helpers
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


def _char_width_to_pixels(width_chars: float) -> int:
    return int(width_chars * 7 + 5)


def _wide_logo_data() -> Tuple[Optional[Path], Optional[str]]:
    wl = _wide_logo_path()
    return wl, _img_to_data_uri(wl) if wl else (None, None)


def _auto_excel_sheet(writer, df: pd.DataFrame, sheet_name: str):
    """Write a DataFrame with formats + optional wide logo."""
    engine = writer.engine
    df_x = df.copy()
    wide_logo_path, _ = _wide_logo_data()

    if engine == "xlsxwriter":
        wb = writer.book
        ws = wb.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = ws

        hdr_fmt = wb.add_format({"align": "right", "bold": True})
        fmt_text = wb.add_format({"align": "right"})
        fmt_date = wb.add_format({"align": "right", "num_format": "yyyy-mm-dd"})
        fmt_num  = wb.add_format({"align": "right", "num_format": "#,##0.00"})
        fmt_link = wb.add_format({"font_color": "blue", "underline": 1, "align": "right"})

        char_widths = []
        for idx, col in enumerate(df_x.columns):
            series = df_x[col]
            max_len = max([len(str(col))] + [len(str(v)) for v in series.values])
            width_chars = min(max_len + 4, 60)
            char_widths.append(width_chars)
            if pd.api.types.is_datetime64_any_dtype(series):
                ws.set_column(idx, idx, max(14, width_chars), fmt_date)
            elif pd.api.types.is_numeric_dtype(series):
                ws.set_column(idx, idx, max(14, width_chars), fmt_num)
            elif "رابط" in col:
                ws.set_column(idx, idx, max(20, width_chars), fmt_link)
            else:
                ws.set_column(idx, idx, width_chars, fmt_text)

        header_row = 0
        if wide_logo_path and wide_logo_path.exists():
            img_w, img_h = _image_size(wide_logo_path)
            total_pixels = sum(_char_width_to_pixels(w) for w in char_widths)
            x_scale = (total_pixels / float(img_w)) if img_w else 1.0
            y_scale = x_scale
            ws.insert_image(header_row, 0, str(wide_logo_path), {"x_scale": x_scale, "y_scale": y_scale, "object_position": 1})
            approx_row_height_px = 20
            header_row = int((img_h * y_scale) / approx_row_height_px) + 1
        else:
            header_row = 2

        for col_num, col_name in enumerate(df_x.columns):
            ws.write(header_row, col_num, col_name, hdr_fmt)

        for idx, col in enumerate(df_x.columns):
            series = df_x[col]
            if "رابط" in col:
                for r, val in enumerate(series, start=header_row + 1):
                    sval = "" if pd.isna(val) else str(val)
                    if sval.startswith(("http://", "https://")):
                        ws.write_url(r, idx, sval, fmt_link, string="فتح الرابط")
                    else:
                        ws.write(r, idx, sval, fmt_text)
            elif pd.api.types.is_datetime64_any_dtype(series):
                for r, val in enumerate(series, start=header_row + 1):
                    if pd.notna(val):
                        ws.write_datetime(r, idx, pd.to_datetime(val), fmt_date)
                    else:
                        ws.write_blank(r, idx, None, fmt_text)
            elif pd.api.types.is_numeric_dtype(series):
                for r, val in enumerate(series, start=header_row + 1):
                    if pd.notna(val):
                        ws.write_number(r, idx, float(val), fmt_num)
                    else:
                        ws.write_blank(r, idx, None, fmt_text)
            else:
                for r, val in enumerate(series, start=header_row + 1):
                    ws.write(r, idx, "" if pd.isna(val) else str(val), fmt_text)

    else:  # openpyxl
        df_x.to_excel(writer, index=False, sheet_name=sheet_name)


def make_excel_bytes(df: pd.DataFrame, sheet_name: str = "البيانات") -> Optional[bytes]:
    engine = _pick_excel_engine()
    if engine is None:
        return None
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine=engine) as writer:
        _auto_excel_sheet(writer, df, sheet_name)
    buf.seek(0)
    return buf.getvalue()


def make_excel_combined(dfs: Dict[str, pd.DataFrame]) -> Optional[bytes]:
    """Create one Excel workbook with multiple sheets (ordered by dict insertion)."""
    engine = _pick_excel_engine()
    if engine is None:
        return None
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine=engine) as writer:
        for sheet, df in dfs.items():
            _auto_excel_sheet(writer, df, sheet)
    buf.seek(0)
    return buf.getvalue()


def make_csv_utf8(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


# =========================================================
# PDF builders
# =========================================================
def _pdf_table(df: pd.DataFrame, title: str = "", max_col_width: int = 120) -> list:
    """Return a list of flowables representing a styled table section."""
    font_name, _ = register_arabic_font()
    hdr_style = ParagraphStyle(name="Hdr", fontName=font_name, fontSize=10, textColor=colors.whitesmoke, alignment=1)
    cell_rtl  = ParagraphStyle(name="CellR", fontName=font_name, fontSize=9, leading=12, alignment=2)
    cell_ltr  = ParagraphStyle(name="CellL", fontName=font_name, fontSize=9, leading=12, alignment=0)

    blocks = []
    if title:
        tstyle = ParagraphStyle(name="Sec", fontName=font_name, fontSize=13, alignment=2, textColor=colors.HexColor("#1E3A8A"))
        blocks += [Paragraph(shape_arabic(title), tstyle), Spacer(1, 6)]

    headers = [Paragraph(shape_arabic(c) if looks_arabic(c) else str(c), hdr_style) for c in df.columns]
    rows = [headers]
    for _, r in df.iterrows():
        cells = []
        for c in df.columns:
            sval = "" if pd.isna(r[c]) else str(r[c])
            is_ar = looks_arabic(sval)
            cells.append(Paragraph(shape_arabic(sval) if is_ar else sval, cell_rtl if is_ar else cell_ltr))
        rows.append(cells)

    # col widths
    col_widths = []
    max_col_width = max_col_width
    for c in df.columns:
        max_len = max(len(str(c)), df[c].astype(str).map(len).max())
        col_widths.append(min(max_len * 7, max_col_width))

    table = Table(rows, repeatRows=1, colWidths=col_widths)
    table.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), font_name),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1E3A8A")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.whitesmoke),
        ("BOTTOMPADDING", (0,0), (-1,0), 6),
        ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
    ]))
    blocks.append(table)
    return blocks


def make_pdf_bytes(df: pd.DataFrame, pdf_name: str = "", max_col_width: int = 120) -> bytes:
    """Generic single-table PDF."""
    buf = BytesIO()
    font_name, arabic_ok = register_arabic_font()

    page = landscape(A4)
    left_margin, right_margin, top_margin, bottom_margin = 20, 20, 28, 20
    doc = SimpleDocTemplate(buf, pagesize=page, rightMargin=right_margin, leftMargin=left_margin, topMargin=top_margin, bottomMargin=bottom_margin)

    title_style = ParagraphStyle(name="Title", fontName=font_name, fontSize=15, leading=18, alignment=1)
    base_title = "قاعدة البيانات والتقارير المالية"
    title_text = f"{base_title} ({pdf_name})" if pdf_name else base_title
    if arabic_ok:
        title_text = shape_arabic(title_text)

    elements = []

    # wide logo
    wlp = _wide_logo_path()
    avail_w = page[0] - left_margin - right_margin
    if wlp and wlp.exists():
        try:
            if PILImage:
                w_px, h_px = _image_size(wlp)
                ratio = h_px / float(w_px) if w_px else 0.2
                img_h = max(28, avail_w * ratio)
            else:
                img_h = 50
            logo_img = RLImage(str(wlp), hAlign="CENTER")
            logo_img.drawWidth = avail_w
            logo_img.drawHeight = img_h
            elements.append(logo_img)
            elements.append(Spacer(1, 6))
        except Exception:
            pass

    elements.append(Paragraph(title_text, title_style))
    elements.append(Spacer(1, 10))
    elements += _pdf_table(df)
    doc.build(elements)
    buf.seek(0)
    return buf.getvalue()


def make_pdf_combined(summary_df: pd.DataFrame, flow_df: pd.DataFrame, header_text: str = "") -> bytes:
    """One PDF file: summary (as table) + page break + flow (table)."""
    buf = BytesIO()
    font_name, arabic_ok = register_arabic_font()

    page = landscape(A4)
    left_margin, right_margin, top_margin, bottom_margin = 20, 20, 28, 20
    doc = SimpleDocTemplate(buf, pagesize=page, rightMargin=right_margin, leftMargin=left_margin, topMargin=top_margin, bottomMargin=bottom_margin)

    title_style = ParagraphStyle(name="Title", fontName=font_name, fontSize=16, leading=20, alignment=1)
    head_style  = ParagraphStyle(name="Head",  fontName=font_name, fontSize=13, leading=16, alignment=2, textColor=colors.HexColor("#1E3A8A"))

    base_title = "التقرير المالي"
    if arabic_ok:
        base_title = shape_arabic(base_title)
        header_text = shape_arabic(header_text)

    elements = []
    # wide logo
    wlp = _wide_logo_path()
    avail_w = page[0] - left_margin - right_margin
    if wlp and wlp.exists():
        try:
            if PILImage:
                w_px, h_px = _image_size(wlp)
                ratio = h_px / float(w_px) if w_px else 0.2
                img_h = max(28, avail_w * ratio)
            else:
                img_h = 50
            logo_img = RLImage(str(wlp), hAlign="CENTER")
            logo_img.drawWidth = avail_w
            logo_img.drawHeight = img_h
            elements.append(logo_img)
            elements.append(Spacer(1, 6))
        except Exception:
            pass

    elements.append(Paragraph(base_title, title_style))
    if header_text:
        elements.append(Spacer(1, 6))
        elements.append(Paragraph(header_text, head_style))
    elements.append(Spacer(1, 8))

    # Summary
    elements += _pdf_table(summary_df, title="ملخص المشروع")
    elements.append(PageBreak())

    # Flow (IDs already excluded by caller)
    elements += _pdf_table(flow_df, title="دفتر التدفق")
    doc.build(elements)
    buf.seek(0)
    return buf.getvalue()


# =========================================================
# UI helpers
# =========================================================
def kpi_card(title: str, value: str):
    st.markdown(
        f"""
        <div class="kpi">
            <h4>{title}</h4>
            <div class="val">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def fin_panel_two_tables(left_items, right_items):
    """Render two side-by-side financial tables (label/value pairs)."""
    # Build HTML table for each list of tuples [(label, value), ...]
    def _build(items):
        rows = []
        for label, value in items:
            rows.append(f"<tr><td>{value}</td><td>{label}</td></tr>")
        return f'<table class="fin-table">{"".join(rows)}</table>'
    html = f'<div class="fin-panel"><div>{_build(left_items)}</div><div>{_build(right_items)}</div></div>'
    st.markdown(html, unsafe_allow_html=True)


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

        date_from, date_to = (None, None)
        if type_key == "financial_report":
            st.subheader("نطاق التاريخ (اختياري)")
            date_from, date_to = create_date_range()

    if not company_name or not project_name or not type_key:
        st.info("برجاء اختيار الشركة والمشروع ونوع البيانات من الشريط الجانبي لعرض النتائج.")
        return

    # =======================
    # Financial Report Mode
    # =======================
    if type_key == "financial_report":
        # Summary
        df_summary = fetch_contract_summary_view(conn, company_name, project_name)

        # Header line: Company | Project | Contract Date (bigger)
        header_company = company_name or "—"
        header_project = project_name or "—"
        header_date = "—"
        if not df_summary.empty and "تاريخ التعاقد" in df_summary.columns:
            header_date = str(df_summary.iloc[0].get("تاريخ التعاقد", "—"))

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

        if df_summary.empty:
            st.warning("لم يتم العثور على ملخص العقد لهذا المشروع.")
            return

        # Fixed order values
        row = df_summary.iloc[0].to_dict()

        def _fmt(v):
            try:
                if isinstance(v, str) and v.strip().endswith("%"):
                    return v
                f = float(str(v).replace(",", ""))
                return f"{f:,.2f}"
            except Exception:
                return str(v)

        # A) Professional KPI grid (2 per row)
        st.markdown('<h3 class="hsec">ملخص المشروع</h3>', unsafe_allow_html=True)
        st.markdown('<div class="kpi-grid">', unsafe_allow_html=True)
        kpi_card("قيمة التعاقد", _fmt(row.get("قيمة التعاقد", 0)))
        kpi_card("حجم الأعمال المنفذة", _fmt(row.get("حجم الاعمال المنفذة", 0)))
        kpi_card("نسبة الأعمال المنفذة", _fmt(row.get("نسبة الاعمال المنفذة", "0%")))
        kpi_card("الدفعة المقدمة", _fmt(row.get("الدفعه المقدمه", 0)))
        kpi_card("التحصيلات", _fmt(row.get("التحصيلات", 0)))
        kpi_card("المستحق صرفه", _fmt(row.get("المستحق صرفه", 0)))
        st.markdown('</div>', unsafe_allow_html=True)

        # B) Two-column financial panel (shape like screenshot)
        left_items = [
            ("مواد أولية", _fmt(row.get("حجم الاعمال المنفذة", 0))),    # مثال: ضع قيمتك الحقيقية هنا إن أردت
            ("مصروفات غير مباشرة", _fmt(0)),
            ("مصروفات تشغيلية", _fmt(0)),
            ("إجمالي المصروفات", _fmt(row.get("حجم الاعمال المنفذة", 0))),  # مثال توضيحي
            ("الحد الائتماني", _fmt(row.get("قيمة التعاقد", 0))),
            ("المستحق صرفه", _fmt(row.get("المستحق صرفه", 0))),
        ]
        right_items = [
            ("تاريخ التعاقد", header_date),
            ("قيمة التعاقد", _fmt(row.get("قيمة التعاقد", 0))),
            ("حجم الاعمال المنفذة", _fmt(row.get("حجم الاعمال المنفذة", 0))),
            ("نسبة الاعمال المنفذة", _fmt(row.get("نسبة الاعمال المنفذة", "0%"))),
            ("الدفعة المقدمة", _fmt(row.get("الدفعه المقدمه", 0))),
            ("إجمالي التحصيلات", _fmt(row.get("التحصيلات", 0))),
        ]
        fin_panel_two_tables(left_items, right_items)

        # Downloads for summary only
        df_summary_out = df_summary.copy()
        xlsx_sum = make_excel_bytes(df_summary_out, sheet_name="ملخص")
        if xlsx_sum:
            st.download_button(
                label="تنزيل الملخص كـ Excel",
                data=xlsx_sum,
                file_name=_safe_filename(f"ملخص_{company_name}_{project_name}.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        pdf_sum = make_pdf_bytes(df_summary_out, pdf_name=_safe_filename(f"ملخص_{company_name}_{project_name}"))
        st.download_button(
            label="تنزيل الملخص كـ PDF",
            data=pdf_sum,
            file_name=_safe_filename(f"ملخص_{company_name}_{project_name}.pdf"),
            mime="application/pdf",
        )

        st.markdown("---")

        # 2) Ledger table from v_financial_flow
        st.markdown('<h3 class="hsec">الدفتر الزمني (v_financial_flow)</h3>', unsafe_allow_html=True)
        df_flow = fetch_financial_flow_view(conn, company_name, project_name, date_from, date_to)
        if df_flow.empty:
            st.info("لا توجد حركات مطابقة ضمن النطاق المحدد.")
            return

        # Search in flow
        col_search, term = create_column_search(df_flow)
        if col_search and term:
            df_flow = df_flow[df_flow[col_search].astype(str).str.contains(str(term), case=False, na=False)]
            if df_flow.empty:
                st.info("لا توجد نتائج بعد تطبيق البحث.")
                return

        # On-screen: hide IDs
        df_flow_display = df_flow.drop(columns=["companyid", "contractid"], errors="ignore")
        st.dataframe(df_flow_display, use_container_width=True, hide_index=True)

        # Individual downloads
        xlsx_flow = make_excel_bytes(df_flow_display, sheet_name="دفتر_التدفق")
        if xlsx_flow:
            st.download_button(
                label="تنزيل الدفتر كـ Excel",
                data=xlsx_flow,
                file_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        csv_flow = make_csv_utf8(df_flow_display)
        st.download_button(
            label="تنزيل الدفتر كـ CSV",
            data=csv_flow,
            file_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}.csv"),
            mime="text/csv",
        )
        pdf_flow = make_pdf_bytes(df_flow_display, pdf_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}"))
        st.download_button(
            label="تنزيل الدفتر كـ PDF",
            data=pdf_flow,
            file_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}.pdf"),
            mime="application/pdf",
        )

        # ✅ Combined downloads (one Excel with 2 sheets, one PDF with both)
        st.markdown("### تنزيل تقرير موحّد")
        excel_all = make_excel_combined({
            "ملخص": df_summary_out,
            "دفتر_التدفق": df_flow_display,
        })
        if excel_all:
            st.download_button(
                label="تنزيل التقرير المالي (Excel واحد)",
                data=excel_all,
                file_name=_safe_filename(f"تقرير_مالي_{company_name}_{project_name}.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        header_line = f"الشركة: {header_company} | المشروع: {header_project} | تاريخ التعاقد: {header_date}"
        pdf_all = make_pdf_combined(df_summary_out, df_flow_display, header_text=header_line)
        st.download_button(
            label="تنزيل التقرير المالي (PDF واحد)",
            data=pdf_all,
            file_name=_safe_filename(f"تقرير_مالي_{company_name}_{project_name}.pdf"),
            mime="application/pdf",
        )
        return

    # =======================
    # Classic table mode (raw tables)
    # =======================
    df = fetch_data(conn, company_name, project_name, type_key)
    if df.empty:
        st.warning("لا توجد بيانات مطابقة للاختيارات المحددة.")
        return

    search_column, search_term = create_column_search(df)
    if search_column and search_term:
        df = df[df[search_column].astype(str).str.contains(str(search_term), case=False, na=False)]
        if df.empty:
            st.info("لا توجد نتائج بعد تطبيق معيار البحث.")
            return

    column_config = {}
    for col in df.columns:
        if "رابط" in col:
            column_config[col] = st.column_config.LinkColumn(label=col, display_text="فتح الرابط")

    st.markdown('<h3 class="hsec">البيانات</h3>', unsafe_allow_html=True)
    st.dataframe(df, column_config=column_config, use_container_width=True, hide_index=True)

    xlsx_bytes = make_excel_bytes(df)
    if xlsx_bytes is not None:
        st.download_button(
            label="تنزيل كـ Excel (XLSX) – مُوصى به",
            data=xlsx_bytes,
            file_name=_safe_filename(f"{type_key}_{company_name}_{project_name}.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    csv_bytes = make_csv_utf8(df)
    st.download_button(
        label="تنزيل كـ CSV (UTF-8)",
        data=csv_bytes,
        file_name=_safe_filename(f"{type_key}_{company_name}_{project_name}.csv"),
        mime="text/csv",
    )

    pdf_title = _safe_filename(f"{type_key}_{company_name}_{project_name}")
    pdf_bytes = make_pdf_bytes(df, pdf_name=pdf_title)
    st.download_button(
        label="تنزيل كـ PDF",
        data=pdf_bytes,
        file_name=f"{pdf_title}.pdf",
        mime="application/pdf",
    )


if __name__ == "__main__":
    main()
