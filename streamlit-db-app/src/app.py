# app.py
import os
import re
import base64
from io import BytesIO
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd
import streamlit as st

# PDF & Arabic
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    Image as RLImage,
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

SITE_LOGO_PREFERRED = ASSETS_DIR / "logo.png"
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
_AR_RE = re.compile(r"[\u0600-\u06FF]")

def _first_existing(paths) -> Optional[str]:
    for p in paths:
        if Path(p).exists():
            return str(p)
    return None

def get_site_logo_path() -> Optional[str]:
    return str(SITE_LOGO_PREFERRED) if SITE_LOGO_PREFERRED.exists() else None

def get_wide_logo_path() -> Optional[str]:
    return _first_existing(WIDE_LOGO_CANDIDATES)

def _first_existing_font_path() -> Optional[str]:
    return _first_existing(AR_FONT_CANDIDATES)

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

def _img_to_data_uri(path: str) -> str:
    ext = Path(path).suffix.lower().lstrip(".") or "png"
    mime = f"image/{'jpeg' if ext in ('jpg','jpeg') else ext}"
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return f"data:{mime};base64,{b64}"

SITE_LOGO = get_site_logo_path()
WIDE_LOGO = get_wide_logo_path()

def _image_size(path: str) -> Tuple[int, int]:
    if PILImage:
        try:
            with PILImage.open(path) as im:
                return im.size
        except Exception:
            pass
    return (600, 120)

# =========================================================
# Page config
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
[data-testid="stSidebar"] { transform:none !important; visibility:visible !important; width:340px !important; min-width:340px !important; }
[data-testid="stSidebar"][aria-expanded="false"] { transform:none !important; visibility:visible !important; }
[data-testid="collapsedControl"],
button[kind="header"],
button[title="Expand sidebar"],
button[title="Collapse sidebar"],
[data-testid="stSidebarCollapseButton"] { display:none !important; }

html, body {
    direction: rtl !important;
    text-align: right !important;
    font-family: "Cairo","Noto Kufi Arabic","Segoe UI",Tahoma,sans-serif !important;
    white-space: normal !important;
    word-wrap: break-word !important;
    overflow-x: hidden !important;
}

[data-testid="stDataFrame"] thead tr th {
    position: sticky; top: 0; background: #1f2937; color: #f9fafb; z-index: 2;
    font-weight: 700; font-size: 16px;
}
[data-testid="stDataFrame"] div[role="row"] { font-size: 15px; }
[data-testid="stDataFrame"] div[role="row"]:nth-child(even) { background-color: rgba(255,255,255,0.04); }

.report-card {
    display:flex; gap: 24px; flex-wrap: wrap;
}
.report-kpi {
    flex: 1 1 200px;
    background: #0b1220;
    border: 1px solid #1e3a8a33;
    border-radius: 14px;
    padding: 16px;
    min-width: 220px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
}
.report-kpi h4 {
    margin: 0 0 8px 0; font-size: 14px; color: #93c5fd; font-weight: 700;
}
.report-kpi .val {
    font-size: 22px; font-weight: 800; color: #e5e7eb;
}
.report-header {
    display:flex; justify-content: space-between; align-items: center;
    padding: 10px 14px; border: 1px dashed #1e3a8a55; border-radius: 12px; margin-bottom: 10px;
}
.badge {
    display:inline-block; background:#1e3a8a; color:white; font-weight:700; padding:4px 10px; border-radius: 999px;
}
</style>
""",
    unsafe_allow_html=True,
)

# =========================================================
# Header
# =========================================================
c_logo, c_title = st.columns([1, 6], gap="small")
with c_logo:
    if SITE_LOGO:
        st.image(SITE_LOGO, width=64)
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
        import xlsxwriter  # noqa
        return "xlsxwriter"
    except Exception:
        pass
    try:
        import openpyxl  # noqa
        return "openpyxl"
    except Exception:
        return None

def _char_width_to_pixels(width_chars: float) -> int:
    return int(width_chars * 7 + 5)

def make_excel_bytes(df: pd.DataFrame, sheet_name: str = "البيانات") -> Optional[bytes]:
    engine = _pick_excel_engine()
    if engine is None:
        return None

    buf = BytesIO()
    df_x = df.copy()

    wide_logo = WIDE_LOGO
    if engine == "xlsxwriter":
        with pd.ExcelWriter(buf, engine=engine) as writer:
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
            if wide_logo and Path(wide_logo).exists():
                img_w, img_h = _image_size(wide_logo)
                total_pixels = sum(_char_width_to_pixels(w) for w in char_widths)
                x_scale = (total_pixels / float(img_w)) if img_w else 1.0
                y_scale = x_scale
                ws.insert_image(header_row, 0, wide_logo, {"x_scale": x_scale, "y_scale": y_scale, "object_position": 1})
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
    else:
        with pd.ExcelWriter(buf, engine=engine) as writer:
            df_x.head(0).to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.book[sheet_name]
            header_row = 2
            if WIDE_LOGO and Path(WIDE_LOGO).exists():
                try:
                    from openpyxl.drawing.image import Image as XLImage
                    char_widths = []
                    for idx, col in enumerate(df_x.columns):
                        series = df_x[col]
                        max_len = max([len(str(col))] + [len(str(v)) for v in series.values])
                        width_chars = min(max_len + 4, 60)
                        char_widths.append(width_chars)
                        ws.column_dimensions[chr(ord("A") + idx)].width = width_chars
                    total_pixels = sum(_char_width_to_pixels(w) for w in char_widths)
                    img = XLImage(WIDE_LOGO)
                    img_w, img_h = _image_size(WIDE_LOGO)
                    if img_w:
                        scale = total_pixels / float(img_w)
                        img.width = int(img_w * scale)
                        img.height = int(img_h * scale)
                    ws.add_image(img, "A1")
                    header_row = 3 + max(int(img.height // 18), 1)
                except Exception:
                    pass
            df_x.to_excel(writer, index=False, sheet_name=sheet_name, startrow=header_row)

    buf.seek(0)
    return buf.getvalue()

def make_csv_utf8(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

# =========================================================
# PDF generation (same as قبل)
# =========================================================
def make_pdf_bytes(df: pd.DataFrame, pdf_name: str = "", max_col_width: int = 120) -> bytes:
    buf = BytesIO()
    font_name, arabic_ok = register_arabic_font()
    link_label = "فتح الرابط" if arabic_ok else "Open link"

    page = landscape(A4)
    left_margin, right_margin, top_margin, bottom_margin = 20, 20, 28, 20
    doc = SimpleDocTemplate(
        buf, pagesize=page,
        rightMargin=right_margin, leftMargin=left_margin,
        topMargin=top_margin, bottomMargin=bottom_margin,
    )

    title_style = ParagraphStyle(name="Title", fontName=font_name, fontSize=15, leading=18, alignment=1)
    hdr_style   = ParagraphStyle(name="Hdr",   fontName=font_name, fontSize=10, textColor=colors.whitesmoke, alignment=1)
    cell_rtl    = ParagraphStyle(name="CellR", fontName=font_name, fontSize=9, leading=12, alignment=2)
    cell_ltr    = ParagraphStyle(name="CellL", fontName=font_name, fontSize=9, leading=12, alignment=0)
    cell_link   = ParagraphStyle(name="CellK", fontName=font_name, fontSize=9, leading=12, alignment=1, textColor=colors.HexColor("#1E3A8A"), underline=True)

    base_title = "قاعدة البيانات والتقارير المالية"
    title_text = f"{base_title} ({pdf_name})" if pdf_name else base_title
    if arabic_ok:
        title_text = shape_arabic(title_text)

    elements = []

    avail_w = page[0] - left_margin - right_margin
    if WIDE_LOGO and Path(WIDE_LOGO).exists():
        try:
            if PILImage:
                w_px, h_px = _image_size(WIDE_LOGO)
                ratio = h_px / float(w_px) if w_px else 0.2
                img_h = max(28, avail_w * ratio)
            else:
                img_h = 50
            logo_img = RLImage(WIDE_LOGO, hAlign="CENTER")
            logo_img.drawWidth = avail_w
            logo_img.drawHeight = img_h
            elements.append(logo_img)
            elements.append(Spacer(1, 6))
        except Exception:
            pass

    elements.append(Paragraph(title_text, title_style))
    elements.append(Spacer(1, 10))

    header_paragraphs = []
    for col in df.columns:
        text = shape_arabic(col) if arabic_ok and looks_arabic(col) else str(col)
        header_paragraphs.append(Paragraph(text, hdr_style))
    rows = [header_paragraphs]

    link_cols_idx = [i for i, c in enumerate(df.columns) if ("رابط" in str(c)) or ("link" in str(c).lower())]

    for _, row in df.iterrows():
        cells = []
        for i, col in enumerate(df.columns):
            sval = "" if pd.isna(row[col]) else str(row[col])
            if i in link_cols_idx and sval.startswith(("http://", "https://")):
                label = shape_arabic(link_label) if looks_arabic(link_label) else link_label
                cells.append(Paragraph(f'<link href="{sval}">{label}</link>', cell_link))
            else:
                is_ar = looks_arabic(sval)
                cells.append(Paragraph(shape_arabic(sval) if is_ar else sval, cell_rtl if is_ar else cell_ltr))
        rows.append(cells)

    col_widths = []
    for idx, col in enumerate(df.columns):
        if idx in link_cols_idx:
            max_len = max(len(str(col)), len(link_label))
        else:
            max_len = max(len(str(col)), df[col].astype(str).map(len).max())
        col_widths.append(min(max_len * 7, max_col_width))

    total_w = sum(col_widths)
    if total_w > avail_w and total_w > 0:
        scale = avail_w / total_w
        col_widths = [w * scale for w in col_widths]

    table = Table(rows, repeatRows=1, colWidths=col_widths)
    style_cmds = [
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
    ]
    for idx, col in enumerate(df.columns):
        if ("تاريخ" in str(col)) or ("إصدار" in str(col)) or ("date" in str(col).lower()):
            align = "CENTER"
        elif ("قيمة" in str(col)) or ("المستحق" in str(col)) or ("شيك" in str(col)) or pd.api.types.is_numeric_dtype(df[col]):
            align = "RIGHT"
        elif idx in link_cols_idx:
            align = "CENTER"
        else:
            align = "RIGHT" if looks_arabic(col) else "LEFT"
        style_cmds.append(("ALIGN", (idx, 1), (idx, -1), align))

    table.setStyle(TableStyle(style_cmds))
    elements.append(table)

    doc.build(elements)
    buf.seek(0)
    return buf.getvalue()

# =========================================================
# UI helpers
# =========================================================
def kpi_card(title: str, value: str):
    st.markdown(f"""
    <div class="report-kpi">
      <h4>{title}</h4>
      <div class="val">{value}</div>
    </div>
    """, unsafe_allow_html=True)

# =========================================================
# Main
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

        # عند اختيار تقرير مالي نُظهر مرشح التاريخ لدفتر التدفق
        date_from, date_to = (None, None)
        if type_key == "financial_report":
            st.subheader("نطاق التاريخ (اختياري)")
            date_from, date_to = create_date_range()

    if not company_name or not project_name or not type_key:
        st.info("برجاء اختيار الشركة والمشروع ونوع البيانات من الشريط الجانبي لعرض النتائج.")
        return

    # =======================
    # وضع "تقرير مالي"
    # =======================
    if type_key == "financial_report":
        # بطاقة ملخص المشروع من v_contract_summary
        df_summary = fetch_contract_summary_view(conn, company_name, project_name)
        if df_summary.empty:
            st.warning("لم يتم العثور على ملخص العقد لهذا المشروع.")
        else:
            row = df_summary.iloc[0].to_dict()
            st.markdown(
                f"""
                <div class="report-header">
                    <div>
                        <strong>المشروع:</strong> {row.get("اسم المشروع","—")} &nbsp;|&nbsp;
                        <strong>تاريخ التعاقد:</strong> {row.get("تاريخ التعاقد","—")}
                    </div>
                    <span class="badge">تقرير مالي</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

            st.markdown('<div class="report-card">', unsafe_allow_html=True)
            kpi_card("قيمة التعاقد", f'{row.get("قيمة التعاقد", 0):,}')
            kpi_card("حجم الأعمال المنفذة", f'{row.get("حجم الاعمال المنفذة", 0):,}')
            kpi_card("نسبة الأعمال المنفذة", str(row.get("نسبة الاعمال المنفذة", "0%")))
            kpi_card("الدفعة المقدمة", f'{row.get("الدفعه المقدمه", 0):,}')
            kpi_card("التحصيلات", f'{row.get("التحصيلات", 0):,}')
            kpi_card("المستحق صرفه", f'{row.get("المستحق صرفه", 0):,}')
            st.markdown('</div>', unsafe_allow_html=True)

            # تنزيل بطاقة الملخص كملف (نحوّل السطر DataFrame للتنزيل)
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

        # الدفتر الزمني من v_financial_flow
        df_flow = fetch_financial_flow_view(conn, company_name, project_name, date_from, date_to)
        if df_flow.empty:
            st.info("لا توجد حركات مطابقة في الدفتر الزمني ضمن النطاق المحدد.")
            return

        # بحث اختياري داخل الدفتر
        st.markdown("### الدفتر الزمني (v_financial_flow)")
        col_search, term = create_column_search(df_flow)
        if col_search and term:
            df_flow = df_flow[df_flow[col_search].astype(str).str.contains(str(term), case=False, na=False)]
            if df_flow.empty:
                st.info("لا توجد نتائج بعد تطبيق البحث.")
                return

        # عرض الدفتر
        st.dataframe(
            df_flow.drop(columns=["companyid", "contractid"], errors="ignore"),
            use_container_width=True, hide_index=True
        )

        # تنزيلات
        xlsx_flow = make_excel_bytes(df_flow, sheet_name="دفتر_التدفق")
        if xlsx_flow:
            st.download_button(
                label="تنزيل الدفتر كـ Excel",
                data=xlsx_flow,
                file_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        csv_flow = make_csv_utf8(df_flow)
        st.download_button(
            label="تنزيل الدفتر كـ CSV",
            data=csv_flow,
            file_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}.csv"),
            mime="text/csv",
        )
        pdf_flow = make_pdf_bytes(df_flow, pdf_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}"))
        st.download_button(
            label="تنزيل الدفتر كـ PDF",
            data=pdf_flow,
            file_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}.pdf"),
            mime="application/pdf",
        )
        return

    # =======================
    # الوضع التقليدي (الجداول الخام)
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

    st.markdown("### البيانات")
    st.dataframe(df, column_config=column_config, use_container_width=True, hide_index=True)

    xlsx_bytes = make_excel_bytes(df)
    if xlsx_bytes is not None:
        st.download_button(
            label="تنزيل كـ Excel (XLSX) – مُوصى به",
            data=xlsx_bytes,
            file_name=_safe_filename(f"{type_key}_{company_name}_{project_name}.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("لتمكين تنزيل Excel مع شعار ممتد، أضف إلى requirements: `xlsxwriter` أو `openpyxl`.")

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
