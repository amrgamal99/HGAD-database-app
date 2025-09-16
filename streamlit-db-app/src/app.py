# app.py
import os
import re
import base64
from io import BytesIO
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd
import streamlit as st
from streamlit.components.v1 import html as st_html

# PDF & Arabic
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
)
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import arabic_reshaper
from bidi.algorithm import get_display

from db.connection import get_db_connection, fetch_data
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

# SQUARE site/logo (requested)
SITE_LOGO_CANDIDATES = [
    ASSETS_DIR / "logo.png",
    ASSETS_DIR / "logo.jpg",
    ASSETS_DIR / "logo.jpeg",
    ASSETS_DIR / "logo.webp",
]

# Arabic fonts (best effort)
AR_FONT_CANDIDATES = [
    ASSETS_DIR / "Cairo-Regular.ttf",
    ASSETS_DIR / "Amiri-Regular.ttf",
    Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
]

_AR_RE = re.compile(r"[\u0600-\u06FF]")  # Arabic block


def _first_existing(paths) -> Optional[str]:
    for p in paths:
        if Path(p).exists():
            return str(p)
    return None


def get_site_logo_path() -> Optional[str]:
    return _first_existing(SITE_LOGO_CANDIDATES)


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
    mime = f"image/{'jpeg' if ext in ('jpg', 'jpeg') else ext}"
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return f"data:{mime};base64,{b64}"


SITE_LOGO = get_site_logo_path()  # square (UI bottom-left + PDF + Excel)

# =========================================================
# Streamlit Page Config
# =========================================================
st.set_page_config(
    page_title="قاعدة البيانات والتقارير المالية | HGAD",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# Global Styles (CSS). No sidebar logo. No footer element.
# We later inject a floating square logo (bottom-left) via st_html.
# =========================================================
st.markdown(
    """
<style>
/* Sidebar always open */
[data-testid="stSidebar"] { transform:none !important; visibility:visible !important; width:340px !important; min-width:340px !important; }
[data-testid="stSidebar"][aria-expanded="false"] { transform:none !important; visibility:visible !important; }
[data-testid="collapsedControl"],
button[kind="header"],
button[title="Expand sidebar"],
button[title="Collapse sidebar"],
[data-testid="stSidebarCollapseButton"] { display:none !important; }

/* RTL root */
html, body {
    direction: rtl !important; text-align: right !important;
    font-family: "Cairo","Noto Kufi Arabic","Segoe UI",Tahoma,sans-serif !important;
    white-space: normal !important; word-wrap: break-word !important; overflow-x: hidden !important;
}

/* DataFrame readability */
[data-testid="stDataFrame"] thead tr th {
    position: sticky; top: 0; background: #1f2937; color: #f9fafb; z-index: 2;
    font-weight: 700; font-size: 16px;
}
[data-testid="stDataFrame"] div[role="row"] { font-size: 15px; }
[data-testid="stDataFrame"] div[role="row"]:nth-child(even) { background-color: rgba(255,255,255,0.04); }
</style>
""",
    unsafe_allow_html=True,
)

# Inject square logo pinned to bottom-left (NOT sidebar, NOT a "footer" element)
if SITE_LOGO and Path(SITE_LOGO).exists():
    st_html(
        f"""
<div id="hgad-fixed-logo">
  <img src="{_img_to_data_uri(SITE_LOGO)}" alt="HGAD" loading="lazy" decoding="async" />
</div>
<style>
  #hgad-fixed-logo {{
    position: fixed; left: 16px; bottom: 14px; z-index: 9999;
    pointer-events: none; /* purely decorative */
  }}
  #hgad-fixed-logo img {{
    height: 48px; width: auto; opacity: 0.95; border-radius: 6px; box-shadow: 0 2px 10px rgba(0,0,0,.2);
  }}
  @media print {{
    #hgad-fixed-logo {{ position: fixed; left: 16px; bottom: 14px; opacity: .98; }}
  }}
</style>
""",
        height=0,  # just inject
    )

# =========================================================
# Header (title only)
# =========================================================
st.markdown(
    """
<h1 style="color:#1E3A8A; font-weight:800; margin:0; text-align:center;">
    قاعدة البيانات والتقارير المالية
    <span style="font-size:20px; color:#4b5563;">| HGAD Company</span>
</h1>
<hr style="border:0; height:2px; background:linear-gradient(to left, transparent, #1E3A8A, transparent);"/>
""",
    unsafe_allow_html=True,
)

# =========================================================
# Helpers & Hints
# =========================================================
DATE_HINTS = ("تاريخ", "إصدار", "date")
NUM_HINTS  = ("قيمة", "المستحق", "شيك", "التحويل", "USD", ")USD")


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

# =========================================================
# Excel (square logo, small, table clear)
# =========================================================
def make_excel_bytes(df: pd.DataFrame) -> Optional[bytes]:
    engine = _pick_excel_engine()
    if engine is None:
        return None

    buf = BytesIO()
    df_x = df.copy()
    sheet = "البيانات"
    startrow = 6
    logo_path = SITE_LOGO
    has_logo = bool(logo_path and Path(logo_path).exists())

    if engine == "xlsxwriter":
        with pd.ExcelWriter(buf, engine=engine) as writer:
            wb = writer.book
            ws = wb.add_worksheet(sheet)
            writer.sheets[sheet] = ws

            # Insert square logo in the top-left
            if has_logo:
                ws.insert_image("A1", logo_path, {"x_scale": 0.45, "y_scale": 0.45})
            else:
                startrow = 2

            # Formats
            hdr_fmt = wb.add_format({"align": "right", "bold": True})
            fmt_text = wb.add_format({"align": "right"})
            fmt_date = wb.add_format({"align": "right", "num_format": "yyyy-mm-dd"})
            fmt_num  = wb.add_format({"align": "right", "num_format": "#,##0.00"})
            fmt_link = wb.add_format({"font_color": "blue", "underline": 1, "align": "right"})

            # Headers
            for col_num, col_name in enumerate(df_x.columns):
                ws.write(startrow, col_num, col_name, hdr_fmt)

            # Body
            for idx, col in enumerate(df_x.columns):
                series = df_x[col]
                max_len = max([len(str(col))] + [len(str(v)) for v in series.values])
                width = min(max_len + 4, 60)

                if "رابط" in col:
                    for row_num, val in enumerate(series, start=startrow + 1):
                        sval = "" if pd.isna(val) else str(val)
                        if sval.startswith(("http://", "https://")):
                            ws.write_url(row_num, idx, sval, fmt_link, string="فتح الرابط")
                        else:
                            ws.write(row_num, idx, sval, fmt_text)
                    ws.set_column(idx, idx, max(20, width), fmt_link)

                elif pd.api.types.is_datetime64_any_dtype(series):
                    for row_num, val in enumerate(series, start=startrow + 1):
                        if pd.notna(val):
                            ws.write_datetime(row_num, idx, pd.to_datetime(val), fmt_date)
                        else:
                            ws.write_blank(row_num, idx, None, fmt_text)
                    ws.set_column(idx, idx, max(14, width), fmt_date)

                elif pd.api.types.is_numeric_dtype(series):
                    for row_num, val in enumerate(series, start=startrow + 1):
                        if pd.notna(val):
                            ws.write_number(row_num, idx, float(val), fmt_num)
                        else:
                            ws.write_blank(row_num, idx, None, fmt_text)
                    ws.set_column(idx, idx, max(14, width), fmt_num)

                else:
                    for row_num, val in enumerate(series, start=startrow + 1):
                        ws.write(row_num, idx, "" if pd.isna(val) else str(val), fmt_text)
                    ws.set_column(idx, idx, width, fmt_text)

    else:
        # openpyxl fallback
        with pd.ExcelWriter(buf, engine=engine) as writer:
            df_x.head(0).to_excel(writer, index=False, sheet_name=sheet)
            ws = writer.book[sheet]

            if has_logo:
                try:
                    from openpyxl.drawing.image import Image as XLImage  # Pillow
                    img = XLImage(logo_path)
                    ws.add_image(img, "A1")
                except Exception:
                    pass
            else:
                startrow = 2

            df_x.to_excel(writer, index=False, sheet_name=sheet, startrow=startrow + 1)
            for col_num, col_name in enumerate(df_x.columns, start=1):
                ws.cell(row=startrow + 1, column=col_num, value=col_name)

    buf.seek(0)
    return buf.getvalue()


def make_csv_utf8(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

# =========================================================
# PDF (square logo at bottom-left on every page)
# =========================================================
def make_pdf_bytes(df: pd.DataFrame, pdf_name: str = "", max_col_width: int = 120) -> bytes:
    buf = BytesIO()

    font_name, arabic_ok = register_arabic_font()
    link_label = "فتح الرابط" if arabic_ok else "Open link"

    page = landscape(A4)
    left_margin, right_margin, top_margin, bottom_margin = 20, 20, 28, 28
    doc = SimpleDocTemplate(
        buf,
        pagesize=page,
        rightMargin=right_margin,
        leftMargin=left_margin,
        topMargin=top_margin,
        bottomMargin=bottom_margin,
    )

    # Styles
    title_style = ParagraphStyle(name="Title", fontName=font_name, fontSize=15, leading=18, alignment=1)
    hdr_style   = ParagraphStyle(name="Hdr",   fontName=font_name, fontSize=10, textColor=colors.whitesmoke, alignment=1)
    cell_rtl    = ParagraphStyle(name="CellR", fontName=font_name, fontSize=9, leading=12, alignment=2)
    cell_ltr    = ParagraphStyle(name="CellL", fontName=font_name, fontSize=9, leading=12, alignment=0)
    cell_link   = ParagraphStyle(name="CellK", fontName=font_name, fontSize=9, leading=12, alignment=1,
                                 textColor=colors.HexColor("#1E3A8A"), underline=True)

    base_title = "قاعدة البيانات والتقارير المالية"
    title_text = f"{base_title} ({pdf_name})" if pdf_name else base_title
    if arabic_ok:
        title_text = shape_arabic(title_text)

    elements = []
    elements.append(Paragraph(title_text, title_style))
    elements.append(Spacer(1, 10))

    # headers
    header_paragraphs = []
    for col in df.columns:
        text = shape_arabic(col) if arabic_ok and looks_arabic(col) else str(col)
        header_paragraphs.append(Paragraph(text, hdr_style))
    rows = [header_paragraphs]

    # link columns
    link_cols_idx = [i for i, c in enumerate(df.columns) if ("رابط" in str(c)) or ("link" in str(c).lower())]

    # body
    for _, row in df.iterrows():
        cells = []
        for i, col in enumerate(df.columns):
            sval = "" if pd.isna(row[col]) else str(row[col])
            if i in link_cols_idx and sval.startswith(("http://", "https://")):
                label = shape_arabic(link_label) if arabic_ok else link_label
                cells.append(Paragraph(f'<link href="{sval}">{label}</link>', cell_link))
            else:
                is_ar = arabic_ok and looks_arabic(sval)
                cells.append(Paragraph(shape_arabic(sval) if is_ar else sval, cell_rtl if is_ar else cell_ltr))
        rows.append(cells)

    # column widths
    avail_w = page[0] - left_margin - right_margin
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
        if any(h in str(col) for h in DATE_HINTS):
            align = "CENTER"
        elif any(h in str(col) for h in NUM_HINTS) or pd.api.types.is_numeric_dtype(df[col]):
            align = "RIGHT"
        elif idx in link_cols_idx:
            align = "CENTER"
        else:
            align = "RIGHT" if looks_arabic(col) else "LEFT"
        style_cmds.append(("ALIGN", (idx, 1), (idx, -1), align))

    table.setStyle(TableStyle(style_cmds))
    elements.append(table)

    # Draw square logo bottom-left on every page
    site_logo_path = SITE_LOGO if (SITE_LOGO and Path(SITE_LOGO).exists()) else None

    def _draw_logo(canvas, _doc):
        if not site_logo_path:
            return
        try:
            # Keep proportions; small height near margin
            img_h = 18  # points
            # width is calculated by keeping a square feel (safe)
            img_w = img_h
            x = left_margin
            y = 8  # a little above the page edge
            canvas.drawImage(site_logo_path, x, y, width=img_w, height=img_h, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    doc.build(elements, onFirstPage=_draw_logo, onLaterPages=_draw_logo)
    buf.seek(0)
    return buf.getvalue()

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
        type_label, target_table = create_type_dropdown()

    if not company_name or not project_name or not target_table:
        st.info("برجاء اختيار الشركة والمشروع ونوع البيانات من الشريط الجانبي لعرض النتائج.")
        return

    df = fetch_data(conn, company_name, project_name, target_table)
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

    # Downloads
    xlsx_bytes = make_excel_bytes(df)
    if xlsx_bytes is not None:
        st.download_button(
            label="تنزيل كـ Excel (XLSX) – مُوصى به",
            data=xlsx_bytes,
            file_name=_safe_filename(f"{target_table}_{company_name}_{project_name}.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("لتمكين تنزيل Excel مع شعار وروابط قابلة للنقر، أضف إلى requirements: `xlsxwriter` أو `openpyxl`.")

    csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        label="تنزيل كـ CSV (UTF-8)",
        data=csv_bytes,
        file_name=_safe_filename(f"{target_table}_{company_name}_{project_name}.csv"),
        mime="text/csv",
    )

    pdf_title = _safe_filename(f"{target_table}_{company_name}_{project_name}")
    pdf_bytes = make_pdf_bytes(df, pdf_name=pdf_title)
    st.download_button(
        label="تنزيل كـ PDF (عربي)",
        data=pdf_bytes,
        file_name=f"{pdf_title}.pdf",
        mime="application/pdf",
    )


if __name__ == "__main__":
    main()
