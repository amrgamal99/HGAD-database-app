import re
import streamlit as st
import pandas as pd
from io import BytesIO

# PDF & Arabic
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import arabic_reshaper
from bidi.algorithm import get_display

# ---------- Arabic PDF helpers (robust fonts + clean links) ----------
import os, re
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
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
    create_column_search
)

# ---------- Page config ----------
st.set_page_config(
    page_title=" قاعدة البيانات والتقارير المالية | HGAD",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------- Styles ----------
st.markdown("""
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
</style>
""", unsafe_allow_html=True)

# ---------- Header ----------
st.markdown("""
    <h1 style="text-align:center; color:#1E3A8A; font-weight:800;">
     قاعدة البيانات والتقارير المالية
        <span style="font-size:20px; color:#4b5563;">| HGAD Company</span>
    </h1>
    <hr style="border:0; height:2px; background:linear-gradient(to left, transparent, #1E3A8A, transparent);"/>
""", unsafe_allow_html=True)

# ---------- Download helpers ----------
DATE_HINTS = ("تاريخ", "إصدار", "date")
NUM_HINTS  = ("قيمة", "المستحق", "شيك", "التحويل", "USD", ")USD")

def _pick_excel_engine():
    """Return 'xlsxwriter' or 'openpyxl' if available; else None."""
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

def make_excel_bytes(df: pd.DataFrame) -> bytes | None:
    """Create XLSX with hyperlinks (if engine exists). Returns None if no engine."""
    engine = _pick_excel_engine()
    if engine is None:
        return None

    buf = BytesIO()
    df_x = df.copy()
    with pd.ExcelWriter(buf, engine=engine) as writer:
        sheet = "البيانات"
        df_x.to_excel(writer, index=False, sheet_name=sheet, startrow=1, header=False)

        if engine == "xlsxwriter":
            wb  = writer.book
            ws  = writer.sheets[sheet]
            fmt_text = wb.add_format({"align": "right"})
            fmt_date = wb.add_format({"align": "right", "num_format": "yyyy-mm-dd"})
            fmt_num  = wb.add_format({"align": "right", "num_format": "#,##0.00"})
            fmt_link = wb.add_format({"font_color": "blue", "underline": 1, "align": "right"})

            # headers
            for col_num, col_name in enumerate(df_x.columns):
                ws.write(0, col_num, col_name, fmt_text)

            # cells
            for idx, col in enumerate(df_x.columns):
                series = df_x[col]
                max_len = max([len(str(col))] + [len(str(v)) for v in series.values])
                width = min(max_len + 4, 60)

                if "رابط" in col:
                    for row_num, val in enumerate(series, start=1):
                        sval = "" if pd.isna(val) else str(val)
                        if sval.startswith(("http://", "https://")):
                            ws.write_url(row_num, idx, sval, fmt_link, string="فتح الرابط")
                        else:
                            ws.write(row_num, idx, sval, fmt_text)
                    ws.set_column(idx, idx, max(20, width), fmt_link)

                elif pd.api.types.is_datetime64_any_dtype(series):
                    for row_num, val in enumerate(series, start=1):
                        if pd.notna(val):
                            ws.write_datetime(row_num, idx, pd.to_datetime(val), fmt_date)
                        else:
                            ws.write_blank(row_num, idx, None, fmt_text)
                    ws.set_column(idx, idx, max(14, width), fmt_date)

                elif pd.api.types.is_numeric_dtype(series):
                    for row_num, val in enumerate(series, start=1):
                        if pd.notna(val):
                            ws.write_number(row_num, idx, float(val), fmt_num)
                        else:
                            ws.write_blank(row_num, idx, None, fmt_text)
                    ws.set_column(idx, idx, max(14, width), fmt_num)

                else:
                    for row_num, val in enumerate(series, start=1):
                        ws.write(row_num, idx, "" if pd.isna(val) else str(val), fmt_text)
                    ws.set_column(idx, idx, width, fmt_text)

        else:
            # openpyxl simple path
            df_x.to_excel(writer, index=False, sheet_name=sheet)

    buf.seek(0)
    return buf.getvalue()

def make_csv_utf8(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")




# try these fonts in order
AR_FONT_CANDIDATES = [
    "assets/Cairo-Regular.ttf",
    "assets/Amiri-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",  # often present on Linux
]
_AR_RE = re.compile(r'[\u0600-\u06FF]')

def _first_existing_font_path():
    for p in AR_FONT_CANDIDATES:
        if os.path.exists(p):
            return p
    return None

def register_arabic_font() -> tuple[str, bool]:
    """
    Register an Arabic-capable TTF if found.
    Returns (font_name, supports_arabic).
    """
    p = _first_existing_font_path()
    if p:
        name = os.path.splitext(os.path.basename(p))[0]  # e.g., Cairo-Regular -> Cairo-Regular
        try:
            pdfmetrics.registerFont(TTFont(name, p))
            return name, True
        except Exception:
            pass
    # fallback
    return "Helvetica", False  # will render but no Arabic glyphs

def looks_arabic(s: str) -> bool:
    return bool(_AR_RE.search(str(s or "")))

def shape_arabic(s: str) -> str:
    try:
        return get_display(arabic_reshaper.reshape(str(s)))
    except Exception:
        return str(s)

def make_pdf_bytes(df: pd.DataFrame, max_col_width: int = 120) -> bytes:
    """
    Landscape A4 PDF mirroring the on-screen table:
    - Preserves DataFrame column order
    - Shapes Arabic text when an Arabic font is available
    - Columns with 'رابط' or 'link' are rendered as clickable 'فتح الرابط' (or 'Open link' if Arabic font missing)
    - Auto column widths, scaled to fit page
    """
    buf = BytesIO()

    font_name, arabic_ok = register_arabic_font()
    link_label = "فتح الرابط" if arabic_ok else "Open link"

    page = landscape(A4)
    doc = SimpleDocTemplate(
        buf,
        pagesize=page,
        rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20
    )

    # Styles
    title_style = ParagraphStyle(
        name="Title",
        fontName=font_name,
        fontSize=16,
        leading=20,
        alignment=1  # CENTER
    )

    # Title text with file name between brackets
    base_title = "قاعدة البيانات والتقارير المالية"
    title_text = f"{base_title} ({pdf_name})" if pdf_name else base_title
    if arabic_ok:
        title_text = shape_arabic(title_text)

    # ---------------- rest of the code stays the same ----------------
    header_paragraphs = []
    for col in df.columns:
        text = shape_arabic(col) if arabic_ok and looks_arabic(col) else str(col)
        header_paragraphs.append(Paragraph(text, title_style))

    page = landscape(A4)
    doc = SimpleDocTemplate(
        buf,
        pagesize=page,
        rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20
    )

    # Styles
    title_style = ParagraphStyle(name="Title", fontName=font_name, fontSize=16, leading=20, alignment=1)  # CENTER
    hdr_style   = ParagraphStyle(name="Hdr",   fontName=font_name, fontSize=10, textColor=colors.whitesmoke, alignment=1)
    cell_rtl    = ParagraphStyle(name="CellR", fontName=font_name, fontSize=9, leading=12, alignment=2)  # RIGHT
    cell_ltr    = ParagraphStyle(name="CellL", fontName=font_name, fontSize=9, leading=12, alignment=0)  # LEFT
    cell_link   = ParagraphStyle(name="CellK", fontName=font_name, fontSize=9, leading=12, alignment=1,
                                 textColor=colors.HexColor("#1E3A8A"), underline=True)

    # Header row: shape only if we can display Arabic glyphs
    header_paragraphs = []
    for col in df.columns:
        text = shape_arabic(col) if arabic_ok and looks_arabic(col) else str(col)
        header_paragraphs.append(Paragraph(text, hdr_style))

    rows = [header_paragraphs]

    # Determine link columns (Arabic or English header)
    link_cols_idx = [i for i, c in enumerate(df.columns) if ("رابط" in str(c)) or ("link" in str(c).lower())]

    # Body rows
    for _, row in df.iterrows():
        cells = []
        for i, col in enumerate(df.columns):
            val = row[col]
            sval = "" if pd.isna(val) else str(val)

            # clickable link cell
            if i in link_cols_idx and sval.startswith(("http://", "https://")):
                label = shape_arabic(link_label) if arabic_ok else link_label
                cells.append(Paragraph(f'<link href="{sval}">{label}</link>', cell_link))
                continue

            # Arabic vs non-Arabic cell text
            if arabic_ok and looks_arabic(sval):
                cells.append(Paragraph(shape_arabic(sval), cell_rtl))
            else:
                cells.append(Paragraph(sval, cell_ltr))
        rows.append(cells)

    # Column widths (use link label length for link columns)
    avail_w = page[0] - doc.leftMargin - doc.rightMargin
    col_widths = []
    for idx, col in enumerate(df.columns):
        if idx in link_cols_idx:
            max_len = max(len(str(col)), len(link_label))
        else:
            max_len = max(len(str(col)), df[col].astype(str).map(len).max())
        width = min(max_len * 7, max_col_width)  # ~7 px per char
        col_widths.append(width)

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

    # Per-column alignment similar to XLSX
    for idx, col in enumerate(df.columns):
        if any(h in str(col) for h in DATE_HINTS):
            align = "CENTER"
        elif any(h in str(col) for h in NUM_HINTS) or pd.api.types.is_numeric_dtype(df[col]):
            align = "RIGHT"
        elif idx in link_cols_idx:
            align = "CENTER"
        else:
            align = "RIGHT" if (arabic_ok and looks_arabic(col)) else "LEFT"
        style_cmds.append(("ALIGN", (idx, 1), (idx, -1), align))

    table.setStyle(TableStyle(style_cmds))

    elems = [Paragraph(title_text, title_style), Spacer(1, 12), table]

    doc.build(elems)
    buf.seek(0)
    return buf.getvalue()


# ---------- App ----------
def main():
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

    # --- Fetch data from Supabase ---
    df = fetch_data(conn, company_name, project_name, target_table)

    if df.empty:
        st.warning("لا توجد بيانات مطابقة للاختيارات المحددة.")
        return

    # --- Column search ---
    search_column, search_term = create_column_search(df)
    if search_column and search_term:
        df = df[df[search_column].astype(str).str.contains(str(search_term), case=False, na=False)]
        if df.empty:
            st.info("لا توجد نتائج بعد تطبيق معيار البحث.")
            return

    # --- Column config: روابط فقط ---
    column_config = {}
    for col in df.columns:
        if "رابط" in col:
            column_config[col] = st.column_config.LinkColumn(label=col, display_text="فتح الرابط")

    # --- Show table ---
    st.markdown("### البيانات")
    st.dataframe(df, column_config=column_config, use_container_width=True, hide_index=True)

    # --- Downloads ---
    # Excel (only if engine exists)
    xlsx_bytes = make_excel_bytes(df)
    if xlsx_bytes is not None:
        st.download_button(
            label="تنزيل كـ Excel (XLSX) – مُوصى به",
            data=xlsx_bytes,
            file_name=f"{target_table}_{company_name}_{project_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("لتمكين تنزيل Excel مع روابط قابلة للنقر، أضف إلى requirements: `xlsxwriter` أو `openpyxl`.")

    # CSV (UTF-8 with BOM for Arabic in Excel)
    csv_bytes = make_csv_utf8(df)
    st.download_button(
        label="تنزيل كـ CSV (UTF-8)",
        data=csv_bytes,
        file_name=f"{target_table}_{company_name}_{project_name}.csv",
        mime="text/csv",
    )
    pdf_title = f"{target_table}_{company_name}_{project_name}"
    pdf_bytes = make_pdf_bytes(df, pdf_name=pdf_title)
    st.download_button(
    label="تنزيل كـ PDF (عربي)",
    data=pdf_bytes,
    file_name=f"{pdf_title}.pdf",
    mime="application/pdf",
)




if __name__ == "__main__":
    main()
