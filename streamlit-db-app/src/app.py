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

# ---------- Arabic PDF helpers (make it look like XLSX) ----------
AR_FONT_PATH = "assets/Cairo-Regular.ttf"  # change to your font path if needed
_AR_RE = re.compile(r'[\u0600-\u06FF]')

def register_arabic_font() -> str | None:
    """Register an Arabic TTF font once; return its name or None if missing."""
    try:
        pdfmetrics.registerFont(TTFont("Cairo", AR_FONT_PATH))
        return "Cairo"
    except Exception:
        return None

def looks_arabic(s: str) -> bool:
    return bool(_AR_RE.search(str(s)))

def shape_arabic(s: str) -> str:
    try:
        return get_display(arabic_reshaper.reshape(str(s)))
    except Exception:
        return str(s)

def make_pdf_bytes(df: pd.DataFrame, max_col_width: int = 120) -> bytes:
    """
    Generate a landscape A4 PDF similar to the XLSX export:
    - keeps DataFrame column order,
    - Arabic shaping for Arabic text only,
    - link columns render as clickable 'فتح الرابط',
    - auto column widths with page-fit scaling.
    """
    buf = BytesIO()
    font_name = register_arabic_font() or "Helvetica"

    # Page + doc
    page = landscape(A4)
    doc = SimpleDocTemplate(
        buf,
        pagesize=page,
        rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20
    )

    # Styles
    title_style = ParagraphStyle(name="TitleAR", fontName=font_name, fontSize=16, leading=20, alignment=1)  # CENTER
    hdr_style   = ParagraphStyle(name="HdrAR",   fontName=font_name, fontSize=10, textColor=colors.whitesmoke, alignment=1)
    rtl_cell    = ParagraphStyle(name="CellRTL", fontName=font_name, fontSize=9, leading=12, alignment=2)  # RIGHT
    ltr_cell    = ParagraphStyle(name="CellLTR", fontName=font_name, fontSize=9, leading=12, alignment=0)  # LEFT
    link_cell   = ParagraphStyle(name="CellLink",fontName=font_name, fontSize=9, leading=12, alignment=1, textColor=colors.HexColor("#1E3A8A"), underline=True)

    # 1) Header row (Paragraphs so we can shape Arabic)
    header_paragraphs = []
    for col in df.columns:
        h = shape_arabic(col) if looks_arabic(col) else str(col)
        header_paragraphs.append(Paragraph(h, hdr_style))

    # 2) Body rows
    table_rows = [header_paragraphs]
    link_cols_idx = [i for i, c in enumerate(df.columns) if "رابط" in c]

    for _, row in df.iterrows():
        cells = []
        for i, col in enumerate(df.columns):
            val = row[col]
            sval = "" if pd.isna(val) else str(val)

            # render links as clickable short label
            if i in link_cols_idx and sval.startswith(("http://", "https://")):
                cells.append(Paragraph(f'<link href="{sval}">{shape_arabic("فتح الرابط")}</link>', link_cell))
                continue

            # Arabic vs non-Arabic text
            if looks_arabic(sval):
                cells.append(Paragraph(shape_arabic(sval), rtl_cell))
            else:
                cells.append(Paragraph(sval, ltr_cell))
        table_rows.append(cells)

    # 3) Compute column widths based on visible content (use "فتح الرابط" for link columns)
    avail_w = page[0] - doc.leftMargin - doc.rightMargin
    col_widths = []
    for idx, col in enumerate(df.columns):
        if idx in link_cols_idx:
            max_len = max(len(str(col)), len("فتح الرابط"))
        else:
            col_len = max(len(str(col)), df[col].astype(str).map(len).max())
            max_len = col_len
        width = min(max_len * 7, max_col_width)  # ~7 px per char
        col_widths.append(width)

    total_w = sum(col_widths)
    if total_w > avail_w and total_w > 0:
        scale = avail_w / total_w
        col_widths = [w * scale for w in col_widths]

    # 4) Build table
    table = Table(table_rows, repeatRows=1, colWidths=col_widths)
    style_cmds = [
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
    ]

    # Per-column alignment similar to Excel
    for idx, col in enumerate(df.columns):
        align = "CENTER"
        if any(h in col for h in DATE_HINTS):
            align = "CENTER"
        elif any(h in col for h in NUM_HINTS) or pd.api.types.is_numeric_dtype(df[col]):
            align = "RIGHT"
        elif idx in link_cols_idx:
            align = "CENTER"
        else:
            # if column header is Arabic, default RIGHT; otherwise LEFT
            align = "RIGHT" if looks_arabic(col) else "LEFT"
        style_cmds.append(("ALIGN", (idx, 1), (idx, -1), align))

    table.setStyle(TableStyle(style_cmds))

    elems = [Paragraph(shape_arabic("قاعدة البيانات والتقارير المالية"), title_style), Spacer(1, 12), table]
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

    # PDF (Arabic RTL with clickable links, tuned to look like XLSX)
    pdf_bytes = make_pdf_bytes(df)
    st.download_button(
        label="تنزيل كـ PDF (عربي)",
        data=pdf_bytes,
        file_name=f"{target_table}_{company_name}_{project_name}.pdf",
        mime="application/pdf",
    )

if __name__ == "__main__":
    main()
