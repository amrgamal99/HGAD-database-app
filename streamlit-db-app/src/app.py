import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
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
        # write data from row 1; write headers manually
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
            # openpyxl path: simple write (no hyperlink styling API here)
            df_x.to_excel(writer, index=False, sheet_name=sheet)

    buf.seek(0)
    return buf.getvalue()

def make_csv_utf8(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

def make_tsv_utf16(df: pd.DataFrame) -> bytes:
    tsv_str = df.to_csv(index=False, sep="\t")
    return tsv_str.encode("utf-16")
# ======= PDF (RTL Arabic) helpers with column chunking =======
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import arabic_reshaper
from bidi.algorithm import get_display

AR_FONT_PATH = "assets/Cairo-Regular.ttf"  # عدّل المسار لو لزم

def register_arabic_font():
    """Register an Arabic TTF font once; fallback to Helvetica if not found."""
    try:
        pdfmetrics.registerFont(TTFont("Cairo", AR_FONT_PATH))
        return "Cairo"
    except Exception:
        return None

def ar(text: str) -> str:
    """Shape + bidi for Arabic strings."""
    if text is None:
        return ""
    try:
        reshaped = arabic_reshaper.reshape(str(text))
        return get_display(reshaped)
    except Exception:
        return str(text)

def chunk_columns(df: pd.DataFrame, max_cols: int) -> list[list[str]]:
    """Split DataFrame columns into chunks of at most max_cols (from right to left)."""
    cols = df.columns.tolist()
    # للـ RTL: نبدأ من اليمين، فنقسّم من النهاية
    chunks = []
    while cols:
        chunk = cols[-max_cols:]  # آخر max_cols (يمين)
        cols = cols[:-max_cols]
        chunks.append(chunk)
    return chunks  # كل عنصر قائمة بأسماء الأعمدة

def df_to_rtl_table_data_for_cols(df: pd.DataFrame, cols: list[str]) -> list[list[str]]:
    """Build a RTL table (headers + rows) for a subset of columns."""
    headers = [ar(c) for c in cols][::-1]  # RTL: اعكس العناوين
    data = [headers]
    for _, row in df.iterrows():
        cells = [ar(row[c]) for c in cols]     # نصوص مع تشكيل
        data.append(cells[::-1])               # RTL: اعكس ترتيب الخلايا
    return data

def make_pdf_bytes(df: pd.DataFrame, max_cols_per_table: int = 8) -> bytes:
    """
    Create a landscape A4 PDF with Arabic RTL table(s).
    If the DataFrame is wider than max_cols_per_table, it will be split across multiple tables/pages.
    """
    buf = BytesIO()
    font_name = register_arabic_font()

    # Page + styles
    page = landscape(A4)
    doc = SimpleDocTemplate(
        buf,
        pagesize=page,
        rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24
    )

    title_style = ParagraphStyle(
        name="TitleAR",
        fontName=font_name or "Helvetica",
        fontSize=18,
        leading=22,
        alignment=2,  # RIGHT
    )
    table_font = font_name or "Helvetica"

    # Split into column chunks
    col_chunks = chunk_columns(df, max_cols=max_cols_per_table)

    elements = []
    # Title once
    elements.append(Paragraph(ar("قاعدة البيانات والتقارير المالية"), title_style))
    elements.append(Spacer(1, 10))

    for idx, cols in enumerate(col_chunks, start=1):
        data = df_to_rtl_table_data_for_cols(df, cols)

        # Compute equal col widths to fit page width
        avail_w = page[0] - doc.leftMargin - doc.rightMargin
        ncols = len(cols)
        col_width = max(60, avail_w / max(ncols, 1))  # حد أدنى 60
        col_widths = [col_width] * ncols

        tbl = Table(data, repeatRows=1, colWidths=col_widths)
        tbl.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), table_font),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),

            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 6),

            ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ]))

        # Subtitle for chunk if multiple pages
        if len(col_chunks) > 1:
            sub = Paragraph(ar(f"الأعمدة {idx} من {len(col_chunks)}"), ParagraphStyle(
                name="SubAR",
                fontName=font_name or "Helvetica",
                fontSize=10,
                alignment=2,
            ))
            elements.append(sub)
            elements.append(Spacer(1, 6))

        elements.append(tbl)

        # Page break between chunks (except after last one)
        if idx < len(col_chunks):
            elements.append(PageBreak())

    doc.build(elements)
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

    # CSV / TSV
    csv_bytes = make_csv_utf8(df)
    st.download_button(
        label="تنزيل كـ CSV (UTF-8)",
        data=csv_bytes,
        file_name=f"{target_table}_{company_name}_{project_name}.csv",
        mime="text/csv",
    )
    # --- PDF (Arabic RTL with embedded Cairo font) ---
    pdf_bytes = make_pdf_bytes(df)
    st.download_button(
        label="تنزيل كـ PDF (عربي)",
        data=pdf_bytes,
        file_name=f"{target_table}_{company_name}_{project_name}.pdf",
        mime="application/pdf",
    )



if __name__ == "__main__":
    main()
