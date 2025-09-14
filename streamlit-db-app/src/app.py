import streamlit as st
import pandas as pd
from io import BytesIO

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
/* ===== Force the sidebar ALWAYS open ===== */
[data-testid="stSidebar"] { transform: none !important; visibility: visible !important; width: 340px !important; min-width: 340px !important; }
[data-testid="stSidebar"][aria-expanded="false"] { transform: none !important; visibility: visible !important; }
[data-testid="collapsedControl"],
button[kind="header"],
button[title="Expand sidebar"],
button[title="Collapse sidebar"],
[data-testid="stSidebarCollapseButton"] { display: none !important; }

/* ===== Root RTL & wrapping ===== */
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
[data-testid="stDataFrame"] div[role="row"] {
    font-size: 15px;
}
[data-testid="stDataFrame"] div[role="row"]:nth-child(even) {
    background-color: rgba(255,255,255,0.04);
}
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

# ---------- Download helpers (Excel with real hyperlinks + CSV/TSV encodings) ----------
DATE_HINTS = ("تاريخ", "إصدار", "date")
NUM_HINTS  = ("قيمة", "المستحق", "شيك", "التحويل", "USD", ")USD")

def make_excel_bytes(df: pd.DataFrame) -> bytes:
    """Create XLSX with Arabic-friendly formatting and clickable hyperlinks for any column containing 'رابط'."""
    buf = BytesIO()
    try:
        import xlsxwriter  # noqa: F401
        engine = "xlsxwriter"
    except Exception:
        engine = "openpyxl"

    df_x = df.copy()

    with pd.ExcelWriter(buf, engine=engine) as writer:
        sheet = "البيانات"
        # write data starting from row 1; we'll write headers ourselves to control format
        df_x.to_excel(writer, index=False, sheet_name=sheet, startrow=1, header=False)

        if engine == "xlsxwriter":
            wb  = writer.book
            ws  = writer.sheets[sheet]

            fmt_text = wb.add_format({"align": "right"})
            fmt_date = wb.add_format({"align": "right", "num_format": "yyyy-mm-dd"})
            fmt_num  = wb.add_format({"align": "right", "num_format": "#,##0.00"})
            fmt_link = wb.add_format({"font_color": "blue", "underline": 1, "align": "right"})

            # headers (row 0)
            for col_num, col_name in enumerate(df_x.columns):
                ws.write(0, col_num, col_name, fmt_text)

            # write cells with types & hyperlinks
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

    buf.seek(0)
    return buf.getvalue()

def make_csv_utf8(df: pd.DataFrame) -> bytes:
    """CSV with UTF-8 BOM (works in most modern Excel)."""
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

def make_tsv_utf16(df: pd.DataFrame) -> bytes:
    """TSV with UTF-16 (Excel on Windows opens Arabic correctly almost always)."""
    tsv_str = df.to_csv(index=False, sep="\t")
    return tsv_str.encode("utf-16")

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
            column_config[col] = st.column_config.LinkColumn(
                label=col,
                display_text="فتح الرابط"
            )

    # --- Show table exactly as in DB ---
    st.markdown("### البيانات")
    st.dataframe(df, column_config=column_config, use_container_width=True, hide_index=True)

    # --- Download buttons (Excel with hyperlinks + CSV/TSV encodings) ---
    xlsx_bytes = make_excel_bytes(df)
    st.download_button(
        label="تنزيل كـ Excel (XLSX) – مُوصى به",
        data=xlsx_bytes,
        file_name=f"{target_table}_{company_name}_{project_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    csv_bytes = make_csv_utf8(df)
    st.download_button(
        label="تنزيل كـ CSV (UTF-8)",
        data=csv_bytes,
        file_name=f"{target_table}_{company_name}_{project_name}.csv",
        mime="text/csv",
    )

    tsv_bytes = make_tsv_utf16(df)
    st.download_button(
        label="تنزيل كـ TSV (UTF-16) – فتح مباشر في Excel",
        data=tsv_bytes,
        file_name=f"{target_table}_{company_name}_{project_name}.tsv",
        mime="text/tab-separated-values",
    )

if __name__ == "__main__":
    main()
