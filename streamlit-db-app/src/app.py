import streamlit as st
import pandas as pd
from db.connection import get_db_connection, fetch_data, fetch_companies, fetch_projects_by_company
from components.filters import (
    create_company_dropdown,
    create_project_dropdown,
    create_type_dropdown,
    create_column_search
)

def main():
    st.set_page_config(
        page_title="عارض قاعدة البيانات والتقارير المالية | HGAD",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # ===== Global Styles (RTL + anti-vertical text + toolbar/side fixes) =====
    st.markdown("""
        <style>
            /* Root: enforce RTL + normal wrapping, avoid horizontal overflow */
            html, body {
                direction: rtl !important;
                text-align: right !important;
                font-family: "Cairo","Noto Kufi Arabic","Segoe UI",Tahoma,sans-serif !important;
                white-space: normal !important;
                word-wrap: break-word !important;
                overflow-x: hidden !important; /* ✅ remove horizontal scroll/ghost columns */
            }

            /* App container content only (safe to style deeply) */
            [data-testid="stAppViewContainer"] * {
                direction: rtl !important;
                text-align: right !important;
                white-space: normal !important;
                word-wrap: break-word !important;
                writing-mode: horizontal-tb !important;
                text-orientation: mixed !important;
            }

            /* ✅ Keep Streamlit top toolbar LTR so it doesn't stack vertically */
            [data-testid="stToolbar"], [data-testid="stToolbar"] * {
                direction: ltr !important;
                text-align: left !important;
                writing-mode: horizontal-tb !important;
                text-orientation: mixed !important;
                white-space: normal !important;
            }

            /* Sidebar theme + fix vertical text */
            [data-testid="stSidebar"] {
                background: #0f172a; /* slate-900 */
                color: #e5e7eb;      /* gray-200 */
                border-left: 1px solid #1f2937;
            }
            [data-testid="stSidebar"] * {
                direction: rtl !important;
                text-align: right !important;
                white-space: normal !important;
                writing-mode: horizontal-tb !important;
                text-orientation: mixed !important;
            }
            /* Inputs inside sidebar: ensure RTL typing */
            [data-testid="stSidebar"] input,
            [data-testid="stSidebar"] textarea,
            [data-testid="stSidebar"] select {
                direction: rtl !important;
                text-align: right !important;
            }

            /* Alerts wrap nicely */
            [data-testid="stAlert"] { white-space: normal !important; }

            /* Container spacing */
            .block-container { padding-top: 1rem; }

            /* Header styles */
            .hgad-title {
                text-align: center;
                direction: rtl;
                color: #1E3A8A;
                font-weight: 800;
                margin-top: .25rem;
                margin-bottom: .25rem;
                line-height: 1.3;
            }
            .hgad-sub {
                display: inline-block;
                font-size: 20px;
                color: #4b5563;  /* gray-600 */
                font-weight: 600;
                margin-inline-start: .35rem;
            }
            .hgad-divider {
                border: 0;
                height: 2px;
                background: linear-gradient(to left, transparent, #1E3A8A, transparent);
                margin: 10px auto 20px auto;
                max-width: 900px;
                border-radius: 6px;
            }
        </style>
    """, unsafe_allow_html=True)

    # ===== Header (formal & beautiful) =====
    st.markdown("""
        <h1 class="hgad-title">
            عارض قاعدة البيانات والتقارير المالية
            <span class="hgad-sub">| HGAD Company</span>
        </h1>
        <hr class="hgad-divider"/>
    """, unsafe_allow_html=True)

    # ===== DB Connection =====
    conn = get_db_connection()
    if conn is None:
        st.error("فشل الاتصال بقاعدة البيانات. يُرجى مراجعة بيانات الاتصال والتأكد من تشغيل الخادم.")
        return

    # ===== Sidebar Filters =====
    with st.sidebar:
        st.title("عوامل التصفية")
        company_name = create_company_dropdown(conn)
        project_name = create_project_dropdown(conn, company_name)
        type_selection, target_table = create_type_dropdown()

    # ===== Main Content =====
    if not company_name or not project_name or not target_table:
        st.info("برجاء اختيار الشركة والمشروع ونوع البيانات من الشريط الجانبي لعرض النتائج.")
        return

    # Fetch data
    df = fetch_data(conn, company_name, project_name, target_table)

    if df.empty:
        st.warning("لا توجد بيانات مطابقة للاختيارات المحددة.")
        return

    # Column search
    search_column, search_term = create_column_search(df)
    if search_column and search_term:
        df = df[df[search_column].astype(str).str.contains(str(search_term), case=False, na=False)]

    # Make Arabic link columns clickable (any column containing 'رابط')
    column_config = {}
    for col in df.columns:
        if 'رابط' in col:
            column_config[col] = st.column_config.LinkColumn(
                label=col,
                display_text="فتح الرابط"
            )

    st.dataframe(df, column_config=column_config, use_container_width=True)

if __name__ == "__main__":
    main()
