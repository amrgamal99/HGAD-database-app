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
    st.markdown("""
        <style>
            html, body, [data-testid="stAppViewContainer"] * {
                direction: rtl;
                text-align: right;
                font-family: "Segoe UI", "Cairo", "Noto Kufi Arabic", Tahoma, sans-serif;
                white-space: normal !important;
                word-wrap: break-word !important;
            }
            .block-container {
                padding-top: 1rem;
            }
            [data-testid="stAlert"] {
                white-space: normal !important;
            }
        </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <h1 style="text-align:center; direction:rtl; color:#1E3A8A; font-family:'Cairo', sans-serif;">
        عارض قاعدة البيانات والتقارير المالية  
        <span style="font-size:20px; color:#555;">(HGAD Company)</span>
    </h1>
    <hr style="border: 2px solid #1E3A8A; border-radius: 5px;"/>
""", unsafe_allow_html=True)

    conn = get_db_connection()
    if conn is None:
        st.error("فشل الاتصال بقاعدة البيانات. يُرجى مراجعة بيانات الاتصال والتأكد من تشغيل الخادم.")
        return

    # الشريط الجانبي لعوامل التصفية
    with st.sidebar:
        st.title("عوامل التصفية")
        company_name = create_company_dropdown(conn)
        project_name = create_project_dropdown(conn, company_name)
        type_selection, target_table = create_type_dropdown()

    # منطقة المحتوى الرئيسية
    if not company_name or not project_name or not target_table:
        st.info("برجاء اختيار الشركة والمشروع ونوع البيانات من الشريط الجانبي لعرض النتائج.")
        return

    # جلب البيانات
    df = fetch_data(conn, company_name, project_name, target_table)

    if df.empty:
        st.warning("لا توجد بيانات مطابقة للاختيارات المحددة.")
        return

    # فلترة بالبحث في عمود
    search_column, search_term = create_column_search(df)
    if search_column and search_term:
        df = df[df[search_column].astype(str).str.contains(str(search_term), case=False, na=False)]

    # جعل أعمدة الروابط قابلة للنقر (أي عمود يحتوي كلمة 'رابط')
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
