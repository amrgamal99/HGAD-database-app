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
    st.set_page_config(layout="wide", page_title="نظام التقارير المالية")
    # تفعيل اتجاه RTL وخط عربي بسيط
    st.markdown("""
        <style>
            html, body, [data-testid="stAppViewContainer"] * {
                direction: rtl; text-align: right;
                font-family: "Segoe UI", "Cairo", "Noto Kufi Arabic", Tahoma, sans-serif;
            }
            .block-container { padding-top: 1rem; }
        </style>
    """, unsafe_allow_html=True)

    st.title("عارض قاعدة البيانات والتقارير المالية")

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
