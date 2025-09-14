import streamlit as st
import pandas as pd
from db.connection import fetch_companies, fetch_projects_by_company

def create_company_dropdown(conn):
    companies_df = fetch_companies(conn)
    companies = companies_df["اسم الشركة"].tolist()
    return st.selectbox("اختر الشركة", options=companies, index=0 if companies else None, placeholder="— اختر —")

def create_project_dropdown(conn, company_name: str):
    if not company_name:
        return None
    projects_df = fetch_projects_by_company(conn, company_name)
    projects = projects_df["اسم المشروع"].tolist()
    return st.selectbox("اختر المشروع", options=projects, index=0 if projects else None, placeholder="— اختر —")

def create_type_dropdown():
    # الأسماء المعروضة بالعربي ← تُحوّل لاسم الجدول الفعلي
    display_to_table = {
        "العقود": "contract",
        "خطابات الضمان": "guarantee",
        "المستخلصات": "invoice",
        "الشيكات / التحويلات": "checks",
    }
    display_list = list(display_to_table.keys())
    display_choice = st.selectbox("اختر نوع البيانات", options=display_list, index=0 if display_list else None, placeholder="— اختر —")
    target_table = display_to_table.get(display_choice)
    return display_choice, target_table

def create_column_search(df: pd.DataFrame):
    if df.empty:
        return None, None
    col = st.selectbox("اختَر عمودًا للبحث", options=df.columns.tolist(), index=0)
    term = st.text_input("كلمة/عبارة للبحث")
    return col, term
