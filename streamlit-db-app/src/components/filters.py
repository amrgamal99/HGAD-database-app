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
    # إضافة "تقرير مالي" كخيار جديد يفعّل عرض الـ Views
    display_to_key = {
        "تقرير مالي": "financial_report",
        "العقود": "contract",
        "خطابات الضمان": "guarantee",
        "المستخلصات": "invoice",
        "الشيكات / التحويلات": "checks",
        "شهادة تامينات": "social_insurance_certificate",  # <-- note space: "شهادة تامينات"
    }
    display_list = list(display_to_key.keys())
    display_choice = st.selectbox("اختر نوع البيانات", options=display_list, index=0 if display_list else None, placeholder="— اختر —")
    return display_choice, display_to_key.get(display_choice)

def create_column_search(df: pd.DataFrame):
    if df.empty:
        return None, None
    col = st.selectbox("اختَر عمودًا للبحث", options=df.columns.tolist(), index=0)
    term = st.text_input("كلمة/عبارة للبحث")
    return col, term

def create_date_range():
    c1, c2 = st.columns(2)
    with c1:
        d_from = st.date_input("من تاريخ", value=None, format="YYYY-MM-DD")
    with c2:
        d_to = st.date_input("إلى تاريخ", value=None, format="YYYY-MM-DD")
    # إرجاع None لو لم يُحدد المستخدم
    d_from = pd.to_datetime(d_from).date() if d_from else None
    d_to = pd.to_datetime(d_to).date() if d_to else None
    return d_from, d_to
