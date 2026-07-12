import re
import streamlit as st
import pandas as pd
from typing import Optional, Tuple
from db.connection import fetch_companies, fetch_projects_by_company, fetch_type_last_edit_dates


def _normalize_last_edit(value):
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "nat"}:
        return None
    try:
        parsed = pd.to_datetime(text, errors="coerce")
        if pd.isna(parsed):
            return text
        return parsed.strftime("%d-%m-%Y")
    except Exception:
        return text


def _contains_arabic(text: str) -> bool:
    return bool(re.search(r"[\u0600-\u06FF]", str(text)))


def _format_option_label(name, last_edit):
    last_edit = _normalize_last_edit(last_edit)
    return name


def _render_last_edit_label(last_edit):
    if not last_edit:
        return
    st.markdown(
        f'<div class="last-edit-note">{last_edit}</div>',
        unsafe_allow_html=True,
    )


def create_factory_dropdown() -> Optional[str]:
    display_to_factory = {
        "الكل": None,
        "التجمع": "التجمع",
        "بدر": "بدر",
    }
    display_choice = st.selectbox(
        "اسم المصنع",
        options=list(display_to_factory.keys()),
        index=0,        key="factory_select",        help="اختر المصنع أو اترك الكل لعرض جميع الشركات",
    )
    return display_to_factory.get(display_choice)


def create_company_dropdown(conn, factory_name: Optional[str] = None):
    companies_df = fetch_companies(conn, factory_name=factory_name)
    if companies_df.empty or "اسم الشركة" not in companies_df.columns:
        message = "لا توجد شركات مطابقة للمصنع المحدد." if factory_name else "لا توجد شركات." 
        st.info(message)
        return None

    companies_df = companies_df.copy()
    companies_df["آخر تعديل"] = companies_df.get("آخر تعديل", None)
    companies_df["آخر تعديل"] = companies_df["آخر تعديل"].replace({pd.NaT: None}).astype(object)

    rows = companies_df.loc[companies_df["اسم الشركة"].notna(), ["اسم الشركة", "آخر تعديل"]]
    rows = rows.drop_duplicates(subset=["اسم الشركة"], keep="last")
    rows = rows.sort_values("اسم الشركة", key=lambda s: s.str.lower())
    options = [ (row["اسم الشركة"], row["آخر تعديل"]) for _, row in rows.iterrows() ]

    query = st.text_input("🔍 اكتب بداية اسم الشركة", value="", placeholder="اكتب بداية اسم الشركة ...", key="company_search")
    if query:
        q = str(query).strip().lower()
        filtered = [opt for opt in options if opt[0].lower().startswith(q)]
    else:
        filtered = options

    if not filtered:
        st.info(f"لا توجد شركات تبدأ بـ «{query}»") if query else st.info("لا توجد شركات.")
        return None

    selected = st.selectbox(
        "اختر الشركة",
        options=filtered,
        index=0 if filtered else None,
        format_func=lambda x: _format_option_label(x[0], x[1]),
        key="company_select",
    )
    if selected and selected[1]:
        _render_last_edit_label(selected[1])
    return selected[0]

def create_project_dropdown(conn, company_name: str):
    if not company_name:
        return None
    projects_df = fetch_projects_by_company(conn, company_name)
    projects_df = projects_df.copy()
    projects_df["آخر تعديل"] = projects_df.get("آخر تعديل", None)
    projects_df["آخر تعديل"] = projects_df["آخر تعديل"].replace({pd.NaT: None}).astype(object)

    rows = projects_df.loc[projects_df["اسم المشروع"].notna(), ["اسم المشروع", "آخر تعديل"]]
    rows = rows.drop_duplicates(subset=["اسم المشروع"], keep="first")
    rows = rows.sort_values("اسم المشروع", key=lambda s: s.str.lower())
    options = [(row["اسم المشروع"], row["آخر تعديل"]) for _, row in rows.iterrows()]

    if not options:
        return None

    selected = st.selectbox(
        "اختر المشروع",
        options=options,
        index=0,
        format_func=lambda x: _format_option_label(x[0], x[1]),
        placeholder="— اختر —",
    )
    if selected and selected[1]:
        _render_last_edit_label(selected[1])
    return selected[0]

def create_type_dropdown(conn) -> Tuple[Optional[str], Optional[str]]:
    # إضافة "تقرير مالي" كخيار جديد يفعّل عرض الـ Views
    display_to_key = {
        "تقرير مالي": "financial_report",
        "العقود": "contract",
        "خطابات الضمان": "guarantee",
        "المستخلصات": "invoice",
        "الشيكات / التحويلات": "checks",
        "مواد اوليه و مقاولين باطن": "supplier_costs",
        "شهادة تامينات": "social_insurance_certificate",  # <-- note space: "شهادة تامينات"
    }
    options = [
        (display_name, key)
        for display_name, key in display_to_key.items()
    ]
    selected = st.selectbox(
        "اختر نوع البيانات",
        options=options,
        index=0 if options else None,
        format_func=lambda x: x[0],
        key="type_select",
    )
    return selected[0], selected[1]

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
