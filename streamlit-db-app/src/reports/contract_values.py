from __future__ import annotations

from typing import Optional, Tuple

import pandas as pd
import streamlit as st
from supabase import Client

from db.connection import fetch_data
from utils.data_helpers import (
    build_contract_value_details,
    build_contract_value_summary,
    prepare_contract_values_dataframe,
)


@st.cache_data(show_spinner=False)
def fetch_contract_value_report_data(
    supabase: Client,
    company_name: Optional[str] = None,
    project_name: Optional[str] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
) -> pd.DataFrame:
    """Fetch and normalize contract-value rows from the database."""
    if supabase is None:
        return pd.DataFrame()

    try:
        raw_df = fetch_data(supabase, company_name or "", project_name or "", "contract")
    except Exception:
        return pd.DataFrame()

    if raw_df.empty:
        return pd.DataFrame()

    return prepare_contract_values_dataframe(raw_df, date_from=date_from, date_to=date_to)


def _render_summary_cards(summary_df: pd.DataFrame) -> None:
    if summary_df is None or summary_df.empty:
        return

    total_value = float(summary_df["قيمة العقود"].sum()) if "قيمة العقود" in summary_df.columns else 0.0
    st.metric("إجمالي قيمة العقود", f"{total_value:,.2f}")

    cols = st.columns(2)
    for idx, row in summary_df.iterrows():
        col = cols[idx % 2]
        col.metric(row["اسم المصنع"], f"{float(row['قيمة العقود']):,.2f}")


def _render_table(df: pd.DataFrame, title: str) -> None:
    if df is None or df.empty:
        st.info(f"لا توجد بيانات متاحة لـ {title}.")
        return

    st.markdown(f"<h3 style='text-align:right'>{title}</h3>", unsafe_allow_html=True)
    st.dataframe(df, use_container_width=True, hide_index=True)


def render_contract_values_report(
    conn: Optional[Client],
    company_name: Optional[str] = None,
    project_name: Optional[str] = None,
) -> None:
    """Render the contract values report within the financial reports mode."""
    st.markdown("<h2 style='text-align:right'>حصر قيمة العقود</h2>", unsafe_allow_html=True)

    if conn is None:
        st.error("تعذر الاتصال بقاعدة البيانات.")
        return

    if not company_name or not project_name:
        st.info("يرجى اختيار الشركة والمشروع من الشريط الجانبي لعرض هذا التقرير.")
        return

    c1, c2 = st.columns(2)
    with c1:
        date_from = st.date_input("من تاريخ", value=None, key="contract_values_from", format="YYYY-MM-DD")
    with c2:
        date_to = st.date_input("إلى تاريخ", value=None, key="contract_values_to", format="YYYY-MM-DD")

    date_from_value = date_from.isoformat() if date_from else None
    date_to_value = date_to.isoformat() if date_to else None

    try:
        df = fetch_contract_value_report_data(
            conn,
            company_name=company_name,
            project_name=project_name,
            date_from=date_from_value,
            date_to=date_to_value,
        )
    except Exception as exc:  # pragma: no cover - UI guard
        st.exception(exc)
        return

    if df.empty:
        st.warning("لا توجد عقود ضمن النطاق المحدد.")
        return

    summary_df = build_contract_value_summary(df)
    details_df = build_contract_value_details(df)

    _render_summary_cards(summary_df)
    _render_table(summary_df, "مقارنة المصانع")
    _render_table(details_df, "تفاصيل العقود")

    export_df = details_df.copy()
    export_df["رابط نسخة العقد"] = export_df["رابط نسخة العقد"].apply(
        lambda value: value if pd.notna(value) and str(value).strip() else ""
    )

    csv_bytes = export_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        label="⬇️ تنزيل تفاصيل العقود (CSV)",
        data=csv_bytes,
        file_name="حصر_قيمه_عقود.csv",
        mime="text/csv",
    )
