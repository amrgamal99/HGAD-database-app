from __future__ import annotations

from io import BytesIO
from typing import Optional

import pandas as pd
import streamlit as st
from supabase import Client

from db.connection import fetch_invoice_report_data
from utils.data_helpers import invoice_latest_rows, invoice_numeric_summary


def _iso_date(value) -> Optional[str]:
    if value is None or value == "":
        return None
    parsed = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(parsed) else parsed.date().isoformat()


def _date_filters(date_from, date_to):
    left, right = st.columns(2)
    with left:
        start = st.date_input("من تاريخ", value=pd.to_datetime(date_from).date() if date_from else None, key="invoice_from")
    with right:
        end = st.date_input("إلى تاريخ", value=pd.to_datetime(date_to).date() if date_to else None, key="invoice_to")
    return _iso_date(start), _iso_date(end)


def _show_summary(df: pd.DataFrame) -> None:
    overall = invoice_numeric_summary(df)
    volume = float(overall.get("حجم الأعمال", pd.Series([0])).iloc[0]) if not overall.empty else 0.0
    st.metric("حجم الأعمال الإجمالي", f"{volume:,.2f}")
    factory_summary = invoice_numeric_summary(df, group_by_factory=True)
    if not factory_summary.empty:
        st.markdown("### إجمالي كل العواميد على حسب كل مصنع")
        st.dataframe(factory_summary, use_container_width=True, hide_index=True)


def _download_excel(df: pd.DataFrame, filename: str):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="المستخلصات")
        writer.sheets["المستخلصات"].right_to_left()
    st.download_button("تنزيل Excel", buffer.getvalue(), filename=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def render_invoices_report(conn: Optional[Client], view_key: str = "period", date_from=None, date_to=None) -> None:
    st.markdown("<h2 style='text-align:right'>المستخلصات</h2>", unsafe_allow_html=True)
    if conn is None:
        st.error("تعذر الاتصال بقاعدة البيانات.")
        return
    if view_key == "period":
        date_from, date_to = _date_filters(date_from, date_to)
        if date_from and date_to and date_from > date_to:
            st.error("تاريخ البداية يجب أن يسبق تاريخ النهاية.")
            return
        df = fetch_invoice_report_data(conn, date_from=date_from, date_to=date_to)
        title = "مستخلصات خلال فترة زمنية"
    else:
        df = fetch_invoice_report_data(conn)
        df = invoice_latest_rows(df)
        title = "آخر مستخلص"
    if df.empty:
        st.info("لا توجد مستخلصات مطابقة.")
        return
    st.markdown(f"### {title}")
    _show_summary(df)
    display_df = df.copy()
    if "تاريخ إصدار المستخلص" in display_df.columns:
        display_df["تاريخ إصدار المستخلص"] = display_df["تاريخ إصدار المستخلص"].dt.strftime("%Y-%m-%d")
    st.markdown("### تفاصيل المستخلصات")
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    _download_excel(display_df, "تقرير_المستخلصات.xlsx")