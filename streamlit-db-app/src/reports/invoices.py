"""
تقرير المستخلصات (Invoices report).

Two views, both fed by the same query + the same summary helpers:
  - "period": مستخلصات خلال فترة زمنية
  - "latest": آخر مستخلص (مع عرض كل المستخلصات المتعادلة في التاريخ)
"""

from __future__ import annotations

from io import BytesIO
from typing import Optional

import pandas as pd
import streamlit as st
from supabase import Client

from db.connection import fetch_invoice_report_data
from utils.data_helpers import (
    DATE_COLUMN,
    FACTORY_COLUMN,
    drop_id_columns,
    invoice_latest_rows,
    invoice_numeric_summary,
    invoice_work_volume,
)

VIEW_PERIOD = "period"
VIEW_LATEST = "latest"

VIEW_TITLES = {
    VIEW_PERIOD: "مستخلصات خلال فترة زمنية",
    VIEW_LATEST: "آخر مستخلص",
}


# ─────────────────────────────────────────────────────────────────────────
# Small UI building blocks
# ─────────────────────────────────────────────────────────────────────────

def _to_iso_date(value) -> Optional[str]:
    if not value:
        return None
    parsed = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(parsed) else parsed.date().isoformat()


def _period_inputs() -> tuple[Optional[str], Optional[str]]:
    start_col, end_col = st.columns(2)
    with start_col:
        start = st.date_input("من تاريخ", value=None, key="invoice_date_from")
    with end_col:
        end = st.date_input("إلى تاريخ", value=None, key="invoice_date_to")
    return _to_iso_date(start), _to_iso_date(end)


def _format_display_dates(df: pd.DataFrame) -> pd.DataFrame:
    display_df = df.copy()
    if DATE_COLUMN in display_df.columns:
        display_df[DATE_COLUMN] = (
            pd.to_datetime(display_df[DATE_COLUMN], errors="coerce")
            .dt.strftime("%Y-%m-%d")
            .fillna("")
        )
    return display_df


def _to_excel_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        writer.sheets[sheet_name].right_to_left()
    return buffer.getvalue()


# ─────────────────────────────────────────────────────────────────────────
# Summary sections
# ─────────────────────────────────────────────────────────────────────────

def _render_totals_section(df: pd.DataFrame) -> None:
    st.markdown("#### إجمالي كل الأعمدة")
    overall_totals = invoice_numeric_summary(df, group_by_factory=False)
    st.dataframe(overall_totals, use_container_width=True, hide_index=True)

    st.markdown("#### إجمالي كل الأعمدة على حسب كل مصنع")
    factory_totals = invoice_numeric_summary(df, group_by_factory=True)
    if factory_totals.empty:
        st.info("لا تتوفر بيانات المصنع لهذه المستخلصات.")
    else:
        st.dataframe(factory_totals, use_container_width=True, hide_index=True)


def _render_work_volume_section(df: pd.DataFrame) -> None:
    st.markdown("#### حجم الأعمال")
    overall_volume = invoice_work_volume(df, group_by_factory=False)
    st.metric("حجم الأعمال الإجمالي", f"{overall_volume:,.2f}")

    factory_volume = invoice_work_volume(df, group_by_factory=True)
    st.markdown("##### حجم الأعمال لكل مصنع")
    if isinstance(factory_volume, pd.DataFrame) and not factory_volume.empty:
        st.dataframe(factory_volume, use_container_width=True, hide_index=True)
    else:
        st.info("لا تتوفر بيانات المصنع لهذه المستخلصات.")


def _render_summary(df: pd.DataFrame) -> None:
    _render_totals_section(df)
    st.divider()
    _render_work_volume_section(df)


# ─────────────────────────────────────────────────────────────────────────
# View-specific data loading
# ─────────────────────────────────────────────────────────────────────────

def _load_period_view(conn: Client) -> tuple[pd.DataFrame, bool]:
    """Returns (dataframe, ok) — ok=False means a validation error was shown."""
    date_from, date_to = _period_inputs()
    if date_from and date_to and date_from > date_to:
        st.error("تاريخ البداية يجب أن يسبق تاريخ النهاية.")
        return pd.DataFrame(), False

    df = fetch_invoice_report_data(conn, date_from=date_from, date_to=date_to)
    return df, True


def _load_latest_view(conn: Client) -> tuple[pd.DataFrame, bool]:
    df = fetch_invoice_report_data(conn)
    return invoice_latest_rows(df), True


# ─────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────

def render_invoices_report(conn: Optional[Client], view_key: str = VIEW_PERIOD) -> None:
    st.markdown("<h2 style='text-align:right'>المستخلصات</h2>", unsafe_allow_html=True)

    if conn is None:
        st.error("تعذر الاتصال بقاعدة البيانات.")
        return

    if view_key == VIEW_PERIOD:
        df, ok = _load_period_view(conn)
    else:
        df, ok = _load_latest_view(conn)

    if not ok:
        return

    st.markdown(f"### {VIEW_TITLES.get(view_key, '')}")

    if df.empty:
        st.info("لا توجد مستخلصات مطابقة.")
        return

    _render_summary(df)

    st.divider()
    st.markdown("### تفاصيل المستخلصات")
    display_df = drop_id_columns(_format_display_dates(df))
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    st.download_button(
        label="تنزيل Excel",
        data=_to_excel_bytes(display_df, sheet_name="المستخلصات"),
        file_name="تقرير_المستخلصات.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="invoice_excel_download",
    )