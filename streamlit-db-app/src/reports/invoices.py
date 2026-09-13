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
    invoice_latest_rows,
    invoice_numeric_summary,
)

DATE_COLUMN = "تاريخ إصدار المستخلص"
FACTORY_COLUMN = "مصنع"
VIEW_PERIOD = "period"
VIEW_LATEST = "latest"

VIEW_TITLES = {
    VIEW_PERIOD: "مستخلصات خلال فترة زمنية",
    VIEW_LATEST: "آخر مستخلص",
}

_INVOICE_LINK_ALIASES = (
    "رابط نسخة مستخلص",
    "رابط نسخة المستخلص",
    "رابط المستخلص",
    "رابط invoice",
)


# ─────────────────────────────────────────────────────────────────────────
# Small UI building blocks
# ─────────────────────────────────────────────────────────────────────────

def _to_iso_date(value) -> Optional[str]:
    if not value:
        return None
    parsed = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(parsed) else parsed.date().isoformat()


def _period_inputs(
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
) -> tuple[Optional[str], Optional[str]]:
    start_col, end_col = st.columns(2)
    with start_col:
        start = st.date_input(
            "من تاريخ",
            value=pd.to_datetime(date_from).date() if date_from else None,
            key="invoice_date_from",
        )
    with end_col:
        end = st.date_input(
            "إلى تاريخ",
            value=pd.to_datetime(date_to).date() if date_to else None,
            key="invoice_date_to",
        )
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


def _drop_id_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Remove database identifiers from the user-facing invoice table."""
    return df.drop(
        columns=[
            column
            for column in df.columns
            if str(column).strip().lower().endswith("id")
            or str(column).strip().lower() in {"id", "invoiceid"}
        ],
        errors="ignore",
    )


def _prepare_display_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Expose business names and links while hiding internal join columns."""
    display_df = _drop_id_columns(_format_display_dates(df))
    rename_map = {}
    if "اسم المشروع" in display_df.columns:
        rename_map["اسم المشروع"] = "اسم العقد"
    for column in _INVOICE_LINK_ALIASES:
        if column in display_df.columns:
            rename_map[column] = "رابط نسخة مستخلص"
            break
    display_df = display_df.rename(columns=rename_map)
    link_columns = [
        "رابط نسخة مستخلص"
        if "رابط نسخة مستخلص" in display_df.columns
        else column
        for column in ("رابط نسخة مستخلص",)
        if column in display_df.columns
    ]
    return display_df, link_columns


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
    st.markdown(
        "<div class='cv-section-title'>إجمالي كل الأعمدة</div>",
        unsafe_allow_html=True,
    )
    overall_totals = invoice_numeric_summary(df, group_by_factory=False)
    st.dataframe(overall_totals, use_container_width=True, hide_index=True)

    st.markdown(
        "<div class='cv-section-title'>إجمالي كل الأعمدة على حسب كل مصنع</div>",
        unsafe_allow_html=True,
    )
    factory_totals = invoice_numeric_summary(df, group_by_factory=True)
    if factory_totals.empty:
        st.info("لا تتوفر بيانات المصنع لهذه المستخلصات.")
    else:
        st.dataframe(factory_totals, use_container_width=True, hide_index=True)


def _render_work_volume_section(df: pd.DataFrame) -> None:
    st.markdown(
        "<div class='cv-section-title'>حجم الأعمال</div>",
        unsafe_allow_html=True,
    )
    overall_totals = invoice_numeric_summary(df, group_by_factory=False)
    overall_volume = (
        float(overall_totals["حجم الأعمال"].iloc[0])
        if "حجم الأعمال" in overall_totals.columns and not overall_totals.empty
        else 0.0
    )
    st.markdown(
        f"""
        <div class="invoice-total-card">
            <div class="invoice-total-label">حجم الأعمال الإجمالي</div>
            <div class="invoice-total-value">{overall_volume:,.2f}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    factory_volume = invoice_numeric_summary(df, group_by_factory=True)
    if "حجم الأعمال" in factory_volume.columns:
        factory_volume = factory_volume[[FACTORY_COLUMN, "حجم الأعمال"]]
    st.markdown("##### حجم الأعمال لكل مصنع")
    if isinstance(factory_volume, pd.DataFrame) and not factory_volume.empty:
        st.dataframe(factory_volume, use_container_width=True, hide_index=True)
    else:
        st.info("لا تتوفر بيانات المصنع لهذه المستخلصات.")


def _render_summary(df: pd.DataFrame) -> None:
    st.markdown(
        """
        <style>
        .cv-section-title {
            direction: rtl;
            text-align: right;
            color: #e5e7eb;
            font-size: 1.15rem;
            font-weight: 800;
            margin: 18px 0 10px;
            padding-right: 10px;
            border-right: 4px solid #60a5fa;
        }
        .invoice-total-card {
            direction: rtl;
            border-radius: 20px;
            padding: 20px 22px;
            margin: 8px 0 18px;
            background: linear-gradient(135deg, #1d3a63 0%, #0c1728 100%);
            border: 1px solid rgba(148,163,184,0.18);
            box-shadow: 0 12px 30px rgba(2,6,23,0.4);
        }
        .invoice-total-label {
            color: #a5b4cf;
            font-size: 14px;
            font-weight: 700;
            margin-bottom: 8px;
        }
        .invoice-total-value {
            color: #f8fafc;
            font-size: 36px;
            font-weight: 800;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    _render_totals_section(df)
    st.divider()
    _render_work_volume_section(df)


# ─────────────────────────────────────────────────────────────────────────
# View-specific data loading
# ─────────────────────────────────────────────────────────────────────────

def _load_period_view(
    conn: Client,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
) -> tuple[pd.DataFrame, bool]:
    """Returns (dataframe, ok) — ok=False means a validation error was shown."""
    selected_from = _to_iso_date(date_from)
    selected_to = _to_iso_date(date_to)
    if selected_from and selected_to and selected_from > selected_to:
        st.error("تاريخ البداية يجب أن يسبق تاريخ النهاية.")
        return pd.DataFrame(), False

    df = fetch_invoice_report_data(conn, date_from=selected_from, date_to=selected_to)
    return df, True


def _load_latest_view(conn: Client) -> tuple[pd.DataFrame, bool]:
    df = fetch_invoice_report_data(conn)
    return invoice_latest_rows(df), True


# ─────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────

def render_invoices_report(
    conn: Optional[Client],
    view_key: str = VIEW_PERIOD,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
) -> None:
    st.markdown("<h2 style='text-align:right'>المستخلصات</h2>", unsafe_allow_html=True)

    if conn is None:
        st.error("تعذر الاتصال بقاعدة البيانات.")
        return

    if view_key == VIEW_PERIOD:
        df, ok = _load_period_view(conn, date_from=date_from, date_to=date_to)
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
    display_df, link_columns = _prepare_display_dataframe(df)
    column_config = {
        column: st.column_config.LinkColumn(
            label="رابط نسخة مستخلص",
            display_text="فتح المستخلص",
            max_chars=50,
        )
        for column in link_columns
    }
    st.dataframe(
        display_df,
        column_config=column_config or None,
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        label="تنزيل Excel",
        data=_to_excel_bytes(display_df, sheet_name="المستخلصات"),
        file_name="تقرير_المستخلصات.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="invoice_excel_download",
    )