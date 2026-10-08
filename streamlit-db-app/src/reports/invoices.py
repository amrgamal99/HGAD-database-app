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

DATE_COLUMN = "تاريخ إصدار المستخلص"
VIEW_PERIOD = "period"
VIEW_TOTAL = "total"

VIEW_TITLES = {
    VIEW_PERIOD: "مستخلصات خلال فترة زمنية",
    VIEW_TOTAL: "حجم الأعمال ككل",
}

# Label of the standalone single-select invoice view filter.
VIEW_FILTER_LABEL = "طريقة العرض"

CARDS_PER_ROW = 3

_INVOICE_LINK_ALIASES = (
    "رابط نسخة مستخلص",
    "رابط نسخة المستخلص",
    "رابط المستخلص",
    "رابط invoice",
)


# ─────────────────────────────────────────────────────────────────────────
# Small UI building blocks
# ─────────────────────────────────────────────────────────────────────────

def render_view_filter(default: str = VIEW_PERIOD) -> str:
    """Dropdown filter (same style as the other filters), not a radio/multi-select."""
    keys = list(VIEW_TITLES.keys())
    return st.selectbox(
        VIEW_FILTER_LABEL,
        options=keys,
        index=keys.index(default) if default in keys else 0,
        format_func=lambda key: VIEW_TITLES[key],
        key="invoice_view_filter",
    )


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


def _first_existing(df: pd.DataFrame, candidates: tuple[str, ...]) -> Optional[str]:
    return next((column for column in candidates if column in df.columns), None)


def _prepare_display_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Return the complete invoice data unchanged, including every column and order."""
    display_df = df.copy()
    return display_df, []


def _to_excel_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        writer.sheets[sheet_name].right_to_left()
    return buffer.getvalue()


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


def _load_total_view(conn: Client) -> tuple[pd.DataFrame, bool]:
    return fetch_invoice_report_data(conn), True


# ─────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────

def render_invoices_report(
    conn: Optional[Client],
    view_key: Optional[str] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
) -> None:
    """
    If view_key is None the dropdown filter is rendered here;
    if the caller already chose a view, pass it in and no filter is shown.
    """
    st.markdown("<h2 style='text-align:right'>المستخلصات</h2>", unsafe_allow_html=True)

    if conn is None:
        st.error("تعذر الاتصال بقاعدة البيانات.")
        return

    if view_key is None:
        view_key = render_view_filter()

    if view_key == VIEW_PERIOD:
        df, ok = _load_period_view(conn, date_from=date_from, date_to=date_to)
    else:
        df, ok = _load_total_view(conn)

    if not ok:
        return

    st.markdown(f"### {VIEW_TITLES.get(view_key, '')}")

    if df.empty:
        st.info("لا توجد مستخلصات مطابقة.")
        return

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