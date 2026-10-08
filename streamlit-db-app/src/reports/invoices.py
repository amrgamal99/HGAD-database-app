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

# Label of the view filter (change this text to rename the filter).
VIEW_FILTER_LABEL = "طريقة عرض المستخلصات"

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
    """
    Build the invoice details table.

    Only the whitelisted columns below are shown, in this order:
    اسم المصنع ← اسم الشركة ← اسم العقد ← تاريخ إصدار المستخلص ← الرابط.
    Any other column (e.g. companyid_contract or other ids) is never included.
    Rows are sorted ascending by تاريخ إصدار المستخلص.
    """
    work = df.copy()
    if DATE_COLUMN in work.columns:
        work[DATE_COLUMN] = pd.to_datetime(work[DATE_COLUMN], errors="coerce")

    factory_source = _first_existing(work, ("factoryname", "مصنع", "اسم المصنع"))
    company_source = _first_existing(work, ("companyname", "اسم الشركة"))
    contract_source = _first_existing(work, ("اسم المشروع", "اسم العقد", "contractname"))
    link_source = _first_existing(work, _INVOICE_LINK_ALIASES)

    display_df = pd.DataFrame({
        "اسم المصنع": work[factory_source] if factory_source else None,
        "اسم الشركة": work[company_source] if company_source else None,
        "اسم العقد": work[contract_source] if contract_source else None,
        DATE_COLUMN: work[DATE_COLUMN] if DATE_COLUMN in work.columns else pd.NaT,
    })
    if link_source:
        display_df["رابط نسخة المستخلص"] = work[link_source]

    # Oldest → newest, rows with no date go last.
    display_df = display_df.sort_values(
        by=DATE_COLUMN,
        ascending=True,
        na_position="last",
        kind="mergesort",
    ).reset_index(drop=True)

    display_df = _format_display_dates(display_df)
    link_columns = ["رابط نسخة المستخلص"] if link_source else []
    return display_df, link_columns


def _to_excel_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        writer.sheets[sheet_name].right_to_left()
    return buffer.getvalue()


# ─────────────────────────────────────────────────────────────────────────
# Summary section: حجم الأعمال لكل مصنع (flash cards)
# ─────────────────────────────────────────────────────────────────────────

_CARD_CSS = """
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
.invoice-factory-card {
    direction: rtl;
    border-radius: 20px;
    padding: 20px 22px;
    margin: 8px 0 18px;
    background: linear-gradient(135deg, #1d3a63 0%, #0c1728 100%);
    border: 1px solid rgba(148,163,184,0.18);
    box-shadow: 0 12px 30px rgba(2,6,23,0.4);
    text-align: center;
}
.invoice-factory-label {
    color: #cbd5e1;
    font-size: 15px;
    font-weight: 700;
    margin-bottom: 8px;
}
.invoice-factory-value {
    color: #f8fafc;
    font-size: 28px;
    font-weight: 800;
}
.invoice-factory-caption {
    color: #60a5fa;
    font-size: 12px;
    font-weight: 700;
    margin-top: 6px;
}
</style>
"""


def _render_work_volume_section(df: pd.DataFrame) -> None:
    st.markdown(
        "<div class='cv-section-title'>حجم الأعمال لكل مصنع</div>",
        unsafe_allow_html=True,
    )
    factory_volume = invoice_numeric_summary(df, group_by_factory=True)
    if factory_volume.empty or FACTORY_COLUMN not in factory_volume.columns:
        st.info("لا تتوفر بيانات المصنع لهذه المستخلصات.")
        return

    factory_volume = factory_volume[[FACTORY_COLUMN, "حجم الأعمال"]].dropna(
        subset=[FACTORY_COLUMN]
    )
    factory_volume["حجم الأعمال"] = pd.to_numeric(
        factory_volume["حجم الأعمال"], errors="coerce"
    ).fillna(0)
    factory_volume = factory_volume.sort_values(
        by="حجم الأعمال", ascending=False, na_position="last"
    ).reset_index(drop=True)
    if factory_volume.empty:
        st.info("لا تتوفر بيانات المصنع لهذه المستخلصات.")
        return

    # One flash card per factory, laid out in rows of CARDS_PER_ROW.
    records = factory_volume.to_dict("records")
    for start in range(0, len(records), CARDS_PER_ROW):
        row_records = records[start:start + CARDS_PER_ROW]
        columns = st.columns(CARDS_PER_ROW, gap="small")
        for column, record in zip(columns, row_records):
            factory_name = str(record[FACTORY_COLUMN] or "مصنع غير معروف")
            volume = float(record["حجم الأعمال"] or 0)
            with column:
                st.markdown(
                    f"""
                    <div class="invoice-factory-card">
                        <div class="invoice-factory-label">{factory_name}</div>
                        <div class="invoice-factory-value">{volume:,.2f}</div>
                        <div class="invoice-factory-caption">حجم الأعمال</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )


def _render_summary(df: pd.DataFrame) -> None:
    """Only the per-factory work-volume cards (no column totals)."""
    st.markdown(_CARD_CSS, unsafe_allow_html=True)
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