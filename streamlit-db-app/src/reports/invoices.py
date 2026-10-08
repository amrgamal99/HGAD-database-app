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
    """Build public invoice details from the requested canonical columns."""
    work = df.copy()
    if DATE_COLUMN in work.columns:
        work[DATE_COLUMN] = pd.to_datetime(work[DATE_COLUMN], errors="coerce")

    factory_source = next(
        (column for column in ("factoryname", "مصنع", "اسم المصنع") if column in work.columns),
        None,
    )
    company_source = next(
        (column for column in ("companyname", "اسم الشركة") if column in work.columns),
        None,
    )
    contract_source = next(
        (column for column in ("اسم المشروع", "اسم العقد", "contractname") if column in work.columns),
        None,
    )
    link_source = next(
        (column for column in _INVOICE_LINK_ALIASES if column in work.columns),
        None,
    )

    display_df = pd.DataFrame({
        "اسم المصنع": work[factory_source] if factory_source else None,
        "اسم الشركة": work[company_source] if company_source else None,
        "اسم العقد": work[contract_source] if contract_source else None,
        DATE_COLUMN: work[DATE_COLUMN] if DATE_COLUMN in work.columns else None,
    })
    if link_source:
        display_df["رابط نسخة المستخلص"] = work[link_source]

    if not display_df.empty:
        display_df = display_df.sort_values(
            by=DATE_COLUMN,
            ascending=True,
            na_position="last",
        )
    display_df = _format_display_dates(display_df)
    return display_df, ["رابط نسخة المستخلص"] if "رابط نسخة المستخلص" in display_df.columns else []


def _to_excel_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        writer.sheets[sheet_name].right_to_left()
    return buffer.getvalue()


# ─────────────────────────────────────────────────────────────────────────
# Summary sections
# ─────────────────────────────────────────────────────────────────────────

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
    )
    if factory_volume.empty:
        st.info("لا تتوفر بيانات المصنع هذه المستخلصات.")
        return

    card_columns = st.columns(
        max(1, min(len(factory_volume), 3)),
        gap="small",
    )
    for card_index, row in factory_volume.iterrows():
        factory_name = str(row[FACTORY_COLUMN] or "مصنع غير معروف")
        volume = float(row["حجم الأعمال"] or 0)
        with card_columns[card_index % len(card_columns)]:
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
        """,
        unsafe_allow_html=True,
    )
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