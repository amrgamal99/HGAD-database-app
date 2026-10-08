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
from utils.data_helpers import invoice_latest_rows, invoice_numeric_summary

DATE_COLUMN = "تاريخ إصدار المستخلص"
VIEW_PERIOD = "period"
VIEW_LATEST = "latest"

VIEW_TITLES = {
    VIEW_PERIOD: "مستخلصات خلال فترة زمنية",
    VIEW_LATEST: "آخر مستخلص",
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


def _should_render_work_volume_card(view_key: str) -> bool:
    """Show the work-volume metric only for the period view."""
    return view_key == VIEW_PERIOD


def _render_work_volume_cards(df: pd.DataFrame) -> None:
    """Render work-volume values as attractive cards for each factory."""
    summary = invoice_numeric_summary(df, group_by_factory=True)
    if summary.empty or "مصنع" not in summary.columns:
        st.info("لا توجد بيانات لحجم الأعمال لكل مصنع.")
        return

    summary = summary[["مصنع", "حجم الأعمال"]].dropna(subset=["مصنع"])
    summary = summary.reset_index(drop=True)
    if summary.empty:
        st.info("لا توجد بيانات لحجم الأعمال لكل مصنع.")
        return

    st.markdown(
        """
        <style>
        .work-card {
            background: linear-gradient(135deg, #1f4e79, #173b5e);
            border: 1px solid rgba(209, 229, 244, 0.22);
            border-radius: 14px;
            padding: 15px 18px;
            box-shadow: 0 8px 25px rgba(0, 0, 0, 0.25);
            text-align: right;
        }
        .work-card-factory {
            color: #dceeff;
            font-size: 15px;
            font-weight: 700;
            margin-bottom: 6px;
        }
        .work-card-value {
            color: #ffcf7a;
            font-size: 22px;
            font-weight: 800;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    columns = st.columns(max(1, min(3, len(summary))))
    for column, row in zip(columns, summary.to_dict("records")):
        with column:
            st.markdown(
                f"""
                <div class="work-card">
                    <div class="work-card-factory">{str(row['مصنع'])}</div>
                    <div class="work-card-value">{float(row['حجم الأعمال'] or 0):,.2f}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def _prepare_display_dataframe(
    df: pd.DataFrame,
    remove_factory_column: bool = False,
) -> tuple[pd.DataFrame, list[str]]:
    """Keep usable data columns, normalize dates, and make the invoice link clickable."""
    work = df.copy()
    excluded_columns = {
        "invoiceid",
        "companyid",
        "contractid",
        "companyid_coontract",
        "companyid_contract",
    }
    work = work.loc[:, [column for column in work.columns if column not in excluded_columns]]
    work = work.loc[:, ~pd.Index(work.columns).duplicated()]
    work = work.loc[:, work.notna().any(axis=0)]

    factory_source = _first_existing(work, ("factoryname", "مصنع"))
    company_source = _first_existing(work, ("companyname",))
    contract_source = _first_existing(work, ("اسم المشروع", "contractname"))
    link_source = _first_existing(work, _INVOICE_LINK_ALIASES)

    display_df = work.copy()
    if factory_source and factory_source != "اسم المصنع":
        if "اسم المصنع" in display_df.columns:
            display_df = display_df.drop(columns=factory_source)
        else:
            display_df = display_df.rename(columns={factory_source: "اسم المصنع"})

    if company_source and company_source != "اسم الشركة":
        if "اسم الشركة" in display_df.columns:
            display_df = display_df.drop(columns=company_source)
        else:
            display_df = display_df.rename(columns={company_source: "اسم الشركة"})

    if contract_source and contract_source != "اسم العقد":
        if "اسم العقد" in display_df.columns:
            display_df = display_df.drop(columns=contract_source)
        else:
            display_df = display_df.rename(columns={contract_source: "اسم العقد"})

    if link_source:
        display_df = display_df.rename(columns={link_source: "رابط نسخة المستخلص"})

    if remove_factory_column:
        display_df = display_df.drop(
            columns=[
                column
                for column in ("مصنع", "اسم المصنع")
                if column in display_df.columns
            ],
            errors="ignore",
        )

    renamed_columns = [
        column for column in ("اسم المصنع", "اسم الشركة", "اسم العقد")
        if column in display_df.columns
    ]
    remaining_columns = [
        column for column in display_df.columns
        if column not in renamed_columns
    ]
    display_df = display_df.loc[:, renamed_columns + remaining_columns]

    if DATE_COLUMN in display_df.columns:
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

    if _should_render_work_volume_card(view_key):
        st.markdown("### حجم الأعمال لكل مصنع")
        _render_work_volume_cards(df)

    st.divider()
    st.markdown("### تفاصيل المستخلصات")
    display_df, link_columns = _prepare_display_dataframe(
        df,
        remove_factory_column=view_key == VIEW_LATEST,
    )
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