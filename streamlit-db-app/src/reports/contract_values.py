from __future__ import annotations

from io import BytesIO
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


def _excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    try:
        import xlsxwriter  # noqa: F401
        engine = "xlsxwriter"
    except Exception:
        engine = "openpyxl"

    with pd.ExcelWriter(buffer, engine=engine) as writer:
        df.to_excel(writer, index=False, sheet_name="حصر_القيمة")

    return buffer.getvalue()


def _pdf_bytes(df: pd.DataFrame) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    table_data = [list(df.columns)] + df.fillna("").astype(str).values.tolist()
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), title="حصر قيمة العقود")
    style = ParagraphStyle("ArabicTitle", fontName="Helvetica", fontSize=16, alignment=1, textColor=colors.HexColor("#1f2937"))
    story = [Paragraph("حصر قيمة العقود", style), Spacer(1, 14)]
    table = Table(table_data, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e5e7eb")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#111827")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
                ("FONTSIZE", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story.append(table)
    doc.build(story)
    return buffer.getvalue()


def _link_anchor(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    url = str(value).strip()
    if not url or not url.startswith(("http://", "https://")):
        return ""
    return f'<a href="{url}" target="_blank" rel="noopener noreferrer">فتح رابط العقد</a>'


@st.cache_data(show_spinner=False)
def fetch_contract_value_report_data(
    _supabase: Client,
    company_name: Optional[str] = None,
    project_name: Optional[str] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
) -> pd.DataFrame:
    """Fetch and normalize contract-value rows from the database.

    The leading underscore tells Streamlit not to hash the client object itself,
    which is required because Supabase client instances are not hashable.
    """
    if _supabase is None:
        return pd.DataFrame()

    try:
        if company_name or project_name:
            raw_df = fetch_data(_supabase, company_name or "", project_name or "", "contract")
        else:
            query = _supabase.table("contract").select("*, company!inner(companyname, factoryname)")
            if date_from:
                query = query.gte("تاريخ التعاقد", date_from)
            if date_to:
                query = query.lte("تاريخ التعاقد", date_to)
            resp = query.execute()
            raw_df = pd.DataFrame(resp.data or [])
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
    if "رابط نسخة العقد" in df.columns:
        display_df = df.copy()
        display_df["رابط نسخة العقد"] = display_df["رابط نسخة العقد"].map(_link_anchor)
        st.markdown(
            display_df.to_html(index=False, escape=False, border=0, justify="right"),
            unsafe_allow_html=True,
        )
        return

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

    excel_bytes = _excel_bytes(export_df)
    pdf_bytes = _pdf_bytes(export_df)

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="⬇️ تنزيل Excel",
            data=excel_bytes,
            file_name="حصر_قيمه_عقود.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col2:
        st.download_button(
            label="⬇️ تنزيل PDF",
            data=pdf_bytes,
            file_name="حصر_قيمه_عقود.pdf",
            mime="application/pdf",
        )
