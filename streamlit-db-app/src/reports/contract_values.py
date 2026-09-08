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


def _normalize_manual_date(raw_value) -> Optional[str]:
    if raw_value is None:
        return None

    text = str(raw_value).strip()
    if not text:
        return None

    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d", "%d-%m-%y", "%d/%m/%y"):
        try:
            return pd.to_datetime(text, format=fmt).date().isoformat()
        except Exception:
            continue

    try:
        parsed = pd.to_datetime(text, errors="coerce")
        if pd.notna(parsed):
            return parsed.date().isoformat()
    except Exception:
        pass

    return None


def _excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    try:
        import xlsxwriter  # noqa: F401
        engine = "xlsxwriter"
    except Exception:
        engine = "openpyxl"

    with pd.ExcelWriter(buffer, engine=engine) as writer:
        ws = writer.book.add_worksheet("حصر_القيمة")
        header_format = writer.book.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "vcenter",
            "align": "center",
            "bg_color": "#10213a",
            "font_color": "#f8fafc",
            "border": 1,
        })
        cell_format = writer.book.add_format({
            "align": "right",
            "valign": "vcenter",
            "border": 1,
            "bg_color": "#0f172a",
            "font_color": "#e2e8f0",
        })

        for col_idx, col_name in enumerate(df.columns):
            ws.write(0, col_idx, str(col_name), header_format)
            ws.set_column(col_idx, col_idx, 24)

        for row_idx, row in enumerate(df.fillna("").astype(str).itertuples(index=False, name=None), start=1):
            for col_idx, value in enumerate(row):
                ws.write(row_idx, col_idx, str(value), cell_format)

    return buffer.getvalue()


def _pdf_bytes(df: pd.DataFrame) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    table_data = [list(df.columns)] + df.fillna("").astype(str).values.tolist()
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), title="حصر قيمة العقود")
    style = ParagraphStyle(
        "ArabicTitle",
        fontName="Helvetica",
        fontSize=16,
        alignment=1,
        textColor=colors.HexColor("#e5e7eb"),
        backColor=colors.HexColor("#0b1220"),
    )
    story = [Paragraph("حصر قيمة العقود", style), Spacer(1, 14)]
    table = Table(table_data, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#16263f")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#f8fafc")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#2b3d5c")),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#0f172a"), colors.HexColor("#111827")]),
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
    total_html = f"""
    <div style="
        background: linear-gradient(135deg, #122338, #0c1728);
        border: 1px solid rgba(148,163,184,0.25);
        border-radius: 18px;
        padding: 22px 24px;
        margin-bottom: 18px;
        box-shadow: 0 10px 30px rgba(2,6,23,0.35);
        direction: rtl;
    ">
        <div style="font-size: 14px; color: #a5b4cf; margin-bottom: 8px;">إجمالي قيمة العقود</div>
        <div style="font-size: 34px; font-weight: 800; color: #f8fafc;">{total_value:,.2f}</div>
    </div>
    """
    st.markdown(total_html, unsafe_allow_html=True)

    factory_rows = []
    for _, row in summary_df.iterrows():
        factory_name = str(row.get("اسم المصنع", "")).strip() or "غير محدد"
        factory_value = float(row.get("قيمة العقود", 0) or 0)
        factory_rows.append(f"""
            <div style="
                background: linear-gradient(135deg, rgba(15,23,42,0.95), rgba(17,24,39,0.95));
                border: 1px solid rgba(148,163,184,0.2);
                border-radius: 16px;
                padding: 18px 20px;
                min-height: 120px;
                box-shadow: 0 8px 22px rgba(15,23,42,0.25);
                direction: rtl;
            ">
                <div style="font-size: 14px; color: #93c5fd; margin-bottom: 10px;">{factory_name}</div>
                <div style="font-size: 28px; font-weight: 800; color: #f8fafc;">{factory_value:,.2f}</div>
            </div>
        """)

    if factory_rows:
        st.markdown(
            f"<div style='display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:16px; margin-bottom:20px;'>{''.join(factory_rows)}</div>",
            unsafe_allow_html=True,
        )


def _render_table(df: pd.DataFrame, title: str) -> None:
    if df is None or df.empty:
        st.info(f"لا توجد بيانات متاحة لـ {title}.")
        return

    display_df = df.copy()
    if "رابط نسخة العقد" in display_df.columns:
        display_df["رابط نسخة العقد"] = display_df["رابط نسخة العقد"].map(_link_anchor)

    html = display_df.to_html(index=False, escape=False, border=0, justify="right")
    html = f"""
    <div style="direction: rtl; text-align: right; margin-top: 18px;">
      <h3 style="margin: 0 0 10px; color: #e5e7eb; font-weight: 800;">{title}</h3>
      <div style="overflow:auto; border: 1px solid rgba(148,163,184,0.18); border-radius: 14px; background: rgba(15,23,42,0.9); padding: 10px;">
        {html}
      </div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)


def render_contract_values_report(
    conn: Optional[Client],
    company_name: Optional[str] = None,
    project_name: Optional[str] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
) -> None:
    """Render the contract values report within the financial reports mode."""
    st.markdown("<h2 style='text-align:right'>حصر قيمة العقود</h2>", unsafe_allow_html=True)

    if conn is None:
        st.error("تعذر الاتصال بقاعدة البيانات.")
        return

    date_from_value = _normalize_manual_date(date_from) if date_from is not None else None
    date_to_value = _normalize_manual_date(date_to) if date_to is not None else None

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
