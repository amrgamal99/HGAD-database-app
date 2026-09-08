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

# Invisible bidi/formatting characters that often sneak into Arabic strings
# when copy-pasted between files/editors. Stripping them prevents column
# names that *look* identical from silently failing an equality check.
_BIDI_CHARS = ("\u200e", "\u200f", "\u202a", "\u202b", "\u202c", "\ufeff")


def _clean_label(value: object) -> str:
    text = str(value)
    for ch in _BIDI_CHARS:
        text = text.replace(ch, "")
    return text.strip()


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
    """Turn a raw URL into a clickable anchor. Safe to call on a value that
    is already an anchor tag (returns it unchanged) so this can never
    double-wrap or blank out a link that was already built upstream."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if not text or text.lower() == "none":
        return ""

    # Already rendered as an <a> tag somewhere upstream — leave it alone.
    if text.startswith("<a "):
        return text

    if not text.startswith(("http://", "https://")):
        return ""

    return f'<a href="{text}" target="_blank" rel="noopener noreferrer">فتح رابط العقد</a>'


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


def _render_date_filters(
    date_from: Optional[str],
    date_to: Optional[str],
) -> Tuple[Optional[str], Optional[str]]:
    """Renders 'من تاريخ' / 'إلى تاريخ' inputs and returns normalized ISO dates.
    Falls back to whatever was passed in from the caller if the user hasn't
    picked anything in the widgets."""
    st.markdown(
        "<div style='direction:rtl; text-align:right; margin-bottom:6px; "
        "color:#93c5fd; font-weight:700;'>تصفية حسب التاريخ</div>",
        unsafe_allow_html=True,
    )
    col_from, col_to = st.columns(2)

    default_from = pd.to_datetime(date_from).date() if date_from else None
    default_to = pd.to_datetime(date_to).date() if date_to else None

    with col_from:
        picked_from = st.date_input(
            "من تاريخ",
            value=default_from,
            format="DD/MM/YYYY",
            key="contract_value_date_from",
        )
    with col_to:
        picked_to = st.date_input(
            "إلى تاريخ",
            value=default_to,
            format="DD/MM/YYYY",
            key="contract_value_date_to",
        )

    resolved_from = picked_from.isoformat() if picked_from else date_from
    resolved_to = picked_to.isoformat() if picked_to else date_to
    return resolved_from, resolved_to


def _render_summary_cards(summary_df: pd.DataFrame) -> None:
    if summary_df is None or summary_df.empty:
        return

    total_value = float(summary_df["قيمة العقود"].sum()) if "قيمة العقود" in summary_df.columns else 0.0

    st.markdown(
        """
        <style>
        .cv-card {
            position: relative;
            border-radius: 20px;
            padding: 20px 22px;
            direction: rtl;
            overflow: hidden;
            border: 1px solid rgba(148,163,184,0.18);
            transition: transform 0.18s ease, box-shadow 0.18s ease;
        }
        .cv-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 16px 34px rgba(2,6,23,0.45);
        }
        .cv-card::before {
            content: "";
            position: absolute;
            top: -30%;
            left: -20%;
            width: 65%;
            height: 160%;
            background: rgba(255,255,255,0.06);
            transform: rotate(18deg);
            pointer-events: none;
        }
        .cv-total {
            background: linear-gradient(135deg, #1d3a63 0%, #0c1728 100%);
            box-shadow: 0 12px 30px rgba(2,6,23,0.4);
            margin-bottom: 20px;
        }
        .cv-icon { font-size: 22px; margin-inline-start: 8px; }
        .cv-label { font-size: 14px; color: #a5b4cf; margin-bottom: 8px; font-weight: 600; }
        .cv-value-total { font-size: 36px; font-weight: 800; color: #f8fafc; letter-spacing: 0.3px; }
        .cv-value-factory { font-size: 27px; font-weight: 800; color: #f8fafc; }
        .cv-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 16px;
            margin-bottom: 22px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    total_html = f"""
    <div class="cv-card cv-total">
        <div class="cv-label"><span class="cv-icon">💰</span>إجمالي قيمة العقود</div>
        <div class="cv-value-total">{total_value:,.2f}</div>
    </div>
    """
    st.markdown(total_html, unsafe_allow_html=True)

    palette = [
        ("#1e3a5f", "#0f2038", "#60a5fa"),
        ("#3b2a5e", "#170f2e", "#c084fc"),
        ("#1f4b3f", "#0c211b", "#34d399"),
        ("#5c3620", "#28160b", "#fb923c"),
        ("#4a1f3d", "#1f0c19", "#f472b6"),
        ("#1c3d4d", "#0a1a21", "#22d3ee"),
    ]

    factory_rows = []
    for idx, (_, row) in enumerate(summary_df.iterrows()):
        factory_name = str(row.get("اسم المصنع", "")).strip() or "غير محدد"
        factory_value = float(row.get("قيمة العقود", 0) or 0)
        color_start, color_end, accent = palette[idx % len(palette)]
        factory_rows.append(f"""
            <div class="cv-card" style="background: linear-gradient(135deg, {color_start}, {color_end});">
                <div class="cv-label" style="color:{accent};"><span class="cv-icon">🏭</span>{factory_name}</div>
                <div class="cv-value-factory">{factory_value:,.2f}</div>
            </div>
        """)

    if factory_rows:
        st.markdown(f"<div class='cv-grid'>{''.join(factory_rows)}</div>", unsafe_allow_html=True)


def _render_table(df: pd.DataFrame, title: str) -> None:
    if df is None or df.empty:
        st.info(f"لا توجد بيانات متاحة لـ {title}.")
        return

    display_df = df.copy()

    # Robust column match: find the link column even if its name picked up
    # invisible bidi/formatting characters that make a plain `==` fail.
    link_col = None
    for col in display_df.columns:
        if _clean_label(col) == "رابط نسخة العقد":
            link_col = col
            break

    if link_col is not None:
        display_df[link_col] = display_df[link_col].map(_link_anchor)

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

    # "من تاريخ" / "إلى تاريخ" filter widgets. Whatever the user picks here
    # takes priority over values passed in from the caller.
    date_from_value, date_to_value = _render_date_filters(date_from_value, date_to_value)

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
    export_link_col = None
    for col in export_df.columns:
        if _clean_label(col) == "رابط نسخة العقد":
            export_link_col = col
            break
    if export_link_col is not None:
        export_df[export_link_col] = export_df[export_link_col].apply(
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