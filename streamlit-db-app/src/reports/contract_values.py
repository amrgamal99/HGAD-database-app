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


def _flatten_html(html: str) -> str:
    """Collapse a multi-line, indented HTML/CSS template into one line.

    Streamlit's markdown renderer treats any line indented 4+ spaces as a
    Markdown "indented code block" *before* it ever gets to parsing HTML.
    A plain triple-quoted f-string written inside nested function bodies
    ends up with exactly that kind of indentation, so a card can randomly
    get rendered as raw text inside a code block instead of as HTML. Every
    HTML/CSS snippet built as an f-string in this file should be passed
    through this helper before being handed to st.markdown().
    """
    return "".join(line.strip() for line in html.strip().splitlines())


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


def _format_period_label(date_from: Optional[str], date_to: Optional[str]) -> str:
    if date_from and date_to:
        return f"\nمن {date_from} إلى {date_to}"
    if date_from:
        return f"\nمن {date_from}"
    if date_to:
        return f"\nحتى {date_to}"
    return ""


def _write_summary_sheet(
    writer: "pd.ExcelWriter",
    summary_df: pd.DataFrame,
    date_from: Optional[str],
    date_to: Optional[str],
    formats: dict,
) -> None:
    from xlsxwriter.utility import xl_col_to_name

    if summary_df is None or summary_df.empty:
        return

    df = summary_df.copy()
    sheet_name = "مقارنة المصانع"
    df.to_excel(writer, sheet_name=sheet_name, startrow=5, index=False)
    ws = writer.sheets[sheet_name]
    ws.right_to_left()
    ws.hide_gridlines(2)
    ws.set_zoom(110)

    rows, cols = df.shape
    last_col = xl_col_to_name(max(cols - 1, 0))
    ws.merge_range(
        f"A1:{last_col}2",
        f"مقارنة قيمة العقود بين المصانع{_format_period_label(date_from, date_to)}",
        formats["title"],
    )

    if rows > 0:
        ws.add_table(5, 0, rows + 5, cols - 1, {
            "style": "Table Style Medium 2",
            "columns": [{"header": str(col)} for col in df.columns],
        })

    value_col_idx = None
    for idx, col in enumerate(df.columns):
        if _clean_label(col) == "قيمة العقود":
            value_col_idx = idx
            break

    ws.set_column(0, 0, 28)
    if value_col_idx is not None:
        ws.set_column(value_col_idx, value_col_idx, 24, formats["money"])
        total_value = pd.to_numeric(df.iloc[:, value_col_idx], errors="coerce").fillna(0).sum()
        total_row = rows + 8
        label_col = max(value_col_idx - 1, 0)
        ws.write(total_row, label_col, "إجمالي قيمة العقود", formats["total_label"])
        ws.write(total_row, value_col_idx, total_value, formats["total_money"])

    ws.freeze_panes(6, 0)


def _write_details_sheet(
    writer: "pd.ExcelWriter",
    details_df: pd.DataFrame,
    date_from: Optional[str],
    date_to: Optional[str],
    formats: dict,
) -> None:
    from xlsxwriter.utility import xl_col_to_name

    if details_df is None or details_df.empty:
        return

    df = details_df.copy()

    link_col_idx = None
    for idx, col in enumerate(df.columns):
        if _clean_label(col) == "رابط نسخة العقد":
            link_col_idx = idx
            # Export the plain URL, not the clickable <a> markup used on the page.
            df[col] = df[col].apply(
                lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
            )
            break

    sheet_name = "تفاصيل العقود"
    df.to_excel(writer, sheet_name=sheet_name, startrow=5, index=False)
    ws = writer.sheets[sheet_name]
    ws.right_to_left()
    ws.hide_gridlines(2)
    ws.set_zoom(110)

    rows, cols = df.shape
    last_col = xl_col_to_name(max(cols - 1, 0))
    ws.merge_range(
        f"A1:{last_col}2",
        f"تفاصيل العقود{_format_period_label(date_from, date_to)}",
        formats["title"],
    )

    if rows > 0:
        ws.add_table(5, 0, rows + 5, cols - 1, {
            "style": "Table Style Medium 2",
            "columns": [{"header": str(col)} for col in df.columns],
        })

    for idx, col in enumerate(df.columns):
        label = _clean_label(col)
        if idx == link_col_idx:
            for row_num, url in enumerate(df[col], start=6):
                if url:
                    ws.write_url(row_num, idx, url, formats["hyperlink"], "فتح العقد")
            ws.set_column(idx, idx, 28)
        elif "قيمة" in label or "قيمه" in label:
            ws.set_column(idx, idx, 22, formats["money"])
        elif label == "اسم المشروع":
            ws.set_column(idx, idx, 42)
        elif label in ("اسم الشركة", "اسم المصنع"):
            ws.set_column(idx, idx, 26)
        else:
            ws.set_column(idx, idx, 20)

    ws.freeze_panes(6, 0)


def _excel_bytes(
    summary_df: pd.DataFrame,
    details_df: pd.DataFrame,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
) -> bytes:
    """Build the two-sheet workbook: مقارنة المصانع then تفاصيل العقود,
    each with a merged RTL title banner, an Excel Table style, and (for
    the details sheet) clickable hyperlinks in the link column."""
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book
        formats = {
            "title": workbook.add_format({
                "bold": True,
                "font_size": 16,
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#1F4E78",
                "font_color": "white",
                "border": 1,
                "text_wrap": True,
            }),
            "money": workbook.add_format({"num_format": "#,##0.00"}),
            "total_label": workbook.add_format({
                "bold": True, "bg_color": "#D9EAD3", "border": 1, "align": "right",
            }),
            "total_money": workbook.add_format({
                "bold": True, "bg_color": "#D9EAD3", "border": 1, "num_format": "#,##0.00",
            }),
            "hyperlink": workbook.add_format({"font_color": "blue", "underline": 1}),
        }

        _write_summary_sheet(writer, summary_df, date_from, date_to, formats)
        _write_details_sheet(writer, details_df, date_from, date_to, formats)

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


# =============================================================================
#  MANUAL DATA LIST (قائمة البيانات اليدوية كاملة)
#
#  Contracts that exist in real life but aren't in Supabase yet. Each entry
#  is only added to the report if its contractid does NOT already appear
#  among the rows fetched from the database — so a manual row is safe to
#  reuse a contractid that hasn't been entered into Supabase, but it never
#  duplicates one that has.
# =============================================================================
MANUAL_CONTRACT_ENTRIES = [
    {
        "contractid": 52, "companyid": 12, "اسم المشروع": "variaton 01 فيلات مدينة نور",
        "تاريخ التعاقد": "2026-04-30", "قيمة التعاقد": 7703458,
        "رابط نسخة العقد": "https://drive.google.com/file/d/1ph_pGd4-D7IaM-CMGE4y1g3VJivX6i2o/view?usp=drive_link",
        "قيمه التعاقد شامله الضريبه": "نعم", "الملاحظات": None, "company": None, "factoryname": "بدر", "companyname": "اتريم"
    },
    {
        "contractid": 21, "companyid": 12, "اسم المشروع": "VO1 عمارات حدائق نور",
        "تاريخ التعاقد": "2025-09-29", "قيمة التعاقد": 54017429,
        "رابط نسخة العقد": "https://drive.google.com/file/d/1CXnPwVop7UT_oQzLpTjd3kLNFJOGH6kP/view?usp=drive_link",
        "قيمه التعاقد شامله الضريبه": "نعم", "الملاحظات": None, "company": None, "factoryname": "بدر", "companyname": "اتريم"
    },
    {
        "contractid": 21, "companyid": 12, "اسم المشروع": "VO2 عمارات حدائق نور",
        "تاريخ التعاقد": "2025-12-01", "قيمة التعاقد": 19937600,
        "رابط نسخة العقد": "https://drive.google.com/file/d/1r8zEI59IFFWC5NROFKjpFdDcatyYnF40/view?usp=drive_link",
        "قيمه التعاقد شامله الضريبه": "نعم", "الملاحظات": None, "company": None, "factoryname": "بدر", "companyname": "اتريم"
    },
    {
        "contractid": 43, "companyid": 12, "اسم المشروع": "VO - Doors&windows PH3CLU3-لوفر و المظلات مرحله ثالثه",
        "تاريخ التعاقد": "2026-03-31", "قيمة التعاقد": 19531858,
        "رابط نسخة العقد": "https://drive.google.com/file/d/1uE8QNlGhz78K3zXOVybliXkUi-dKi2zf/view?usp=drive_link",
        "قيمه التعاقد شامله الضريبه": "نعم", "الملاحظات": None, "company": None, "factoryname": "بدر", "companyname": "اتريم"
    },
    {
        "contractid": 43, "companyid": 12, "اسم المشروع": "VO - Doors&windows PH3CLU4-لوفر و المظلات مرحله ثالثه",
        "تاريخ التعاقد": "2026-03-30", "قيمة التعاقد": 18474287,
        "رابط نسخة العقد": "https://drive.google.com/file/d/10BIVVWA89o_gt1vPCkNmcZH8nJEILKPq/view?usp=drive_link",
        "قيمه التعاقد شامله الضريبه": "نعم", "الملاحظات": None, "company": None, "factoryname": "بدر", "companyname": "اتريم"
    },
    {
        "contractid": 55, "companyid": 12, "اسم المشروع": "vo2 بريفادو المرحله الثانيه بمدينتي",
        "تاريخ التعاقد": "2026-06-09", "قيمة التعاقد": 4420742,
        "رابط نسخة العقد": "https://drive.google.com/file/d/16-oXXP7-i4i9pVQzh1N_vn8m90bQDQrD/view?usp=drive_link",
        "قيمه التعاقد شامله الضريبه": "نعم", "الملاحظات": None, "company": None, "factoryname": "بدر", "companyname": "اتريم"
    },
    {
        "contractid": 70, "companyid": 23, "اسم المشروع": "الوميتال داود بلوم فيلد",
        "تاريخ التعاقد": "2025-08-25", "قيمة التعاقد": 7406888,
        "رابط نسخة العقد": "https://drive.google.com/file/d/1jM3pEjuSZiR-U7UKdu14z1ih9TM9syRd/view?usp=drive_link",
        "قيمه التعاقد شامله الضريبه": "لا", "الملاحظات": None, "company": None, "factoryname": "بدر", "companyname": "داود النصر"
    },
    {
        "contractid": 70, "companyid": 23, "اسم المشروع": "كلادينج داود بلوم فيلد",
        "تاريخ التعاقد": "2025-11-18", "قيمة التعاقد": 2609374,
        "رابط نسخة العقد": "https://drive.google.com/file/d/1CPQi1L4dcKcI2Xc3K-0TD4aUWlswmgxr/view?usp=drive_link",
        "قيمه التعاقد شامله الضريبه": "لا", "الملاحظات": None, "company": None, "factoryname": "بدر", "companyname": "داود النصر"
    },
    {
        "contractid": 70, "companyid": 23, "اسم المشروع": "ملحق رقم 1 لعقد الوميتال",
        "تاريخ التعاقد": "2026-08-13", "قيمة التعاقد": 4456656,
        "رابط نسخة العقد": "https://drive.google.com/file/d/19UTnx6jar9FLV7Rcjr0f3G94Xb-CGxED/view?usp=drive_link",
        "قيمه التعاقد شامله الضريبه": "لا", "الملاحظات": None, "company": None, "factoryname": "بدر", "companyname": "داود النصر"
    },
    {
        "contractid": 70, "companyid": 23, "اسم المشروع": "ملحق رقم 1 لعقد كلادينج",
        "تاريخ التعاقد": "2026-08-13", "قيمة التعاقد": 2142860,
        "رابط نسخة العقد": "https://drive.google.com/file/d/1Dw3vomdhAV8YuvtybGCfn5wr6TJRvR5s/view?usp=drive_link",
        "قيمه التعاقد شامله الضريبه": "لا", "الملاحظات": None, "company": None, "factoryname": "بدر", "companyname": "داود النصر"
    },
    {
        "contractid": 57, "companyid": 37, "اسم المشروع": "نور هاندريل",
        "تاريخ التعاقد": "2026-04-29", "قيمة التعاقد": 15962242.95,
        "رابط نسخة العقد": "https://drive.google.com/file/d/1Kg9__Iiqf9itDa0frGHdAJlk_BCKx1qm/view?usp=drive_link",
        "قيمه التعاقد شامله الضريبه": "نعم", "الملاحظات": None, "company": None, "factoryname": "التجمع", "companyname": "الاتحاد المصري  الحاذق"
    },
    {
        "contractid": 57, "companyid": 37, "اسم المشروع": "نور لوفر",
        "تاريخ التعاقد": "2026-05-30", "قيمة التعاقد": 4292159,
        "رابط نسخة العقد": None,
        "قيمه التعاقد شامله الضريبه": "نعم", "الملاحظات": None, "company": None, "factoryname": "التجمع", "companyname": "الاتحاد المصري  الحاذق"
    },
    {
        "contractid": 74, "companyid": 35, "اسم المشروع": "موقع القمزي الاول",
        "تاريخ التعاقد": "2026-05-30", "قيمة التعاقد": 987450,
        "رابط نسخة العقد": "https://drive.google.com/file/d/1Fr7mi1pBqpaNI7ugNlojUz0F3FFVgyA9/view?usp=drive_link",
        "قيمه التعاقد شامله الضريبه": "لا", "الملاحظات": None, "company": None, "factoryname": "التجمع", "companyname": "GRID"
    },
    {
        "contractid": 74, "companyid": 35, "اسم المشروع": "موقع القمزي الثاني",
        "تاريخ التعاقد": "2026-01-28", "قيمة التعاقد": 3711736,
        "رابط نسخة العقد": None,
        "قيمه التعاقد شامله الضريبه": "لا", "الملاحظات": None, "company": None, "factoryname": "التجمع", "companyname": "GRID"
    },
    {
        "contractid": 44, "companyid": 12, "اسم المشروع": "ميفيدا جاردن",
        "تاريخ التعاقد": "2026-01-22", "قيمة التعاقد": 34415689.50,
        "رابط نسخة العقد": "https://drive.google.com/file/d/1HnfVeJaILScGmvPFT9HIvUcJELnSF9So/view?usp=drive_link",
        "قيمه التعاقد شامله الضريبه": "نعم", "الملاحظات": None, "company": None, "factoryname": "بدر", "companyname": "اتريم"
    },
    {
        "contractid": 44, "companyid": 12, "اسم المشروع": "VO1 ميفيدا جاردن",
        "تاريخ التعاقد": "2026-08-11", "قيمة التعاقد": 37114316.60,
        "رابط نسخة العقد": "https://drive.google.com/file/d/1Mwjro2QLUCNRImkQIC-vgooO-w5qRxmk/view?usp=drive_link",
        "قيمه التعاقد شامله الضريبه": "نعم", "الملاحظات": None, "company": None, "factoryname": "بدر", "companyname": "اتريم"
    },
]


def _merge_manual_entries(
    raw_df: pd.DataFrame,
    date_from: Optional[str],
    date_to: Optional[str],
) -> pd.DataFrame:
    """Append MANUAL_CONTRACT_ENTRIES rows that fall inside [date_from, date_to]
    and whose contractid isn't already present among the rows pulled from
    Supabase. A manual entry is skipped only if its contractid was already
    fetched from the database — it's fine for two manual entries, or a
    manual entry and an unrelated db row, to share a contractid value."""
    df = raw_df.copy() if raw_df is not None else pd.DataFrame()

    db_contract_ids = set()
    if not df.empty and "contractid" in df.columns:
        db_contract_ids = set(df["contractid"].dropna().unique())

    for col in ["factoryname", "companyname", "contractid", "companyid", "قيمة التعاقد"]:
        if col not in df.columns:
            df[col] = None

    if "تاريخ التعاقد" in df.columns and not df.empty:
        df["تاريخ التعاقد"] = pd.to_datetime(df["تاريخ التعاقد"], errors="coerce")
    else:
        df["تاريخ التعاقد"] = pd.to_datetime(pd.Series([], dtype="object"))

    date_from_dt = pd.to_datetime(date_from) if date_from else None
    date_to_dt = pd.to_datetime(date_to) if date_to else None

    rows_to_add = []
    for item in MANUAL_CONTRACT_ENTRIES:
        entry = dict(item)
        item_date = pd.to_datetime(entry["تاريخ التعاقد"])

        if date_from_dt is not None and item_date < date_from_dt:
            continue
        if date_to_dt is not None and item_date > date_to_dt:
            continue

        if entry["contractid"] in db_contract_ids:
            continue

        entry["تاريخ التعاقد"] = item_date
        rows_to_add.append(entry)

    if rows_to_add:
        df = pd.concat([df, pd.DataFrame(rows_to_add)], ignore_index=True)

    if date_from_dt is not None:
        df = df[df["تاريخ التعاقد"] >= date_from_dt]
    if date_to_dt is not None:
        df = df[df["تاريخ التعاقد"] <= date_to_dt]

    df["قيمة التعاقد"] = pd.to_numeric(df.get("قيمة التعاقد"), errors="coerce").fillna(0)
    df["تاريخ التعاقد"] = df["تاريخ التعاقد"].dt.strftime("%Y-%m-%d")

    return df


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

            # Flatten the nested `company` join into plain columns so the
            # manual entries (which use flat factoryname/companyname keys)
            # line up with the rows fetched from Supabase.
            if not raw_df.empty and "company" in raw_df.columns:
                raw_df["factoryname"] = raw_df["company"].apply(
                    lambda x: x.get("factoryname") if isinstance(x, dict) else None
                )
                raw_df["companyname"] = raw_df["company"].apply(
                    lambda x: x.get("companyname") if isinstance(x, dict) else None
                )

            raw_df = _merge_manual_entries(raw_df, date_from, date_to)
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

    style_block = _flatten_html("""
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
    """)
    st.markdown(style_block, unsafe_allow_html=True)

    total_html = _flatten_html(f"""
        <div class="cv-card cv-total">
            <div class="cv-label"><span class="cv-icon">💰</span>إجمالي قيمة العقود</div>
            <div class="cv-value-total">{total_value:,.2f}</div>
        </div>
    """)
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
        card_html = _flatten_html(f"""
            <div class="cv-card" style="background: linear-gradient(135deg, {color_start}, {color_end});">
                <div class="cv-label" style="color:{accent};"><span class="cv-icon">🏭</span>{factory_name}</div>
                <div class="cv-value-factory">{factory_value:,.2f}</div>
            </div>
        """)
        factory_rows.append(card_html)

    if factory_rows:
        grid_html = _flatten_html(
            "<div class='cv-grid'>" + "".join(factory_rows) + "</div>"
        )
        st.markdown(grid_html, unsafe_allow_html=True)


def _render_table(df: pd.DataFrame, title: str) -> None:
    if df is None or df.empty:
        st.info(f"لا توجد بيانات متاحة لـ {title}.")
        return

    display_df = df.copy()
    link_col = None
    for col in display_df.columns:
        if _clean_label(col) == "رابط نسخة العقد":
            link_col = col
            break

    column_config = {}
    if link_col is not None:
        display_df[link_col] = display_df[link_col].apply(
            lambda value: str(value).strip() if pd.notna(value) and str(value).strip() else None
        )
        column_config[link_col] = st.column_config.LinkColumn(
            "رابط نسخة العقد",
            max_chars=50,
            display_text="فتح الرابط",
        )

    st.markdown(
        f"<div style='direction: rtl; text-align: right; margin-top: 18px;'><h3 style='margin: 0 0 12px; color: #e5e7eb; font-weight: 800;'>{title}</h3></div>",
        unsafe_allow_html=True,
    )
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config=column_config or None,
        height=440,
    )


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

    # Summary is already covered by the cards above, so only the detailed
    # table is shown on the page itself.
    _render_summary_cards(summary_df)
    _render_table(details_df, "تفاصيل العقود")

    excel_bytes = _excel_bytes(summary_df, details_df, date_from_value, date_to_value)

    st.download_button(
        label="⬇️ تنزيل Excel",
        data=excel_bytes,
        file_name="حصر_قيمه_عقود.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )