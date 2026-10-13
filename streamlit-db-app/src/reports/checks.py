from __future__ import annotations

import html
from io import BytesIO
from typing import Optional

import pandas as pd
import streamlit as st
from supabase import Client

from db.connection import fetch_checks_report_data

CHECK_DATE = "تاريخ الشيك"
RECEIPT_DATE = "تاريخ استلام شيك"
DATE_FIELDS = (CHECK_DATE, RECEIPT_DATE)

CHECK_COLUMNS = (
    "اسم المصنع",
    "اسم الشركة",
    "اسم العقد",
    "رقم الشيك",
    CHECK_DATE,
    "البنك",
    "قيمة الشيك",
    "رابط نسخة الشيك",
    "الغرض من إصدار الشيك",
    RECEIPT_DATE,
    "السنة_المالية_المحاسبية",
)

PURPOSE_LABELS = {
    "دفعة مقدمة": "دفعة مقدمة",
    "رد تامينات اجتماعيه": "رد تامينات اجتماعيه",
    "سداد مستخلص": "سداد مستخلص",
}


def _to_iso_date(value) -> Optional[str]:
    if value is None:
        return None
    parsed = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(parsed) else parsed.date().isoformat()


def _normalize_purpose(value: object) -> str:
    text = str(value).strip().replace("أ", "ا").replace("إ", "ا")
    text = text.replace("ة", "ه").replace("ى", "ي")
    return " ".join(text.split())


def _prepare_display_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()
    aliases = {
        "factoryname": "اسم المصنع",
        "companyname": "اسم الشركة",
        "اسم المشروع": "اسم العقد",
    }
    for source, target in aliases.items():
        if source in work.columns and target not in work.columns:
            work = work.rename(columns={source: target})

    link_source = next(
        (column for column in ("رابط الشيك", "رابط نسخة الشيك") if column in work.columns),
        None,
    )
    if link_source and link_source != "رابط نسخة الشيك":
        work = work.rename(columns={link_source: "رابط نسخة الشيك"})

    for column in CHECK_COLUMNS:
        if column not in work.columns:
            work[column] = pd.NA

    display_df = work.loc[:, list(CHECK_COLUMNS)].copy()
    for column in (CHECK_DATE, RECEIPT_DATE):
        display_df[column] = (
            pd.to_datetime(display_df[column], errors="coerce")
            .dt.strftime("%Y-%m-%d")
            .fillna("")
        )
    return display_df


def _factory_check_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "اسم المصنع" not in df.columns or "قيمة الشيك" not in df.columns:
        return pd.DataFrame(columns=["اسم المصنع", "الإجمالي", *PURPOSE_LABELS.values()])

    work = df.copy()
    work["قيمة الشيك"] = pd.to_numeric(work["قيمة الشيك"], errors="coerce").fillna(0)
    work["_purpose"] = work.get("الغرض من إصدار الشيك", pd.Series(index=work.index, dtype=object)).map(
        _normalize_purpose
    )
    rows = []
    for factory, factory_df in work.groupby("اسم المصنع", dropna=True, sort=True):
        row = {"اسم المصنع": factory, "الإجمالي": factory_df["قيمة الشيك"].sum()}
        for normalized, label in PURPOSE_LABELS.items():
            target = _normalize_purpose(normalized)
            row[label] = factory_df.loc[
                factory_df["_purpose"] == target, "قيمة الشيك"
            ].sum()
        rows.append(row)
    return pd.DataFrame(rows, columns=["اسم المصنع", "الإجمالي", *PURPOSE_LABELS.values()])


def _factory_cards_html(summary: pd.DataFrame) -> str:
    cards = []
    for row in summary.to_dict("records"):
        categories = "".join(
            f'<div class="checks-category"><div class="checks-category-label">{html.escape(label)}</div>'
            f'<div class="checks-category-value" dir="ltr">{float(row[label]):,.2f}</div></div>'
            for label in PURPOSE_LABELS.values()
        )
        cards.append(
            f'<article class="checks-factory-card" dir="rtl">'
            f'<h4 class="checks-factory-name">{html.escape(str(row["اسم المصنع"]))}</h4>'
            f'<div class="checks-total-label">إجمالي الشيكات</div>'
            f'<div class="checks-total-value" dir="ltr">{float(row["الإجمالي"]):,.2f}</div>'
            f'<div class="checks-category-grid">{categories}</div>'
            f'</article>'
        )

    return (
        "<style>"
        ".checks-factory-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,360px),1fr));gap:14px;direction:rtl;}"
        ".checks-factory-card{box-sizing:border-box;min-width:0;padding:18px;border:1px solid #3b4048;border-radius:8px;background:#111419;color:#f4f5f7;text-align:right;}"
        ".checks-factory-name{margin:0 0 18px;font-size:19px;font-weight:750;line-height:1.4;}"
        ".checks-total-label{margin-bottom:3px;color:#b7bdc7;font-size:12px;font-weight:600;}"
        ".checks-total-value{margin-bottom:18px;font-size:28px;font-weight:700;line-height:1.25;font-variant-numeric:tabular-nums;overflow-wrap:anywhere;}"
        ".checks-category-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;}"
        ".checks-category{min-width:0;padding:10px 7px;border-top:1px solid #343943;text-align:center;}"
        ".checks-category-label{min-height:34px;color:#b7bdc7;font-size:12px;font-weight:600;line-height:1.4;overflow-wrap:anywhere;}"
        ".checks-category-value{font-size:15px;font-weight:700;line-height:1.4;font-variant-numeric:tabular-nums;overflow-wrap:anywhere;}"
        "@media(max-width:420px){.checks-factory-card{padding:14px;}.checks-category-grid{gap:3px;}.checks-category{padding-inline:3px;}.checks-category-value{font-size:13px;}}"
        "</style>"
        f'<div class="checks-factory-grid">{"".join(cards)}</div>'
    )


def _render_factory_cards(df: pd.DataFrame) -> None:
    summary = _factory_check_summary(df)
    if summary.empty:
        st.info("لا توجد بيانات للشيكات خلال الفترة المحددة.")
        return
    st.markdown(_factory_cards_html(summary), unsafe_allow_html=True)


def _to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        from xlsxwriter.utility import xl_col_to_name

        workbook = writer.book
        sheet_name = "الشيكات"
        start_row = 5
        df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=start_row)
        worksheet = writer.sheets[sheet_name]
        worksheet.right_to_left()
        worksheet.hide_gridlines(2)
        worksheet.set_zoom(110)
        worksheet.set_row(0, 25)
        worksheet.set_row(1, 25)

        column_count = len(df.columns)
        last_column = xl_col_to_name(column_count - 1)
        title_format = workbook.add_format({
            "bold": True,
            "font_size": 16,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#1F4E78",
            "font_color": "white",
            "border": 1,
            "text_wrap": True,
        })
        worksheet.merge_range(
            f"A1:{last_column}2",
            "تقرير الشيكات",
            title_format,
        )

        row_count = len(df)
        worksheet.add_table(
            start_row,
            0,
            start_row + row_count,
            column_count - 1,
            {
                "name": "ChecksTable",
                "style": "Table Style Medium 2",
                "columns": [{"header": str(column)} for column in df.columns],
            },
        )

        widths = {
            "اسم المصنع": 24,
            "اسم الشركة": 28,
            "اسم العقد": 34,
            "رقم الشيك": 16,
            CHECK_DATE: 16,
            "البنك": 22,
            "قيمة الشيك": 18,
            "رابط نسخة الشيك": 18,
            "الغرض من إصدار الشيك": 30,
            RECEIPT_DATE: 18,
            "السنة_المالية_المحاسبية": 22,
        }
        amount_format = workbook.add_format({"num_format": "#,##0.00"})
        hyperlink_format = workbook.add_format({
            "font_color": "blue",
            "underline": 1,
        })

        link_column = None
        for column_index, column in enumerate(df.columns):
            worksheet.set_column(
                column_index,
                column_index,
                widths.get(str(column), 20),
                amount_format if column == "قيمة الشيك" else None,
            )
            if column == "رابط نسخة الشيك":
                link_column = column_index

        if link_column is not None:
            for row_index, value in enumerate(df.iloc[:, link_column], start=start_row + 1):
                if pd.isna(value):
                    continue
                url = str(value).strip()
                if url.startswith(("http://", "https://")):
                    worksheet.write_url(
                        row_index,
                        link_column,
                        url,
                        hyperlink_format,
                        string="فتح الشيك",
                    )
                elif url:
                    worksheet.write(row_index, link_column, url)

        worksheet.freeze_panes(start_row + 1, 0)
    return buffer.getvalue()


def render_checks_report(
    conn: Optional[Client],
    date_field: str = CHECK_DATE,
    date_from=None,
    date_to=None,
) -> None:
    st.markdown("<h2 style='text-align:right'>التحويلات والشيكات</h2>", unsafe_allow_html=True)
    if conn is None:
        st.error("تعذر الاتصال بقاعدة البيانات.")
        return
    if date_field not in DATE_FIELDS:
        date_field = CHECK_DATE

    selected_from = _to_iso_date(date_from)
    selected_to = _to_iso_date(date_to)
    if selected_from and selected_to and selected_from > selected_to:
        st.error("تاريخ البداية يجب أن يسبق تاريخ النهاية.")
        return

    df = fetch_checks_report_data(
        conn,
        date_field=date_field,
        date_from=selected_from,
        date_to=selected_to,
    )
    st.markdown("### الشيكات خلال فتره زمنيه")
    if df.empty:
        st.info("لا توجد شيكات مطابقة.")
        return

    st.markdown("### الشيكات خلال فترة لكل مصنع")
    display_df = _prepare_display_dataframe(df)
    _render_factory_cards(display_df)
    st.divider()
    st.markdown("### تفاصيل الشيكات")
    st.dataframe(
        display_df,
        column_config={
            "رابط نسخة الشيك": st.column_config.LinkColumn(
                label="رابط نسخة الشيك",
                display_text="فتح الشيك",
                max_chars=50,
            ),
        },
        use_container_width=True,
        hide_index=True,
    )
    st.download_button(
        label="تنزيل Excel",
        data=_to_excel_bytes(display_df),
        file_name="تقرير_الشيكات.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="checks_excel_download",
    )