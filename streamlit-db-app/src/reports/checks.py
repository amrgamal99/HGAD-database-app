from __future__ import annotations

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


def _render_factory_cards(df: pd.DataFrame) -> None:
    summary = _factory_check_summary(df)
    if summary.empty:
        st.info("لا توجد بيانات للشيكات خلال الفترة المحددة.")
        return

    for start in range(0, len(summary), 3):
        columns = st.columns(3)
        for column, (_, row) in zip(columns, summary.iloc[start:start + 3].iterrows()):
            with column:
                with st.container(border=True):
                    st.markdown(f"#### {row['اسم المصنع']}")
                    st.metric("إجمالي الشيكات", f"{row['الإجمالي']:,.2f}")
                    metric_columns = st.columns(3)
                    for metric_column, label in zip(metric_columns, PURPOSE_LABELS.values()):
                        metric_column.metric(label, f"{row[label]:,.2f}")


def _to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="الشيكات")
        writer.sheets["الشيكات"].right_to_left()
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