import streamlit as st
import pandas as pd
from db.connection import get_db_connection, fetch_data
from components.filters import (
    create_company_dropdown,
    create_project_dropdown,
    create_type_dropdown,
    create_column_search
)

# ---------- Page config ----------
st.set_page_config(
    page_title="عارض قاعدة البيانات والتقارير المالية | HGAD",
    layout="wide",
    initial_sidebar_state="expanded",  # open on load
)

def _global_styles():
    st.markdown("""
    <style>
        /* ===== Force sidebar open & remove collapse controls ===== */
        [data-testid="stSidebar"] { transform: none !important; }
        [data-testid="stSidebar"][aria-expanded="false"] { transform: none !important; }
        [data-testid="collapsedControl"] { display:none !important; }       /* bottom chevron */
        button[kind="header"] { display:none !important; }                   /* header hamburger */

        /* ===== Root RTL & wrap ===== */
        html, body {
            direction: rtl !important;
            text-align: right !important;
            font-family: "Cairo","Noto Kufi Arabic","Segoe UI",Tahoma,sans-serif !important;
            white-space: normal !important;
            word-wrap: break-word !important;
            overflow-x: hidden !important;
        }

        /* App content */
        [data-testid="stAppViewContainer"] * {
            direction: rtl !important;
            text-align: right !important;
            white-space: normal !important;
            writing-mode: horizontal-tb !important;
            text-orientation: mixed !important;
        }

        /* Keep Streamlit toolbar LTR so it never stacks vertically */
        [data-testid="stToolbar"], [data-testid="stToolbar"] * {
            direction: ltr !important;
            text-align: left !important;
            writing-mode: horizontal-tb !important;
            text-orientation: mixed !important;
            white-space: normal !important;
        }

        /* Sidebar theme + inputs */
        [data-testid="stSidebar"] {
            background:#0f172a; color:#e5e7eb; border-left:1px solid #1f2937;
        }
        [data-testid="stSidebar"] input,
        [data-testid="stSidebar"] textarea,
        [data-testid="stSidebar"] select {
            direction: rtl !important; text-align: right !important;
        }

        /* Alerts wrap nicely */
        [data-testid="stAlert"] { white-space: normal !important; }

        /* Header */
        .hgad-title { text-align:center; color:#1E3A8A; font-weight:800; line-height:1.3; margin:.25rem 0; }
        .hgad-sub { display:inline-block; font-size:20px; color:#4b5563; font-weight:600; margin-inline-start:.35rem; }
        .hgad-divider { border:0; height:2px; background:linear-gradient(to left, transparent, #1E3A8A, transparent);
                        margin:10px auto 20px; max-width:900px; border-radius:6px; }

        /* Dataframe header */
        [data-testid="stDataFrame"] thead tr th {
            position:sticky; top:0; background:#111827; color:#e5e7eb; z-index:2;
        }
    </style>
    """, unsafe_allow_html=True)

def _header():
    st.markdown("""
        <h1 class="hgad-title">
            عارض قاعدة البيانات والتقارير المالية
            <span class="hgad-sub">| HGAD Company</span>
        </h1>
        <hr class="hgad-divider"/>
    """, unsafe_allow_html=True)

# ---------- Utilities: clean/format data ----------
DATE_HINTS = ("تاريخ", "إصدار", "date")

def normalize_df_types(df: pd.DataFrame) -> pd.DataFrame:
    """Parse Arabic date columns, coerce numbers, keep everything display-friendly."""
    if df.empty: return df
    out = df.copy()

    # 1) Dates → pandas datetime (so we can render with DateColumn cleanly)
    date_cols = [c for c in out.columns if any(h in c for h in DATE_HINTS)]
    for col in date_cols:
        out[col] = pd.to_datetime(out[col], errors="coerce")

    # 2) Common numeric Arabic columns → numbers (for pretty formatting)
    num_hints = ("قيمة", "المستحق", "شيك", "التحويل", "USD", "USD)")
    num_cols = [c for c in out.columns if any(h in c for h in num_hints)]
    for col in num_cols:
        # strip commas/spaces if any came as strings
        out[col] = pd.to_numeric(
            out[col].astype(str).str.replace(",", "").str.replace("\u00a0", ""), errors="coerce"
        )

    return out

def build_column_config(df: pd.DataFrame):
    """Link columns, Date columns (YYYY-MM-DD), Number columns with thousands."""
    config = {}

    # Links (any column containing 'رابط')
    for col in df.columns:
        if "رابط" in col:
            config[col] = st.column_config.LinkColumn(label=col, display_text="فتح الرابط")

    # Dates
    for col in df.columns:
        if any(h in col for h in DATE_HINTS):
            config[col] = st.column_config.DateColumn(label=col, format="YYYY-MM-DD")

    # Numbers (pretty)
    num_hints = ("قيمة", "المستحق", "شيك", "التحويل", "USD", "USD)")
    for col in df.columns:
        if any(h in col for h in num_hints):
            config[col] = st.column_config.NumberColumn(label=col, format="%,.2f")

    return config

# ---------- App ----------
def main():
    _global_styles()
    _header()

    conn = get_db_connection()
    if conn is None:
        st.error("فشل الاتصال بقاعدة البيانات. يرجى مراجعة بيانات الاتصال والتأكد من تشغيل الخادم.")
        return

    # Sidebar (forced open by CSS above)
    with st.sidebar:
        st.title("عوامل التصفية")
        company_name = create_company_dropdown(conn)
        project_name = create_project_dropdown(conn, company_name)
        type_label, target_table = create_type_dropdown()

    if not company_name or not project_name or not target_table:
        st.info("برجاء اختيار الشركة والمشروع ونوع البيانات من الشريط الجانبي لعرض النتائج.")
        return

    # Fetch + normalize
    df = fetch_data(conn, company_name, project_name, target_table)
    if df.empty:
        st.warning("لا توجد بيانات مطابقة للاختيارات المحددة.")
        return

    df = normalize_df_types(df)

    # Optional search
    search_column, search_term = create_column_search(df)
    if search_column and search_term:
        # For dates, cast to string for contains; otherwise string repr
        df = df[df[search_column].astype(str).str.contains(str(search_term), case=False, na=False)]

    # Column config (links/dates/numbers)
    column_config = build_column_config(df)

    # Table
    st.markdown("### البيانات")
    st.dataframe(df, column_config=column_config, use_container_width=True, hide_index=True)

    # CSV download (UTF-8 BOM for Arabic/Excel)
    csv = df.to_csv(index=False, encoding="utf-8-sig")
    st.download_button(
        label="تنزيل كملف CSV",
        data=csv,
        file_name=f"{target_table}_{company_name}_{project_name}.csv",
        mime="text/csv",
    )

if __name__ == "__main__":
    main()
