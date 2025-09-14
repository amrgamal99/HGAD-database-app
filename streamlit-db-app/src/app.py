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
    initial_sidebar_state="expanded",
)

# ---------- Styles ----------
st.markdown("""
<style>
/* ===== Force the sidebar ALWAYS open ===== */
/* Keep it expanded and visible */
[data-testid="stSidebar"] { transform: none !important; visibility: visible !important; width: 340px !important; min-width: 340px !important; }
/* Prevent any collapse transforms (new/old Streamlit versions) */
[data-testid="stSidebar"][aria-expanded="false"] { transform: none !important; visibility: visible !important; }
/* Hide collapse toggles (old & new selectors) */
[data-testid="collapsedControl"],
button[kind="header"],
button[title="Expand sidebar"],
button[title="Collapse sidebar"],
[data-testid="stSidebarCollapseButton"] {
    display: none !important;
}
/* Prevent accidental hotkey collapsing */
body { overscroll-behavior: none; }

/* ===== Root RTL & wrapping ===== */
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

/* Keep Streamlit top toolbar LTR so it never stacks vertically */
[data-testid="stToolbar"], [data-testid="stToolbar"] * {
    direction: ltr !important;
    text-align: left !important;
}

/* Sidebar theme */
[data-testid="stSidebar"] {
    background:#0f172a; color:#e5e7eb; border-left:1px solid #1f2937;
}

/* Header */
.hgad-title { text-align:center; color:#1E3A8A; font-weight:800; line-height:1.3; margin:.25rem 0; }
.hgad-sub { display:inline-block; font-size:20px; color:#4b5563; font-weight:600; margin-inline-start:.35rem; }
.hgad-divider { border:0; height:2px; background:linear-gradient(to left, transparent, #1E3A8A, transparent);
                margin:10px auto 20px; max-width:900px; border-radius:6px; }

/* DataFrame readability */
[data-testid="stDataFrame"] thead tr th {
    position: sticky; top: 0; background: #111827; color: #e5e7eb; z-index: 2;
    font-weight: 700;
}
[data-testid="stDataFrame"] div[role="row"] {
    font-size: 15px;
}
[data-testid="stDataFrame"] div[role="row"]:nth-child(even) {
    background-color: rgba(255,255,255,0.02);
}
</style>
""", unsafe_allow_html=True)

# ---------- Header ----------
st.markdown("""
    <h1 class="hgad-title">
        عارض قاعدة البيانات والتقارير المالية
        <span class="hgad-sub">| HGAD Company</span>
    </h1>
    <hr class="hgad-divider"/>
""", unsafe_allow_html=True)

# ---------- Helpers: normalize data ----------
DATE_HINTS = ("تاريخ", "إصدار", "date")  # matches e.g. "تاريخ انتهاء الضمان", "إصدار"
NUM_HINTS  = ("قيمة", "المستحق", "شيك", "التحويل", "USD", ")USD")

def normalize_df_types(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()

    # Dates → real datetime (to let Streamlit DateColumn render cleanly)
    date_cols = [c for c in out.columns if any(h in c for h in DATE_HINTS)]
    for col in date_cols:
        out[col] = pd.to_datetime(out[col], errors="coerce")

    # Numbers → numeric with thousands (strip commas/nbsp if strings)
    for col in [c for c in out.columns if any(h in c for h in NUM_HINTS)]:
        out[col] = (
            out[col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("\u00a0", "", regex=False)
        )
        out[col] = pd.to_numeric(out[col], errors="coerce")

    return out

def build_column_config(df: pd.DataFrame):
    cfg = {}
    # Links (any Arabic column containing 'رابط')
    for col in df.columns:
        if "رابط" in col:
            cfg[col] = st.column_config.LinkColumn(label=col, display_text="فتح الرابط")
    # Dates
    for col in df.columns:
        if any(h in col for h in DATE_HINTS):
            cfg[col] = st.column_config.DateColumn(label=col, format="YYYY-MM-DD")
    # Numbers
    for col in df.columns:
        if any(h in col for h in NUM_HINTS):
            cfg[col] = st.column_config.NumberColumn(label=col, format="%,.2f")
    return cfg

# ---------- App ----------
def main():
    conn = get_db_connection()
    if conn is None:
        st.error("فشل الاتصال بقاعدة البيانات. يرجى مراجعة بيانات الاتصال والتأكد من تشغيل الخادم.")
        return

    with st.sidebar:
        st.title("عوامل التصفية")
        company_name = create_company_dropdown(conn)
        project_name = create_project_dropdown(conn, company_name)
        type_label, target_table = create_type_dropdown()

    if not company_name or not project_name or not target_table:
        st.info("برجاء اختيار الشركة والمشروع ونوع البيانات من الشريط الجانبي لعرض النتائج.")
        return

    df = fetch_data(conn, company_name, project_name, target_table)

    if df.empty:
        st.warning("لا توجد بيانات مطابقة للاختيارات المحددة.")
        return

    # --- Normalize & search ---
    df = normalize_df_types(df)

    search_column, search_term = create_column_search(df)
    if search_column and search_term:
        df = df[df[search_column].astype(str).str.contains(str(search_term), case=False, na=False)]
        if df.empty:
            st.info("لا توجد نتائج بعد تطبيق معيار البحث.")
            return

    # --- Column config (links/dates/numbers) ---
    column_config = build_column_config(df)

    # --- Show table ---
    st.markdown("### البيانات")
    st.dataframe(df, column_config=column_config, use_container_width=True, hide_index=True)

    # --- CSV download (UTF-8 BOM for Arabic/Excel) ---
    csv = df.to_csv(index=False, encoding="utf-8-sig")
    st.download_button(
        label="تنزيل كملف CSV",
        data=csv,
        file_name=f"{target_table}_{company_name}_{project_name}.csv",
        mime="text/csv",
    )

if __name__ == "__main__":
    main()
