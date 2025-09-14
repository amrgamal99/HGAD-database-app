import streamlit as st
import pandas as pd
from supabase import create_client, Client

# تهيئة عميل Supabase مرة واحدة
@st.cache_resource
def get_db_connection() -> Client | None:
    try:
        url = st.secrets["supabase_url"]
        key = st.secrets["supabase_key"]
        supabase_client: Client = create_client(url, key)
        return supabase_client
    except Exception as e:
        st.error(f"فشل تهيئة عميل Supabase. راجع secrets. الخطأ: {e}")
        return None

# الشركات
def fetch_companies(supabase: Client) -> pd.DataFrame:
    try:
        resp = supabase.table("company").select("companyname").execute()
        df = pd.DataFrame(resp.data or [])
        df = df.rename(columns={"companyname": "اسم الشركة"})
        return df.drop_duplicates()
    except Exception:
        return pd.DataFrame(columns=["اسم الشركة"])

# المشاريع بحسب الشركة (يعتمد على الحقل العربي "اسم المشروع")
def fetch_projects_by_company(supabase: Client, company_name: str) -> pd.DataFrame:
    if not company_name:
        return pd.DataFrame(columns=["اسم المشروع"])
    try:
        company_resp = (
            supabase.table("company")
            .select("companyid")
            .eq("companyname", company_name)
            .single()
            .execute()
        )
        if not company_resp.data:
            return pd.DataFrame(columns=["اسم المشروع"])
        company_id = company_resp.data["companyid"]

        projects_resp = (
            supabase.table("contract")
            .select('اسم المشروع')
            .eq("companyid", company_id)
            .execute()
        )
        df = pd.DataFrame(projects_resp.data or [])
        return df.drop_duplicates()
    except Exception:
        return pd.DataFrame(columns=["اسم المشروع"])

# جلب البيانات بحسب (شركة + مشروع + جدول)
def fetch_data(supabase: Client, company_name: str, project_name: str, target_table: str) -> pd.DataFrame:
    """
    target_table one of: 'contract' | 'guarantee' | 'invoice' | 'checks'
    يطابق المخطط الجديد بأسماء الأعمدة العربية داخل الجداول.
    """
    try:
        # احصل على companyid
        company_resp = (
            supabase.table("company")
            .select("companyid")
            .eq("companyname", company_name)
            .single()
            .execute()
        )
        if not company_resp.data:
            return pd.DataFrame()
        company_id = company_resp.data["companyid"]

        # احصل على contractid للمشروع المحدد
        contract_resp = (
            supabase.table("contract")
            .select("contractid, companyid, اسم المشروع")
            .eq("companyid", company_id)
            .eq("اسم المشروع", project_name)
            .single()
            .execute()
        )
        if not contract_resp.data:
            return pd.DataFrame()
        contract_id = contract_resp.data["contractid"]

        # احضار البيانات
        tbl = target_table.lower()
        if tbl == "contract":
            resp = (
                supabase.table("contract")
                .select("*")
                .eq("contractid", contract_id)
                .single()
                .execute()
            )
            data = [resp.data] if resp.data else []
        elif tbl in ["guarantee", "invoice", "checks"]:
            resp = (
                supabase.table(tbl)
                .select("*")
                .eq("companyid", company_id)
                .eq("contractid", contract_id)
                .execute()
            )
            data = resp.data or []
        else:
            st.error("نوع البيانات غير صالح.")
            return pd.DataFrame()

        df = pd.DataFrame(data)

        # حذف أعمدة IDs من العرض
        if not df.empty:
            cols_to_drop = [c for c in df.columns if c.lower().endswith("id")]
            df.drop(columns=cols_to_drop, inplace=True, errors="ignore")

        # تنسيقات طفيفة (تواريخ/أرقام/روابط)
        df = _format_columns_for_display(df)
        return df

    except Exception as e:
        st.error(f"حدث خطأ أثناء جلب البيانات: {e}")
        return pd.DataFrame()

def _format_columns_for_display(df: pd.DataFrame) -> pd.DataFrame:
    df2 = df.copy()
    # تحويل أعمدة التواريخ العربية إن وُجدت إلى صيغة YYYY-MM-DD
    date_like_cols = [c for c in df2.columns if "تاريخ" in c or "إصدار" in c]
    for col in date_like_cols:
        try:
            df2[col] = pd.to_datetime(df2[col]).dt.strftime("%Y-%m-%d")
        except Exception:
            pass
    # لا تغيّر أعمدة الأرقام؛ تترك كما هي (Postgres numeric → float/decimal)
    return df2
