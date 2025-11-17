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
        # sort alphabetically
        if not df.empty:
            df = df.drop_duplicates().sort_values("اسم الشركة", key=lambda s: s.str.lower())
        else:
            df = df.drop_duplicates()
        return df
    except Exception:
        return pd.DataFrame(columns=["اسم الشركة"])

# المشاريع بحسب الشركة
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
            .select('"اسم المشروع"')
            .eq("companyid", company_id)
            .execute()
        )
        df = pd.DataFrame(projects_resp.data or [])
        df = df.rename(columns={'"اسم المشروع"': 'اسم المشروع'})
        return df.drop_duplicates()

    except Exception as e:
        st.caption(f"⚠️ fetch_projects_by_company error: {e}")
        return pd.DataFrame(columns=["اسم المشروع"])

def _get_company_and_contract_ids(supabase: Client, company_name: str, project_name: str) -> tuple[int | None, int | None]:
    try:
        company_resp = (
            supabase.table("company")
            .select("companyid, companyname")
            .eq("companyname", company_name)
            .single()
            .execute()
        )
        if not company_resp.data:
            return None, None
        company_id = company_resp.data["companyid"]

        contract_resp = (
            supabase.table("contract")
            .select('contractid, companyid, "اسم المشروع"')
            .eq("companyid", company_id)
            .filter('اسم المشروع', 'eq', project_name)
            .single()
            .execute()
        )
        if not contract_resp.data:
            return company_id, None
        return company_id, contract_resp.data["contractid"]
    except Exception as e:
        st.caption(f"⚠️ _get_company_and_contract_ids error: {e}")
        return None, None

# البيانات الخام (للأنواع الأخرى)
def fetch_data(supabase: Client, company_name: str, project_name: str, target_table: str) -> pd.DataFrame:
    try:
        company_id, contract_id = _get_company_and_contract_ids(supabase, company_name, project_name)
        if not company_id or not contract_id:
            return pd.DataFrame()

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
        elif tbl in ["guarantee", "invoice", "checks", "social_insurance_certificate"]:
            # treat social_insurance_certificate same as invoice (filtered by companyid+contractid)
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
        if not df.empty:
            df.drop(columns=[c for c in df.columns if c.lower().endswith("id")],
                    inplace=True, errors="ignore")

        # format date-like columns
        for col in df.columns:
            if "تاريخ" in col or "إصدار" in col:
                try:
                    df[col] = pd.to_datetime(df[col]).dt.strftime("%Y-%m-%d")
                except Exception:
                    pass

        # If it's the social_insurance_certificate table, normalize the "اسم الشهادة" values
        if tbl == "social_insurance_certificate" and "اسم الشهادة" in df.columns:
            try:
                df["اسم الشهادة"] = "شهاده تامينات جاري"
            except Exception:
                pass

        # Sort by the first date-like column (old → new) for invoice/checks/social_insurance_certificate
        if tbl in ["invoice", "checks", "social_insurance_certificate"] and not df.empty:
            date_cols = [c for c in df.columns if "تاريخ" in c or "إصدار" in c]
            if date_cols:
                try:
                    df = df.sort_values(by=date_cols[0], ascending=True, na_position="last")
                except Exception:
                    pass

        return df
    except Exception as e:
        st.caption(f"⚠️ fetch_data error: {e}")
        return pd.DataFrame()

# === NEW: جلب v_financial_flow مع فلاتر التاريخ ===
def fetch_financial_flow_view(supabase: Client, company_name: str, project_name: str,
                              date_from=None, date_to=None) -> pd.DataFrame:
    company_id, contract_id = _get_company_and_contract_ids(supabase, company_name, project_name)
    if not company_id or not contract_id:
        return pd.DataFrame()

    try:
        q = supabase.table("v_financial_flow").select(
            'companyid, contractid, "التاريخ", "نوع العملية", "اسم المستخلص", '
            '"قيمة المستخلص قبل الخصومات", "صافي المستحق بعد الخصومات", '
            '"رقم الشيك", "البنك", "قيمة الشيك", "الغرض من إصدار الشيك", "المتبقي"'
        ).eq("contractid", contract_id)

        # استخدم filter مع أسماء الأعمدة العربية
        if date_from:
            q = q.filter('التاريخ', 'gte', str(date_from))
        if date_to:
            q = q.filter('التاريخ', 'lte', str(date_to))

        resp = q.order("التاريخ", desc=False).execute()
        df = pd.DataFrame(resp.data or [])
        # تنسيقات
        if not df.empty:
            for col in ["التاريخ"]:
                try:
                    df[col] = pd.to_datetime(df[col]).dt.strftime("%Y-%m-%d")
                except Exception:
                    pass
        return df
    except Exception as e:
        st.caption(f"⚠️ fetch_financial_flow_view error: {e}")
        return pd.DataFrame()

# === NEW: جلب v_contract_summary كسطر ملخّص واحد للمشروع ===
def fetch_contract_summary_view(supabase: Client, company_name: str, project_name: str) -> pd.DataFrame:
    company_id, contract_id = _get_company_and_contract_ids(supabase, company_name, project_name)
    if not company_id or not contract_id:
        return pd.DataFrame()
    try:
        # نختار بالاسم والتعاقد لضمان صف واحد
        resp = (
            supabase.table("v_contract_summary")
            .select('"اسم المشروع","تاريخ التعاقد","قيمة التعاقد","حجم الاعمال المنفذة",'
                    '"نسبة الاعمال المنفذة","الدفعه المقدمه","التحصيلات","المستحق صرفه"')
            .filter('اسم المشروع', 'eq', project_name)
            .execute()
        )
        df = pd.DataFrame(resp.data or [])
        if not df.empty:
            if "تاريخ التعاقد" in df.columns:
                try:
                    df["تاريخ التعاقد"] = pd.to_datetime(df["تاريخ التعاقد"]).dt.strftime("%Y-%m-%d")
                except Exception:
                    pass
        return df.head(1)  # بطاقة واحدة
    except Exception as e:
        st.caption(f"⚠️ fetch_contract_summary_view error: {e}")
        return pd.DataFrame()
