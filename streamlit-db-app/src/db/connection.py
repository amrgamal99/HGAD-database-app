import streamlit as st
import pandas as pd
from supabase import create_client, Client
from typing import Dict, Optional, Tuple
import re

# تهيئة عميل Supabase مرة واحدة
@st.cache_resource
def get_db_connection() -> Optional[Client]:
    try:
        url = st.secrets["supabase_url"]
        key = st.secrets["supabase_key"]
        supabase_client: Client = create_client(url, key)
        return supabase_client
    except Exception as e:
        st.error(f"فشل تهيئة عميل Supabase. راجع secrets. الخطأ: {e}")
        return None

# الشركات
def fetch_companies(supabase: Client, factory_name: Optional[str] = None) -> pd.DataFrame:
    try:
        query = supabase.table("company").select("companyid, companyname, \"اخر تعديل\"")
        if factory_name and factory_name != "الكل":
            query = query.eq("factoryname", factory_name)
        resp = query.execute()
        df = pd.DataFrame(resp.data or [])
    except Exception:
        query = supabase.table("company").select("companyid, companyname")
        if factory_name and factory_name != "الكل":
            query = query.eq("factoryname", factory_name)
        resp = query.execute()
        df = pd.DataFrame(resp.data or [])

    if "companyname" not in df.columns:
        df = pd.DataFrame(columns=["companyname"])
    if "اخر تعديل" in df.columns:
        try:
            df["اخر تعديل"] = pd.to_datetime(df["اخر تعديل"], errors="coerce").dt.date.astype(str)
        except Exception:
            pass
    df = df.rename(columns={"companyname": "اسم الشركة", "اخر تعديل": "آخر تعديل"})
    if "اسم الشركة" not in df.columns:
        df = pd.DataFrame(columns=["اسم الشركة"])
    if not df.empty:
        df = df.drop_duplicates(subset=["اسم الشركة"], keep="last")
        df = df.sort_values("اسم الشركة", key=lambda s: s.str.lower())
    return df

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

        try:
            projects_resp = (
                supabase.table("contract")
                .select('"اسم المشروع", "اخر تعديل"')
                .eq("companyid", company_id)
                .execute()
            )
            df = pd.DataFrame(projects_resp.data or [])
        except Exception:
            projects_resp = (
                supabase.table("contract")
                .select('"اسم المشروع"')
                .eq("companyid", company_id)
                .execute()
            )
            df = pd.DataFrame(projects_resp.data or [])

        if "اخر تعديل" in df.columns:
            try:
                df["اخر تعديل"] = pd.to_datetime(df["اخر تعديل"], errors="coerce").dt.date.astype(str)
            except Exception:
                pass
        df = df.rename(columns={'"اسم المشروع"': 'اسم المشروع', 'اخر تعديل': 'آخر تعديل'})
        if "اسم المشروع" not in df.columns:
            return pd.DataFrame(columns=["اسم المشروع"])
        if not df.empty:
            if "آخر تعديل" in df.columns:
                df = df.sort_values(["اسم المشروع", "آخر تعديل"], ascending=[True, False], na_position="last")
            df = df.drop_duplicates(subset=["اسم المشروع"], keep="first")
        return df

    except Exception as e:
        st.caption(f"⚠️ fetch_projects_by_company error: {e}")
        return pd.DataFrame(columns=["اسم المشروع"])


def _fetch_latest_last_edit(supabase: Client, table_name: str) -> Optional[str]:
    try:
        resp = (
            supabase.table(table_name)
            .select('"اخر تعديل"')
            .order('"اخر تعديل"', desc=True)
            .limit(1)
            .execute()
        )
        if resp.data and isinstance(resp.data, list):
            val = resp.data[0].get("اخر تعديل")
            if val:
                try:
                    return pd.to_datetime(val, errors="coerce").date().isoformat()
                except Exception:
                    return str(val)[:10]
    except Exception:
        pass
    return None


def fetch_type_last_edit_dates(supabase: Client) -> Dict[str, Optional[str]]:
    mapping = {
        "financial_report": "contract",
        "contract": "contract",
        "guarantee": "guarantee",
        "invoice": "invoice",
        "checks": "checks",
        "social_insurance_certificate": "social_insurance_certificate",
        "supplier_costs": "supplier_monthly_cost",
    }
    return {key: _fetch_latest_last_edit(supabase, table) for key, table in mapping.items()}


def _get_company_and_contract_ids(
    supabase: Client, company_name: str, project_name: str
) -> Tuple[Optional[int], Optional[int]]:
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
            resp = (
                supabase.table(tbl)
                .select("*")
                .eq("companyid", company_id)
                .eq("contractid", contract_id)
                .execute()
            )
            data = resp.data or []
        elif tbl == "supplier_costs":
            resp = supabase.table("supplier_monthly_cost").select("*").eq("companyid", company_id)
            if contract_id is not None:
                resp = resp.eq("contractid", contract_id)
            resp = resp.execute()
            data = resp.data or []
            df = pd.DataFrame(data)
            if not df.empty:
                supplier_ids = df.get("supplierid")
                if supplier_ids is not None:
                    try:
                        supplier_ids = pd.Series(supplier_ids).dropna().astype(int).unique().tolist()
                    except Exception:
                        supplier_ids = []
                else:
                    supplier_ids = []

                if supplier_ids:
                    suppliers_resp = (
                        supabase.table("supplier")
                        .select("supplierid, \"اسم المورد\", \"المواد الخام\"")
                        .in_("supplierid", supplier_ids)
                        .execute()
                    )
                    suppliers = suppliers_resp.data or []
                    supplier_df = pd.DataFrame(suppliers)
                    if not supplier_df.empty:
                        df = df.merge(supplier_df, on="supplierid", how="left")

                desired_cols = [
                    "اسم المورد",
                    "المواد الخام",
                    "من تاريخ",
                    "الي تاريخ",
                    "القيمة خلال الفترة",
                ]
                cols_present = [c for c in desired_cols if c in df.columns]
                df = df[cols_present] if cols_present else df
        else:
            st.error("نوع البيانات غير صالح.")
            return pd.DataFrame()

        if tbl != "supplier_costs":
            df = pd.DataFrame(data)
            if not df.empty:
                df.drop(columns=[c for c in df.columns if c.lower().endswith("id")],
                        inplace=True, errors="ignore")
                df.dropna(axis=1, how="all", inplace=True)
        else:
            if not df.empty:
                df.dropna(axis=1, how="all", inplace=True)

            if tbl == "guarantee":
                guarantee_cols = [
                    "رقم خطاب الضمان",
                    "البنك المصدر",
                    "تاريخ إصدار الضمان",
                    "تاريخ انتهاء الضمان",
                    "رابط نسخة الضمان",
                    "الغرض من اصدار خطاب ضمان",
                    "قيمة خطاب الضمان الحالية",
                    "قيمه ما تم تخفيضه في خطاب ضمان",
                ]
                existing_cols = [c for c in guarantee_cols if c in df.columns]
                other_cols = [c for c in df.columns if c not in existing_cols]
                df = df[existing_cols + other_cols]

        if tbl == "guarantee" and not df.empty and "رقم خطاب الضمان" in df.columns:
            sort_by = ["رقم خطاب الضمان"]
            ascending = [True]
            if "تاريخ إصدار الضمان" in df.columns:
                sort_by.append("تاريخ إصدار الضمان")
                ascending.append(True)
            if "تاريخ انتهاء الضمان" in df.columns:
                sort_by.append("تاريخ انتهاء الضمان")
                ascending.append(True)
            try:
                df = df.sort_values(by=sort_by, ascending=ascending, na_position="last")
            except Exception:
                pass

        for col in df.columns:
            if "تاريخ" in col or "إصدار" in col:
                try:
                    df[col] = pd.to_datetime(df[col]).dt.strftime("%Y-%m-%d")
                except Exception:
                    pass

        if tbl == "social_insurance_certificate" and "اسم الشهادة" in df.columns:
            try:
                def normalize_name(val):
                    if pd.isna(val):
                        return "شهاده تامينات جاري"
                    s = str(val).strip()
                    m = re.search(r'([0-9\u0660-\u0669\u06F0-\u06F9]+)', s)
                    if m:
                        num = m.group(1)
                        return f"شهاده تامينات جاري {num}"
                    return "شهاده تامينات جاري"
                df["اسم الشهادة"] = df["اسم الشهادة"].apply(normalize_name)
            except Exception:
                pass

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


# ── جلب v_financial_flow ──────────────────────────────────────────────────────
def fetch_financial_flow_view(
    supabase: Client,
    company_name: str,
    project_name: str,
    date_from=None,
    date_to=None,
) -> pd.DataFrame:
    company_id, contract_id = _get_company_and_contract_ids(supabase, company_name, project_name)
    if not company_id or not contract_id:
        return pd.DataFrame()

    try:
        # ✅ Column order matches the updated view:
        # صافي المستحق بعد الخصومات moved after اجمالي خصومات, before رقم الشيك
        select_cols = (
            'companyid,'
            'contractid,'
            '"التاريخ",'
            '"نوع العملية",'
            '"اسم المستخلص",'
            '"إجمالي المستخلص شامل الضريبة",'
            '"قيمة المستخلص قبل الخصومات",'
            '"ضريبة قيمة مضافة",'
            '"تأمين ابتدائي",'
            '"تأمين نهائي",'
            '"(ضريبة خصم %1) خصم منبع",'
            '"دمغة هندسية",'
            '"خصم دفعة مقدمة",'
            '"خصم دفعة مقدمة 2",'
            '"خصومات موقع",'
            '"تامينات اجتماعية",'
            '"عمالة غير منتظمة",'
            '"استهلاك استشاري و مالي",'
            '"خصم تعاقدي",'
            '"دمغة اتحاد وتشيد",'
            '"ضرائب عامة",'
            '"اجمالي خصومات",'
            '"صافي المستحق بعد الخصومات",'
            '"رقم الشيك",'
            '"البنك",'
            '"قيمة الشيك",'
            '"الغرض من إصدار الشيك",'
            '"المتبقي",'
            '"المستحق صرفه من تامينات اجتماعيه"'
        )

        q = (
            supabase.table("v_financial_flow")
            .select(select_cols)
            .eq("contractid", contract_id)
        )

        if date_from:
            q = q.filter('التاريخ', 'gte', str(date_from))
        if date_to:
            q = q.filter('التاريخ', 'lte', str(date_to))

        resp = q.order("التاريخ", desc=False).execute()
        df = pd.DataFrame(resp.data or [])

        if not df.empty:
            if "التاريخ" in df.columns:
                try:
                    df["التاريخ"] = pd.to_datetime(df["التاريخ"]).dt.strftime("%Y-%m-%d")
                except Exception:
                    pass
            # Auto-drops fully-NULL columns (e.g. insurance col for contracts with no insurance)
            df.dropna(axis=1, how="all", inplace=True)

        return df
    except Exception as e:
        st.caption(f"⚠️ fetch_financial_flow_view error: {e}")
        return pd.DataFrame()


# ── جلب v_contract_summary ────────────────────────────────────────────────────
def fetch_contract_summary_view(
    supabase: Client,
    company_name: str,
    project_name: str,
) -> pd.DataFrame:
    company_id, contract_id = _get_company_and_contract_ids(supabase, company_name, project_name)
    if not company_id or not contract_id:
        return pd.DataFrame()
    try:
        resp = (
            supabase.table("v_contract_summary")
            .select(
                '"اسم المشروع",'
                '"تاريخ التعاقد",'
                '"قيمة التعاقد",'
                '"حجم الاعمال المنفذة",'
                '"نسبة الاعمال المنفذة",'
                '"حجم الاعمال المنفذة 2026",'
                '"نسبة الاعمال المنفذة 2026",'
                '"حجم الاعمال المنفذة الفعليه 2026",'
                '"نسبة الاعمال المنفذة الفعليه 2026",'
                '"الدفعه المقدمه",'
                '"التحصيلات",'
                '"التحصيلات 2026",'
                '"التحصيلات الفعليه 2026",'
                '"مواد اوليه",'
                '"مواد اوليه 2026",'
                '"مقاولين باطن",'
                '"مقاولين باطن 2026",'
                '"مصروفات تامينات اجتماعية",'
                '"مصروفات تامينات اجتماعية 2026",'
                '"المستحق صرفه من تامينات اجتماعيه",'
                '"المتبقي من دفعات المقدمه",'
                '"المستحق صرفه"'
            )
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
            # NULL columns auto-dropped for contracts with no insurance/advance data
            df.dropna(axis=1, how="all", inplace=True)
        return df.head(1)
    except Exception as e:
        st.caption(f"⚠️ fetch_contract_summary_view error: {e}")
        return pd.DataFrame()