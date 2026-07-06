import datetime
import pandas as pd

def normalize_date_for_supabase(value):
    if value is None:
        return None
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.isoformat()
    try:
        ts = pd.to_datetime(value, errors="coerce")
        if pd.isna(ts):
            return None
        return ts.date().isoformat()
    except Exception:
        return None

def prepare_payload_dates(payload: dict, date_fields: list) -> dict:
    out = payload.copy()
    for field in date_fields:
        if field in out:
            out[field] = normalize_date_for_supabase(out[field])
    return out

def format_data_for_display(data: pd.DataFrame) -> pd.DataFrame:
    """تنسيق عام إن رغبت باستخدامه لاحقًا."""
    df = data.copy()
    # مثال: محاولة تنسيق الأعمدة التي تبدو كتواريخ
    for col in df.columns:
        if "تاريخ" in col or "إصدار" in col:
            try:
                df[col] = pd.to_datetime(df[col]).dt.strftime("%Y-%m-%d")
            except Exception:
                pass
    return df

def filter_data_by_company(data: pd.DataFrame, company_name: str) -> pd.DataFrame:
    return data[data["اسم الشركة"] == company_name]

def filter_data_by_project(data: pd.DataFrame, project_name: str) -> pd.DataFrame:
    return data[data["اسم المشروع"] == project_name]
