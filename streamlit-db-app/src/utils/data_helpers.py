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


def prepare_contract_values_dataframe(raw_df: pd.DataFrame, date_from=None, date_to=None) -> pd.DataFrame:
    """Normalize and filter the contract-values dataset for report rendering."""
    df = raw_df.copy() if raw_df is not None else pd.DataFrame()
    if df.empty:
        return df

    if "company" in df.columns:
        df["factoryname"] = df["company"].apply(
            lambda value: value.get("factoryname") if isinstance(value, dict) else None
        )
        df["companyname"] = df["company"].apply(
            lambda value: value.get("companyname") if isinstance(value, dict) else None
        )

    for col in [
        "factoryname",
        "companyname",
        "contractid",
        "companyid",
        "اسم المشروع",
        "قيمة التعاقد",
        "تاريخ التعاقد",
        "رابط نسخة العقد",
        "قيمه التعاقد شامله الضريبه",
    ]:
        if col not in df.columns:
            df[col] = None

    if "تاريخ التعاقد" in df.columns and not df.empty:
        df["تاريخ التعاقد"] = pd.to_datetime(df["تاريخ التعاقد"], errors="coerce")
    else:
        df["تاريخ التعاقد"] = pd.to_datetime([])

    if date_from:
        date_from_dt = pd.to_datetime(date_from)
        df = df[df["تاريخ التعاقد"] >= date_from_dt]
    if date_to:
        date_to_dt = pd.to_datetime(date_to)
        df = df[df["تاريخ التعاقد"] <= date_to_dt]

    df["قيمة التعاقد"] = pd.to_numeric(df["قيمة التعاقد"], errors="coerce").fillna(0)
    df["تاريخ التعاقد"] = df["تاريخ التعاقد"].dt.strftime("%Y-%m-%d")
    return df


def build_contract_value_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    comparison = (
        df[df["factoryname"].isin(["التجمع", "بدر"])]
        .groupby("factoryname", as_index=False)["قيمة التعاقد"]
        .sum()
    )
    comparison = comparison.rename(columns={"factoryname": "اسم المصنع", "قيمة التعاقد": "قيمة العقود"})
    return comparison


def build_contract_value_details(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    required_cols = [
        "factoryname",
        "companyname",
        "اسم المشروع",
        "تاريخ التعاقد",
        "قيمة التعاقد",
        "قيمه التعاقد شامله الضريبه",
        "رابط نسخة العقد",
    ]
    for col in required_cols:
        if col not in df.columns:
            df[col] = None

    details = df.rename(columns={
        "factoryname": "اسم المصنع",
        "companyname": "اسم الشركة",
        "اسم المشروع": "اسم المشروع",
        "تاريخ التعاقد": "تاريخ التعاقد",
        "قيمة التعاقد": "قيمة التعاقد",
        "قيمه التعاقد شامله الضريبه": "قيمه التعاقد شامله الضريبه",
        "رابط نسخة العقد": "رابط نسخة العقد",
    }).copy()
    details = details[required_cols].copy()
    details = details.rename(columns={
        "factoryname": "اسم المصنع",
        "companyname": "اسم الشركة",
    })
    return details.sort_values(["اسم المصنع", "اسم الشركة"], na_position="last")
