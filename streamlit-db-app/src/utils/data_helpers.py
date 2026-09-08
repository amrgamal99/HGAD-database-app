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


def _normalize_factory_name(value):
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    return text.replace("\u200f", "").replace("\u200e", "")


def prepare_payload_dates(payload: dict, date_fields: list) -> dict:
    out = payload.copy()
    for field in date_fields:
        if field in out:
            out[field] = normalize_date_for_supabase(out[field])
    return out


def format_data_for_display(data: pd.DataFrame) -> pd.DataFrame:
    """تنسيق عام إن رغبت باستخدامه لاحقًا."""
    df = data.copy()
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

    if "اسم المصنع" in df.columns and "factoryname" in df.columns:
        df["factoryname"] = df["factoryname"].combine_first(df["اسم المصنع"])
    if "اسم الشركة" in df.columns and "companyname" in df.columns:
        df["companyname"] = df["companyname"].combine_first(df["اسم الشركة"])

    if "اسم المصنع" in df.columns and "factoryname" not in df.columns:
        df["factoryname"] = df["اسم المصنع"]
    if "اسم الشركة" in df.columns and "companyname" not in df.columns:
        df["companyname"] = df["اسم الشركة"]

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

    df["factoryname"] = df["factoryname"].map(_normalize_factory_name)
    df["companyname"] = df["companyname"].map(_normalize_factory_name)
    df["اسم المشروع"] = df["اسم المشروع"].map(_normalize_factory_name)

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

    work = df.copy()
    if "factoryname" not in work.columns and "اسم المصنع" in work.columns:
        work = work.rename(columns={"اسم المصنع": "factoryname"})
    if "factoryname" not in work.columns or "قيمة التعاقد" not in work.columns:
        return pd.DataFrame()

    work["factoryname"] = work["factoryname"].map(_normalize_factory_name)
    work = work[work["factoryname"].isin(["التجمع", "بدر"])].copy()
    if work.empty:
        return pd.DataFrame(columns=["اسم المصنع", "قيمة العقود"])

    comparison = (
        work.groupby("factoryname", as_index=False)["قيمة التعاقد"]
        .sum()
        .rename(columns={"factoryname": "اسم المصنع", "قيمة التعاقد": "قيمة العقود"})
    )

    factory_order = ["التجمع", "بدر"]
    comparison["اسم المصنع"] = comparison["اسم المصنع"].map(_normalize_factory_name)
    comparison = comparison.set_index("اسم المصنع").reindex(factory_order).reset_index().rename(columns={"index": "اسم المصنع"})
    comparison["قيمة العقود"] = pd.to_numeric(comparison["قيمة العقود"], errors="coerce").fillna(0)
    comparison = comparison[comparison["اسم المصنع"].isin(factory_order)]
    return comparison


def build_contract_value_details(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    work = df.copy()
    if "factoryname" not in work.columns and "اسم المصنع" in work.columns:
        work["factoryname"] = work["اسم المصنع"]
    if "companyname" not in work.columns and "اسم الشركة" in work.columns:
        work["companyname"] = work["اسم الشركة"]
    if "factoryname" not in work.columns:
        work["factoryname"] = None
    if "companyname" not in work.columns:
        work["companyname"] = None

    work["factoryname"] = work["factoryname"].map(_normalize_factory_name)
    work["companyname"] = work["companyname"].map(_normalize_factory_name)
    work["اسم المشروع"] = work.get("اسم المشروع", pd.Series([None] * len(work))).map(_normalize_factory_name)

    work["اسم المصنع"] = work["factoryname"].combine_first(work.get("اسم المصنع", pd.Series([None] * len(work))))
    work["اسم الشركة"] = work["companyname"].combine_first(work.get("اسم الشركة", pd.Series([None] * len(work))))

    final_columns = [
        "اسم المصنع",
        "اسم الشركة",
        "اسم المشروع",
        "تاريخ التعاقد",
        "قيمة التعاقد",
        "قيمه التعاقد شامله الضريبه",
        "رابط نسخة العقد",
    ]

    details = work.copy()
    for col in final_columns:
        if col not in details.columns:
            details[col] = None

    details = details[final_columns].copy()
    return details.sort_values(["اسم المصنع", "اسم الشركة"], na_position="last")
