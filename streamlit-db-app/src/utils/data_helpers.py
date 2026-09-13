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
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    text = str(value).strip()
    if not text or text.lower() in {"none", "nan", "null"}:
        return None
    text = text.replace("\u200f", "").replace("\u200e", "")
    return text if text and text.lower() not in {"none", "nan", "null"} else None


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
        nested_factory = df["company"].apply(
            lambda value: value.get("factoryname") if isinstance(value, dict) else None
        )
        nested_company = df["company"].apply(
            lambda value: value.get("companyname") if isinstance(value, dict) else None
        )

        if "factoryname" not in df.columns:
            df["factoryname"] = nested_factory
        else:
            df["factoryname"] = df["factoryname"].combine_first(nested_factory)

        if "companyname" not in df.columns:
            df["companyname"] = nested_company
        else:
            df["companyname"] = df["companyname"].combine_first(nested_company)

    for target_col, aliases in [
        ("factoryname", ["factoryname", "اسم المصنع", "factory"]),
        ("companyname", ["companyname", "اسم الشركة", "company"]),
    ]:
        values = []
        for alias in aliases:
            if alias in df.columns:
                values.append(df[alias].map(_normalize_factory_name))
        if values:
            merged = values[0]
            for value in values[1:]:
                merged = merged.combine_first(value)
            df[target_col] = merged
        elif target_col not in df.columns:
            df[target_col] = None

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
    work = work[work["factoryname"].notna()].copy()
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

    work["اسم المصنع"] = work["factoryname"].combine_first(work.get("اسم المصنع", pd.Series([None] * len(work)))).map(_normalize_factory_name)
    work["اسم الشركة"] = work["companyname"].combine_first(work.get("اسم الشركة", pd.Series([None] * len(work)))).map(_normalize_factory_name)

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


def _invoice_value_column(df: pd.DataFrame) -> str | None:
    for column in ("إجمالي المستخلص شامل الضريبة", "قيمة المستخلص قبل الخصومات"):
        if column in df.columns:
            return column
    return None


def prepare_invoice_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize invoice dates, labels, and numeric fields for reporting."""
    work = df.copy() if df is not None else pd.DataFrame()
    if work.empty:
        return work
    for source, target in (("factoryname", "مصنع"), ("companyname", "اسم الشركة")):
        if source in work.columns and target not in work.columns:
            work[target] = work[source]
    if "factoryname" in work.columns:
        work["مصنع"] = work["مصنع"].fillna(work["factoryname"])
    if "تاريخ إصدار المستخلص" in work.columns:
        work["تاريخ إصدار المستخلص"] = pd.to_datetime(
            work["تاريخ إصدار المستخلص"], errors="coerce"
        )
    for column in work.columns:
        if column not in {"contractid", "companyid"} and column not in {"تاريخ إصدار المستخلص"}:
            converted = pd.to_numeric(work[column], errors="coerce")
            if converted.notna().any():
                work[column] = converted
    return work


def invoice_latest_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Return every row tied for the latest invoice date per contract/factory."""
    work = prepare_invoice_dataframe(df)
    if work.empty or "تاريخ إصدار المستخلص" not in work.columns:
        return work
    group_columns = [column for column in ("contractid", "مصنع") if column in work.columns]
    if not group_columns:
        return work.loc[work["تاريخ إصدار المستخلص"].eq(work["تاريخ إصدار المستخلص"].max())].copy()
    max_dates = work.groupby(group_columns, dropna=False)["تاريخ إصدار المستخلص"].transform("max")
    return work.loc[work["تاريخ إصدار المستخلص"].eq(max_dates)].copy()


def invoice_numeric_summary(df: pd.DataFrame, group_by_factory: bool = False) -> pd.DataFrame:
    """Sum numeric invoice columns, including the requested work-volume metric."""
    work = prepare_invoice_dataframe(df)
    if work.empty:
        return pd.DataFrame()
    value_column = _invoice_value_column(work)
    if value_column:
        work["حجم الأعمال"] = work[value_column].fillna(0)
        fallback = work.get("قيمة المستخلص قبل الخصومات")
        if fallback is not None and value_column != "قيمة المستخلص قبل الخصومات":
            work["حجم الأعمال"] = work[value_column].where(work[value_column].notna(), fallback).fillna(0)
    numeric_columns = work.select_dtypes(include="number").columns.tolist()
    numeric_columns = [column for column in numeric_columns if not column.lower().endswith("id")]
    if "حجم الأعمال" in work.columns and "حجم الأعمال" not in numeric_columns:
        numeric_columns.append("حجم الأعمال")
    if not numeric_columns:
        return pd.DataFrame()
    if group_by_factory and "مصنع" in work.columns:
        return work.groupby("مصنع", dropna=False)[numeric_columns].sum().reset_index()
    return pd.DataFrame([work[numeric_columns].sum(numeric_only=True)])
