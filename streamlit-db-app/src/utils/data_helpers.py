import pandas as pd

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
