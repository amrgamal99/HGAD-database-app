from pathlib import Path
import sys

import pandas as pd

PROJECT_SRC = Path(__file__).resolve().parents[1] / "src"
sys.path.insert(0, str(PROJECT_SRC))

from reports.checks import (
    CHECK_COLUMNS,
    _factory_check_summary,
    _prepare_display_dataframe,
)


def test_check_details_have_requested_columns_order_and_clean_dates():
    source = pd.DataFrame(
        {
            "factoryname": ["بدر"],
            "companyname": ["شركة مثال"],
            "اسم المشروع": ["عقد تجريبي"],
            "رقم الشيك": ["123"],
            "تاريخ الشيك": ["2026-10-03 00:00:00"],
            "البنك": ["بنك مثال"],
            "قيمة الشيك": [250],
            "رابط الشيك": ["https://example.test/check"],
            "الغرض من إصدار الشيك": ["دفعة مقدمة"],
            "تاريخ استلام شيك": ["2026-10-04"],
            "السنة_المالية_المحاسبية": [2026],
        }
    )

    display_df = _prepare_display_dataframe(source)

    assert display_df.columns.tolist() == list(CHECK_COLUMNS)
    assert display_df.loc[0, "اسم المصنع"] == "بدر"
    assert display_df.loc[0, "اسم العقد"] == "عقد تجريبي"
    assert display_df.loc[0, "رابط نسخة الشيك"] == "https://example.test/check"
    assert display_df.loc[0, "تاريخ الشيك"] == "2026-10-03"


def test_factory_summary_totals_and_normalizes_purpose_variants():
    source = pd.DataFrame(
        {
            "اسم المصنع": ["بدر", "بدر", "بدر", "التجمع"],
            "قيمة الشيك": [100, 200, 300, 50],
            "الغرض من إصدار الشيك": [
                "دفعه مقدمه",
                "رد تأمينات اجتماعية",
                "سداد مستخلص",
                "سداد مستخلص",
            ],
        }
    )

    summary = _factory_check_summary(source).set_index("اسم المصنع")

    assert summary.loc["بدر", "الإجمالي"] == 600
    assert summary.loc["بدر", "دفعة مقدمة"] == 100
    assert summary.loc["بدر", "رد تامينات اجتماعيه"] == 200
    assert summary.loc["بدر", "سداد مستخلص"] == 300
