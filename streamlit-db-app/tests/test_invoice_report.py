from pathlib import Path
import sys

import pandas as pd

PROJECT_SRC = Path(__file__).resolve().parents[1] / "src"
sys.path.insert(0, str(PROJECT_SRC))

from reports.invoices import (
    VIEW_LATEST,
    VIEW_PERIOD,
    _prepare_display_dataframe,
    _should_render_work_volume_card,
)


def test_work_volume_card_is_visible_only_for_period_view():
    assert _should_render_work_volume_card(VIEW_PERIOD) is True
    assert _should_render_work_volume_card(VIEW_LATEST) is False


def test_invoice_details_reorder_names_keep_all_columns_and_sort_by_date():
    source = pd.DataFrame(
        {
            "companyid_contract": 1,
            "factoryname": "التجمع",
            "companyname": "شركة مثال",
            "اسم المشروع": "عقد تجريبي",
            "تاريخ إصدار المستخلص": ["2026-10-02", "2026-10-01"],
            "رابط نسخة المستخلص": "https://example.test/invoice",
        },
        columns=[
            "companyid_contract",
            "factoryname",
            "companyname",
            "اسم المشروع",
            "تاريخ إصدار المستخلص",
            "رابط نسخة المستخلص",
        ],
    )

    display_df, link_columns = _prepare_display_dataframe(source)

    assert list(display_df.columns) == [
        "اسم المصنع",
        "اسم الشركة",
        "اسم العقد",
        "companyid_contract",
        "تاريخ إصدار المستخلص",
        "رابط نسخة المستخلص",
    ]
    assert display_df["اسم المصنع"].tolist() == ["التجمع", "التجمع"]
    assert display_df["اسم الشركة"].tolist() == ["شركة مثال", "شركة مثال"]
    assert display_df["اسم العقد"].tolist() == ["عقد تجريبي", "عقد تجريبي"]
    assert display_df["تاريخ إصدار المستخلص"].tolist() == [
        "2026-10-01",
        "2026-10-02",
    ]
    assert link_columns == []


if __name__ == "__main__":
    test_invoice_details_reorder_names_keep_all_columns_and_sort_by_date()
    print("PASS: invoice details contract")
