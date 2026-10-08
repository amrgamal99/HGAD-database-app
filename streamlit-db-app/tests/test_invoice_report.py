from pathlib import Path
import sys

import pandas as pd

PROJECT_SRC = Path(__file__).resolve().parents[1] / "src"
sys.path.insert(0, str(PROJECT_SRC))

from reports.invoices import _prepare_display_dataframe


def test_invoice_details_preserve_all_columns_in_original_order():
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

    assert list(display_df.columns) == list(source.columns)
    assert display_df["companyid_contract"].tolist() == [1, 1]
    assert display_df["تاريخ إصدار المستخلص"].tolist() == [
        "2026-10-02",
        "2026-10-01",
    ]
    assert link_columns == []


if __name__ == "__main__":
    test_invoice_details_preserve_all_columns_in_original_order()
    print("PASS: invoice details contract")
