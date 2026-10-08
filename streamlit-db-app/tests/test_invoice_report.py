from pathlib import Path
import sys

import pandas as pd

PROJECT_SRC = Path(__file__).resolve().parents[1] / "src"
sys.path.insert(0, str(PROJECT_SRC))

from reports.invoices import _prepare_display_dataframe


def test_invoice_details_use_requested_columns_and_ascending_date():
    source = pd.DataFrame(
        {
            "companyid_contract": 1,
            "factoryname": "التجمع",
            "companyname": "شركة مثال",
            "اسم المشروع": "عقد تجريبي",
            "تاريخ إصدار المستخلص": ["2026-10-02", "2026-10-01"],
            "رابط نسخة المستخلص": "https://example.test/invoice",
        }
    )

    display_df, link_columns = _prepare_display_dataframe(source)

    assert list(display_df.columns[:4]) == [
        "اسم المصنع",
        "اسم الشركة",
        "اسم العقد",
        "تاريخ إصدار المستخلص",
    ]
    assert "companyid_contract" not in display_df.columns
    assert display_df["تاريخ إصدار المستخلص"].tolist() == [
        "2026-10-01",
        "2026-10-02",
    ]
    assert link_columns == ["رابط نسخة المستخلص"]


if __name__ == "__main__":
    test_invoice_details_use_requested_columns_and_ascending_date()
    print("PASS: invoice details contract")
