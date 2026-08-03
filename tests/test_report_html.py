import datetime

from src.judge import judge_record
from src.models import OrderRecord
from src.report_data import build_report_data
from src.report_html import write_html_report

BASE_DATE = datetime.date(2026, 8, 3)


def make_judged_record(**overrides) -> OrderRecord:
    defaults = dict(
        order_no="PO-000001",
        drawing_no="312127272601",
        product_name="SPB M10x13T L",
        qty=1,
        company_deadline=datetime.date(2026, 9, 1),
        remaining_business_days=20,
    )
    defaults.update(overrides)
    record = OrderRecord(**defaults)
    judge_record(record)
    return record


def test_write_html_report_is_self_contained_and_has_no_external_refs(tmp_path):
    records = [
        make_judged_record(order_no="ORD-DELAYED", remaining_business_days=-3),
        make_judged_record(order_no="ORD-RISK", remaining_business_days=2),
    ]
    data = build_report_data(records, BASE_DATE)
    output_path = tmp_path / "report.html"

    write_html_report(data, output_path)

    assert output_path.exists()
    content = output_path.read_text(encoding="utf-8")
    assert "http://" not in content
    assert "https://" not in content
    assert "<script src=" not in content
    assert "ORD-DELAYED" in content
    assert "ORD-RISK" in content
