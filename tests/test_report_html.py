import datetime

from src.judge import judge_record
from src.models import OrderRecord
from src.report_data import build_report_data
from src.report_html import write_html_report

BASE_DATE = datetime.date(2026, 8, 3)


def make_judged_record(**overrides) -> OrderRecord:
    defaults = dict(
        order_no="ORD-0001",
        drawing_no="DWG-0001",
        qty=10,
        order_date=datetime.date(2026, 6, 1),
        company_deadline=datetime.date(2026, 9, 1),
        total_process_count=5,
        process_seq=2,
        process_code="SAMPLE_G02",
        all_process_complete_flag=0,
        current_process_standard_lt=5,
        process_category="歯切り系",
        is_bottleneck_process=True,
    )
    defaults.update(overrides)
    record = OrderRecord(**defaults)
    judge_record(record, BASE_DATE)
    return record


def test_write_html_report_is_self_contained_and_has_no_external_refs(tmp_path):
    records = [
        make_judged_record(order_no="ORD-DELAYED", company_deadline=datetime.date(2026, 7, 1)),
        make_judged_record(order_no="ORD-RISK", company_deadline=datetime.date(2026, 8, 6)),
    ]
    data = build_report_data(records, BASE_DATE, unknown_process_codes=["WEIRD_CODE"])
    output_path = tmp_path / "report.html"

    write_html_report(data, output_path)

    assert output_path.exists()
    content = output_path.read_text(encoding="utf-8")
    assert "http://" not in content
    assert "https://" not in content
    assert "<script src=" not in content
    assert "ORD-DELAYED" in content
    assert "ORD-RISK" in content
    assert "WEIRD_CODE" in content
    assert "<svg" in content
