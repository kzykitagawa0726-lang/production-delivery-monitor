import datetime

from src.judge import judge_record
from src.models import OrderRecord, ProcessStep
from src.process_history import ProcessHistory
from src.process_master import ProcessMaster
from src.report_data import build_report_data
from src.report_html import write_html_report

BASE_DATE = datetime.date(2026, 8, 3)


def make_judged_record(**overrides) -> OrderRecord:
    defaults = dict(
        order_no="PO-000001",
        drawing_no="312127272601",
        product_name="SPB M10x13T L",
        qty=1,
        order_date=datetime.date(2026, 8, 1),
        company_deadline=datetime.date(2026, 9, 1),
        customer_deadline=datetime.date(2026, 9, 5),
        remaining_business_days=20,
        processes=[
            ProcessStep(slot_index=2, process_code="CR", destination="F1028",
                        date1="08/11", date2="08/18", status="オーダー確定前"),
        ],
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
    data = build_report_data(records, BASE_DATE, ProcessMaster())
    output_path = tmp_path / "report.html"

    write_html_report(data, output_path)

    assert output_path.exists()
    content = output_path.read_text(encoding="utf-8")
    assert "http://" not in content
    assert "https://" not in content
    assert "<script src=" not in content
    assert "ORD-DELAYED" in content
    assert "ORD-RISK" in content
    assert "<svg" in content  # 工程別仕掛中ランキングのグラフ


def test_write_html_report_shows_feasibility_and_weekly_load(tmp_path):
    record = make_judged_record(order_no="ORD-RISK", remaining_business_days=2)

    process_history = ProcessHistory()
    process_history._actual_durations["CR"] = [200]  # 客先納期に対して大幅に不足する想定
    process_history._weekly_load_by_department[(datetime.date(2026, 7, 27), "1223")] = 12.5

    data = build_report_data([record], BASE_DATE, ProcessMaster(), process_history=process_history)
    output_path = tmp_path / "report.html"

    write_html_report(data, output_path)

    content = output_path.read_text(encoding="utf-8")
    assert "実績ベースで納期に間に合わない見込み" in content
    assert "週別負荷" in content
    assert "部署1223" in content
