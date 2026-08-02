import datetime

from openpyxl import load_workbook

from src.judge import judge_record
from src.models import OrderRecord
from src.report_excel import write_excel_report

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
        process_code="SAMPLE_G01",
        all_process_complete_flag=0,
        current_process_standard_lt=5,
        process_category="歯切り系",
    )
    defaults.update(overrides)
    record = OrderRecord(**defaults)
    judge_record(record, BASE_DATE)
    return record


def test_write_excel_report_creates_expected_sheets(tmp_path):
    records = [
        make_judged_record(order_no="ORD-DELAYED", company_deadline=datetime.date(2026, 7, 1)),
        make_judged_record(order_no="ORD-RISK", company_deadline=datetime.date(2026, 8, 6)),
        make_judged_record(order_no="ORD-UNDETERMINED", total_process_count=0),
    ]
    output_path = tmp_path / "report.xlsx"

    write_excel_report(records, output_path, generated_at=BASE_DATE, unknown_process_codes=[])

    assert output_path.exists()
    wb = load_workbook(output_path)
    assert wb.sheetnames == [
        "メインサマリー",
        "①納期遅延リスト",
        "②納期遅延リスクリスト",
        "判定不能・要確認",
        "工程別混雑ランキング",
    ]
    assert wb["①納期遅延リスト"]["A2"].value == "ORD-DELAYED"
    assert wb["②納期遅延リスクリスト"]["A2"].value == "ORD-RISK"
    assert wb["判定不能・要確認"]["A2"].value == "ORD-UNDETERMINED"
