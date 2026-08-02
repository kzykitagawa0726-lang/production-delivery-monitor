import datetime

from src.judge import judge_record
from src.models import OrderRecord
from src.process_master import ProcessMaster
from src.report_data import build_report_data
from src.report_excel import write_excel_report
from src.report_html import write_html_report

BASE_DATE = datetime.date(2026, 8, 3)


def synthetic_orders() -> list[OrderRecord]:
    return [
        OrderRecord(
            order_no="ORD-1001", drawing_no="DWG-A1", qty=20,
            order_date=datetime.date(2026, 5, 1), company_deadline=datetime.date(2026, 7, 20),
            total_process_count=6, process_seq=3, process_code="SAMPLE_H01",
            all_process_complete_flag=0, assignee="山田", product_name="平歯車A",
        ),
        OrderRecord(
            order_no="ORD-1002", drawing_no="DWG-B2", qty=5,
            order_date=datetime.date(2026, 6, 10), company_deadline=datetime.date(2026, 8, 6),
            total_process_count=8, process_seq=2, process_code="SAMPLE_G02",
            all_process_complete_flag=0, assignee="鈴木", product_name="はすば歯車B",
        ),
        OrderRecord(
            order_no="ORD-1003", drawing_no="DWG-C3", qty=100,
            order_date=datetime.date(2026, 7, 1), company_deadline=datetime.date(2026, 12, 1),
            total_process_count=4, process_seq=1, process_code="SAMPLE_P01",
            all_process_complete_flag=0, assignee="佐藤", product_name="ウォームギヤC",
        ),
        OrderRecord(
            order_no="ORD-1004", drawing_no="DWG-D4", qty=1,
            order_date=datetime.date(2026, 6, 20), company_deadline=datetime.date(2026, 1, 1),  # 納期未更新想定
            total_process_count=5, process_seq=1, process_code="SAMPLE_I01",
            all_process_complete_flag=0, assignee="田中", product_name="スパイラルギヤD",
        ),
        OrderRecord(
            order_no="ORD-1005", drawing_no="DWG-E5", qty=3,
            order_date=datetime.date(2026, 7, 15), company_deadline=datetime.date(2026, 10, 1),
            total_process_count=0, process_seq=0, process_code="UNKNOWN_XX",  # 工程未展開+未知コード
            all_process_complete_flag=0, assignee="田中", product_name="スパイラルギヤE",
        ),
        OrderRecord(
            order_no="ORD-1006", drawing_no="DWG-F6", qty=50,
            order_date=datetime.date(2026, 7, 20), company_deadline=datetime.date(2026, 8, 5),
            total_process_count=3, process_seq=3, process_code="SAMPLE_G02",
            all_process_complete_flag=1, assignee="山田", product_name="平歯車F",  # 完了なので対象外
        ),
    ]


def test_full_pipeline_generates_valid_excel_and_html(tmp_path):
    orders = synthetic_orders()
    process_master = ProcessMaster()

    for order in orders:
        result = process_master.categorize(order.process_code)
        order.process_category = result.category
        order.current_process_standard_lt = result.standard_lt_business_days
        order.is_unknown_process_code = result.is_unknown_code
        order.is_bottleneck_process = result.is_bottleneck

    in_progress = [o for o in orders if o.all_process_complete_flag == 0]
    assert len(in_progress) == 5  # ORD-1006(完了)は除外

    for order in in_progress:
        judge_record(order, BASE_DATE)

    unknown_codes = process_master.get_unknown_codes()
    assert unknown_codes == ["UNKNOWN_XX"]

    data = build_report_data(in_progress, BASE_DATE, unknown_codes)

    # ORD-1001: 自社納期2026-07-20は基準日2026-08-03より前 -> 納期遅延
    # ORD-1002: 納期まで3営業日, 残必要日数 = (8-2)*5+5=35 -> リスクあり
    # ORD-1003: 納期まで十分猶予あり -> 正常(いずれにも含まれない)
    # ORD-1004: 納期未更新プレースホルダ(20260101) -> 判定不能
    # ORD-1005: 全工程数0 -> 判定不能
    assert {r.order_no for r in data.delayed} == {"ORD-1001"}
    assert {r.order_no for r in data.at_risk} == {"ORD-1002"}
    assert {r.order_no for r in data.undetermined} == {"ORD-1004", "ORD-1005"}

    excel_path = tmp_path / "report.xlsx"
    html_path = tmp_path / "report.html"
    write_excel_report(in_progress, excel_path, BASE_DATE, unknown_codes)
    write_html_report(data, html_path)

    assert excel_path.exists()
    assert html_path.exists()
    html_content = html_path.read_text(encoding="utf-8")
    assert "http://" not in html_content and "https://" not in html_content
