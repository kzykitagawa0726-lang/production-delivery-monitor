import datetime

from src.judge import judge_record
from src.models import OrderRecord, ProcessStep
from src.process_master import ProcessMaster
from src.report_data import build_report_data
from src.report_excel import write_excel_report
from src.report_html import write_html_report

BASE_DATE = datetime.date(2026, 8, 3)


def synthetic_orders() -> list[OrderRecord]:
    return [
        OrderRecord(
            order_no="PO-1001", drawing_no="DWG-A1", product_name="平歯車A", qty=20,
            order_date=datetime.date(2026, 6, 1),
            company_deadline=datetime.date(2026, 7, 20), customer_deadline=datetime.date(2026, 7, 18),
            remaining_business_days=-10,
            customer_order_no="C-001", customer_no="G6000", old_system_no="BA9201",
            processes=[ProcessStep(2, "CR", "F1028", "07/01", "07/05", "オーダー確定前")],
        ),
        OrderRecord(
            order_no="PO-1002", drawing_no="DWG-B2", product_name="はすば歯車B", qty=5,
            order_date=datetime.date(2026, 7, 1),
            company_deadline=datetime.date(2026, 8, 6), customer_deadline=datetime.date(2026, 8, 10),
            remaining_business_days=3,
            customer_order_no="C-002", customer_no="G6010", old_system_no="AZ8701",
            processes=[ProcessStep(2, "CH", "F1028", "07/29", "08/01", "作業完了")],
        ),
        OrderRecord(
            order_no="PO-1003", drawing_no="DWG-C3", product_name="ウォームギヤC", qty=100,
            order_date=datetime.date(2026, 6, 1),
            company_deadline=datetime.date(2026, 12, 1), customer_deadline=datetime.date(2026, 12, 1),
            remaining_business_days=80,
            customer_order_no="C-003", customer_no="G6000", old_system_no="BC6301",
        ),
        OrderRecord(
            order_no="PO-1004", drawing_no="DWG-D4", product_name="スパイラルギヤD", qty=1,
            order_date=None, company_deadline=None, customer_deadline=None, remaining_business_days=None,
            customer_order_no="C-004", customer_no="G6000", old_system_no="BC8501",
        ),
        OrderRecord(
            order_no="PO-1005", drawing_no="DWG-E5", product_name="内示案件E", qty=50,
            order_date=None, company_deadline=datetime.date(2026, 9, 1), customer_deadline=None,
            remaining_business_days=15,
            customer_order_no=None, customer_no="G6000", old_system_no=None,
            is_forecast_order=True,
        ),
    ]


def test_full_pipeline_generates_valid_excel_and_html(tmp_path):
    orders = synthetic_orders()

    for order in orders:
        judge_record(order)

    process_master = ProcessMaster()
    data = build_report_data(orders, BASE_DATE, process_master)

    # PO-1001: 残日-10 -> 納期遅延
    # PO-1002: 残日3(<=5) -> リスクあり
    # PO-1003: 残日80 -> 正常(いずれにも含まれない)
    # PO-1004: 自社納期・残日とも欠損 -> 判定不能
    # PO-1005: 内示・先行手配だが自社納期・残日はあるため通常通り判定(正常)
    assert {r.order_no for r in data.delayed} == {"PO-1001"}
    assert {r.order_no for r in data.at_risk} == {"PO-1002"}
    assert {r.order_no for r in data.undetermined} == {"PO-1004"}
    assert data.forecast_order_count == 1

    excel_path = tmp_path / "report.xlsx"
    html_path = tmp_path / "report.html"
    write_excel_report(orders, excel_path, BASE_DATE, process_master)
    write_html_report(data, html_path)

    assert excel_path.exists()
    assert html_path.exists()
    html_content = html_path.read_text(encoding="utf-8")
    assert "http://" not in html_content and "https://" not in html_content
