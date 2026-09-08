import datetime

from openpyxl import Workbook

from src.excel_io import load_orders

HEADERS = [
    "図番/品名", "製造オーダー№", "客先注番", "数量", "受注日\n自社納期\n顧客納期", "残日",
    "得意先NO", "停滞日", "製番", "MEMO1", "MEMO2", "MEMO3",
    "材料状況",
    "工程2", "工程3",
]


def make_workbook(rows: list[tuple], path, headers=HEADERS):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(path)


def test_load_orders_parses_basic_columns_and_deadline_triplet(tmp_path):
    path = tmp_path / "sample.xlsx"
    make_workbook(
        [
            (
                "312127272601\nSPB M10x13T L", "PO-017070", "191404", 1,
                "26年05月22日\n26年09月30日\n26年10月01日", 21, "G1013", 24, "JR7601",
                None, 0, "9月3週A", "材料入荷済み", None, None,
            ),
        ],
        path,
    )

    orders = load_orders(path)
    assert len(orders) == 1
    o = orders[0]
    assert o.order_no == "PO-017070"
    assert o.drawing_no == "312127272601"
    assert o.product_name == "SPB M10x13T L"
    assert o.qty == 1
    assert o.order_date == datetime.date(2026, 5, 22)
    assert o.company_deadline == datetime.date(2026, 9, 30)
    assert o.customer_deadline == datetime.date(2026, 10, 1)
    assert o.remaining_business_days == 21
    assert o.customer_no == "G1013"
    assert o.stagnation_days == 24
    assert o.stagnation_note is None
    assert o.old_system_no == "JR7601"
    assert o.production_note is None  # MEMO2=0はNone扱い
    assert o.shipment_plan_bucket == "9月3週A"
    assert o.material_arrived is True
    assert o.is_forecast_order is False


def test_load_orders_treats_00_placeholder_as_forecast_order(tmp_path):
    path = tmp_path / "sample.xlsx"
    make_workbook(
        [
            (
                "A1\n品名1", "PO-1", "", 10,
                "00年01月00日\n26年09月29日\n00年01月00日", 17, "G1018", "前工程実績無し", "JV7301",
                None, 0, "10月2週B", "材料入荷済み", None, None,
            ),
        ],
        path,
    )

    orders = load_orders(path)
    o = orders[0]
    assert o.order_date is None
    assert o.company_deadline == datetime.date(2026, 9, 29)  # 自社納期は常に実データ
    assert o.customer_deadline is None
    assert o.is_forecast_order is True
    assert o.stagnation_days is None
    assert o.stagnation_note == "前工程実績無し"


def test_load_orders_skips_rows_without_order_no(tmp_path):
    path = tmp_path / "sample.xlsx"
    make_workbook(
        [
            (
                "A1\n品名1", "PO-1", "", 1, "26年08月24日\n26年09月01日\n26年09月05日", 5,
                "G1", 1, "H1", None, 0, "9月1週A", "材料入荷済み", None, None,
            ),
            (None,) * len(HEADERS),
        ],
        path,
    )

    orders = load_orders(path)
    assert len(orders) == 1


def test_load_orders_handles_missing_deadline_value(tmp_path):
    path = tmp_path / "sample.xlsx"
    make_workbook(
        [
            ("A1\n品名1", "PO-1", "", 1, None, None, "G1", None, "H1", None, 0, None, None, None, None),
        ],
        path,
    )

    orders = load_orders(path)
    assert orders[0].company_deadline is None
    assert orders[0].remaining_business_days is None
    assert orders[0].order_date is None


def test_load_orders_parses_memo1_as_date_when_8digit_number(tmp_path):
    path = tmp_path / "sample.xlsx"
    make_workbook(
        [
            (
                "A1\n品名1", "PO-1", "", 1, "26年08月24日\n26年09月01日\n26年09月05日", 5,
                "G1", 1, "H1", 20260925, 0, "9月1週A", "材料入荷済み", None, None,
            ),
        ],
        path,
    )

    orders = load_orders(path)
    o = orders[0]
    assert o.sales_agreed_deadline == datetime.date(2026, 9, 25)
    assert o.sales_note is None


def test_load_orders_parses_memo1_as_free_text_note(tmp_path):
    path = tmp_path / "sample.xlsx"
    make_workbook(
        [
            (
                "A1\n品名1", "PO-1", "", 1, "26年08月24日\n26年09月01日\n26年09月05日", 5,
                "G1", 1, "H1", "納期確認お願い致します。", 0, "9月1週A", "材料入荷済み", None, None,
            ),
        ],
        path,
    )

    orders = load_orders(path)
    o = orders[0]
    assert o.sales_agreed_deadline is None
    assert o.sales_note == "納期確認お願い致します。"


def test_load_orders_parses_process_slots(tmp_path):
    path = tmp_path / "sample.xlsx"
    make_workbook(
        [
            (
                "A1\n品名1", "PO-1", "", 1, "26年08月24日\n26年09月01日\n26年09月05日", 5,
                "G1", 1, "H1", None, 0, "9月1週A",
                "材料入荷済み",
                "LD\nF1028\n↓実績日↓\n07/29\n作業完了",
                "IL\nTVG35C\n08/11\n08/18\nオーダー確定前",
            ),
        ],
        path,
    )

    orders = load_orders(path)
    o = orders[0]
    assert o.material_arrived is True
    assert len(o.processes) == 2
    first, second = o.processes
    assert first.process_code == "LD"
    assert first.destination == "F1028"
    assert first.status == "作業完了"
    assert first.is_completed is True
    assert second.process_code == "IL"
    assert second.status == "オーダー確定前"
    assert second.is_completed is False
    assert o.current_process is second
    assert o.remaining_process_count == 1
