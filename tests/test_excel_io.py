import datetime

from openpyxl import Workbook

from src.excel_io import load_orders

HEADERS = [
    "図番/品名", "製造オーダー№", "客先注番", "数量", "自社納期\n受注日", "残日",
    "取引先NO", "製番",
]


def make_workbook(rows: list[tuple], path, headers=HEADERS):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(path)


def test_load_orders_parses_basic_columns(tmp_path):
    path = tmp_path / "sample.xlsx"
    make_workbook(
        [
            (
                "312127272601\nSPB M10x13T L", "PO-017070", "", 1,
                "26年08月24日\n00年01月00日\nLT=33041", 21, "G6000", "BA9201（926201）",
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
    assert o.company_deadline == datetime.date(2026, 8, 24)
    assert o.remaining_business_days == 21
    assert o.customer_no == "G6000"
    assert o.old_system_no == "BA9201（926201）"
    assert o.customer_order_no is None  # 空文字はNone扱い


def test_load_orders_skips_rows_without_order_no(tmp_path):
    path = tmp_path / "sample.xlsx"
    make_workbook(
        [
            ("A1\n品名1", "PO-1", "", 1, "26年08月24日\n00年01月00日\nLT=1", 5, "G1", "H1"),
            (None, None, None, None, None, None, None, None),
        ],
        path,
    )

    orders = load_orders(path)
    assert len(orders) == 1


def test_load_orders_handles_unparseable_deadline(tmp_path):
    path = tmp_path / "sample.xlsx"
    make_workbook(
        [
            ("A1\n品名1", "PO-1", "", 1, None, None, "G1", "H1"),
        ],
        path,
    )

    orders = load_orders(path)
    assert orders[0].company_deadline is None
    assert orders[0].remaining_business_days is None


def test_load_orders_drawing_no_without_product_name(tmp_path):
    path = tmp_path / "sample.xlsx"
    make_workbook(
        [
            ("DWG-ONLY", "PO-1", "", 1, "26年08月24日\n00年01月00日\nLT=1", 5, "G1", "H1"),
        ],
        path,
    )

    orders = load_orders(path)
    assert orders[0].drawing_no == "DWG-ONLY"
    assert orders[0].product_name is None
