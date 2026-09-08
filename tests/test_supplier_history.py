import datetime

from openpyxl import Workbook

from src.supplier_history import HEADER_19, HEADER_55, SupplierHistory


def make_workbook_19(rows: list[tuple], path):
    wb = Workbook()
    ws = wb.active
    ws.append(HEADER_19)
    for row in rows:
        ws.append(row)
    wb.save(path)


def make_workbook_55(rows: list[dict], path):
    wb = Workbook()
    ws = wb.active
    ws.append(HEADER_55)
    for values in rows:
        row = [values.get(h) for h in HEADER_55]
        ws.append(row)
    wb.save(path)


def _row19(supplier="2103", drawing="DWG-A1", process="HB", date=20230115):
    # HEADER_19: 得意先ＣＤ,仕入先ＣＤ,担当CD,商品名,数量,仕入単価,金額（支払額）,受注№,行№,情報ＮＯ,
    #            注文№,図番,材質,寸法,伝票日付,工程コード,明細１,明細２,明細４
    return (1000, supplier, 10, "GEAR", 5, 100, 500, "AB123", 1, 1, 0, drawing, "SCM415", "50X50", date, process, "", "", "")


def test_load_aggregates_by_drawing_and_process(tmp_path):
    path = tmp_path / "supplier2023.xlsx"
    make_workbook_19(
        [
            _row19(supplier="2103", drawing="DWG-A1", process="HB", date=20230110),
            _row19(supplier="2103", drawing="DWG-A1", process="HB", date=20230301),
            _row19(supplier="4016", drawing="DWG-A1", process="MC", date=20230115),
            _row19(supplier="2103", drawing="DWG-B2", process="HB", date=20230201),
        ],
        path,
    )

    history = SupplierHistory.load([path], cache_path=tmp_path / "cache.json")

    suggestions = history.suggest("DWG-A1", "HB")
    assert len(suggestions) == 2  # 2103(2件), 4016(1件)
    assert suggestions[0].supplier_code == "2103"
    assert suggestions[0].match_type == "drawing"
    assert suggestions[0].count == 2
    assert suggestions[0].last_used == datetime.date(2023, 3, 1)


def test_suggest_falls_back_to_process_code_when_drawing_unknown():
    history = SupplierHistory()
    history._by_process["HB"]["9999"] = {"count": 3, "last_used": datetime.date(2024, 5, 1)}

    suggestions = history.suggest("UNKNOWN-DRAWING", "HB")
    assert len(suggestions) == 1
    assert suggestions[0].match_type == "process_code"
    assert suggestions[0].supplier_code == "9999"


def test_blank_process_code_rows_are_ignored(tmp_path):
    path = tmp_path / "supplier2023.xlsx"
    make_workbook_19(
        [
            _row19(supplier="2103", drawing="DWG-A1", process=""),  # 材料そのものの仕入。無視対象
            _row19(supplier="2103", drawing="DWG-A1", process="HB"),
        ],
        path,
    )

    history = SupplierHistory.load([path], cache_path=tmp_path / "cache.json")
    suggestions = history.suggest("DWG-A1", "HB")
    assert len(suggestions) == 1
    assert suggestions[0].count == 1  # 空欄行はカウントされない


def test_numeric_process_code_is_normalized(tmp_path):
    path = tmp_path / "supplier2022.xlsx"
    make_workbook_19([_row19(supplier="2103", drawing="DWG-A1", process=6)], path)  # Excel上は数値の"6"

    history = SupplierHistory.load([path], cache_path=tmp_path / "cache.json")
    suggestions = history.suggest(None, "06")  # マスタ側はゼロ埋め表記
    assert len(suggestions) == 1
    assert suggestions[0].count == 1


def test_cache_is_reused_when_source_files_unchanged(tmp_path):
    path = tmp_path / "supplier2023.xlsx"
    make_workbook_19([_row19()], path)
    cache_path = tmp_path / "cache.json"

    history1 = SupplierHistory.load([path], cache_path=cache_path)
    assert cache_path.exists()

    # ファイルを一切変更せず再ロード -> キャッシュから復元されても同じ結果になること
    history2 = SupplierHistory.load([path], cache_path=cache_path)
    assert history2.suggest("DWG-A1", "HB")[0].count == history1.suggest("DWG-A1", "HB")[0].count
