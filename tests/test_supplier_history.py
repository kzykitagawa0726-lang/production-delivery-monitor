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


def _row19(supplier="2103", drawing="DWG-A1", process="HB", date=20230115, amount=500):
    # HEADER_19: 得意先ＣＤ,仕入先ＣＤ,担当CD,商品名,数量,仕入単価,金額（支払額）,受注№,行№,情報ＮＯ,
    #            注文№,図番,材質,寸法,伝票日付,工程コード,明細１,明細２,明細４
    return (1000, supplier, 10, "GEAR", 5, 100, amount, "AB123", 1, 1, 0, drawing, "SCM415", "50X50", date, process, "", "", "")


def test_suggest_requires_drawing_and_process_code_to_both_match(tmp_path):
    # 図番だけの一致では、別の工程を担当しただけの仕入先が紛れ込んでしまうため、
    # 図番+工程コードの組み合わせで一致した実績のみを返す(kii-san指摘: 2026-09-08)。
    path = tmp_path / "supplier2023.xlsx"
    make_workbook_19(
        [
            _row19(supplier="2103", drawing="DWG-A1", process="HB", date=20230110),
            _row19(supplier="2103", drawing="DWG-A1", process="HB", date=20230301),
            _row19(supplier="4016", drawing="DWG-A1", process="MC", date=20230115),  # 同じ図番だが別工程
            _row19(supplier="2103", drawing="DWG-B2", process="HB", date=20230201),  # 別図番
        ],
        path,
    )

    history = SupplierHistory.load([path], cache_path=tmp_path / "cache.json")

    suggestions = history.suggest("DWG-A1", "HB")
    assert len(suggestions) == 1
    assert suggestions[0].supplier_code == "2103"
    assert suggestions[0].match_type == "drawing_process"
    assert suggestions[0].count == 2
    assert suggestions[0].last_used == datetime.date(2023, 3, 1)

    mc_suggestions = history.suggest("DWG-A1", "MC")
    assert len(mc_suggestions) == 1
    assert mc_suggestions[0].supplier_code == "4016"


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


def test_average_amount_and_supplier_amount_shares_for_process(tmp_path):
    # 仕入先週別予測(相対比較)専用の集計。代替候補・ランキングには一切影響しない。
    path = tmp_path / "supplier2023.xlsx"
    make_workbook_19(
        [
            _row19(supplier="2103", process="HB", amount=600),
            _row19(supplier="2103", process="HB", amount=400),  # 2103合計1000
            _row19(supplier="4016", process="HB", amount=1000),  # 4016合計1000
        ],
        path,
    )

    history = SupplierHistory.load([path], cache_path=tmp_path / "cache.json")

    assert history.average_amount_for_process("HB") == 2000 / 3  # 全行の平均

    shares = history.supplier_amount_shares_for_process("HB")
    assert shares == {"2103": 0.5, "4016": 0.5}  # 金額比率(件数比率ではない)


def test_average_amount_is_none_when_no_data():
    history = SupplierHistory()
    assert history.average_amount_for_process("HB") is None
    assert history.supplier_amount_shares_for_process("HB") == {}


def test_weekly_amount_by_supplier(tmp_path):
    path = tmp_path / "supplier2023.xlsx"
    make_workbook_19(
        [
            # 2023-01-02(月)〜01-08(日)の週。伝票日付01-05(木)。
            _row19(supplier="2103", process="HB", date=20230105, amount=600),
            _row19(supplier="2103", process="HB", date=20230106, amount=400),
            _row19(supplier="4016", process="HB", date=20230105, amount=300),
            # 翌週(01-16、月)
            _row19(supplier="2103", process="HB", date=20230116, amount=100),
        ],
        path,
    )

    history = SupplierHistory.load([path], cache_path=tmp_path / "cache.json")
    points = {(w, s): amt for w, s, amt in history.weekly_amount_by_supplier()}

    assert points[(datetime.date(2023, 1, 2), "2103")] == 1000
    assert points[(datetime.date(2023, 1, 2), "4016")] == 300
    assert points[(datetime.date(2023, 1, 16), "2103")] == 100


def test_cache_round_trip_preserves_amount_data(tmp_path):
    path = tmp_path / "supplier2023.xlsx"
    make_workbook_19([_row19(supplier="2103", process="HB", date=20230105, amount=777)], path)
    cache_path = tmp_path / "cache.json"

    history1 = SupplierHistory.load([path], cache_path=cache_path)
    history2 = SupplierHistory.load([path], cache_path=cache_path)  # キャッシュから復元

    assert history2.average_amount_for_process("HB") == history1.average_amount_for_process("HB") == 777
    assert history2.weekly_amount_by_supplier() == history1.weekly_amount_by_supplier()
    assert history2.weekly_amount_by_supplier() == [(datetime.date(2023, 1, 2), "2103", 777)]
