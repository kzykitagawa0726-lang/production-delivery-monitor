import datetime

from openpyxl import Workbook

from src.process_history import (
    IDX_ACTUAL_COMPLETE,
    IDX_ACTUAL_MANHOURS,
    IDX_DEPARTMENT,
    IDX_DRAWING,
    IDX_LINE,
    IDX_MACHINE,
    IDX_PROC_CODE,
    IDX_PROC_SEQ,
    IDX_PRODUCT_NAME,
    IDX_START,
    ProcessHistory,
)

# 工程累積データの98列ヘッダーを1件だけ実データと同じ位置関係で再現する
# (テストで使う列だけ意味のある値を入れ、それ以外はダミー値で埋める)
NUM_COLS = 98


def _row(
    order_no="AB123", line="1", drawing="DWG-A1", process="HB", machine="KW451",
    start=20230101, complete=20230110, seq=1, department="1223", manhours=0, product_name="",
):
    values = [0] * NUM_COLS
    values[0] = order_no
    values[IDX_LINE] = line
    values[IDX_PROC_SEQ] = seq
    values[IDX_DRAWING] = drawing
    values[IDX_PROC_CODE] = process
    values[7] = 1  # 内外作区分
    values[IDX_DEPARTMENT] = department
    values[IDX_MACHINE] = machine
    values[67] = 1  # 作業完了
    values[IDX_ACTUAL_MANHOURS] = manhours
    values[IDX_PRODUCT_NAME] = product_name
    values[IDX_START] = start
    values[IDX_ACTUAL_COMPLETE] = complete
    return tuple(values)


def make_workbook(rows: list[tuple], path):
    wb = Workbook()
    ws = wb.active
    ws.append([f"col{i}" for i in range(NUM_COLS)])  # ヘッダーは列名参照しないのでダミーでよい
    for row in rows:
        ws.append(row)
    wb.save(path)


def test_suggest_requires_drawing_and_process_code_to_both_match(tmp_path):
    # 図番だけの一致では、別の工程を担当しただけの設備が紛れ込んでしまうため、
    # 図番+工程コードの組み合わせで一致した実績のみを返す(kii-san指摘: 2026-09-08)。
    path = tmp_path / "process2023.xlsx"
    make_workbook(
        [
            _row(line="1", machine="KW451", drawing="DWG-A1", process="HB", start=20230101, complete=20230105),
            _row(line="2", machine="KW451", drawing="DWG-A1", process="HB", start=20230201, complete=20230210),
            _row(line="3", machine="KB400", drawing="DWG-A1", process="MC", start=20230101, complete=20230103),  # 同じ図番だが別工程
        ],
        path,
    )

    history = ProcessHistory.load([path], cache_path=tmp_path / "cache.json")

    suggestions = history.suggest_machines("DWG-A1", "HB")
    assert len(suggestions) == 1
    assert suggestions[0].machine_code == "KW451"
    assert suggestions[0].match_type == "drawing_process"
    assert suggestions[0].count == 2
    assert suggestions[0].last_used == datetime.date(2023, 2, 10)

    mc_suggestions = history.suggest_machines("DWG-A1", "MC")
    assert len(mc_suggestions) == 1
    assert mc_suggestions[0].machine_code == "KB400"


def test_suggest_falls_back_to_process_code_when_drawing_unknown():
    history = ProcessHistory()
    history._by_process["HB"]["KW451"] = {"count": 4, "last_used": datetime.date(2024, 5, 1)}

    suggestions = history.suggest_machines("UNKNOWN-DRAWING", "HB")
    assert len(suggestions) == 1
    assert suggestions[0].match_type == "process_code"
    assert suggestions[0].machine_code == "KW451"


def test_actual_lt_median_is_reference_only(tmp_path):
    path = tmp_path / "process2023.xlsx"
    make_workbook(
        [
            _row(line="1", process="HB", start=20230101, complete=20230111),  # 10日
            _row(line="2", process="HB", start=20230101, complete=20230121),  # 20日
            _row(line="3", process="HB", start=20230101, complete=20230131),  # 30日
        ],
        path,
    )

    history = ProcessHistory.load([path], cache_path=tmp_path / "cache.json")
    assert history.actual_lt_calendar_days_median("HB") == 20


def test_actual_lt_prefers_drawing_specific_data(tmp_path):
    path = tmp_path / "process2023.xlsx"
    make_workbook(
        [
            # DWG-A1自体の実績: 5日固定
            _row(line="1", drawing="DWG-A1", process="HB", start=20230101, complete=20230106),
            _row(line="2", drawing="DWG-A1", process="HB", start=20230201, complete=20230206),
            # 他図番の実績: 50日(全体平均を大きく引き上げる)
            _row(line="3", drawing="DWG-B2", process="HB", start=20230101, complete=20230220),
        ],
        path,
    )

    history = ProcessHistory.load([path], cache_path=tmp_path / "cache.json")
    # DWG-A1を指定すればDWG-A1自身の実績(5日)が優先される。工程コード全体の中央値ではない。
    assert history.actual_lt_calendar_days_median("HB", drawing_no="DWG-A1") == 5
    # 図番未指定なら工程コード全体(5,5,50日)の中央値
    assert history.actual_lt_calendar_days_median("HB") == 5


def test_actual_lt_is_none_when_no_data():
    history = ProcessHistory()
    assert history.actual_lt_calendar_days_median("HB") is None


def test_estimate_route_orders_steps_by_typical_seq_and_prefers_order_level_total(tmp_path):
    # 同じ受注№+行の2工程(1インスタンス)。合計は「工程別実績の単純合計(2+5=7)」ではなく、
    # 「そのインスタンス自体の最初の着手〜最後の完成(5日)」を優先する
    # (kii-san指摘: 2026-09-08。工程別実績には他案件との待ち時間が含まれ、単純合計すると
    # 残り工程が多い受注ほど過大評価になる異常値が実データで見つかったため)。
    path = tmp_path / "process2023.xlsx"
    make_workbook(
        [
            _row(order_no="AB123", line="1", drawing="DWG-A1", process="HG", seq=2, start=20230103, complete=20230106),
            _row(order_no="AB123", line="1", drawing="DWG-A1", process="L0", seq=1, start=20230101, complete=20230103),
        ],
        path,
    )

    history = ProcessHistory.load([path], cache_path=tmp_path / "cache.json")
    estimate = history.estimate_route_for_drawing("DWG-A1")

    assert estimate is not None
    assert [s.process_code for s in estimate.steps] == ["L0", "HG"]  # 内訳は工程順どおり(参考情報)
    assert estimate.total_calendar_days == 5  # 2023-01-01 〜 2023-01-06 の通し実績
    assert estimate.total_is_order_level is True
    assert estimate.total_sample_count == 1


def test_estimate_route_falls_back_to_summed_steps_when_no_order_level_data(tmp_path):
    # 受注№または行が空欄で、インスタンス単位の実績が組み立てられない場合のみ、
    # 工程別実績LTの単純合計にフォールバックする(通常の実データでは起きにくい経路)。
    path = tmp_path / "process2023.xlsx"
    make_workbook(
        [
            _row(order_no="", line="", drawing="DWG-A1", process="L0", seq=1, start=20230101, complete=20230103),  # 2日
            _row(order_no="", line="", drawing="DWG-A1", process="HG", seq=2, start=20230101, complete=20230106),  # 5日
        ],
        path,
    )

    history = ProcessHistory.load([path], cache_path=tmp_path / "cache.json")
    estimate = history.estimate_route_for_drawing("DWG-A1")
    assert estimate is not None
    assert estimate.total_is_order_level is False
    assert estimate.total_calendar_days == 7  # 2 + 5 (フォールバックの単純合計)


def test_estimate_route_for_drawing_flags_missing_actual_data(tmp_path):
    path = tmp_path / "process2023.xlsx"
    make_workbook(
        [
            _row(order_no="AB123", line="1", drawing="DWG-A1", process="L0", seq=1, start=20230101, complete=20230103),
            _row(order_no="AB123", line="1", drawing="DWG-A1", process="HG", seq=2, start=0, complete=0),  # 着手日・完成日なし
        ],
        path,
    )

    history = ProcessHistory.load([path], cache_path=tmp_path / "cache.json")
    estimate = history.estimate_route_for_drawing("DWG-A1")

    assert estimate is not None
    assert estimate.missing_process_codes == ["HG"]  # 工程別内訳としては参考表示
    assert estimate.total_is_order_level is True
    assert estimate.total_calendar_days == 2  # L0の着手日(01-01)〜完成日(01-03)の通し実績
    assert estimate.has_full_data is False


def test_estimate_route_for_drawing_returns_none_for_unknown_drawing():
    history = ProcessHistory()
    assert history.estimate_route_for_drawing("UNKNOWN") is None


def test_order_level_lt_uses_median_across_multiple_instances(tmp_path):
    path = tmp_path / "process2023.xlsx"
    make_workbook(
        [
            _row(order_no="AB123", line="1", drawing="DWG-A1", process="L0", start=20230101, complete=20230111),  # 10日
            _row(order_no="CD456", line="1", drawing="DWG-A1", process="L0", start=20230101, complete=20230121),  # 20日
            _row(order_no="EF789", line="1", drawing="DWG-A1", process="L0", start=20230101, complete=20230131),  # 30日
        ],
        path,
    )

    history = ProcessHistory.load([path], cache_path=tmp_path / "cache.json")
    assert history.order_level_lt_calendar_days_median("DWG-A1") == 20
    assert history.order_level_lt_sample_count("DWG-A1") == 3


def test_order_level_lt_is_none_when_no_data():
    history = ProcessHistory()
    assert history.order_level_lt_calendar_days_median("DWG-A1") is None
    assert history.order_level_lt_sample_count("DWG-A1") == 0


def test_weekly_load_by_department_and_machine(tmp_path):
    path = tmp_path / "process2023.xlsx"
    make_workbook(
        [
            # 2023-01-02(月)〜01-08(日)の週。完成日01-05(木)。
            _row(line="1", machine="KW451", department="1223", complete=20230105, manhours=2.0),
            _row(line="2", machine="KW451", department="1223", complete=20230106, manhours=1.5),
            _row(line="3", machine="KB400", department="9211", complete=20230105, manhours=3.0),
        ],
        path,
    )

    history = ProcessHistory.load([path], cache_path=tmp_path / "cache.json")
    week = datetime.date(2023, 1, 2)  # その週の月曜日

    by_dept = {(w, d): h for w, d, h in history.weekly_load_by_department()}
    assert by_dept[(week, "1223")] == 3.5
    assert by_dept[(week, "9211")] == 3.0

    by_machine = {(w, m): h for w, m, h in history.weekly_load_by_machine()}
    assert by_machine[(week, "KW451")] == 3.5
    assert by_machine[(week, "KB400")] == 3.0


def test_monthly_load_by_product_token(tmp_path):
    path = tmp_path / "process2023.xlsx"
    make_workbook(
        [
            _row(complete=20230105, manhours=2.0, product_name="ＧＥＡＲ　Ｍ６ｘ６０　Ｓ"),
            _row(complete=20230120, manhours=3.0, product_name="ＧＥＡＲ　Ｍ３Ｘ２４Ｔ"),  # 同月・同トークン
            _row(complete=20230201, manhours=1.5, product_name="ＳＰＢ　Ｍ３Ｘ２４Ｔ　Ｌ"),  # 翌月・別トークン
        ],
        path,
    )

    history = ProcessHistory.load([path], cache_path=tmp_path / "cache.json")
    points = {(m, t): h for m, t, h in history.monthly_load_by_product_token()}

    assert points[("2023-01", "ＧＥＡＲ")] == 5.0  # 同月・同トークンは合算される
    assert points[("2023-02", "ＳＰＢ")] == 1.5


def test_cache_is_reused_when_source_files_unchanged(tmp_path):
    path = tmp_path / "process2023.xlsx"
    make_workbook([_row()], path)
    cache_path = tmp_path / "cache.json"

    history1 = ProcessHistory.load([path], cache_path=cache_path)
    assert cache_path.exists()

    history2 = ProcessHistory.load([path], cache_path=cache_path)
    assert (
        history2.suggest_machines("DWG-A1", "HB")[0].count
        == history1.suggest_machines("DWG-A1", "HB")[0].count
    )
