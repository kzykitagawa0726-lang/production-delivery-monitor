import datetime

from openpyxl import Workbook

from src.process_history import (
    IDX_ACTUAL_COMPLETE,
    IDX_DRAWING,
    IDX_MACHINE,
    IDX_PROC_CODE,
    IDX_START,
    ProcessHistory,
)

# 工程累積データの98列ヘッダーを1件だけ実データと同じ位置関係で再現する
# (テストで使う列だけ意味のある値を入れ、それ以外はダミー値で埋める)
NUM_COLS = 98


def _row(order_no="AB123", drawing="DWG-A1", process="HB", machine="KW451", start=20230101, complete=20230110):
    values = [0] * NUM_COLS
    values[0] = order_no
    values[2] = 1
    values[IDX_DRAWING] = drawing
    values[IDX_PROC_CODE] = process
    values[7] = 1  # 内外作区分
    values[IDX_MACHINE] = machine
    values[67] = 1  # 作業完了
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


def test_load_aggregates_by_drawing_and_process(tmp_path):
    path = tmp_path / "process2023.xlsx"
    make_workbook(
        [
            _row(machine="KW451", drawing="DWG-A1", process="HB", start=20230101, complete=20230105),
            _row(machine="KW451", drawing="DWG-A1", process="HB", start=20230201, complete=20230210),
            _row(machine="KB400", drawing="DWG-A1", process="MC", start=20230101, complete=20230103),
        ],
        path,
    )

    history = ProcessHistory.load([path], cache_path=tmp_path / "cache.json")

    # 図番一致は工程コードを問わず集計する(KW451:HB×2件, KB400:MC×1件)。件数順。
    suggestions = history.suggest_machines("DWG-A1", "HB")
    assert len(suggestions) == 2
    assert suggestions[0].machine_code == "KW451"
    assert suggestions[0].match_type == "drawing"
    assert suggestions[0].count == 2
    assert suggestions[0].last_used == datetime.date(2023, 2, 10)
    assert suggestions[1].machine_code == "KB400"
    assert suggestions[1].count == 1


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
            _row(process="HB", start=20230101, complete=20230111),  # 10日
            _row(process="HB", start=20230101, complete=20230121),  # 20日
            _row(process="HB", start=20230101, complete=20230131),  # 30日
        ],
        path,
    )

    history = ProcessHistory.load([path], cache_path=tmp_path / "cache.json")
    assert history.actual_lt_calendar_days_median("HB") == 20


def test_actual_lt_is_none_when_no_data():
    history = ProcessHistory()
    assert history.actual_lt_calendar_days_median("HB") is None


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
