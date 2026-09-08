from openpyxl import Workbook

from lookup_lt import main
from src.process_history import (
    IDX_ACTUAL_COMPLETE,
    IDX_DRAWING,
    IDX_LINE,
    IDX_MACHINE,
    IDX_ORDER_NO,
    IDX_PROC_CODE,
    IDX_PROC_SEQ,
    IDX_START,
)

NUM_COLS = 98


def _row(drawing, process, seq, start, complete, order_no="AB123", line="1", machine="KW451"):
    values = [0] * NUM_COLS
    values[IDX_ORDER_NO] = order_no
    values[IDX_LINE] = line
    values[IDX_PROC_SEQ] = seq
    values[IDX_DRAWING] = drawing
    values[IDX_PROC_CODE] = process
    values[IDX_MACHINE] = machine
    values[IDX_START] = start
    values[IDX_ACTUAL_COMPLETE] = complete
    return tuple(values)


def make_workbook(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.append([f"col{i}" for i in range(NUM_COLS)])
    for row in rows:
        ws.append(row)
    wb.save(path)


def test_lookup_lt_prints_order_level_total_when_available(tmp_path, capsys):
    path = tmp_path / "process.xlsx"
    make_workbook(
        [
            # 同じ受注(order_no+line)の2工程 -> 1インスタンスとして通しの実績(10日)が使われる
            _row("DWG-A1", "L0", 1, 20230101, 20230103, order_no="AB123", line="1"),  # 単体では2日
            _row("DWG-A1", "HG", 2, 20230101, 20230111, order_no="AB123", line="1"),  # 単体では10日
        ],
        path,
    )

    exit_code = main(["DWG-A1", "--process-data", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "図番: DWG-A1" in captured.out
    assert "L0" in captured.out
    assert "HG" in captured.out
    # 単純合計(2+10=12日)ではなく、インスタンス全体の通し実績(01-01〜01-11=10日)が使われる
    assert "想定LT(実績ベース、受注〜完成の中央値): 10.0日" in captured.out
    assert "(過去1件の実績より)" in captured.out


def test_lookup_lt_falls_back_to_summed_steps_when_no_order_level_data(tmp_path, capsys):
    path = tmp_path / "process.xlsx"
    make_workbook(
        [
            # 受注№・行が空欄 -> インスタンス実績を組み立てられないため単純合計にフォールバック
            _row("DWG-A1", "L0", 1, 20230101, 20230103, order_no="", line=""),  # 2日
            _row("DWG-A1", "HG", 2, 20230101, 20230111, order_no="", line=""),  # 10日
        ],
        path,
    )

    exit_code = main(["DWG-A1", "--process-data", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "想定LT(参考値、工程別実績の単純合計): 12.0日" in captured.out
    assert "過大評価の可能性" in captured.out


def test_lookup_lt_reports_when_drawing_not_found(tmp_path, capsys):
    path = tmp_path / "process.xlsx"
    make_workbook([_row("DWG-A1", "L0", 1, 20230101, 20230103)], path)

    exit_code = main(["DWG-NOT-FOUND", "--process-data", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "見つかりませんでした" in captured.out
