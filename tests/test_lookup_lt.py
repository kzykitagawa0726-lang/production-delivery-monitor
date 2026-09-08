from openpyxl import Workbook

from lookup_lt import main
from src.process_history import (
    IDX_ACTUAL_COMPLETE,
    IDX_DRAWING,
    IDX_MACHINE,
    IDX_PROC_CODE,
    IDX_PROC_SEQ,
    IDX_START,
)

NUM_COLS = 98


def _row(drawing, process, seq, start, complete, machine="KW451"):
    values = [0] * NUM_COLS
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


def test_lookup_lt_prints_route_and_total(tmp_path, capsys):
    path = tmp_path / "process.xlsx"
    make_workbook(
        [
            _row("DWG-A1", "L0", 1, 20230101, 20230103),  # 2日
            _row("DWG-A1", "HG", 2, 20230101, 20230111),  # 10日
        ],
        path,
    )

    exit_code = main(["DWG-A1", "--process-data", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "図番: DWG-A1" in captured.out
    assert "L0" in captured.out
    assert "HG" in captured.out
    assert "想定LT合計(実績ベース、暦日): 12.0日" in captured.out


def test_lookup_lt_reports_when_drawing_not_found(tmp_path, capsys):
    path = tmp_path / "process.xlsx"
    make_workbook([_row("DWG-A1", "L0", 1, 20230101, 20230103)], path)

    exit_code = main(["DWG-NOT-FOUND", "--process-data", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "見つかりませんでした" in captured.out
