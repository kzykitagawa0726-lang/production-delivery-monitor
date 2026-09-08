#!/usr/bin/env python3
"""図番を指定して、実績データに基づく想定リードタイム(LT)を調べるCLIツール。

使い方:
    python lookup_lt.py <図番> --process-data 工程データ*.xlsx

工程累積データ(社内実績)から、その図番が過去にたどった典型的な工程ルート
(工程順の中央値で並び替え)と、各工程の実績LT(暦日、着手日〜完成日の中央値)を
表示する。図番自体の実績が無い工程は、工程コード単位の全図番平均にフォールバックし、
その旨を明示する。

①②の判定(自社納期・残日ベース)には一切影響しない、独立した参考情報。
外部API・外部通信は使用しない。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from src.process_history import ProcessHistory


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="図番から実績ベースの想定LTを調べます")
    parser.add_argument("drawing_no", help="調べたい図番")
    parser.add_argument(
        "--process-data", type=Path, nargs="+", required=True,
        help="工程累積(社内実績)データExcel(複数可)",
    )
    args = parser.parse_args(argv)

    for p in args.process_data:
        if not p.exists():
            print(f"工程データファイルが見つかりません: {p}", file=sys.stderr)
            return 1

    history = ProcessHistory.load(args.process_data)
    estimate = history.estimate_route_for_drawing(args.drawing_no)

    if estimate is None:
        print(f"図番「{args.drawing_no}」の実績が工程累積データに見つかりませんでした。")
        print("完全一致する図番のみ検索対象です。表記ゆれ(全角/半角、末尾スペース等)をご確認ください。")
        return 1

    print(f"図番: {estimate.drawing_no}")
    print(f"工程数: {len(estimate.steps)}")
    print()
    print(f"{'順':>3} {'工程コード':<10} {'実績LT(暦日)':>12} {'サンプル数':>8}  備考")
    print("-" * 60)
    for i, step in enumerate(estimate.steps, start=1):
        lt_text = f"{step.actual_lt_calendar_days:.1f}" if step.actual_lt_calendar_days is not None else "データなし"
        note = "この図番の実績" if step.is_drawing_specific else "工程コード全体の平均(フォールバック)"
        if step.actual_lt_calendar_days is None:
            note = "⚠ 実績データなし"
        print(f"{i:>3} {step.process_code:<10} {lt_text:>12} {step.sample_count:>8}  {note}")

    print("-" * 60)
    if estimate.total_calendar_days is not None:
        print(f"想定LT合計(実績ベース、暦日): {estimate.total_calendar_days:.1f}日")
    else:
        print("想定LT合計: 算出不可(実績データが1件もありません)")

    if estimate.missing_process_codes:
        print(
            f"\n⚠ 以下の工程は実績データが無いため、上記の合計には含まれていません(過小評価の可能性): "
            f"{', '.join(estimate.missing_process_codes)}"
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
