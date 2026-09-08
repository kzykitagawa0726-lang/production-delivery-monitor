#!/usr/bin/env python3
"""図番を指定して、実績データに基づく想定リードタイム(LT)を調べるCLIツール。

使い方:
    python lookup_lt.py <図番> --process-data 工程データ*.xlsx

工程累積データ(社内実績)から、その図番が過去に「受注(最初の着手)〜完成(最後の完成)」で
実際どれだけかかったか(実績LT中央値、暦日)を表示する。参考として、典型的な工程ルート
(工程順の中央値で並び替え)と各工程ごとの実績LTも合わせて表示する。

【2026-09-08 実データ検証で設計変更】当初は各工程の実績LTを単純合計していたが、
各工程の実績日数には他案件との待ち時間が相当含まれており、工程数の多い図番では
合計が数千日規模になる異常値が実データで見つかった。→ 図番単位の「受注〜完成」
実績日数(1受注インスタンスの通しの実績)の中央値を主役に変更した。

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
    print()

    if estimate.total_calendar_days is not None and estimate.total_is_order_level:
        print(
            f"★ 想定LT(実績ベース、受注〜完成の中央値): {estimate.total_calendar_days:.1f}日"
            f"  (過去{estimate.total_sample_count}件の実績より)"
        )
    elif estimate.total_calendar_days is not None:
        print(
            f"△ 想定LT(参考値、工程別実績の単純合計): {estimate.total_calendar_days:.1f}日"
            "  (この図番の受注〜完成の通し実績が無いためのフォールバック。過大評価の可能性があります)"
        )
    else:
        print("想定LT: 算出不可(実績データが1件もありません)")

    print()
    print(f"--- 参考: 典型的な工程ルート(工程数: {len(estimate.steps)}) ---")
    print(f"{'順':>3} {'工程コード':<10} {'工程別実績LT(暦日)':>16} {'サンプル数':>8}  備考")
    print("-" * 66)
    for i, step in enumerate(estimate.steps, start=1):
        lt_text = f"{step.actual_lt_calendar_days:.1f}" if step.actual_lt_calendar_days is not None else "データなし"
        if step.actual_lt_calendar_days is None:
            note = "⚠ 実績データなし"
        else:
            note = "この図番の実績" if step.is_drawing_specific else "工程コード全体の平均(フォールバック)"
        print(f"{i:>3} {step.process_code:<10} {lt_text:>16} {step.sample_count:>8}  {note}")

    if estimate.missing_process_codes:
        print(
            f"\n⚠ 以下の工程は実績データが無いため、工程別内訳には反映されていません: "
            f"{', '.join(estimate.missing_process_codes)}"
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
