#!/usr/bin/env python3
"""生産管理支援ツール CLIエントリポイント。

使い方:
    python analyze.py 今週の受注データ.xlsx

実行のたびに output/YYYYMMDD/ フォルダを作成し、Excel(5シート)と
オフラインHTMLレポートを出力する。外部API・外部通信は使用しない。
"""
from __future__ import annotations

import argparse
import datetime
import sys
from pathlib import Path

from src.excel_io import load_orders
from src.judge import judge_record
from src.process_master import ProcessMaster
from src.report_data import build_report_data
from src.report_excel import write_excel_report
from src.report_html import write_html_report


def run(input_path: Path, today: datetime.date, output_root: Path) -> Path:
    orders = load_orders(input_path)

    process_master = ProcessMaster()
    for order in orders:
        result = process_master.categorize(order.process_code)
        order.process_category = result.category
        order.current_process_standard_lt = result.standard_lt_business_days
        order.is_unknown_process_code = result.is_unknown_code
        order.is_bottleneck_process = result.is_bottleneck

    in_progress = [o for o in orders if o.all_process_complete_flag == 0]
    for order in in_progress:
        judge_record(order, today)

    unknown_codes = process_master.get_unknown_codes()
    report_data = build_report_data(in_progress, today, unknown_codes)

    output_dir = output_root / today.strftime("%Y%m%d")
    write_excel_report(in_progress, output_dir / "production_delivery_report.xlsx", today, unknown_codes)
    write_html_report(report_data, output_dir / "production_delivery_report.html")

    return output_dir


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="受注データExcelから納期遅延/リスクを分析します")
    parser.add_argument("input_file", type=Path, help="基幹システムからエクスポートした受注データExcelファイル")
    args = parser.parse_args(argv)

    if not args.input_file.exists():
        print(f"入力ファイルが見つかりません: {args.input_file}", file=sys.stderr)
        return 1

    today = datetime.date.today()
    try:
        output_dir = run(args.input_file, today, Path("output"))
    except NotImplementedError as e:
        print(f"未実装: {e}", file=sys.stderr)
        return 2
    print(f"出力しました: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
