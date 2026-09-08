#!/usr/bin/env python3
"""生産管理支援ツール CLIエントリポイント。

使い方:
    python analyze.py 今週の受注データ.xlsx
    python analyze.py 今週の受注データ.xlsx --supplier-data 仕入データ*.xlsx --process-data 工程データ*.xlsx

--supplier-data / --process-data を指定すると、①②の停滞工程について過去の実績から
代替候補(社内設備・外注仕入先)を提示する(図番一致を優先、なければ工程コード一致)。
--process-data は工程別仕掛中ランキングに「参考実績LT(暦日)」も追加する
(標準LTは変更せず、判定ロジックにも使わない。あくまで参考情報)。

いずれも行数が多く読み込みに時間がかかるため、集計結果を .cache/ にキャッシュし、
同じファイル一式であれば2回目以降は高速化される。

実行のたびに output/YYYYMMDD/ フォルダを作成し、Excel(5〜6シート)と
オフラインHTMLレポートを出力する。外部API・外部通信は使用しない。
"""
from __future__ import annotations

import argparse
import datetime
import sys
from pathlib import Path

from src.excel_io import load_orders
from src.judge import judge_record
from src.process_history import ProcessHistory
from src.process_master import ProcessMaster
from src.report_data import build_report_data
from src.report_excel import write_excel_report
from src.report_html import write_html_report
from src.supplier_history import SupplierHistory


def run(
    input_path: Path,
    today: datetime.date,
    output_root: Path,
    supplier_data_paths: list[Path] | None = None,
    process_data_paths: list[Path] | None = None,
) -> Path:
    orders = load_orders(input_path)
    process_master = ProcessMaster()
    supplier_history = SupplierHistory.load(supplier_data_paths) if supplier_data_paths else None
    process_history = ProcessHistory.load(process_data_paths) if process_data_paths else None

    for order in orders:
        judge_record(order)

    report_data = build_report_data(orders, today, process_master, supplier_history, process_history)

    output_dir = output_root / today.strftime("%Y%m%d")
    write_excel_report(
        orders, output_dir / "production_delivery_report.xlsx", today,
        process_master, supplier_history, process_history,
    )
    write_html_report(report_data, output_dir / "production_delivery_report.html")

    return output_dir


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="受注データExcelから納期遅延/リスクを分析します")
    parser.add_argument("input_file", type=Path, help="基幹システムからエクスポートした受注データExcelファイル")
    parser.add_argument(
        "--supplier-data", type=Path, nargs="+", default=None,
        help="仕入(外注)累積データExcel(複数可)。指定すると代替候補仕入先を提示する",
    )
    parser.add_argument(
        "--process-data", type=Path, nargs="+", default=None,
        help="工程累積(社内実績)データExcel(複数可)。指定すると代替候補設備・参考実績LTを提示する",
    )
    args = parser.parse_args(argv)

    if not args.input_file.exists():
        print(f"入力ファイルが見つかりません: {args.input_file}", file=sys.stderr)
        return 1
    for group_name, paths in (("仕入", args.supplier_data), ("工程", args.process_data)):
        if not paths:
            continue
        for p in paths:
            if not p.exists():
                print(f"{group_name}データファイルが見つかりません: {p}", file=sys.stderr)
                return 1

    today = datetime.date.today()
    output_dir = run(args.input_file, today, Path("output"), args.supplier_data, args.process_data)
    print(f"出力しました: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
