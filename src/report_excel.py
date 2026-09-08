"""Excel(5〜6シート)レポート出力。

シート構成:
  1. メインサマリー
  2. ①納期遅延リスト(超過日数順)
  3. ②納期遅延リスクリスト(危険度順)
  4. 判定不能・要確認リスト
  5. 工程別仕掛中ランキング(現在停滞している工程コード別。ボトルネック候補を強調。
     --process-data指定時は「参考実績LT(暦日)」列も追加。標準LT(営業日)は変更しない)
  6. 工程別仕入先実績ランキング(--supplier-data指定時のみ。仕入累積データから
     工程コード別に実績の多い仕入先を集計)

①②リストには「代替候補(社内設備／外注仕入先)」列を追加する
(--process-data / --supplier-data のどちらか、または両方を指定した場合のみ値が入る)。

工程コード→カテゴリの対応表が未整備の間は、シート5のカテゴリ・ボトルネック表示は
すべて「その他(未分類)」となる。届き次第、config/process_code_master.csv を
更新すれば自動的に反映される。
"""
from __future__ import annotations

import datetime
from pathlib import Path
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from src.models import OrderRecord
from src.process_history import ProcessHistory
from src.process_master import ProcessMaster
from src.report_data import ReportData, build_report_data
from src.supplier_history import SupplierHistory

HEADER_FILL = PatternFill(start_color="FF305496", end_color="FF305496", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFFFF")
DELAYED_FILL = PatternFill(start_color="FFF8CBAD", end_color="FFF8CBAD", fill_type="solid")
RISK_FILL = PatternFill(start_color="FFFFE699", end_color="FFFFE699", fill_type="solid")
UNDETERMINED_FILL = PatternFill(start_color="FFD9D9D9", end_color="FFD9D9D9", fill_type="solid")
BOTTLENECK_FILL = PatternFill(start_color="FFE35D5D", end_color="FFE35D5D", fill_type="solid")


def _write_header_row(ws: Worksheet, headers: list[str], row: int = 1) -> None:
    for col, title in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")


def _autosize_columns(ws: Worksheet, headers: list[str], rows: list[tuple]) -> None:
    widths = [len(h) for h in headers]
    for row in rows:
        for i, value in enumerate(row):
            widths[i] = max(widths[i], len(str(value)) if value is not None else 0)
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(width + 2, 40)


def _current_process_label(r: OrderRecord) -> str:
    step = r.current_process
    if step is None:
        return "(全工程完了)"
    return f"{step.process_code} / {step.status}"


def _alternative_suggestions_label(r: OrderRecord) -> Optional[str]:
    """代替候補(社内設備→外注仕入先の順)。どちらも無ければNone。"""
    labels = [s.label for s in r.machine_suggestions] + [s.label for s in r.supplier_suggestions]
    if not labels:
        return None
    return "; ".join(labels)


def _common_row(r: OrderRecord) -> tuple:
    return (
        r.order_no, r.drawing_no, r.product_name, r.qty, r.customer_order_no,
        r.company_deadline, r.customer_deadline, r.deadline_gap_days,
        _current_process_label(r),
        "はい" if r.is_forecast_order else "",
        r.sales_note or (r.sales_agreed_deadline.isoformat() if r.sales_agreed_deadline else None),
        r.customer_no, r.old_system_no, _alternative_suggestions_label(r),
    )


COMMON_HEADERS = [
    "製造オーダー№", "図番", "品名", "数量", "客先注番", "自社納期", "顧客納期",
    "対顧客納期差(日)", "現在工程/ステータス", "内示・先行手配", "営業メモ(納期確認等)", "得意先NO", "製番",
    "代替候補(社内設備／外注仕入先)",
]


def _delayed_row(r: OrderRecord) -> tuple:
    overdue_days = -r.remaining_business_days if r.remaining_business_days is not None else None
    common = _common_row(r)
    return (*common[:7], overdue_days, *common[7:])


def _risk_row(r: OrderRecord) -> tuple:
    common = _common_row(r)
    return (*common[:7], r.remaining_business_days, *common[7:])


def _undetermined_row(r: OrderRecord) -> tuple:
    return (*_common_row(r), r.judgement_reason)


def write_excel_report(
    records: list[OrderRecord],
    output_path: Path,
    generated_at: datetime.date,
    process_master: ProcessMaster | None = None,
    supplier_history: SupplierHistory | None = None,
    process_history: ProcessHistory | None = None,
) -> None:
    data = build_report_data(records, generated_at, process_master, supplier_history, process_history)

    wb = Workbook()

    _write_summary_sheet(wb.active, data)

    _write_list_sheet(
        wb.create_sheet("①納期遅延リスト"),
        headers=[*COMMON_HEADERS[:7], "超過営業日数", *COMMON_HEADERS[7:]],
        rows=[_delayed_row(r) for r in data.delayed],
        highlight_fill=DELAYED_FILL,
    )
    _write_list_sheet(
        wb.create_sheet("②納期遅延リスクリスト"),
        headers=[*COMMON_HEADERS[:7], "残営業日", *COMMON_HEADERS[7:]],
        rows=[_risk_row(r) for r in data.at_risk],
        highlight_fill=RISK_FILL,
    )
    _write_list_sheet(
        wb.create_sheet("判定不能・要確認"),
        headers=[*COMMON_HEADERS, "判定理由"],
        rows=[_undetermined_row(r) for r in data.undetermined],
        highlight_fill=UNDETERMINED_FILL,
    )
    _write_congestion_sheet(wb.create_sheet("工程別仕掛中ランキング"), data)
    if data.supplier_ranking_by_process:
        _write_supplier_ranking_sheet(wb.create_sheet("工程別仕入先実績ランキング"), data)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def _write_summary_sheet(ws: Worksheet, data: ReportData) -> None:
    ws.title = "メインサマリー"
    ws.append(["生産管理支援ツール メインサマリー"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([f"出力日: {data.generated_at.isoformat()}"])
    ws.append([f"対象件数: {data.total_count}"])
    ws.append([])

    ws.append(["区分", "件数", "備考"])
    for col in range(1, 4):
        ws.cell(row=ws.max_row, column=col).font = Font(bold=True)

    row = ws.max_row + 1
    ws.cell(row=row, column=1, value="① 納期遅延(超過)")
    ws.cell(row=row, column=2, value=len(data.delayed))
    ws.cell(row=row, column=1).fill = DELAYED_FILL
    ws.cell(row=row, column=2).fill = DELAYED_FILL

    row += 1
    ws.cell(row=row, column=1, value="② 納期遅延リスク")
    ws.cell(row=row, column=2, value=len(data.at_risk))
    ws.cell(row=row, column=3, value="残営業日5日以内")
    ws.cell(row=row, column=1).fill = RISK_FILL
    ws.cell(row=row, column=2).fill = RISK_FILL

    row += 1
    ws.cell(row=row, column=1, value="判定不能・要確認")
    ws.cell(row=row, column=2, value=len(data.undetermined))
    ws.cell(row=row, column=3, value="自社納期・残日が取得できない、または残日が異常値の案件。①②には含めない")

    row += 2
    ws.cell(row=row, column=1, value="(参考)内示・先行手配と推測される件数")
    ws.cell(row=row, column=2, value=data.forecast_order_count)
    ws.cell(row=row, column=3, value="受注日・顧客納期が未設定。①②の判定には通常通り含めている")

    if data.unknown_process_codes:
        row += 2
        ws.cell(row=row, column=1, value="⚠ 未分類の工程コード数")
        ws.cell(row=row, column=2, value=len(data.unknown_process_codes))
        ws.cell(row=row, column=3, value="工程コード対応表(config/process_code_master.csv)に未登録。シート5参照")

    for col_letter, width in zip("ABC", (32, 12, 55)):
        ws.column_dimensions[col_letter].width = width


def _write_list_sheet(ws: Worksheet, headers: list[str], rows: list[tuple], highlight_fill: PatternFill) -> None:
    _write_header_row(ws, headers)
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
        ws.cell(row=r_idx, column=1).fill = highlight_fill
    _autosize_columns(ws, headers, rows)
    ws.freeze_panes = "A2"


def _write_congestion_sheet(ws: Worksheet, data: ReportData) -> None:
    headers = ["工程コード", "カテゴリ", "仕掛中件数", "標準LT(営業日)", "参考実績LT(暦日)", "ボトルネック", "未分類"]
    _write_header_row(ws, headers)
    for r_idx, entry in enumerate(data.congestion_ranking, start=2):
        ws.cell(row=r_idx, column=1, value=entry.process_code)
        ws.cell(row=r_idx, column=2, value=entry.category)
        ws.cell(row=r_idx, column=3, value=entry.count)
        ws.cell(row=r_idx, column=4, value=entry.standard_lt_business_days)
        ws.cell(row=r_idx, column=5, value=entry.actual_lt_calendar_days)
        ws.cell(row=r_idx, column=6, value="★ボトルネック" if entry.is_bottleneck else "")
        ws.cell(row=r_idx, column=7, value="⚠未分類" if entry.is_unknown_code else "")
        if entry.is_bottleneck:
            for c in range(1, 8):
                ws.cell(row=r_idx, column=c).fill = BOTTLENECK_FILL
    rows = [
        (e.process_code, e.category, e.count, e.standard_lt_business_days, e.actual_lt_calendar_days,
         e.is_bottleneck, e.is_unknown_code)
        for e in data.congestion_ranking
    ]
    _autosize_columns(ws, headers, rows)
    ws.freeze_panes = "A2"

    if data.unknown_process_codes:
        note_row = ws.max_row + 2
        ws.cell(
            row=note_row, column=1,
            value=(
                "⚠ 工程コード対応表(config/process_code_master.csv)が未整備のため、"
                "現時点ではすべての工程コードが「その他(未分類)」として扱われています。"
                "対応表が届き次第、正しいカテゴリ・ボトルネック区分が反映されます。"
            ),
        )
        ws.cell(row=note_row, column=1).font = Font(italic=True, color="FF9C0006")


def _write_supplier_ranking_sheet(ws: Worksheet, data: ReportData) -> None:
    """工程コード別に、過去の仕入(外注)実績の多い仕入先を並べる(--supplier-data指定時のみ)。

    仕入単価・金額は含めない(件数・最終利用日のみ。kii-san合意事項)。
    """
    headers = ["工程コード", "順位", "仕入先CD", "実績件数", "最終利用日"]
    _write_header_row(ws, headers)
    rows: list[tuple] = []
    for process_code, suggestions in data.supplier_ranking_by_process.items():
        for rank, s in enumerate(suggestions, start=1):
            rows.append((process_code, rank, s.supplier_code, s.count, s.last_used))
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    _autosize_columns(ws, headers, rows)
    ws.freeze_panes = "A2"
