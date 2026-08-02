"""Excel(5シート)レポート出力。

シート構成:
  1. メインサマリー
  2. ①納期遅延リスト(超過日数順)
  3. ②納期遅延リスクリスト(危険度順)
  4. 判定不能・要確認リスト
  5. 工程別混雑ランキング(ボトルネック工程を強調)
"""
from __future__ import annotations

import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from src.aggregate import CongestionEntry
from src.models import OrderRecord
from src.report_data import ReportData, build_report_data

HEADER_FILL = PatternFill(start_color="FF305496", end_color="FF305496", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFFFF")
DELAYED_FILL = PatternFill(start_color="FFF8CBAD", end_color="FFF8CBAD", fill_type="solid")
RISK_FILL = PatternFill(start_color="FFFFE699", end_color="FFFFE699", fill_type="solid")
BOTTLENECK_FILL = PatternFill(start_color="FFFF7C80", end_color="FFFF7C80", fill_type="solid")
UNKNOWN_CODE_FILL = PatternFill(start_color="FFD9D9D9", end_color="FFD9D9D9", fill_type="solid")


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


def _delayed_row(r: OrderRecord) -> tuple:
    overdue_days = -r.remaining_business_days_to_deadline if r.remaining_business_days_to_deadline is not None else None
    return (
        r.order_no, r.drawing_no, r.qty, r.assignee, r.product_name,
        r.company_deadline, overdue_days, r.process_code, r.process_category,
        r.judgement_reason,
    )


def _risk_row(r: OrderRecord) -> tuple:
    gap = (
        r.remaining_required_business_days - r.remaining_business_days_to_deadline
        if r.remaining_required_business_days is not None and r.remaining_business_days_to_deadline is not None
        else None
    )
    return (
        r.order_no, r.drawing_no, r.qty, r.assignee, r.product_name,
        r.company_deadline, r.remaining_business_days_to_deadline,
        r.remaining_required_business_days, gap, r.process_code, r.process_category,
        r.judgement_reason,
    )


def _undetermined_row(r: OrderRecord) -> tuple:
    return (
        r.order_no, r.drawing_no, r.qty, r.assignee, r.product_name,
        r.company_deadline, r.total_process_count, r.process_seq,
        r.process_code, r.judgement_reason,
    )


def write_excel_report(
    records: list[OrderRecord],
    output_path: Path,
    generated_at: datetime.date,
    unknown_process_codes: list[str] | None = None,
) -> None:
    data = build_report_data(records, generated_at, unknown_process_codes)

    wb = Workbook()

    _write_summary_sheet(wb.active, data)

    _write_list_sheet(
        wb.create_sheet("①納期遅延リスト"),
        headers=["受注No", "図番", "数量", "担当者", "商品名", "自社納期", "超過営業日数", "工程コード", "工程カテゴリ", "判定理由"],
        rows=[_delayed_row(r) for r in data.delayed],
        highlight_fill=DELAYED_FILL,
    )
    _write_list_sheet(
        wb.create_sheet("②納期遅延リスクリスト"),
        headers=["受注No", "図番", "数量", "担当者", "商品名", "自社納期", "残営業日", "残必要日数", "不足日数", "工程コード", "工程カテゴリ", "判定理由"],
        rows=[_risk_row(r) for r in data.at_risk],
        highlight_fill=RISK_FILL,
    )
    _write_list_sheet(
        wb.create_sheet("判定不能・要確認"),
        headers=["受注No", "図番", "数量", "担当者", "商品名", "自社納期", "全工程数", "工程順", "工程コード", "判定理由"],
        rows=[_undetermined_row(r) for r in data.undetermined],
        highlight_fill=UNKNOWN_CODE_FILL,
    )
    _write_congestion_sheet(wb.create_sheet("工程別混雑ランキング"), data.congestion, data.unknown_process_codes)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def _write_summary_sheet(ws: Worksheet, data: ReportData) -> None:
    ws.title = "メインサマリー"
    ws.append(["生産管理支援ツール メインサマリー"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([f"出力日: {data.generated_at.isoformat()}"])
    ws.append([f"対象件数(仕掛中): {data.total_count}"])
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
    ws.cell(row=row, column=1).fill = RISK_FILL
    ws.cell(row=row, column=2).fill = RISK_FILL

    row += 1
    ws.cell(row=row, column=1, value="判定不能・要確認")
    ws.cell(row=row, column=2, value=len(data.undetermined))
    ws.cell(row=row, column=3, value="工程未展開/納期未更新など。①②には含めない")

    if data.unknown_process_codes:
        row += 2
        ws.cell(row=row, column=1, value="⚠ 未知の工程コード検出").font = Font(bold=True, color="FFC00000")
        row += 1
        ws.cell(row=row, column=1, value=", ".join(data.unknown_process_codes))

    for col_letter, width in zip("ABC", (28, 12, 50)):
        ws.column_dimensions[col_letter].width = width


def _write_list_sheet(ws: Worksheet, headers: list[str], rows: list[tuple], highlight_fill: PatternFill) -> None:
    _write_header_row(ws, headers)
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
        ws.cell(row=r_idx, column=1).fill = highlight_fill
    _autosize_columns(ws, headers, rows)
    ws.freeze_panes = "A2"


def _write_congestion_sheet(ws: Worksheet, congestion: list[CongestionEntry], unknown_process_codes: list[str]) -> None:
    headers = ["工程コード", "工程カテゴリ", "仕掛件数", "ボトルネック工程"]
    _write_header_row(ws, headers)
    rows: list[tuple] = []
    for r_idx, entry in enumerate(congestion, start=2):
        is_unknown = entry["process_code"] in unknown_process_codes
        row = (
            entry["process_code"],
            entry["category"],
            entry["count"],
            "★ボトルネック" if entry["is_bottleneck"] else "",
        )
        rows.append(row)
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
        if entry["is_bottleneck"]:
            for c_idx in range(1, len(headers) + 1):
                ws.cell(row=r_idx, column=c_idx).fill = BOTTLENECK_FILL
        elif is_unknown:
            for c_idx in range(1, len(headers) + 1):
                ws.cell(row=r_idx, column=c_idx).fill = UNKNOWN_CODE_FILL
    _autosize_columns(ws, headers, rows)
    ws.freeze_panes = "A2"
