"""Excel(5〜16シート)レポート出力。

シート構成:
  1. メインサマリー
  2. ①納期遅延リスト(超過日数順)
  3. ②納期遅延リスクリスト(危険度順)
  4. 判定不能・要確認リスト
  5. 工程別仕掛中ランキング(現在停滞している工程コード別。ボトルネック候補を強調。
     --process-data指定時は「参考実績LT(暦日)」列も追加。標準LT(営業日)は変更しない)
  6. 工程別仕入先実績ランキング(--supplier-data指定時のみ。仕入累積データから
     工程コード別に実績の多い仕入先を集計)
  7. 実績ベース納期充足予測(--process-data指定時のみ。図番単位の「受注〜完成」実績日数から
     予測完了日を算出し、顧客納期と比較する。①②の判定[自社納期・残日ベース]とは独立した参考情報)
  8. 週別負荷実績(部署別)/9. 週別負荷実績(設備別)(--process-data指定時のみ。
     過去の実績工数を週別に集計した「振り返り」)
  10. 週別予測負荷(部署別)/11. 週別予測負荷(設備別)(--process-data指定時のみ。
     仕掛中の受注の残り工程を標準LTで先の週へ積み上げた「これから」の需要予測。
     8・9とは向き[過去/未来]が異なる別物なので混同しないよう注意)
  12. 品種別月間キャパシティ(GEAR/BEVEL/WORM)(--process-data指定時のみ。稼働率30〜40%からの
     逆算によるキャパシティ目安と、月別の実績・予測工数を品種別に並べた営業向けのざっくり参考資料)
  13. 週別予測負荷内訳(設備別)(--process-data指定時のみ。10・11の設備別予測を、
     どの製造オーダー・どの工程が積み上がっているかまで1行ずつ分解した内訳。
     2026-09-08 kii-san要望:「調整必要な設備を週単位で、製造オーダー単位まで確認したい」)
  14. 品種別月間キャパシティ内訳(--process-data指定時のみ。12の予測工数(forecast_hours)を
     製造オーダー単位まで分解した内訳。2026-09-08 kii-san要望:「月ごとの品種別キャパシティを、
     負荷分散のため製造オーダーに降りて確認したい」。実績側は受注単位で追えないため予測側のみ)
  15. 仕入先週別予測(--supplier-data と --process-data の両方指定時のみ。仕入累積データに
     工数の記録が無いため金額を使うが、個別受注・仕入先の金額は表示せず仕入先ごとの週次合計
     金額のみを使う。社内設備のような稼働率基準が無いため、絶対的な上限ではなく「普段の
     実績水準との比較倍率」で示す。2026-09-08 kii-san要望:「調整必要な仕入れ先を週単位で」)
  16. 仕入先週別予測内訳(15の予測金額を製造オーダー単位まで分解した内訳)

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
from src.product_category import ProductCategoryClassifier
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
    product_category_classifier: ProductCategoryClassifier | None = None,
) -> None:
    data = build_report_data(
        records, generated_at, process_master, supplier_history, process_history, product_category_classifier
    )

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
    if data.feasibility_estimates:
        _write_feasibility_sheet(wb.create_sheet("実績ベース納期充足予測"), data)
    if data.weekly_load_by_department:
        _write_weekly_department_pivot_sheet(
            wb.create_sheet("週別負荷実績(部署別)"), data.weekly_load_by_department, "実績工数合計"
        )
    if data.weekly_load_by_machine:
        _write_weekly_machine_list_sheet(
            wb.create_sheet("週別負荷実績(設備別)"), data.weekly_load_by_machine, "実績工数合計"
        )
    if data.capacity_forecast_by_department:
        _write_weekly_department_pivot_sheet(
            wb.create_sheet("週別予測負荷(部署別)"), data.capacity_forecast_by_department, "予測工数合計"
        )
        _write_capacity_forecast_note(wb["週別予測負荷(部署別)"])
    if data.capacity_forecast_by_machine:
        _write_weekly_machine_list_sheet(
            wb.create_sheet("週別予測負荷(設備別)"), data.capacity_forecast_by_machine, "予測工数合計"
        )
        _write_capacity_forecast_note(wb["週別予測負荷(設備別)"])
    if data.capacity_forecast_machine_detail:
        _write_machine_forecast_detail_sheet(wb.create_sheet("週別予測負荷内訳(設備別)"), data)
    if data.monthly_category_capacity:
        _write_monthly_category_capacity_sheet(wb.create_sheet("品種別月間キャパシティ"), data)
    if data.monthly_category_capacity_detail:
        _write_monthly_category_capacity_detail_sheet(wb.create_sheet("品種別月間キャパシティ内訳"), data)
    if data.supplier_capacity_forecast:
        _write_supplier_capacity_forecast_sheet(wb.create_sheet("仕入先週別予測"), data)
    if data.supplier_capacity_forecast_detail:
        _write_supplier_capacity_forecast_detail_sheet(wb.create_sheet("仕入先週別予測内訳"), data)

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


def _write_feasibility_sheet(ws: Worksheet, data: ReportData) -> None:
    """図番単位の受注〜完成実績から算出した客先納期充足予測(--process-data指定時のみ)。

    ①②(自社納期・残日ベース)の判定とは独立した参考情報。margin_days昇順
    (間に合わない見込み・データ不足が上に来る)。
    """
    headers = [
        "製造オーダー№", "図番", "品名", "受注日", "顧客納期", "現在工程", "残り工程数",
        "図番実績LT中央値(暦日)", "実績件数", "予測完了日(受注日+実績LT)",
        "余裕日数(顧客納期-予測完了日)", "判定", "備考",
    ]
    _write_header_row(ws, headers)
    rows: list[tuple] = []
    for r_idx, e in enumerate(data.feasibility_estimates, start=2):
        row = (
            e.order_no, e.drawing_no, e.product_name, e.order_date, e.customer_deadline, e.current_process_code,
            e.remaining_step_count, e.typical_total_lt_calendar_days, e.typical_lt_sample_count,
            e.predicted_completion_date, e.margin_days, e.status, e.data_note,
        )
        rows.append(row)
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
        if e.status == "間に合わない見込み":
            ws.cell(row=r_idx, column=1).fill = DELAYED_FILL
        elif e.status == "データ不足":
            ws.cell(row=r_idx, column=1).fill = UNDETERMINED_FILL
    _autosize_columns(ws, headers, rows)
    ws.freeze_panes = "A2"

    note_row = ws.max_row + 2
    ws.cell(
        row=note_row, column=1,
        value=(
            "※ この予測は①②の判定(自社納期・残日ベース)とは独立した参考情報です。"
            "その図番が過去に実際「受注(最初の着手)〜完成(最後の完成)」でどれだけかかったか(実績LT中央値)を"
            "受注日に足して予測完了日を算出し、顧客納期と比較しています(残り工程数による比例配分はしていません)。"
            "実績件数が少ない行(1〜2件)は参考程度にご覧ください。"
        ),
    )
    ws.cell(row=note_row, column=1).font = Font(italic=True, color="FF9C0006")


def _write_weekly_department_pivot_sheet(
    ws: Worksheet, points: list[tuple[datetime.date, str, float]], value_header: str
) -> None:
    """週×部署コードのピボット表(週を行、部署コードを列とする)。

    「加工先」列(数値・13種類)を部署/コストセンターコードとして扱う(kii-san確認)。
    実績シートでは週は完成日が属する月曜始まりの週、予測シートでは標準LTを積み上げた予測日が
    属する週(いずれもMVP。日別按分はしていない)。
    """
    departments = sorted({dept for _, dept, _ in points})
    weeks = sorted({week for week, _, _ in points})
    grid: dict[tuple, float] = {(week, dept): hours for week, dept, hours in points}

    headers = ["週(月曜始まり)", *departments, "合計"]
    _write_header_row(ws, headers)
    rows: list[tuple] = []
    for week in weeks:
        values = [grid.get((week, dept), 0.0) for dept in departments]
        row = (week, *values, sum(values))
        rows.append(row)
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    _autosize_columns(ws, headers, rows)
    ws.freeze_panes = "B2"


def _write_weekly_machine_list_sheet(
    ws: Worksheet, points: list[tuple[datetime.date, str, float]], value_header: str
) -> None:
    """週×設備コードの一覧(設備数が多いため縦持ち形式。週昇順→工数降順)。"""
    headers = ["週(月曜始まり)", "設備コード", value_header]
    _write_header_row(ws, headers)
    rows = sorted(points, key=lambda t: (t[0], -t[2]))
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    _autosize_columns(ws, headers, rows)
    ws.freeze_panes = "A2"


def _write_monthly_category_capacity_sheet(ws: Worksheet, data: ReportData) -> None:
    """品種(GEAR/BEVEL/WORM)別・月別のキャパシティ参考資料(--process-data指定時のみ)。

    ①②の判定・実績ベース納期充足予測・週別(予測)負荷とは独立した、追加の参考情報。
    月×品種の行に、実績工数・予測工数・稼働率30〜40%から逆算したキャパシティ目安を並べる。
    """
    headers = [
        "月", "品種", "実績工数合計", "予測工数合計",
        "キャパシティ目安(稼働率35%)", "キャパシティ下限(稼働率40%・保守的)",
        "キャパシティ上限(稼働率30%・楽観的)", "キャパシティ目安(1営業日あたり、月20日換算)",
        "予測充足率(予測÷目安)",
    ]
    _write_header_row(ws, headers)
    rows: list[tuple] = []
    for r_idx, e in enumerate(data.monthly_category_capacity, start=2):
        row = (
            e.month, e.category, e.actual_hours, e.forecast_hours,
            e.capacity_hours_typical, e.capacity_hours_low, e.capacity_hours_high,
            e.capacity_hours_typical_per_business_day, e.fulfillment_rate,
        )
        rows.append(row)
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
        if e.fulfillment_rate is not None and e.fulfillment_rate > 1.0:
            ws.cell(row=r_idx, column=1).fill = DELAYED_FILL
    _autosize_columns(ws, headers, rows)
    ws.freeze_panes = "A2"

    note_row = ws.max_row + 2
    ws.cell(
        row=note_row, column=1,
        value=(
            "※ これは営業がおおまかに月間キャパシティ感をつかむための参考資料です(受注時のご参考に)。"
            "①②の判定・実績ベース納期充足予測とは独立しています。"
            "キャパシティ目安は、設備の物理的な上限を測定したものではなく、「弊社の稼働率はだいたい"
            "30〜40%(24時間を100%とした場合)」というご申告値をもとに、過去の典型的な実績月間工数を"
            "その稼働率で逆算した推定値です(稼働率が低いほどキャパシティは大きく出るため、"
            "40%を保守的な下限、30%を楽観的な上限としています)。実績月間工数は、算出時点の月(まだ"
            "確定していない)を除いた直近12か月の平均です。予測工数は、仕掛中の受注の残り工程を"
            "標準LTで先の月へ積み上げ、受注自身の品名から判定した品種(GEAR/BEVEL/WORM)で集計した"
            "ものです(部署・設備別の予測とは別の切り口)。品名から品種を判定できない受注"
            "(付随部品や型番のみの表記など)は集計から除外しています。"
            "「1営業日あたり」の列は、月間キャパシティ目安を弊社の平均営業日数(月20日)で単純に"
            "割った参考値です(月ごとの実際の営業日数の違いは考慮していません)。"
        ),
    )
    ws.cell(row=note_row, column=1).font = Font(italic=True, color="FF9C0006")


def _write_machine_forecast_detail_sheet(ws: Worksheet, data: ReportData) -> None:
    """週別予測負荷(設備別)を製造オーダー単位まで分解した内訳(--process-data指定時のみ)。

    1つの工程は過去の実績シェアに応じて複数の設備に按分される設計のため、同じ受注・工程が
    複数の設備にまたがって複数行に現れる(1台に決め打ちしない。kii-san合意)。
    """
    headers = ["週(月曜始まり)", "設備コード", "製造オーダー№", "図番", "品名", "工程コード", "見込み工数(按分後)"]
    _write_header_row(ws, headers)
    rows = [
        (e.week, e.machine_code, e.order_no, e.drawing_no, e.product_name, e.process_code, e.hours)
        for e in data.capacity_forecast_machine_detail
    ]
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    _autosize_columns(ws, headers, rows)
    ws.freeze_panes = "A2"

    note_row = ws.max_row + 2
    ws.cell(
        row=note_row, column=1,
        value=(
            "※ 「週別予測負荷(設備別)」シートの各設備・各週の合計値を、製造オーダー単位まで"
            "分解した内訳です。1つの工程は過去の実績シェアに応じて複数の設備に按分されるため、"
            "同じ受注・工程が複数の設備に分かれて現れます(1台に決め打ちしない設計。"
            "見込み工数はすでに按分後の値のため、同じ受注・工程の行を全部足すと元の工数に戻ります)。"
        ),
    )
    ws.cell(row=note_row, column=1).font = Font(italic=True, color="FF9C0006")


def _write_monthly_category_capacity_detail_sheet(ws: Worksheet, data: ReportData) -> None:
    """品種別月間キャパシティの予測工数を製造オーダー単位まで分解した内訳(--process-data指定時のみ)。"""
    headers = ["月", "品種", "製造オーダー№", "図番", "品名", "工程コード", "見込み工数"]
    _write_header_row(ws, headers)
    rows = [
        (e.month, e.category, e.order_no, e.drawing_no, e.product_name, e.process_code, e.forecast_hours)
        for e in data.monthly_category_capacity_detail
    ]
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    _autosize_columns(ws, headers, rows)
    ws.freeze_panes = "A2"

    note_row = ws.max_row + 2
    ws.cell(
        row=note_row, column=1,
        value=(
            "※ 「品種別月間キャパシティ」シートの予測工数(月×品種)の元になった、"
            "各受注・各残り工程の見込み工数を1行ずつ並べたものです(負荷分散のご検討にご利用ください)。"
            "実績工数(過去分)は集計元データが受注単位で追えないため、内訳はありません。"
        ),
    )
    ws.cell(row=note_row, column=1).font = Font(italic=True, color="FF9C0006")


def _write_supplier_capacity_forecast_sheet(ws: Worksheet, data: ReportData) -> None:
    """仕入先(外注)週別予測(--supplier-data と --process-data の両方指定時のみ)。

    個別受注・仕入先の金額は一切表示せず、仕入先ごとの週次合計金額のみを使う(kii-san合意)。
    """
    headers = ["週(月曜始まり)", "仕入先CD", "予測金額(合計)", "普段の週次実績金額(直近52週平均)", "普段との比率"]
    _write_header_row(ws, headers)
    rows: list[tuple] = []
    for r_idx, e in enumerate(data.supplier_capacity_forecast, start=2):
        row = (e.week, e.supplier_code, e.forecast_amount, e.typical_weekly_amount, e.ratio_to_typical)
        rows.append(row)
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
        if e.ratio_to_typical is not None and e.ratio_to_typical > 1.5:
            ws.cell(row=r_idx, column=1).fill = DELAYED_FILL
    _autosize_columns(ws, headers, rows)
    ws.freeze_panes = "A2"

    note_row = ws.max_row + 2
    ws.cell(
        row=note_row, column=1,
        value=(
            "※ 仕入累積データには工数(時間)の記録が無いため、金額を使って算出しています。"
            "個別の受注・仕入先ごとの金額は一切表示せず、仕入先ごとの週次合計金額のみを使っています。"
            "社内設備のような稼働率基準(30〜40%)が無く絶対的な上限は推定できないため、"
            "「普段の実績水準(算出時点の週を除く直近52週平均)と比べて、予測される依頼金額が"
            "どれだけ多い/少ないか」という相対比較(比率)として示しています。過去に外注実績が"
            "一切ない工程コードは、按分先が無いため対象外です(=社内設備側の予測のみに含まれます)。"
            "比率1.5倍超の行を目安として強調表示していますが、あくまで参考の目安です。"
        ),
    )
    ws.cell(row=note_row, column=1).font = Font(italic=True, color="FF9C0006")


def _write_supplier_capacity_forecast_detail_sheet(ws: Worksheet, data: ReportData) -> None:
    """仕入先週別予測の、製造オーダー単位の内訳。個別金額を含むため、社内利用に留めてください。"""
    headers = ["週(月曜始まり)", "仕入先CD", "製造オーダー№", "図番", "品名", "工程コード", "予測金額(按分後)"]
    _write_header_row(ws, headers)
    rows = [
        (e.week, e.supplier_code, e.order_no, e.drawing_no, e.product_name, e.process_code, e.forecast_amount)
        for e in data.supplier_capacity_forecast_detail
    ]
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    _autosize_columns(ws, headers, rows)
    ws.freeze_panes = "A2"

    note_row = ws.max_row + 2
    ws.cell(
        row=note_row, column=1,
        value=(
            "※ 「仕入先週別予測」シートの予測金額の元になった、各受注・各残り工程の見込み金額を"
            "1行ずつ並べたものです。1つの工程は過去の実績金額シェアに応じて複数の仕入先に按分される"
            "ため、同じ受注・工程が複数の仕入先に分かれて現れます。"
        ),
    )
    ws.cell(row=note_row, column=1).font = Font(italic=True, color="FF9C0006")


def _write_capacity_forecast_note(ws: Worksheet) -> None:
    """週別予測負荷シートの下に、算出方法と限界を明記する注記を付ける。"""
    note_row = ws.max_row + 2
    ws.cell(
        row=note_row, column=1,
        value=(
            "※ これは仕掛中の受注の「残り工程」を、標準LT(config/process_code_master.csv、"
            "営業日、実績日数ではない)で先の週へ積み上げた将来の予測需要です。工数の大きさは"
            "工程コードごとの過去の平均実績工数、部署・設備への配分は過去の実績シェアに基づく"
            "按分(比例配分)です。設備への配分は、将来どの設備が空いているかまでは予測できないため、"
            "過去の使用実績の比率で仮に割り振った参考値です。"
            "設備別のキャパシティ上限(1日あたり稼働可能時間など)が分かれば、この予測工数と比較して"
            "充足率を出せます。"
        ),
    )
    ws.cell(row=note_row, column=1).font = Font(italic=True, color="FF9C0006")
