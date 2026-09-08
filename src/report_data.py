"""Excel/HTML両方の出力が共有するレポート集計データ。"""
from __future__ import annotations

import datetime
from collections import Counter
from dataclasses import dataclass, field

from src.models import Judgement, OrderRecord
from src.process_history import ProcessHistory
from src.process_master import ProcessMaster
from src.supplier_history import SupplierHistory, SupplierSuggestion


@dataclass
class CongestionEntry:
    process_code: str
    category: str
    count: int
    is_bottleneck: bool
    is_unknown_code: bool
    standard_lt_business_days: int
    # 実績リードタイム中央値(暦日、着手日〜完成日)。--process-data指定時のみ、参考情報として設定される。
    # 標準LT(標準リードタイム、営業日)は変更しない(kii-san指示)。判定ロジックにも使わない。
    actual_lt_calendar_days: float | None = None


@dataclass
class FeasibilityEntry:
    """進行中の受注について、図番単位の「受注〜完成」実績日数から客先納期充足を予測する。

    ①②の判定(自社納期・残日ベース)とは独立した、追加の参考情報。
    --process-data指定時のみ算出される。

    【2026-09-08 実データ検証で設計変更】当初は残り工程ごとの実績LT(暦日)を
    単純合計していたが、各工程の実績日数には他案件との待ち時間が相当含まれており、
    残り工程が多い受注では合計が数千日規模になる異常値が実データで頻発した。
    → kii-san合意により、図番単位の「受注(最初の着手)〜完成(最後の完成)」実績日数
    (ProcessHistory.order_level_lt_calendar_days_median)を使う設計に変更。
    自社の受注日(order_date)を起点に、その図番が過去に実際どれだけかかったかの
    中央値を足して予測完了日とする(残り工程数による比例配分はしない。シンプルさを優先)。
    """

    order_no: str
    drawing_no: str
    product_name: str | None
    order_date: datetime.date | None
    customer_deadline: datetime.date
    current_process_code: str
    remaining_step_count: int
    typical_total_lt_calendar_days: float | None  # 図番単位の受注〜完成 実績中央値(暦日)
    typical_lt_sample_count: int  # 何件の過去実績から算出したか
    predicted_completion_date: datetime.date | None  # 受注日 + typical_total_lt
    margin_days: int | None  # 顧客納期 - 予測完了日。正=間に合う見込み、負=間に合わない見込み
    data_note: str | None = None  # 予測不可の理由(データ不足の場合)

    @property
    def status(self) -> str:
        if self.predicted_completion_date is None:
            return "データ不足"
        if self.margin_days is not None and self.margin_days < 0:
            return "間に合わない見込み"
        return "間に合う見込み"


@dataclass
class ReportData:
    generated_at: datetime.date
    total_count: int
    delayed: list[OrderRecord]
    at_risk: list[OrderRecord]
    undetermined: list[OrderRecord]
    forecast_order_count: int  # 内示・先行手配(受注日・顧客納期未設定)と思われる件数。参考情報。
    congestion_ranking: list[CongestionEntry] = field(default_factory=list)
    unknown_process_codes: list[str] = field(default_factory=list)
    # 工程コード別の仕入先実績ランキング(--supplier-data指定時のみ)。上位5社まで。
    supplier_ranking_by_process: dict[str, list[SupplierSuggestion]] = field(default_factory=dict)
    # 実績ベース納期充足予測(--process-data指定時のみ)。margin_days昇順(危険な順)。
    feasibility_estimates: list[FeasibilityEntry] = field(default_factory=list)
    # 週別負荷(--process-data指定時のみ)。(週初め, 部署/設備コード, 実績工数合計)のリスト。
    weekly_load_by_department: list[tuple[datetime.date, str, float]] = field(default_factory=list)
    weekly_load_by_machine: list[tuple[datetime.date, str, float]] = field(default_factory=list)


def build_report_data(
    records: list[OrderRecord],
    generated_at: datetime.date,
    process_master: ProcessMaster | None = None,
    supplier_history: SupplierHistory | None = None,
    process_history: ProcessHistory | None = None,
) -> ReportData:
    delayed = sorted(
        (r for r in records if r.judgement == Judgement.DELAYED),
        key=lambda r: r.remaining_business_days,  # 超過が大きい(より負)ほど先頭
    )
    at_risk = sorted(
        (r for r in records if r.judgement == Judgement.AT_RISK),
        key=lambda r: r.remaining_business_days,  # 残日が少ないほど危険 = 先頭
    )
    undetermined = [r for r in records if r.judgement == Judgement.UNDETERMINED]
    forecast_order_count = sum(1 for r in records if r.is_forecast_order)

    congestion_ranking: list[CongestionEntry] = []
    if process_master is not None:
        counts: Counter[str] = Counter()
        for r in records:
            current = r.current_process
            if current is not None:
                counts[current.process_code] += 1

        for code, count in counts.most_common():
            result = process_master.categorize(code)
            congestion_ranking.append(
                CongestionEntry(
                    process_code=code,
                    category=result.category,
                    count=count,
                    is_bottleneck=result.is_bottleneck,
                    is_unknown_code=result.is_unknown_code,
                    standard_lt_business_days=result.standard_lt_business_days,
                )
            )

    if process_history is not None:
        for entry in congestion_ranking:
            entry.actual_lt_calendar_days = process_history.actual_lt_calendar_days_median(entry.process_code)

    supplier_ranking_by_process: dict[str, list[SupplierSuggestion]] = {}
    if supplier_history is not None:
        # ①②(遅延・リスク)の案件についてのみ、現在工程の代替候補仕入先を算出する。
        for r in (*delayed, *at_risk):
            current = r.current_process
            if current is not None:
                r.supplier_suggestions = supplier_history.suggest(r.drawing_no, current.process_code)

        for entry in congestion_ranking:
            suggestions = supplier_history.suggest(None, entry.process_code, top_n=5)
            if suggestions:
                supplier_ranking_by_process[entry.process_code] = suggestions

    feasibility_estimates: list[FeasibilityEntry] = []
    weekly_load_by_department: list[tuple[datetime.date, str, float]] = []
    weekly_load_by_machine: list[tuple[datetime.date, str, float]] = []
    if process_history is not None:
        # ①②(遅延・リスク)の案件についてのみ、現在工程の代替候補設備(社内)を算出する。
        for r in (*delayed, *at_risk):
            current = r.current_process
            if current is not None:
                r.machine_suggestions = process_history.suggest_machines(r.drawing_no, current.process_code)

        feasibility_estimates = build_feasibility_estimates(records, generated_at, process_history)
        weekly_load_by_department = process_history.weekly_load_by_department()
        weekly_load_by_machine = process_history.weekly_load_by_machine()

    return ReportData(
        generated_at=generated_at,
        total_count=len(records),
        delayed=delayed,
        at_risk=at_risk,
        undetermined=undetermined,
        forecast_order_count=forecast_order_count,
        congestion_ranking=congestion_ranking,
        unknown_process_codes=process_master.get_unknown_codes() if process_master is not None else [],
        supplier_ranking_by_process=supplier_ranking_by_process,
        feasibility_estimates=feasibility_estimates,
        weekly_load_by_department=weekly_load_by_department,
        weekly_load_by_machine=weekly_load_by_machine,
    )


def build_feasibility_estimates(
    records: list[OrderRecord],
    generated_at: datetime.date,
    process_history: ProcessHistory,
) -> list[FeasibilityEntry]:
    """進行中の全受注について、図番単位の受注〜完成実績から客先納期充足予測を作る。

    ①②(自社納期・残日ベース)の判定とは独立(kii-san合意)。対象は
    「現在工程があり、かつ顧客納期が分かっている」受注のみ(それ以外は比較対象なしのため対象外)。
    受注日が不明(内示・先行手配案件など)、またはその図番の受注〜完成実績が
    1件もない場合は「データ不足」として予測日を出さない(誤って安全側に見せない)。
    margin_days(顧客納期 - 予測完了日)が小さい(=危険な)順に並べる。データ不足は最後。
    """
    entries: list[FeasibilityEntry] = []
    for r in records:
        current = r.current_process
        if current is None or r.customer_deadline is None:
            continue

        typical_lt = process_history.order_level_lt_calendar_days_median(r.drawing_no)
        sample_count = process_history.order_level_lt_sample_count(r.drawing_no)

        predicted_date: datetime.date | None = None
        data_note: str | None = None
        if typical_lt is None:
            data_note = "この図番の受注〜完成実績が工程累積データに無いため予測不可"
        elif r.order_date is None:
            data_note = "受注日が不明(内示・先行手配案件等)のため予測不可"
        else:
            predicted_date = r.order_date + datetime.timedelta(days=round(typical_lt))

        margin = (r.customer_deadline - predicted_date).days if predicted_date is not None else None

        entries.append(
            FeasibilityEntry(
                order_no=r.order_no,
                drawing_no=r.drawing_no,
                product_name=r.product_name,
                order_date=r.order_date,
                customer_deadline=r.customer_deadline,
                current_process_code=current.process_code,
                remaining_step_count=sum(1 for s in r.processes if not s.is_completed),
                typical_total_lt_calendar_days=typical_lt,
                typical_lt_sample_count=sample_count,
                predicted_completion_date=predicted_date,
                margin_days=margin,
                data_note=data_note,
            )
        )

    entries.sort(key=lambda e: (e.margin_days is None, e.margin_days if e.margin_days is not None else 0))
    return entries
