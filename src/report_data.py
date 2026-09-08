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

    if process_history is not None:
        # ①②(遅延・リスク)の案件についてのみ、現在工程の代替候補設備(社内)を算出する。
        for r in (*delayed, *at_risk):
            current = r.current_process
            if current is not None:
                r.machine_suggestions = process_history.suggest_machines(r.drawing_no, current.process_code)

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
    )
