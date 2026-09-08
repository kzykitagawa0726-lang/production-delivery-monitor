"""Excel/HTML両方の出力が共有するレポート集計データ。"""
from __future__ import annotations

import datetime
from collections import Counter
from dataclasses import dataclass, field

from src.models import Judgement, OrderRecord
from src.process_master import ProcessMaster


@dataclass
class CongestionEntry:
    process_code: str
    category: str
    count: int
    is_bottleneck: bool
    is_unknown_code: bool
    standard_lt_business_days: int


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


def build_report_data(
    records: list[OrderRecord],
    generated_at: datetime.date,
    process_master: ProcessMaster | None = None,
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

    return ReportData(
        generated_at=generated_at,
        total_count=len(records),
        delayed=delayed,
        at_risk=at_risk,
        undetermined=undetermined,
        forecast_order_count=forecast_order_count,
        congestion_ranking=congestion_ranking,
        unknown_process_codes=process_master.get_unknown_codes() if process_master is not None else [],
    )
