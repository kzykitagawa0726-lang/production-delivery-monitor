"""Excel/HTML両方の出力が共有するレポート集計データ。"""
from __future__ import annotations

import datetime
from dataclasses import dataclass, field

from src.aggregate import CongestionEntry, congestion_ranking
from src.models import Judgement, OrderRecord


@dataclass
class ReportData:
    generated_at: datetime.date
    total_count: int
    delayed: list[OrderRecord]
    at_risk: list[OrderRecord]
    undetermined: list[OrderRecord]
    congestion: list[CongestionEntry]
    unknown_process_codes: list[str] = field(default_factory=list)


def build_report_data(
    records: list[OrderRecord],
    generated_at: datetime.date,
    unknown_process_codes: list[str] | None = None,
) -> ReportData:
    delayed = sorted(
        (r for r in records if r.judgement == Judgement.DELAYED),
        key=lambda r: r.remaining_business_days_to_deadline,  # 最も負(超過大)が先頭
    )
    at_risk = sorted(
        (r for r in records if r.judgement == Judgement.AT_RISK),
        key=lambda r: (r.remaining_required_business_days - r.remaining_business_days_to_deadline),
        reverse=True,
    )
    undetermined = [r for r in records if r.judgement == Judgement.UNDETERMINED]
    congestion = congestion_ranking(records)

    return ReportData(
        generated_at=generated_at,
        total_count=len(records),
        delayed=delayed,
        at_risk=at_risk,
        undetermined=undetermined,
        congestion=congestion,
        unknown_process_codes=unknown_process_codes or [],
    )
