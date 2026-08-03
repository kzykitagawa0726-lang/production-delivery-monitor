"""Excel/HTML両方の出力が共有するレポート集計データ。"""
from __future__ import annotations

import datetime
from dataclasses import dataclass

from src.models import Judgement, OrderRecord


@dataclass
class ReportData:
    generated_at: datetime.date
    total_count: int
    delayed: list[OrderRecord]
    at_risk: list[OrderRecord]
    undetermined: list[OrderRecord]


def build_report_data(records: list[OrderRecord], generated_at: datetime.date) -> ReportData:
    delayed = sorted(
        (r for r in records if r.judgement == Judgement.DELAYED),
        key=lambda r: r.remaining_business_days,  # 超過が大きい(より負)ほど先頭
    )
    at_risk = sorted(
        (r for r in records if r.judgement == Judgement.AT_RISK),
        key=lambda r: r.remaining_business_days,  # 残日が少ないほど危険 = 先頭
    )
    undetermined = [r for r in records if r.judgement == Judgement.UNDETERMINED]

    return ReportData(
        generated_at=generated_at,
        total_count=len(records),
        delayed=delayed,
        at_risk=at_risk,
        undetermined=undetermined,
    )
