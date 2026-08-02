"""工程別の仕掛件数集計(混雑ランキング)。"""
from __future__ import annotations

from collections import defaultdict
from typing import TypedDict

from src.models import OrderRecord


class CongestionEntry(TypedDict):
    process_code: str
    category: str | None
    count: int
    is_bottleneck: bool


def congestion_ranking(records: list[OrderRecord]) -> list[CongestionEntry]:
    """仕掛中の records を工程コード別に集計し、件数降順で返す。"""
    counter: dict[str, CongestionEntry] = defaultdict(
        lambda: {"process_code": "", "category": None, "count": 0, "is_bottleneck": False}
    )
    for r in records:
        entry = counter[r.process_code]
        entry["process_code"] = r.process_code
        entry["category"] = r.process_category
        entry["count"] += 1
        entry["is_bottleneck"] = entry["is_bottleneck"] or r.is_bottleneck_process

    ranking = list(counter.values())
    ranking.sort(key=lambda e: e["count"], reverse=True)
    return ranking
