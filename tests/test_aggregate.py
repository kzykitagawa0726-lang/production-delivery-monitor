import datetime

from src.aggregate import congestion_ranking
from src.models import OrderRecord


def make_record(process_code, category, is_bottleneck=False):
    return OrderRecord(
        order_no="ORD-1",
        drawing_no="DWG-1",
        qty=1,
        order_date=datetime.date(2026, 1, 1),
        company_deadline=datetime.date(2026, 12, 1),
        total_process_count=3,
        process_seq=1,
        process_code=process_code,
        all_process_complete_flag=0,
        process_category=category,
        is_bottleneck_process=is_bottleneck,
    )


def test_congestion_ranking_counts_and_sorts_desc():
    records = [
        make_record("G02", "歯切り系", is_bottleneck=True),
        make_record("G02", "歯切り系", is_bottleneck=True),
        make_record("H01", "熱処理系"),
    ]
    ranking = congestion_ranking(records)
    assert ranking[0]["process_code"] == "G02"
    assert ranking[0]["count"] == 2
    assert ranking[0]["is_bottleneck"] is True
    assert ranking[1]["process_code"] == "H01"
    assert ranking[1]["count"] == 1
