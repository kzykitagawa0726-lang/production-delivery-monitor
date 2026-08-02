import datetime

from src.judge import DEADLINE_PLACEHOLDER, judge_record
from src.models import Judgement, OrderRecord

BASE_DATE = datetime.date(2026, 8, 3)  # Mon


def make_record(**overrides) -> OrderRecord:
    defaults = dict(
        order_no="ORD-0001",
        drawing_no="DWG-0001",
        qty=10,
        order_date=datetime.date(2026, 6, 1),
        company_deadline=datetime.date(2026, 9, 1),
        total_process_count=5,
        process_seq=2,
        process_code="SAMPLE_G01",
        all_process_complete_flag=0,
        current_process_standard_lt=5,
    )
    defaults.update(overrides)
    return OrderRecord(**defaults)


def test_delayed_when_past_deadline():
    record = make_record(company_deadline=datetime.date(2026, 7, 1))
    judge_record(record, BASE_DATE)
    assert record.judgement == Judgement.DELAYED
    assert record.remaining_business_days_to_deadline < 0


def test_at_risk_when_insufficient_time():
    # 残工程数=3(5-2), 残必要日数=3*5+5=20営業日。納期まで数営業日しかない場合はリスクあり
    record = make_record(company_deadline=datetime.date(2026, 8, 6))  # Thu, 3 business days out
    judge_record(record, BASE_DATE)
    assert record.judgement == Judgement.AT_RISK


def test_normal_when_enough_time():
    record = make_record(company_deadline=datetime.date(2026, 12, 1))
    judge_record(record, BASE_DATE)
    assert record.judgement is None


def test_undetermined_when_deadline_is_placeholder():
    record = make_record(company_deadline=DEADLINE_PLACEHOLDER)
    judge_record(record, BASE_DATE)
    assert record.judgement == Judgement.UNDETERMINED
    assert record.is_deadline_placeholder is True


def test_undetermined_when_total_process_count_zero():
    record = make_record(total_process_count=0)
    judge_record(record, BASE_DATE)
    assert record.judgement == Judgement.UNDETERMINED


def test_undetermined_when_deadline_missing():
    record = make_record(company_deadline=None)
    judge_record(record, BASE_DATE)
    assert record.judgement == Judgement.UNDETERMINED
