import datetime

from src.judge import RISK_THRESHOLD_BUSINESS_DAYS, judge_record
from src.models import Judgement, OrderRecord


def make_record(**overrides) -> OrderRecord:
    defaults = dict(
        order_no="PO-000001",
        drawing_no="312127272601",
        product_name="SPB M10x13T L",
        qty=1,
        company_deadline=datetime.date(2026, 9, 1),
        remaining_business_days=20,
    )
    defaults.update(overrides)
    return OrderRecord(**defaults)


def test_delayed_when_remaining_days_negative():
    record = make_record(remaining_business_days=-5)
    judge_record(record)
    assert record.judgement == Judgement.DELAYED
    assert "5営業日超過" in record.judgement_reason


def test_at_risk_when_within_threshold():
    record = make_record(remaining_business_days=RISK_THRESHOLD_BUSINESS_DAYS)
    judge_record(record)
    assert record.judgement == Judgement.AT_RISK


def test_normal_when_beyond_threshold():
    record = make_record(remaining_business_days=RISK_THRESHOLD_BUSINESS_DAYS + 1)
    judge_record(record)
    assert record.judgement is None


def test_undetermined_when_deadline_missing():
    record = make_record(company_deadline=None)
    judge_record(record)
    assert record.judgement == Judgement.UNDETERMINED


def test_undetermined_when_remaining_days_missing():
    record = make_record(remaining_business_days=None)
    judge_record(record)
    assert record.judgement == Judgement.UNDETERMINED


def test_boundary_zero_is_at_risk():
    record = make_record(remaining_business_days=0)
    judge_record(record)
    assert record.judgement == Judgement.AT_RISK
