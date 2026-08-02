import datetime

from src.business_days import add_business_days, is_business_day, remaining_business_days


def test_is_business_day():
    assert is_business_day(datetime.date(2026, 8, 3))  # Mon
    assert not is_business_day(datetime.date(2026, 8, 1))  # Sat
    assert not is_business_day(datetime.date(2026, 8, 2))  # Sun


def test_remaining_business_days_same_week():
    # 月曜から金曜まで: 火水木金の4営業日
    base = datetime.date(2026, 8, 3)  # Mon
    deadline = datetime.date(2026, 8, 7)  # Fri
    assert remaining_business_days(base, deadline) == 4


def test_remaining_business_days_across_weekend():
    # 金曜から翌週月曜まで: 月曜の1営業日のみ(土日はカウントしない)
    base = datetime.date(2026, 8, 7)  # Fri
    deadline = datetime.date(2026, 8, 10)  # Mon
    assert remaining_business_days(base, deadline) == 1


def test_remaining_business_days_same_day():
    d = datetime.date(2026, 8, 3)
    assert remaining_business_days(d, d) == 0


def test_remaining_business_days_overdue_returns_negative():
    base = datetime.date(2026, 8, 10)  # Mon
    deadline = datetime.date(2026, 8, 7)  # Fri (past)
    assert remaining_business_days(base, deadline) == -1


def test_add_business_days_roundtrip():
    base = datetime.date(2026, 8, 3)  # Mon
    result = add_business_days(base, 4)
    assert result == datetime.date(2026, 8, 7)  # Fri
