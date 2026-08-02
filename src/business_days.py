"""営業日数の計算。祝日は考慮せず、土日のみを除外する(叩き台の方針に準拠)。"""
from __future__ import annotations

import datetime


def is_business_day(d: datetime.date) -> bool:
    return d.weekday() < 5  # 0=Mon ... 4=Fri


def remaining_business_days(base_date: datetime.date, deadline: datetime.date) -> int:
    """base_date から deadline までの営業日数(土日を除く)。

    deadline が base_date より前の場合は負の値を返す(超過日数として利用可能)。
    base_date, deadline 自体はカウント対象日数に含めない(区間の営業日数)。
    """
    if deadline >= base_date:
        count = 0
        d = base_date
        while d < deadline:
            d += datetime.timedelta(days=1)
            if is_business_day(d):
                count += 1
        return count
    else:
        # 超過している場合は負数(超過営業日数)を返す
        return -remaining_business_days(deadline, base_date)


def add_business_days(base_date: datetime.date, num_days: int) -> datetime.date:
    """base_date から num_days 営業日後の日付を返す(土日を除く)。"""
    d = base_date
    remaining = num_days
    step = 1 if remaining >= 0 else -1
    remaining = abs(remaining)
    while remaining > 0:
        d += datetime.timedelta(days=step)
        if is_business_day(d):
            remaining -= 1
    return d
