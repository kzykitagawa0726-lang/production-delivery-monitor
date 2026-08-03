"""①納期遅延 / ②納期遅延リスク / 判定不能・要確認 の判定。

工程進捗列(旧基幹システムの状況欄)は移管期間中で信頼できないため使用しない。
自社納期までの残営業日数は、独自計算せず基幹システムが算出した「残日」列を
そのまま用いる(kii-san了承済み: 2026-08-03のやり取り)。
"""
from __future__ import annotations

from src.models import Judgement, OrderRecord

RISK_THRESHOLD_BUSINESS_DAYS = 5  # 残り5営業日以内は②遅延リスクあり(kii-san確定)


def judge_record(record: OrderRecord) -> None:
    if record.company_deadline is None or record.remaining_business_days is None:
        record.judgement = Judgement.UNDETERMINED
        record.judgement_reason = "自社納期または残日数を取得できないため判定不能"
        return

    remaining = record.remaining_business_days
    if remaining < 0:
        record.judgement = Judgement.DELAYED
        record.judgement_reason = f"自社納期を{-remaining}営業日超過"
    elif remaining <= RISK_THRESHOLD_BUSINESS_DAYS:
        record.judgement = Judgement.AT_RISK
        record.judgement_reason = f"残り{remaining}営業日(しきい値{RISK_THRESHOLD_BUSINESS_DAYS}営業日以下)"
    else:
        record.judgement = None
        record.judgement_reason = None
