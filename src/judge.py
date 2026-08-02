"""①納期遅延 / ②納期遅延リスク / 判定不能・要確認 の判定ロジック(叩き台)。

依頼文に記載された叩き台の数式をそのまま実装しているが、以下の点は
実データ確認(Step 1-3)後に調整が必要な暫定仕様であるため、
定数として切り出し、判定理由に明記している。

TODO(実データ確認後に確定させる項目):
  - DEADLINE_PLACEHOLDER: 自社納期の「納期未更新」を示すデフォルト値。
    現時点では依頼文にある "20260101" を仮定しているが、実データでの分布を見て確定する。
  - 全工程数=0 の扱い: 現時点では安全側に倒して「判定不能・要確認」とする
    (依頼文の方針6: 判定に使えない案件は誤って正常に見えないよう判定不能にする、に準拠)。
    実データで比率を確認し、別の推定方法があれば変更する。
  - 古い受注日(内示・先行手配による在庫分)の除外要否は未実装。
    現時点では全件を判定対象とし、除外は行わない(人間の目視判断に委ねる前提)。
"""
from __future__ import annotations

import datetime

from src.business_days import remaining_business_days
from src.models import Judgement, OrderRecord

DAYS_PER_REMAINING_PROCESS = 5  # 叩き台: 残工程数 × 5営業日

# TODO(要実データ確認): 自社納期の「納期未更新」を示す可能性が高いデフォルト値
DEADLINE_PLACEHOLDER = datetime.date(2026, 1, 1)


def calculate_remaining_process_count(total_process_count: int, process_seq: int) -> int:
    return max(total_process_count - process_seq, 0)


def calculate_remaining_required_business_days(remaining_process_count: int, current_process_standard_lt: int) -> int:
    return remaining_process_count * DAYS_PER_REMAINING_PROCESS + current_process_standard_lt


def judge_record(record: OrderRecord, base_date: datetime.date) -> None:
    """record を仕掛中(工程全完了区分=0)の1件として判定し、フィールドを直接更新する。

    工程全完了区分=1(完了)の record を渡すことは想定していない(呼び出し側で事前にフィルタする)。
    """
    record.is_deadline_placeholder = record.company_deadline == DEADLINE_PLACEHOLDER

    if record.company_deadline is None or record.is_deadline_placeholder:
        record.judgement = Judgement.UNDETERMINED
        record.judgement_reason = "自社納期が未設定、または納期未更新の可能性があるデフォルト値のため判定不能"
        return

    if record.total_process_count == 0:
        record.judgement = Judgement.UNDETERMINED
        record.judgement_reason = "全工程数が0のため工程未展開の可能性があり判定不能(暫定仕様: 要実データ確認)"
        return

    if record.current_process_standard_lt is None:
        record.judgement = Judgement.UNDETERMINED
        record.judgement_reason = "工程コードのマスタ照合に失敗したため判定不能"
        return

    record.remaining_process_count = calculate_remaining_process_count(
        record.total_process_count, record.process_seq
    )
    record.remaining_required_business_days = calculate_remaining_required_business_days(
        record.remaining_process_count, record.current_process_standard_lt
    )
    record.remaining_business_days_to_deadline = remaining_business_days(base_date, record.company_deadline)

    if record.remaining_business_days_to_deadline < 0:
        record.judgement = Judgement.DELAYED
        record.judgement_reason = f"自社納期を{-record.remaining_business_days_to_deadline}営業日超過"
    elif record.remaining_required_business_days > record.remaining_business_days_to_deadline:
        record.judgement = Judgement.AT_RISK
        record.judgement_reason = (
            f"残必要日数{record.remaining_required_business_days}営業日 > "
            f"残営業日{record.remaining_business_days_to_deadline}営業日"
        )
    else:
        record.judgement = None
        record.judgement_reason = None
