"""①納期遅延 / ②納期遅延リスク / 判定不能・要確認 の判定。

自社納期までの残営業日数は、独自計算せず基幹システムが算出した「残日」列を
そのまま用いる(kii-san了承済み: 2026-08-03のやり取りを継続)。

2026-09-08 kii-san確認済みの実データで、自社納期の入力誤りと思われる
極端な外れ値(残日が数千営業日規模、未来方向)が確認されたため、上限を超える場合は
「判定不能・要確認」に振り分ける(誤って①②から漏れる/正常に見えるのを防ぐ)。

同日、①納期遅延リストの最上位が「安全在庫」案件(自社納期が数年前の固定値、
出荷計画区分=在庫対象)で占められている問題も見つかった。実際の客先納期ではなく
恒常的な在庫補充用の参照日のため、①に混ざると本当に急ぐべき案件が埋もれる。
kii-san確認済み: 過去方向も同じ閾値で判定不能とし、理由に「自社納期の設定変更が
必要」である旨を明記する(単なる「判定不能」で終わらせず、対応アクションが
分かるようにする)。

内示・先行手配案件(受注日・顧客納期が未設定)は①②の対象から除外しない
(kii-san確認済み: 自社納期・残日は実データが入っているため通常通り判定する)。
"""
from __future__ import annotations

from src.models import Judgement, OrderRecord

RISK_THRESHOLD_BUSINESS_DAYS = 5  # 残り5営業日以内は②遅延リスクあり(kii-san確定、2026-08-03継続)
ANOMALY_THRESHOLD_BUSINESS_DAYS = 500  # 約2年分。残日の絶対値がこれを超えたら自社納期の設定に問題ありとみなす


def judge_record(record: OrderRecord) -> None:
    if record.company_deadline is None or record.remaining_business_days is None:
        record.judgement = Judgement.UNDETERMINED
        record.judgement_reason = "自社納期または残日数を取得できないため判定不能"
        return

    remaining = record.remaining_business_days

    if remaining > ANOMALY_THRESHOLD_BUSINESS_DAYS:
        record.judgement = Judgement.UNDETERMINED
        record.judgement_reason = (
            f"残日が{remaining}営業日と異常に大きく、自社納期の設定誤り(誤入力)の疑いがあるため判定不能"
        )
        return

    if remaining < -ANOMALY_THRESHOLD_BUSINESS_DAYS:
        record.judgement = Judgement.UNDETERMINED
        record.judgement_reason = (
            f"自社納期を{-remaining}営業日(約2年以上)超過しており、安全在庫等の固定参照日の疑いあり。"
            "①遅延として扱わず、自社納期の設定変更が必要か確認してください"
        )
        return

    if remaining < 0:
        record.judgement = Judgement.DELAYED
        record.judgement_reason = f"自社納期を{-remaining}営業日超過"
    elif remaining <= RISK_THRESHOLD_BUSINESS_DAYS:
        record.judgement = Judgement.AT_RISK
        record.judgement_reason = f"残り{remaining}営業日(しきい値{RISK_THRESHOLD_BUSINESS_DAYS}営業日以下)"
    else:
        record.judgement = None
        record.judgement_reason = None
