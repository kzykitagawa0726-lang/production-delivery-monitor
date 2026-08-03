"""内部データモデル。

実データ確認の結果、現行エクスポート(移管期間中の基幹システム)では
工程進捗が信頼できる形で列に残っていないため、A〜H列の受注情報のみを
対象とする。将来「目標納期」列が追加された場合に備え、target_deadline
だけは拡張用に用意してあるが、ギャップ分析ロジックは未実装(現時点では不要)。
"""
from __future__ import annotations

import datetime
import enum
from dataclasses import dataclass
from typing import Optional


class Judgement(str, enum.Enum):
    DELAYED = "納期遅延"  # ①
    AT_RISK = "遅延リスクあり"  # ②
    UNDETERMINED = "判定不能・要確認"


@dataclass
class OrderRecord:
    order_no: str  # 製造オーダー№(B列, 現行基幹システムキー)
    drawing_no: str  # 図番(A列を改行分割)
    product_name: Optional[str]  # 品名(A列を改行分割)
    qty: Optional[float]  # 数量(D列)
    company_deadline: Optional[datetime.date]  # 自社納期(E列1行目)
    remaining_business_days: Optional[int]  # 残日(F列。基幹システム側の計算値をそのまま使用)
    customer_order_no: Optional[str] = None  # 客先注番(C列)
    customer_no: Optional[str] = None  # 取引先NO(G列)
    old_system_no: Optional[str] = None  # 製番(H列, 旧基幹システムキー)

    # 将来拡張: 目標納期(顧客への回答納期)。現行エクスポートには存在しない。
    target_deadline: Optional[datetime.date] = None

    judgement: Optional[Judgement] = None
    judgement_reason: Optional[str] = None
