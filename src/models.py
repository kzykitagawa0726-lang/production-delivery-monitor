"""内部データモデル。

生データ(基幹システムからのExcel)の列名や表記ゆれから切り離した正規化済みの
案件(受注データの1工程分の行)表現。Excel列名 -> OrderRecord へのマッピングは
実データの列名確認後に実装する src/excel_io.py が担当する。
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
    # --- 生データ由来の項目 ---
    order_no: str  # 受注№
    drawing_no: str  # 図番
    qty: Optional[float]  # 数量
    order_date: Optional[datetime.date]  # 受注日
    company_deadline: Optional[datetime.date]  # 自社納期
    total_process_count: int  # 全工程数
    process_seq: int  # 工程順
    process_code: str  # 工程コード
    all_process_complete_flag: int  # 工程全完了区分(1=完了, 0=仕掛中)
    assignee: Optional[str] = None  # 担当者
    product_name: Optional[str] = None  # 商品名

    # 将来拡張: 目標納期(顧客への回答納期)。現行エクスポートには存在しない想定。
    target_deadline: Optional[datetime.date] = None

    # --- 工程マスタ照合結果(process_master.ProcessMasterで付与) ---
    process_category: Optional[str] = None
    current_process_standard_lt: Optional[int] = None  # 当該工程の標準LT(営業日)
    is_unknown_process_code: bool = False
    is_bottleneck_process: bool = False

    # --- 判定結果(judge.pyで付与) ---
    is_deadline_placeholder: bool = False  # 自社納期が「納期未更新」を示すデフォルト値と推定される場合
    remaining_process_count: Optional[int] = None
    remaining_required_business_days: Optional[int] = None
    remaining_business_days_to_deadline: Optional[int] = None
    judgement: Optional[Judgement] = None
    judgement_reason: Optional[str] = None
