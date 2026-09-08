"""内部データモデル。

2026-09-08 kii-san確認済みの実データ(週次エクスポート、2477件)に基づく構造。
2026-08-03時点の旧エクスポートと異なり、以下が確認された:

  - 自社納期(E列2行目)は全件で実データが入っている(未入力プレースホルダーなし)
  - 顧客納期(E列3行目)が存在する(将来拡張として想定していた「目標納期」に相当)
  - 工程進捗(M〜AF列、最大20スロット)が非常にクリーンな形式で存在する
    (スロット1=材料入荷状況の固定文言、スロット2以降=工程コード/加工先/日付/ステータス)

このため、工程進捗データも取り込み対象とする。ただし工程コード→カテゴリの
対応表(PDF由来)がまだ届いていないため、カテゴリ分類・ボトルネック強調は
config/process_code_master.csv が用意され次第有効化する(process_master.py参照)。
"""
from __future__ import annotations

import datetime
import enum
from dataclasses import dataclass, field
from typing import Optional


class Judgement(str, enum.Enum):
    DELAYED = "納期遅延"  # ①
    AT_RISK = "遅延リスクあり"  # ②
    UNDETERMINED = "判定不能・要確認"


# 工程スロットの取り得るステータス(実データで確認された4種のみ)
PROCESS_STATUS_COMPLETED = "作業完了"
PROCESS_STATUS_ORDER_CONFIRMED = "オーダー確定"
PROCESS_STATUS_ORDER_NOT_CONFIRMED = "オーダー確定前"
PROCESS_STATUS_ORDERING = "発注中"


@dataclass
class ProcessStep:
    """工程進捗の1スロット分(スロット2以降。スロット1=材料入荷状況は別扱い)。"""

    slot_index: int  # 2〜20
    process_code: str  # 工程コード(例: CR, CH, HB。工程コード対応表でカテゴリに変換する)
    destination: Optional[str]  # 加工先
    date1: Optional[str]  # 実績日 または 予定日(MM/DD, 年不明のため文字列のまま保持)
    date2: Optional[str]  # 実績日 または 納入日(MM/DD)
    status: str  # 作業完了 / オーダー確定 / オーダー確定前 / 発注中

    @property
    def is_completed(self) -> bool:
        return self.status == PROCESS_STATUS_COMPLETED


@dataclass
class OrderRecord:
    order_no: str  # 製造オーダー№(B列)
    drawing_no: str  # 図番(A列を改行分割)
    product_name: Optional[str]  # 品名(A列を改行分割)
    qty: Optional[float]  # 数量(D列)

    order_date: Optional[datetime.date]  # 受注日(E列1行目)
    company_deadline: Optional[datetime.date]  # 自社納期(E列2行目)
    customer_deadline: Optional[datetime.date]  # 顧客納期(E列3行目)

    remaining_business_days: Optional[int]  # 残日(F列。基幹システム側の計算値をそのまま使用)

    customer_order_no: Optional[str] = None  # 客先注番(C列)
    customer_no: Optional[str] = None  # 得意先NO(G列)

    stagnation_days: Optional[int] = None  # 停滞日(H列。数値の場合のみ)
    stagnation_note: Optional[str] = None  # H列が文字列の場合(例:「前工程実績無し」)

    old_system_no: Optional[str] = None  # 製番(I列。新基幹システム案件はNone=正常)

    sales_note: Optional[str] = None  # MEMO1(J列)が自由記述だった場合の生テキスト
    sales_agreed_deadline: Optional[datetime.date] = None  # MEMO1が客先すり合わせ後の納期(数値)だった場合

    production_note: Optional[str] = None  # MEMO2(K列)。営業/生産管理の自由記載(0は「記載なし」として扱う)
    shipment_plan_bucket: Optional[str] = None  # MEMO3(L列)。月次生産計画表の売上区分(週次出荷予定)

    material_arrived: bool = False  # スロット1の固定文言「材料入荷済み」の有無
    processes: list[ProcessStep] = field(default_factory=list)  # スロット2〜20

    # 仕入(外注)累積データ(--supplier-data)が渡された場合のみ設定される代替候補仕入先。
    # 現在工程(current_process)について、図番一致→工程コード一致の優先順で算出する。
    supplier_suggestions: list = field(default_factory=list)  # list[SupplierSuggestion]

    judgement: Optional[Judgement] = None
    judgement_reason: Optional[str] = None

    # 内示・先行手配(受注日・顧客納期とも未設定)と思われる案件の参考フラグ。
    # kii-san確認済み: ①②の判定からは除外しない(自社納期は実データがあるため通常通り判定)。
    is_forecast_order: bool = False

    @property
    def current_process(self) -> Optional[ProcessStep]:
        """現在停滞している(未完了の)最初の工程。全工程完了ならNone。"""
        for step in self.processes:
            if not step.is_completed:
                return step
        return None

    @property
    def remaining_process_count(self) -> int:
        """現在工程を含む、完了していない工程の残数。"""
        return sum(1 for step in self.processes if not step.is_completed)

    @property
    def deadline_gap_days(self) -> Optional[int]:
        """自社納期と顧客納期のギャップ(自社納期 - 顧客納期)。

        正の値: 自社納期が顧客納期より後(=客先期待より社内予定が遅い。要注意)
        負の値: 自社納期が顧客納期より前(=社内で余裕を見ている)
        """
        if self.company_deadline is None or self.customer_deadline is None:
            return None
        return (self.company_deadline - self.customer_deadline).days
