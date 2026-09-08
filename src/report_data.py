"""Excel/HTML両方の出力が共有するレポート集計データ。"""
from __future__ import annotations

import datetime
import statistics
from collections import Counter, defaultdict
from dataclasses import dataclass, field

from src.models import Judgement, OrderRecord
from src.process_history import ProcessHistory
from src.process_master import ProcessMaster
from src.product_category import ProductCategoryClassifier
from src.supplier_history import SupplierHistory, SupplierSuggestion


@dataclass
class CongestionEntry:
    process_code: str
    category: str
    count: int
    is_bottleneck: bool
    is_unknown_code: bool
    standard_lt_business_days: int
    # 実績リードタイム中央値(暦日、着手日〜完成日)。--process-data指定時のみ、参考情報として設定される。
    # 標準LT(標準リードタイム、営業日)は変更しない(kii-san指示)。判定ロジックにも使わない。
    actual_lt_calendar_days: float | None = None


@dataclass
class FeasibilityEntry:
    """進行中の受注について、図番単位の「受注〜完成」実績日数から客先納期充足を予測する。

    ①②の判定(自社納期・残日ベース)とは独立した、追加の参考情報。
    --process-data指定時のみ算出される。

    【2026-09-08 実データ検証で設計変更】当初は残り工程ごとの実績LT(暦日)を
    単純合計していたが、各工程の実績日数には他案件との待ち時間が相当含まれており、
    残り工程が多い受注では合計が数千日規模になる異常値が実データで頻発した。
    → kii-san合意により、図番単位の「受注(最初の着手)〜完成(最後の完成)」実績日数
    (ProcessHistory.order_level_lt_calendar_days_median)を使う設計に変更。
    自社の受注日(order_date)を起点に、その図番が過去に実際どれだけかかったかの
    中央値を足して予測完了日とする(残り工程数による比例配分はしない。シンプルさを優先)。
    """

    order_no: str
    drawing_no: str
    product_name: str | None
    order_date: datetime.date | None
    customer_deadline: datetime.date
    current_process_code: str
    remaining_step_count: int
    typical_total_lt_calendar_days: float | None  # 図番単位の受注〜完成 実績中央値(暦日)
    typical_lt_sample_count: int  # 何件の過去実績から算出したか
    predicted_completion_date: datetime.date | None  # 受注日 + typical_total_lt
    margin_days: int | None  # 顧客納期 - 予測完了日。正=間に合う見込み、負=間に合わない見込み
    data_note: str | None = None  # 予測不可の理由(データ不足の場合)

    @property
    def status(self) -> str:
        if self.predicted_completion_date is None:
            return "データ不足"
        if self.margin_days is not None and self.margin_days < 0:
            return "間に合わない見込み"
        return "間に合う見込み"


@dataclass
class MonthlyCategoryCapacityEntry:
    """品種(GEAR/BEVEL/WORM)別・月別のキャパシティ参考資料。

    【2026-09-08 kii-san要望】「弊社の稼働率がだいたい30～40%くらい(24時間を100%とした場合)。
    その観点でキャパシティが足りているか、GEAR/BEVEL/WORMの3品種で月単位に見たい。営業として
    ざっくり月のキャパシティ感を、受注時の参考にしたい」を受けて追加。①②の判定や実績LTとは
    独立した、追加の参考情報(--process-data指定時のみ算出)。

    キャパシティ上限の考え方:
    実際の設備上限(24時間×稼働率)そのものはこのツールからは分からないため、代わりに
    「過去の典型的な実績工数は、稼働率30～40%の状態で出ている実績のはずだ」という
    kii-san申告の稼働率を逆算に使う。品種別の直近実績月(算出時点の月は未確定のため除外)
    最大12か月分の平均工数を「典型実績」とし、
        キャパシティ = 典型実績 ÷ 稼働率
    でキャパシティ上限を逆算する(稼働率が低いほどキャパシティは大きく出る)。
    30%〜40%の幅をそのまま「保守的(稼働率40%と仮定)〜楽観的(稼働率30%と仮定)」の
    レンジとして示し、中間の35%を「ざっくりの目安」の見出し数値として使う。
    あくまで申告の稼働率レンジからの逆算であり、設備の物理的な上限を測定したものではない
    (営業のおおまかな参考資料という位置づけ)。

    月間キャパシティ目安は「1営業日あたり」にも換算する(2026-09-08 kii-san補足:
    「弊社の平均営業日は20日」)。月ごとの実際の営業日数の違い(祝日・盆休み・年末年始等)は
    考慮しない単純な換算(キャパシティ目安 ÷ 20日)で、感覚をつかみやすくするための参考値。

    予測工数の考え方:
    build_capacity_forecast()と同じ方式(標準LTで日程を、実績平均工数で大きさを見積もる)で、
    仕掛中の各受注の残り工程を先の月へ積み上げる。ただし配分先は部署/設備ではなく、
    その受注自身の品名(product_name)から分類した品種(GEAR/BEVEL/WORM)。
    品名から品種を判定できない受注(付随部品・型番のみの表記など)は3分類の対象外として除外する
    (実績側の集計方針と揃えている)。
    """

    month: str  # "YYYY-MM"
    category: str  # GEAR / BEVEL / WORM
    actual_hours: float | None  # その月の実績工数合計(品種分類できた分のみ)
    forecast_hours: float | None  # 仕掛中受注の残り工程を積み上げた予測工数合計
    capacity_hours_low: float | None  # 稼働率40%と仮定した場合のキャパシティ上限(保守的)
    capacity_hours_high: float | None  # 稼働率30%と仮定した場合のキャパシティ上限(楽観的)
    capacity_hours_typical: float | None  # 稼働率35%と仮定した場合の目安(headline数値)
    capacity_hours_typical_per_business_day: float | None  # 上記を月平均営業日20日で割った1営業日あたりの目安
    fulfillment_rate: float | None  # 予測工数 ÷ キャパシティ目安。1.0超で目安を超過


@dataclass
class MonthlyCategoryCapacityDetailEntry:
    """品種別月間キャパシティ予測(forecast_hours)の、製造オーダー単位の内訳。

    【2026-09-08 kii-san要望】「月ごとの品種別キャパシティを、負荷分散をする必要があるので
    製造オーダーに降りて確認できる様にしたい」を受けて追加。build_monthly_category_capacityの
    forecast_hoursの元になった、各受注・各残り工程の見込み工数を1行ずつ並べたもの。
    実績側(過去の完了済み工数)は集計元データが受注単位で追えないため、内訳は予測側のみ提供する。
    """

    month: str  # "YYYY-MM"
    category: str  # GEAR / BEVEL / WORM
    order_no: str
    drawing_no: str
    product_name: str | None
    process_code: str
    forecast_hours: float


@dataclass
class MachineForecastDetailEntry:
    """週別予測負荷(設備別)の、製造オーダー単位の内訳。

    【2026-09-08 kii-san要望】「調整必要な設備を週単位で、内訳は製造オーダーで確認したい」を
    受けて追加。build_capacity_forecastの機械別集計の元になった、各受注・各残り工程の
    按分結果を1行ずつ並べたもの。1つの工程は過去の実績シェアに応じて複数の設備に按分される
    設計(1台に決め打ちしない)のため、同じ受注・工程が複数の設備にまたがって複数行に
    現れうる(kii-san合意: 按分候補を全部表示する方針)。
    """

    week: datetime.date
    machine_code: str
    order_no: str
    drawing_no: str
    product_name: str | None
    process_code: str
    hours: float  # 按分後の見込み工数


@dataclass
class SupplierForecastEntry:
    """仕入先(外注)週別予測。①②の判定・設備の工数ベース予測とは別の切り口。

    【2026-09-08 kii-san要望】「調整必要な仕入れ先を週単位で確認したい」を受けて追加。
    仕入累積データには工数(時間)の記録が無いため、金額(仕入先ごとの週次合計のみ。
    個別受注・仕入先の金額は一切表示しない)を使う。また社内設備のような稼働率基準
    (30〜40%)が無く絶対的な上限は推定できないため、「普段の実績水準
    (直近SUPPLIER_WEEKLY_BASELINE_WEEKS週平均)と比べて、予測される依頼金額が
    どれだけ多い/少ないか」という相対比較で示す(kii-san合意)。

    過去に外注実績が一切ない工程コードは、按分先(仕入先シェア)が存在しないため
    自然に対象外となる(=社内設備側の予測のみに現れる)。
    """

    week: datetime.date
    supplier_code: str
    forecast_amount: float  # 仕掛中受注の残り工程を標準LTで積み上げた予測金額(按分後合計)
    typical_weekly_amount: float | None  # 直近実績の週平均金額(算出時点の週は未確定のため除く)
    ratio_to_typical: float | None  # forecast_amount ÷ typical_weekly_amount


@dataclass
class SupplierForecastDetailEntry:
    """仕入先週別予測の、製造オーダー単位の内訳。"""

    week: datetime.date
    supplier_code: str
    order_no: str
    drawing_no: str
    product_name: str | None
    process_code: str
    forecast_amount: float


@dataclass
class ReportData:
    generated_at: datetime.date
    total_count: int
    delayed: list[OrderRecord]
    at_risk: list[OrderRecord]
    undetermined: list[OrderRecord]
    forecast_order_count: int  # 内示・先行手配(受注日・顧客納期未設定)と思われる件数。参考情報。
    congestion_ranking: list[CongestionEntry] = field(default_factory=list)
    unknown_process_codes: list[str] = field(default_factory=list)
    # 工程コード別の仕入先実績ランキング(--supplier-data指定時のみ)。上位5社まで。
    supplier_ranking_by_process: dict[str, list[SupplierSuggestion]] = field(default_factory=dict)
    # 実績ベース納期充足予測(--process-data指定時のみ)。margin_days昇順(危険な順)。
    feasibility_estimates: list[FeasibilityEntry] = field(default_factory=list)
    # 週別負荷・実績(--process-data指定時のみ)。(週初め, 部署/設備コード, 実績工数合計)のリスト。
    weekly_load_by_department: list[tuple[datetime.date, str, float]] = field(default_factory=list)
    weekly_load_by_machine: list[tuple[datetime.date, str, float]] = field(default_factory=list)
    # 週別予測負荷(--process-data指定時のみ)。仕掛中の残り工程を標準LTで先の週へ積み上げた将来予測。
    # (週初め, 部署/設備コード, 予測工数合計)のリスト。
    capacity_forecast_by_department: list[tuple[datetime.date, str, float]] = field(default_factory=list)
    capacity_forecast_by_machine: list[tuple[datetime.date, str, float]] = field(default_factory=list)
    # 品種別(GEAR/BEVEL/WORM)月間キャパシティ(--process-data指定時のみ)。営業向けのざっくり参考資料。
    monthly_category_capacity: list[MonthlyCategoryCapacityEntry] = field(default_factory=list)
    # 品種別月間キャパシティ予測の、製造オーダー単位の内訳(--process-data指定時のみ)。
    monthly_category_capacity_detail: list[MonthlyCategoryCapacityDetailEntry] = field(default_factory=list)
    # 週別予測負荷(設備別)の、製造オーダー単位の内訳(--process-data指定時のみ)。
    capacity_forecast_machine_detail: list[MachineForecastDetailEntry] = field(default_factory=list)
    # 仕入先(外注)週別予測(--supplier-data と --process-data の両方指定時のみ)。相対比較(倍率)。
    supplier_capacity_forecast: list[SupplierForecastEntry] = field(default_factory=list)
    # 仕入先週別予測の、製造オーダー単位の内訳。
    supplier_capacity_forecast_detail: list[SupplierForecastDetailEntry] = field(default_factory=list)


def build_report_data(
    records: list[OrderRecord],
    generated_at: datetime.date,
    process_master: ProcessMaster | None = None,
    supplier_history: SupplierHistory | None = None,
    process_history: ProcessHistory | None = None,
    product_category_classifier: ProductCategoryClassifier | None = None,
) -> ReportData:
    delayed = sorted(
        (r for r in records if r.judgement == Judgement.DELAYED),
        key=lambda r: r.remaining_business_days,  # 超過が大きい(より負)ほど先頭
    )
    at_risk = sorted(
        (r for r in records if r.judgement == Judgement.AT_RISK),
        key=lambda r: r.remaining_business_days,  # 残日が少ないほど危険 = 先頭
    )
    undetermined = [r for r in records if r.judgement == Judgement.UNDETERMINED]
    forecast_order_count = sum(1 for r in records if r.is_forecast_order)

    congestion_ranking: list[CongestionEntry] = []
    if process_master is not None:
        counts: Counter[str] = Counter()
        for r in records:
            current = r.current_process
            if current is not None:
                counts[current.process_code] += 1

        for code, count in counts.most_common():
            result = process_master.categorize(code)
            congestion_ranking.append(
                CongestionEntry(
                    process_code=code,
                    category=result.category,
                    count=count,
                    is_bottleneck=result.is_bottleneck,
                    is_unknown_code=result.is_unknown_code,
                    standard_lt_business_days=result.standard_lt_business_days,
                )
            )

    if process_history is not None:
        for entry in congestion_ranking:
            entry.actual_lt_calendar_days = process_history.actual_lt_calendar_days_median(entry.process_code)

    supplier_ranking_by_process: dict[str, list[SupplierSuggestion]] = {}
    if supplier_history is not None:
        # ①②(遅延・リスク)の案件についてのみ、現在工程の代替候補仕入先を算出する。
        for r in (*delayed, *at_risk):
            current = r.current_process
            if current is not None:
                r.supplier_suggestions = supplier_history.suggest(r.drawing_no, current.process_code)

        for entry in congestion_ranking:
            suggestions = supplier_history.suggest(None, entry.process_code, top_n=5)
            if suggestions:
                supplier_ranking_by_process[entry.process_code] = suggestions

    feasibility_estimates: list[FeasibilityEntry] = []
    weekly_load_by_department: list[tuple[datetime.date, str, float]] = []
    weekly_load_by_machine: list[tuple[datetime.date, str, float]] = []
    capacity_forecast_by_department: list[tuple[datetime.date, str, float]] = []
    capacity_forecast_by_machine: list[tuple[datetime.date, str, float]] = []
    capacity_forecast_machine_detail: list[MachineForecastDetailEntry] = []
    monthly_category_capacity: list[MonthlyCategoryCapacityEntry] = []
    monthly_category_capacity_detail: list[MonthlyCategoryCapacityDetailEntry] = []
    supplier_capacity_forecast: list[SupplierForecastEntry] = []
    supplier_capacity_forecast_detail: list[SupplierForecastDetailEntry] = []
    if process_history is not None:
        # ①②(遅延・リスク)の案件についてのみ、現在工程の代替候補設備(社内)を算出する。
        for r in (*delayed, *at_risk):
            current = r.current_process
            if current is not None:
                r.machine_suggestions = process_history.suggest_machines(r.drawing_no, current.process_code)

        feasibility_estimates = build_feasibility_estimates(records, generated_at, process_history)
        weekly_load_by_department = process_history.weekly_load_by_department()
        weekly_load_by_machine = process_history.weekly_load_by_machine()

        if process_master is not None:
            capacity_forecast_by_department, capacity_forecast_by_machine = build_capacity_forecast(
                records, generated_at, process_master, process_history
            )
            capacity_forecast_machine_detail = build_capacity_forecast_machine_detail(
                records, generated_at, process_master, process_history
            )

            if product_category_classifier is not None:
                monthly_category_capacity = build_monthly_category_capacity(
                    records, generated_at, process_master, process_history, product_category_classifier
                )
                monthly_category_capacity_detail = build_monthly_category_capacity_detail(
                    records, generated_at, process_master, process_history, product_category_classifier
                )

            if supplier_history is not None:
                supplier_capacity_forecast, supplier_capacity_forecast_detail = build_supplier_capacity_forecast(
                    records, generated_at, process_master, process_history, supplier_history
                )

    return ReportData(
        generated_at=generated_at,
        total_count=len(records),
        delayed=delayed,
        at_risk=at_risk,
        undetermined=undetermined,
        forecast_order_count=forecast_order_count,
        congestion_ranking=congestion_ranking,
        unknown_process_codes=process_master.get_unknown_codes() if process_master is not None else [],
        supplier_ranking_by_process=supplier_ranking_by_process,
        feasibility_estimates=feasibility_estimates,
        weekly_load_by_department=weekly_load_by_department,
        weekly_load_by_machine=weekly_load_by_machine,
        capacity_forecast_by_department=capacity_forecast_by_department,
        capacity_forecast_by_machine=capacity_forecast_by_machine,
        capacity_forecast_machine_detail=capacity_forecast_machine_detail,
        monthly_category_capacity=monthly_category_capacity,
        monthly_category_capacity_detail=monthly_category_capacity_detail,
        supplier_capacity_forecast=supplier_capacity_forecast,
        supplier_capacity_forecast_detail=supplier_capacity_forecast_detail,
    )


def build_feasibility_estimates(
    records: list[OrderRecord],
    generated_at: datetime.date,
    process_history: ProcessHistory,
) -> list[FeasibilityEntry]:
    """進行中の全受注について、図番単位の受注〜完成実績から客先納期充足予測を作る。

    ①②(自社納期・残日ベース)の判定とは独立(kii-san合意)。対象は
    「現在工程があり、かつ顧客納期が分かっている」受注のみ(それ以外は比較対象なしのため対象外)。
    受注日が不明(内示・先行手配案件など)、またはその図番の受注〜完成実績が
    1件もない場合は「データ不足」として予測日を出さない(誤って安全側に見せない)。
    margin_days(顧客納期 - 予測完了日)が小さい(=危険な)順に並べる。データ不足は最後。
    """
    entries: list[FeasibilityEntry] = []
    for r in records:
        current = r.current_process
        if current is None or r.customer_deadline is None:
            continue

        typical_lt = process_history.order_level_lt_calendar_days_median(r.drawing_no)
        sample_count = process_history.order_level_lt_sample_count(r.drawing_no)

        predicted_date: datetime.date | None = None
        data_note: str | None = None
        if typical_lt is None:
            data_note = "この図番の受注〜完成実績が工程累積データに無いため予測不可"
        elif r.order_date is None:
            data_note = "受注日が不明(内示・先行手配案件等)のため予測不可"
        else:
            predicted_date = r.order_date + datetime.timedelta(days=round(typical_lt))

        margin = (r.customer_deadline - predicted_date).days if predicted_date is not None else None

        entries.append(
            FeasibilityEntry(
                order_no=r.order_no,
                drawing_no=r.drawing_no,
                product_name=r.product_name,
                order_date=r.order_date,
                customer_deadline=r.customer_deadline,
                current_process_code=current.process_code,
                remaining_step_count=sum(1 for s in r.processes if not s.is_completed),
                typical_total_lt_calendar_days=typical_lt,
                typical_lt_sample_count=sample_count,
                predicted_completion_date=predicted_date,
                margin_days=margin,
                data_note=data_note,
            )
        )

    entries.sort(key=lambda e: (e.margin_days is None, e.margin_days if e.margin_days is not None else 0))
    return entries


def _week_start(d: datetime.date) -> datetime.date:
    return d - datetime.timedelta(days=d.weekday())  # 月曜始まり


def _add_business_days(start: datetime.date, business_days: int) -> datetime.date:
    """土日のみを除いた営業日数を加算する(祝日は考慮しない。原方針どおりシンプルに)。"""
    d = start
    remaining = max(business_days, 0)
    while remaining > 0:
        d += datetime.timedelta(days=1)
        if d.weekday() < 5:
            remaining -= 1
    return d


def build_capacity_forecast(
    records: list[OrderRecord],
    generated_at: datetime.date,
    process_master: ProcessMaster,
    process_history: ProcessHistory,
) -> tuple[list[tuple[datetime.date, str, float]], list[tuple[datetime.date, str, float]]]:
    """仕掛中の全受注について、残り工程が「いつ・どの部署/設備に・どれだけ」の負荷になるかを予測する。

    ②の教訓(実績日数の単純合計は誤差が拡大する)を踏まえ、日程の予測には実績日数ではなく
    標準LT(config/process_code_master.csv、営業日、変更しない)を使う。工数の大きさと、
    どの部署・設備が担当するかは過去実績(平均工数・実績シェア)から推定する(kii-san合意:
    部署単位を主軸としつつ、設備単位も実績シェアで比例配分して算出する)。

    ①②の判定・②の客先納期充足予測とは独立した、追加の参考情報。
    """
    department_totals: dict[tuple[datetime.date, str], float] = {}
    machine_totals: dict[tuple[datetime.date, str], float] = {}

    for r in records:
        current = r.current_process
        if current is None:
            continue

        cursor = generated_at
        for step in (s for s in r.processes if not s.is_completed):
            result = process_master.categorize(step.process_code)
            cursor = _add_business_days(cursor, result.standard_lt_business_days)
            week = _week_start(cursor)

            hours = process_history.average_manhours_for_process(step.process_code)
            if not hours:
                continue

            for dept, share in process_history.department_shares_for_process(step.process_code).items():
                key = (week, dept)
                department_totals[key] = department_totals.get(key, 0.0) + hours * share

            for machine, share in process_history.machine_shares_for_process(step.process_code).items():
                key = (week, machine)
                machine_totals[key] = machine_totals.get(key, 0.0) + hours * share

    by_department = sorted(
        ((week, dept, hours) for (week, dept), hours in department_totals.items()),
        key=lambda t: (t[0], t[1]),
    )
    by_machine = sorted(
        ((week, machine, hours) for (week, machine), hours in machine_totals.items()),
        key=lambda t: (t[0], t[1]),
    )
    return by_department, by_machine


def _walk_remaining_steps(records, generated_at, process_master):
    """仕掛中の各受注について、残り工程を標準LTで先へ進めながら(受注, 工程, 到達日)を列挙する。

    build_capacity_forecast等の集計系と、その内訳(製造オーダー単位)を返す関数群で
    同じ「日程の進め方」を共有するための内部ヘルパー(2026-09-08 kii-san要望対応で追加)。
    """
    for r in records:
        current = r.current_process
        if current is None:
            continue

        cursor = generated_at
        for step in (s for s in r.processes if not s.is_completed):
            result = process_master.categorize(step.process_code)
            cursor = _add_business_days(cursor, result.standard_lt_business_days)
            yield r, step, cursor


def build_capacity_forecast_machine_detail(
    records: list[OrderRecord],
    generated_at: datetime.date,
    process_master: ProcessMaster,
    process_history: ProcessHistory,
) -> list[MachineForecastDetailEntry]:
    """週別予測負荷(設備別)を製造オーダー単位まで分解した内訳。build_capacity_forecastと対になる。

    集計値(build_capacity_forecastのby_machine)と同じ按分ロジック(実績シェアで複数設備に
    比例配分)を使うため、両者の合計値は一致する。
    """
    detail: list[MachineForecastDetailEntry] = []
    for r, step, cursor in _walk_remaining_steps(records, generated_at, process_master):
        hours = process_history.average_manhours_for_process(step.process_code)
        if not hours:
            continue
        week = _week_start(cursor)
        for machine, share in process_history.machine_shares_for_process(step.process_code).items():
            detail.append(
                MachineForecastDetailEntry(
                    week=week, machine_code=machine, order_no=r.order_no, drawing_no=r.drawing_no,
                    product_name=r.product_name, process_code=step.process_code, hours=hours * share,
                )
            )
    detail.sort(key=lambda e: (e.week, e.machine_code, -e.hours))
    return detail


# 品種別月間キャパシティの前提(kii-san申告値、2026-09-08)。稼働率は24時間を100%とした値。
CATEGORY_MONTHLY_UTILIZATION_LOW = 0.30  # 楽観的(この稼働率だったとみなすと、キャパシティ上限は高く出る)
CATEGORY_MONTHLY_UTILIZATION_HIGH = 0.40  # 保守的(この稼働率だったとみなすと、キャパシティ上限は低く出る)
CATEGORY_MONTHLY_UTILIZATION_TYPICAL = 0.35  # ざっくりの目安(headline数値)に使う中間値
CATEGORY_MONTHLY_CAPACITY_BASELINE_MONTHS = 12  # 典型実績の算出に使う直近月数の上限
CATEGORY_MONTHLY_CAPACITY_CATEGORIES = ("GEAR", "BEVEL", "WORM")
CATEGORY_MONTHLY_AVERAGE_BUSINESS_DAYS = 20  # 弊社の平均営業日数/月(kii-san申告値、2026-09-08)


def build_monthly_category_capacity(
    records: list[OrderRecord],
    generated_at: datetime.date,
    process_master: ProcessMaster,
    process_history: ProcessHistory,
    classifier: ProductCategoryClassifier,
) -> list[MonthlyCategoryCapacityEntry]:
    """品種(GEAR/BEVEL/WORM)別・月別に、実績工数・予測工数・キャパシティ目安を並べる。

    詳しい考え方はMonthlyCategoryCapacityEntryのdocstring参照。
    """
    # --- 実績: 月×品種の実績工数を集計(品名から品種を判定できない分は除外) ---
    actual_by_month_category: dict[tuple[str, str], float] = defaultdict(float)
    for month, token, hours in process_history.monthly_load_by_product_token():
        category = classifier.classify(token)
        if category is None:
            continue
        actual_by_month_category[(month, category)] += hours

    # --- キャパシティ目安: 算出時点の月(未確定)を除いた直近12か月の平均実績から稼働率で逆算 ---
    current_month = generated_at.strftime("%Y-%m")
    months_by_category: dict[str, list[str]] = defaultdict(list)
    for month, category in actual_by_month_category:
        if month < current_month:
            months_by_category[category].append(month)

    capacity_by_category: dict[str, tuple[float, float, float]] = {}
    for category in CATEGORY_MONTHLY_CAPACITY_CATEGORIES:
        recent_months = sorted(months_by_category.get(category, []))[-CATEGORY_MONTHLY_CAPACITY_BASELINE_MONTHS:]
        if not recent_months:
            continue
        typical_actual = statistics.mean(actual_by_month_category[(m, category)] for m in recent_months)
        capacity_by_category[category] = (
            typical_actual / CATEGORY_MONTHLY_UTILIZATION_HIGH,  # low(保守的)
            typical_actual / CATEGORY_MONTHLY_UTILIZATION_LOW,  # high(楽観的)
            typical_actual / CATEGORY_MONTHLY_UTILIZATION_TYPICAL,  # typical(目安)
        )

    # --- 予測: 仕掛中受注の残り工程を標準LTで先の月へ積み上げ、受注自身の品種で集計 ---
    forecast_by_month_category: dict[tuple[str, str], float] = defaultdict(float)
    for r in records:
        current = r.current_process
        if current is None:
            continue
        category = classifier.classify(r.product_name)
        if category is None:
            continue

        cursor = generated_at
        for step in (s for s in r.processes if not s.is_completed):
            result = process_master.categorize(step.process_code)
            cursor = _add_business_days(cursor, result.standard_lt_business_days)
            month = cursor.strftime("%Y-%m")

            hours = process_history.average_manhours_for_process(step.process_code)
            if not hours:
                continue
            forecast_by_month_category[(month, category)] += hours

    months = sorted({m for m, _ in actual_by_month_category} | {m for m, _ in forecast_by_month_category})

    entries: list[MonthlyCategoryCapacityEntry] = []
    for month in months:
        for category in CATEGORY_MONTHLY_CAPACITY_CATEGORIES:
            actual = actual_by_month_category.get((month, category))
            forecast = forecast_by_month_category.get((month, category))
            if actual is None and forecast is None:
                continue

            low, high, typical = capacity_by_category.get(category, (None, None, None))
            fulfillment = (forecast / typical) if (forecast is not None and typical) else None
            typical_per_business_day = (
                typical / CATEGORY_MONTHLY_AVERAGE_BUSINESS_DAYS if typical is not None else None
            )

            entries.append(
                MonthlyCategoryCapacityEntry(
                    month=month,
                    category=category,
                    actual_hours=actual,
                    forecast_hours=forecast,
                    capacity_hours_low=low,
                    capacity_hours_high=high,
                    capacity_hours_typical=typical,
                    capacity_hours_typical_per_business_day=typical_per_business_day,
                    fulfillment_rate=fulfillment,
                )
            )
    return entries


def build_monthly_category_capacity_detail(
    records: list[OrderRecord],
    generated_at: datetime.date,
    process_master: ProcessMaster,
    process_history: ProcessHistory,
    classifier: ProductCategoryClassifier,
) -> list[MonthlyCategoryCapacityDetailEntry]:
    """品種別月間キャパシティの予測工数(forecast_hours)を製造オーダー単位まで分解した内訳。

    build_monthly_category_capacityと同じ按分ロジック(標準LTで日程、実績平均工数で大きさ)を
    使うため、両者の合計値は一致する。実績側(過去分)は集計元データが受注単位で追えないため
    内訳は提供しない(MonthlyCategoryCapacityDetailEntryのdocstring参照)。
    """
    detail: list[MonthlyCategoryCapacityDetailEntry] = []
    for r, step, cursor in _walk_remaining_steps(records, generated_at, process_master):
        category = classifier.classify(r.product_name)
        if category is None:
            continue
        hours = process_history.average_manhours_for_process(step.process_code)
        if not hours:
            continue
        detail.append(
            MonthlyCategoryCapacityDetailEntry(
                month=cursor.strftime("%Y-%m"), category=category, order_no=r.order_no,
                drawing_no=r.drawing_no, product_name=r.product_name, process_code=step.process_code,
                forecast_hours=hours,
            )
        )
    detail.sort(key=lambda e: (e.month, e.category, -e.forecast_hours))
    return detail


# 仕入先週別予測の前提(kii-san合意、2026-09-08)。
# 社内設備のような稼働率基準(30〜40%)が無いため絶対的な上限は推定できず、
# 「普段の実績水準」との相対比較(倍率)として示す。
SUPPLIER_WEEKLY_BASELINE_WEEKS = 52  # 「普段の実績水準」の算出に使う直近週数の上限(算出時点の週は除く)


def build_supplier_capacity_forecast(
    records: list[OrderRecord],
    generated_at: datetime.date,
    process_master: ProcessMaster,
    process_history: ProcessHistory,
    supplier_history: SupplierHistory,
) -> tuple[list[SupplierForecastEntry], list[SupplierForecastDetailEntry]]:
    """仕掛中の全受注について、外注実績のある残り工程を標準LTで先の週へ積み上げ、
    仕入先別の予測金額(相対比較)を算出する。詳しい考え方はSupplierForecastEntryのdocstring参照。
    """
    forecast_totals: dict[tuple[datetime.date, str], float] = defaultdict(float)
    detail: list[SupplierForecastDetailEntry] = []

    for r, step, cursor in _walk_remaining_steps(records, generated_at, process_master):
        avg_amount = supplier_history.average_amount_for_process(step.process_code)
        if not avg_amount:
            continue
        shares = supplier_history.supplier_amount_shares_for_process(step.process_code)
        if not shares:
            continue

        week = _week_start(cursor)
        for supplier, share in shares.items():
            amount = avg_amount * share
            forecast_totals[(week, supplier)] += amount
            detail.append(
                SupplierForecastDetailEntry(
                    week=week, supplier_code=supplier, order_no=r.order_no, drawing_no=r.drawing_no,
                    product_name=r.product_name, process_code=step.process_code, forecast_amount=amount,
                )
            )

    # 「普段の実績水準」= 算出時点の週(未確定)より前の、直近最大SUPPLIER_WEEKLY_BASELINE_WEEKS週平均。
    current_week = _week_start(generated_at)
    weeks_by_supplier: dict[str, list[tuple[datetime.date, float]]] = defaultdict(list)
    for week, supplier, amount in supplier_history.weekly_amount_by_supplier():
        if week < current_week:
            weeks_by_supplier[supplier].append((week, amount))

    typical_by_supplier: dict[str, float] = {}
    for supplier, points in weeks_by_supplier.items():
        recent = sorted(points)[-SUPPLIER_WEEKLY_BASELINE_WEEKS:]
        if recent:
            typical_by_supplier[supplier] = statistics.mean(amount for _, amount in recent)

    entries: list[SupplierForecastEntry] = []
    for (week, supplier), forecast_amount in forecast_totals.items():
        typical = typical_by_supplier.get(supplier)
        ratio = (forecast_amount / typical) if typical else None
        entries.append(
            SupplierForecastEntry(
                week=week, supplier_code=supplier, forecast_amount=forecast_amount,
                typical_weekly_amount=typical, ratio_to_typical=ratio,
            )
        )
    entries.sort(key=lambda e: (e.week, e.supplier_code))
    detail.sort(key=lambda e: (e.week, e.supplier_code, -e.forecast_amount))
    return entries, detail
