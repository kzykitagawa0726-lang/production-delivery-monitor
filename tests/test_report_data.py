import datetime

from src.models import OrderRecord, ProcessStep
from src.process_history import ProcessHistory
from src.process_master import ProcessMaster
from src.product_category import ProductCategoryClassifier
from src.report_data import (
    build_capacity_forecast,
    build_capacity_forecast_machine_detail,
    build_feasibility_estimates,
    build_monthly_category_capacity,
    build_monthly_category_capacity_detail,
    build_supplier_capacity_forecast,
)
from src.supplier_history import SupplierHistory

GENERATED_AT = datetime.date(2026, 9, 8)  # 火曜日


def make_process_master(tmp_path, rows: list[tuple[str, str, int]]) -> ProcessMaster:
    """(process_code, category, standard_lt_business_days) からテスト用の対応表を作る。"""
    categories_path = tmp_path / "categories.json"
    categories_path.write_text(
        '{"categories": {"その他": {"standard_lt_business_days": 5}}, "unknown_code_category": "その他"}',
        encoding="utf-8",
    )
    master_path = tmp_path / "master.csv"
    lines = ["process_code,category,note,is_bottleneck,standard_lt_business_days"]
    for code, category, lt in rows:
        lines.append(f"{code},{category},,false,{lt}")
    master_path.write_text("\n".join(lines), encoding="utf-8")
    return ProcessMaster(categories_path=categories_path, master_csv_path=master_path)


def make_order(**overrides) -> OrderRecord:
    defaults = dict(
        order_no="PO-1",
        drawing_no="DWG-A1",
        product_name="ギヤA",
        qty=1,
        order_date=datetime.date(2026, 8, 1),
        company_deadline=datetime.date(2026, 9, 20),
        customer_deadline=datetime.date(2026, 9, 20),
        remaining_business_days=8,
    )
    defaults.update(overrides)
    return OrderRecord(**defaults)


def make_history_with_order_level_lt(drawing_no: str, days: int, sample_count: int = 1) -> ProcessHistory:
    history = ProcessHistory()
    history._order_level_durations_by_drawing[drawing_no] = [days] * sample_count
    return history


def make_classifier(tmp_path, rows: list[tuple[str, str]]) -> ProductCategoryClassifier:
    """(keyword, category) からテスト用の品種対応表を作る。"""
    path = tmp_path / "keywords.csv"
    lines = ["keyword,category"]
    for keyword, category in rows:
        lines.append(f"{keyword},{category}")
    path.write_text("\n".join(lines), encoding="utf-8")
    return ProductCategoryClassifier(keywords_path=path)


def test_feasibility_uses_order_date_plus_drawing_typical_lt():
    order = make_order(
        order_date=datetime.date(2026, 8, 1),
        processes=[ProcessStep(2, "HG", "F2", "d1", "d2", "オーダー確定前")],
    )
    history = make_history_with_order_level_lt("DWG-A1", days=14)

    entries = build_feasibility_estimates([order], GENERATED_AT, history)

    assert len(entries) == 1
    e = entries[0]
    assert e.typical_total_lt_calendar_days == 14
    assert e.typical_lt_sample_count == 1
    assert e.predicted_completion_date == datetime.date(2026, 8, 15)  # 8/1 + 14日
    assert e.margin_days == (order.customer_deadline - e.predicted_completion_date).days
    assert e.data_note is None


def test_feasibility_status_reflects_margin():
    order = make_order(
        order_date=datetime.date(2026, 9, 1),
        customer_deadline=datetime.date(2026, 9, 5),
        processes=[ProcessStep(2, "HG", "F2", "d1", "d2", "オーダー確定前")],
    )
    history = make_history_with_order_level_lt("DWG-A1", days=30)  # 30日かかる想定 vs 納期まで4日

    entries = build_feasibility_estimates([order], GENERATED_AT, history)
    assert entries[0].status == "間に合わない見込み"
    assert entries[0].margin_days < 0


def test_feasibility_excludes_orders_without_customer_deadline():
    order = make_order(
        customer_deadline=None,
        processes=[ProcessStep(2, "HG", "F2", "d1", "d2", "オーダー確定前")],
    )
    history = make_history_with_order_level_lt("DWG-A1", days=10)

    entries = build_feasibility_estimates([order], GENERATED_AT, history)
    assert entries == []


def test_feasibility_excludes_orders_with_no_current_process():
    order = make_order(
        processes=[ProcessStep(2, "HG", "F2", "d1", "d2", "作業完了")],  # 全工程完了
    )
    history = make_history_with_order_level_lt("DWG-A1", days=10)

    entries = build_feasibility_estimates([order], GENERATED_AT, history)
    assert entries == []


def test_feasibility_reports_no_data_when_drawing_has_no_order_level_history():
    order = make_order(
        drawing_no="DWG-UNKNOWN",
        processes=[ProcessStep(2, "HG", "F2", "d1", "d2", "オーダー確定前")],
    )
    history = ProcessHistory()  # 実績データなし

    entries = build_feasibility_estimates([order], GENERATED_AT, history)
    assert len(entries) == 1
    assert entries[0].predicted_completion_date is None
    assert entries[0].status == "データ不足"
    assert "受注〜完成実績が" in entries[0].data_note


def test_feasibility_reports_no_data_when_order_date_missing():
    # 内示・先行手配案件など、受注日が無いと起点が分からず予測できない
    order = make_order(
        order_date=None,
        processes=[ProcessStep(2, "HG", "F2", "d1", "d2", "オーダー確定前")],
    )
    history = make_history_with_order_level_lt("DWG-A1", days=10)

    entries = build_feasibility_estimates([order], GENERATED_AT, history)
    assert len(entries) == 1
    assert entries[0].predicted_completion_date is None
    assert entries[0].status == "データ不足"
    assert "受注日が不明" in entries[0].data_note


def test_feasibility_sorts_most_urgent_first_and_missing_data_last():
    late_order = make_order(
        order_no="PO-LATE",
        order_date=datetime.date(2026, 9, 1),
        customer_deadline=datetime.date(2026, 9, 3),
        processes=[ProcessStep(2, "HG", "F2", "d1", "d2", "オーダー確定前")],
    )
    ontime_order = make_order(
        order_no="PO-ONTIME",
        order_date=datetime.date(2026, 9, 1),
        customer_deadline=datetime.date(2026, 12, 1),
        processes=[ProcessStep(2, "HG", "F2", "d1", "d2", "オーダー確定前")],
    )
    unknown_order = make_order(
        order_no="PO-UNKNOWN",
        drawing_no="DWG-UNKNOWN",
        processes=[ProcessStep(2, "HG", "F2", "d1", "d2", "オーダー確定前")],
    )
    history = make_history_with_order_level_lt("DWG-A1", days=10)

    entries = build_feasibility_estimates([ontime_order, unknown_order, late_order], GENERATED_AT, history)

    assert [e.order_no for e in entries] == ["PO-LATE", "PO-ONTIME", "PO-UNKNOWN"]


def test_capacity_forecast_uses_standard_lt_for_timing_and_shares_for_magnitude(tmp_path):
    # GENERATED_AT(2026-09-08 火)+3営業日 = 09-09(水),09-10(木),09-11(金) -> 週初め09-07(月)
    order = make_order(processes=[ProcessStep(2, "TC", "F1", "d1", "d2", "オーダー確定前")])
    process_master = make_process_master(tmp_path, [("TC", "その他", 3)])

    history = ProcessHistory()
    history._manhours_by_process["TC"] = [4.0]
    history._department_counts_by_process["TC"]["DEPT1"] += 1
    history._by_process["TC"]["M1"] = {"count": 3, "last_used": None}
    history._by_process["TC"]["M2"] = {"count": 2, "last_used": None}

    by_dept, by_machine = build_capacity_forecast([order], GENERATED_AT, process_master, history)

    week = datetime.date(2026, 9, 7)
    assert by_dept == [(week, "DEPT1", 4.0)]
    assert sorted(by_machine) == sorted([(week, "M1", 2.4), (week, "M2", 1.6)])  # シェア0.6/0.4で按分


def test_capacity_forecast_walks_multiple_remaining_steps_forward(tmp_path):
    # 1歩目: 09-08+3営業日=09-11(金,週07-13の週) / 2歩目: さらに+2営業日=09-15(火,週14-20の週)
    order = make_order(
        processes=[
            ProcessStep(2, "TC", "F1", "d1", "d2", "オーダー確定前"),
            ProcessStep(3, "TD", "F2", "d1", "d2", "オーダー確定前"),
        ],
    )
    process_master = make_process_master(tmp_path, [("TC", "その他", 3), ("TD", "その他", 2)])

    history = ProcessHistory()
    history._manhours_by_process["TC"] = [4.0]
    history._department_counts_by_process["TC"]["DEPT1"] += 1
    history._manhours_by_process["TD"] = [6.0]
    history._department_counts_by_process["TD"]["DEPT2"] += 1

    by_dept, _ = build_capacity_forecast([order], GENERATED_AT, process_master, history)

    assert (datetime.date(2026, 9, 7), "DEPT1", 4.0) in by_dept
    assert (datetime.date(2026, 9, 14), "DEPT2", 6.0) in by_dept


def test_capacity_forecast_skips_steps_without_manhours_data(tmp_path):
    order = make_order(processes=[ProcessStep(2, "ZZ", "F1", "d1", "d2", "オーダー確定前")])
    process_master = make_process_master(tmp_path, [("ZZ", "その他", 3)])
    history = ProcessHistory()  # 実績データなし

    by_dept, by_machine = build_capacity_forecast([order], GENERATED_AT, process_master, history)
    assert by_dept == []
    assert by_machine == []


def test_capacity_forecast_excludes_orders_with_no_remaining_steps(tmp_path):
    order = make_order(processes=[ProcessStep(2, "TC", "F1", "d1", "d2", "作業完了")])  # 全工程完了
    process_master = make_process_master(tmp_path, [("TC", "その他", 3)])
    history = ProcessHistory()
    history._manhours_by_process["TC"] = [4.0]
    history._department_counts_by_process["TC"]["DEPT1"] += 1

    by_dept, _ = build_capacity_forecast([order], GENERATED_AT, process_master, history)
    assert by_dept == []


def test_monthly_category_capacity_derives_range_from_utilization_rate(tmp_path):
    # 実績: GEARの過去2か月(いずれも算出時点2026-09より前)の平均10.0h -> 稼働率30-40%で逆算
    history = ProcessHistory()
    history._monthly_load_by_product_token[("2026-07", "ギヤ")] = 8.0
    history._monthly_load_by_product_token[("2026-08", "ギヤ")] = 12.0  # 平均10.0
    classifier = make_classifier(tmp_path, [("ギヤ", "GEAR")])
    process_master = make_process_master(tmp_path, [])

    entries = build_monthly_category_capacity([], GENERATED_AT, process_master, history, classifier)

    aug = next(e for e in entries if e.month == "2026-08" and e.category == "GEAR")
    assert aug.actual_hours == 12.0
    assert round(aug.capacity_hours_low, 2) == round(10.0 / 0.40, 2)  # 保守的(稼働率40%と仮定)
    assert round(aug.capacity_hours_high, 2) == round(10.0 / 0.30, 2)  # 楽観的(稼働率30%と仮定)
    assert round(aug.capacity_hours_typical, 2) == round(10.0 / 0.35, 2)  # 目安(稼働率35%と仮定)
    assert round(aug.capacity_hours_typical_per_business_day, 2) == round((10.0 / 0.35) / 20, 2)  # 月20営業日換算


def test_monthly_category_capacity_excludes_current_partial_month_from_baseline(tmp_path):
    history = ProcessHistory()
    history._monthly_load_by_product_token[("2026-08", "ギヤ")] = 10.0
    history._monthly_load_by_product_token[("2026-09", "ギヤ")] = 1.0  # 算出時点の月(まだ確定していない)
    classifier = make_classifier(tmp_path, [("ギヤ", "GEAR")])
    process_master = make_process_master(tmp_path, [])

    entries = build_monthly_category_capacity([], GENERATED_AT, process_master, history, classifier)

    sep = next(e for e in entries if e.month == "2026-09" and e.category == "GEAR")
    assert sep.actual_hours == 1.0  # 実績としては表示する
    # ただしキャパシティのベースラインには含めない(確定済みの08月10.0hのみで算出)
    assert round(sep.capacity_hours_typical, 2) == round(10.0 / 0.35, 2)


def test_monthly_category_capacity_forecast_uses_order_product_category(tmp_path):
    order = make_order(
        product_name="ギヤA",
        processes=[ProcessStep(2, "TC", "F1", "d1", "d2", "オーダー確定前")],
    )
    process_master = make_process_master(tmp_path, [("TC", "その他", 20)])
    history = ProcessHistory()
    history._manhours_by_process["TC"] = [5.0]
    classifier = make_classifier(tmp_path, [("ギヤ", "GEAR")])

    entries = build_monthly_category_capacity([order], GENERATED_AT, process_master, history, classifier)

    forecast_entry = next(e for e in entries if e.category == "GEAR" and e.forecast_hours is not None)
    assert forecast_entry.forecast_hours == 5.0
    assert forecast_entry.actual_hours is None  # 実績データが無いため


def test_monthly_category_capacity_excludes_unclassified_products(tmp_path):
    # 「フランジ」は品種対応表に無いキーワードなので、実績・予測どちらも3分類の対象外
    history = ProcessHistory()
    history._monthly_load_by_product_token[("2026-08", "フランジ")] = 10.0
    classifier = make_classifier(tmp_path, [("ギヤ", "GEAR")])
    process_master = make_process_master(tmp_path, [("TC", "その他", 3)])

    order = make_order(
        product_name="フランジB",
        processes=[ProcessStep(2, "TC", "F1", "d1", "d2", "オーダー確定前")],
    )

    entries = build_monthly_category_capacity([order], GENERATED_AT, process_master, history, classifier)
    assert entries == []


def test_capacity_forecast_machine_detail_matches_aggregate_total(tmp_path):
    # build_capacity_forecastの集計値(2.4/1.6h、シェア0.6/0.4)と、内訳の合計が一致すること。
    order = make_order(processes=[ProcessStep(2, "TC", "F1", "d1", "d2", "オーダー確定前")])
    process_master = make_process_master(tmp_path, [("TC", "その他", 3)])

    history = ProcessHistory()
    history._manhours_by_process["TC"] = [4.0]
    history._by_process["TC"]["M1"] = {"count": 3, "last_used": None}
    history._by_process["TC"]["M2"] = {"count": 2, "last_used": None}

    detail = build_capacity_forecast_machine_detail([order], GENERATED_AT, process_master, history)

    week = datetime.date(2026, 9, 7)
    assert sorted((e.machine_code, round(e.hours, 2)) for e in detail) == [("M1", 2.4), ("M2", 1.6)]
    for e in detail:
        assert e.week == week
        assert e.order_no == "PO-1"
        assert e.drawing_no == "DWG-A1"
        assert e.process_code == "TC"


def test_monthly_category_capacity_detail_matches_aggregate_total(tmp_path):
    order = make_order(
        product_name="ギヤA",
        processes=[ProcessStep(2, "TC", "F1", "d1", "d2", "オーダー確定前")],
    )
    process_master = make_process_master(tmp_path, [("TC", "その他", 3)])
    history = ProcessHistory()
    history._manhours_by_process["TC"] = [5.0]
    classifier = make_classifier(tmp_path, [("ギヤ", "GEAR")])

    detail = build_monthly_category_capacity_detail([order], GENERATED_AT, process_master, history, classifier)

    assert len(detail) == 1
    e = detail[0]
    assert e.month == "2026-09"
    assert e.category == "GEAR"
    assert e.order_no == "PO-1"
    assert e.drawing_no == "DWG-A1"
    assert e.process_code == "TC"
    assert e.forecast_hours == 5.0


def test_supplier_capacity_forecast_distributes_by_amount_share_and_flags_relative_to_typical(tmp_path):
    # TC: 外注実績あり(仕入先9001/9002、金額シェア0.75/0.25) -> 予測対象。
    # ZZ: 外注実績なし -> 按分先が無く自然に予測対象外(社内設備側のみに現れる想定)。
    order = make_order(
        processes=[
            ProcessStep(2, "TC", "F1", "d1", "d2", "オーダー確定前"),
            ProcessStep(3, "ZZ", "F2", "d1", "d2", "オーダー確定前"),
        ],
    )
    process_master = make_process_master(tmp_path, [("TC", "その他", 3), ("ZZ", "その他", 3)])
    process_history = ProcessHistory()

    supplier_history = SupplierHistory()
    supplier_history._by_process["TC"]["9001"] = {"count": 3, "last_used": None, "amount": 3000.0}
    supplier_history._by_process["TC"]["9002"] = {"count": 1, "last_used": None, "amount": 1000.0}
    # 「普段の実績水準」(算出時点の週より前の週次実績)
    prior_week = datetime.date(2026, 8, 31)
    supplier_history._weekly_amount_by_supplier[(prior_week, "9001")] = 400.0
    supplier_history._weekly_amount_by_supplier[(prior_week, "9002")] = 100.0

    entries, detail = build_supplier_capacity_forecast(
        [order], GENERATED_AT, process_master, process_history, supplier_history
    )

    week = datetime.date(2026, 9, 7)
    by_supplier = {e.supplier_code: e for e in entries}
    assert set(by_supplier) == {"9001", "9002"}  # ZZは対象外(按分先なし)

    assert round(by_supplier["9001"].forecast_amount, 2) == 750.0  # 平均金額1000 × シェア0.75
    assert by_supplier["9001"].typical_weekly_amount == 400.0
    assert round(by_supplier["9001"].ratio_to_typical, 3) == round(750 / 400, 3)

    assert round(by_supplier["9002"].forecast_amount, 2) == 250.0
    assert by_supplier["9002"].typical_weekly_amount == 100.0

    assert all(e.week == week for e in entries)
    assert {e.process_code for e in detail} == {"TC"}  # ZZの内訳は無い
    assert round(sum(e.forecast_amount for e in detail if e.supplier_code == "9001"), 2) == 750.0


def test_supplier_capacity_forecast_excludes_orders_with_no_remaining_steps(tmp_path):
    order = make_order(processes=[ProcessStep(2, "ZZ", "F1", "d1", "d2", "作業完了")])  # 全工程完了
    process_master = make_process_master(tmp_path, [("ZZ", "その他", 3)])
    process_history = ProcessHistory()
    supplier_history = SupplierHistory()
    supplier_history._by_process["ZZ"]["9001"] = {"count": 1, "last_used": None, "amount": 100.0}

    entries, detail = build_supplier_capacity_forecast(
        [order], GENERATED_AT, process_master, process_history, supplier_history
    )
    assert entries == []
    assert detail == []
