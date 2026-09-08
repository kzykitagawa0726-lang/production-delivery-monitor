import datetime

from src.models import OrderRecord, ProcessStep
from src.process_history import ProcessHistory
from src.report_data import build_feasibility_estimates

GENERATED_AT = datetime.date(2026, 9, 8)


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
