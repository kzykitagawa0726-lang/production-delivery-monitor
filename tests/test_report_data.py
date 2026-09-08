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


def make_history_with_durations(durations: dict[str, list[int]]) -> ProcessHistory:
    history = ProcessHistory()
    for code, values in durations.items():
        history._actual_durations[code] = list(values)
    return history


def test_feasibility_sums_remaining_step_durations():
    order = make_order(
        processes=[
            ProcessStep(2, "L0", "F1", "d1", "d2", "作業完了"),  # 完了済みなので合計に含めない
            ProcessStep(3, "HG", "F2", "d1", "d2", "オーダー確定前"),
            ProcessStep(4, "CH", "F3", "d1", "d2", "オーダー確定前"),
        ],
    )
    history = make_history_with_durations({"HG": [10], "CH": [4]})

    entries = build_feasibility_estimates([order], GENERATED_AT, history)

    assert len(entries) == 1
    e = entries[0]
    assert e.predicted_remaining_calendar_days == 14
    assert e.predicted_completion_date == GENERATED_AT + datetime.timedelta(days=14)
    assert e.margin_days == (order.customer_deadline - e.predicted_completion_date).days
    assert e.missing_process_codes == []
    assert e.status in ("間に合う見込み", "間に合わない見込み")


def test_feasibility_flags_missing_data_but_still_sums_known_steps():
    order = make_order(
        processes=[
            ProcessStep(2, "HG", "F2", "d1", "d2", "オーダー確定前"),
            ProcessStep(3, "ZZ", "F3", "d1", "d2", "オーダー確定前"),  # 実績データなし
        ],
    )
    history = make_history_with_durations({"HG": [10]})

    entries = build_feasibility_estimates([order], GENERATED_AT, history)

    assert len(entries) == 1
    e = entries[0]
    assert e.predicted_remaining_calendar_days == 10  # 分かる工程だけの合計(下限値)
    assert e.missing_process_codes == ["ZZ"]
    assert e.predicted_completion_date is not None


def test_feasibility_status_reflects_margin():
    order = make_order(
        customer_deadline=GENERATED_AT + datetime.timedelta(days=3),
        processes=[ProcessStep(2, "HG", "F2", "d1", "d2", "オーダー確定前")],
    )
    history = make_history_with_durations({"HG": [10]})  # 10日かかる予測 vs 納期まで3日

    entries = build_feasibility_estimates([order], GENERATED_AT, history)
    assert entries[0].status == "間に合わない見込み"
    assert entries[0].margin_days < 0


def test_feasibility_excludes_orders_without_customer_deadline():
    order = make_order(
        customer_deadline=None,
        processes=[ProcessStep(2, "HG", "F2", "d1", "d2", "オーダー確定前")],
    )
    history = make_history_with_durations({"HG": [10]})

    entries = build_feasibility_estimates([order], GENERATED_AT, history)
    assert entries == []


def test_feasibility_excludes_orders_with_no_current_process():
    order = make_order(
        processes=[ProcessStep(2, "HG", "F2", "d1", "d2", "作業完了")],  # 全工程完了
    )
    history = make_history_with_durations({"HG": [10]})

    entries = build_feasibility_estimates([order], GENERATED_AT, history)
    assert entries == []


def test_feasibility_reports_no_data_when_all_remaining_steps_unknown():
    order = make_order(
        processes=[ProcessStep(2, "ZZ", "F2", "d1", "d2", "オーダー確定前")],
    )
    history = ProcessHistory()  # 実績データなし

    entries = build_feasibility_estimates([order], GENERATED_AT, history)
    assert len(entries) == 1
    assert entries[0].predicted_completion_date is None
    assert entries[0].status == "データ不足"


def test_feasibility_sorts_most_urgent_first_and_missing_data_last():
    late_order = make_order(
        order_no="PO-LATE",
        customer_deadline=GENERATED_AT + datetime.timedelta(days=1),
        processes=[ProcessStep(2, "HG", "F2", "d1", "d2", "オーダー確定前")],
    )
    ontime_order = make_order(
        order_no="PO-ONTIME",
        customer_deadline=GENERATED_AT + datetime.timedelta(days=30),
        processes=[ProcessStep(2, "HG", "F2", "d1", "d2", "オーダー確定前")],
    )
    unknown_order = make_order(
        order_no="PO-UNKNOWN",
        processes=[ProcessStep(2, "ZZ", "F2", "d1", "d2", "オーダー確定前")],
    )
    history = make_history_with_durations({"HG": [10]})

    entries = build_feasibility_estimates([ontime_order, unknown_order, late_order], GENERATED_AT, history)

    assert [e.order_no for e in entries] == ["PO-LATE", "PO-ONTIME", "PO-UNKNOWN"]
