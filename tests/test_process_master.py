from src.process_master import ProcessMaster


def test_known_code_is_categorized_with_per_code_lt():
    master = ProcessMaster()
    result = master.categorize("HB")  # 浸炭焼入れ（防炭有り）、対応表上のLT=10
    assert result.category == "熱処理系"
    assert result.is_bottleneck is False
    assert result.is_unknown_code is False
    assert result.standard_lt_business_days == 10


def test_bottleneck_code_is_flagged():
    master = ProcessMaster()
    result = master.categorize("GW")  # ウォーム歯研
    assert result.category == "歯切り系"
    assert result.is_bottleneck is True
    assert result.is_unknown_code is False


def test_unknown_code_falls_back_to_other_and_is_tracked():
    master = ProcessMaster()
    result = master.categorize("ZZZ99")  # 対応表に存在しないコード
    assert result.category == "その他"
    assert result.is_bottleneck is False
    assert result.is_unknown_code is True
    assert "ZZZ99" in master.get_unknown_codes()
