from src.process_master import ProcessMaster


def test_known_code_returns_category_and_lt():
    pm = ProcessMaster()
    result = pm.categorize("SAMPLE_H01")
    assert result.category == "熱処理系"
    assert result.standard_lt_business_days == 10
    assert result.is_unknown_code is False


def test_bottleneck_flag():
    pm = ProcessMaster()
    result = pm.categorize("SAMPLE_G02")
    assert result.is_bottleneck is True

    result2 = pm.categorize("SAMPLE_G01")
    assert result2.is_bottleneck is False


def test_unknown_code_falls_back_to_other_and_is_flagged():
    pm = ProcessMaster()
    result = pm.categorize("UNKNOWN_CODE_XYZ")
    assert result.category == "その他"
    assert result.standard_lt_business_days == 5
    assert result.is_unknown_code is True
    assert "UNKNOWN_CODE_XYZ" in pm.get_unknown_codes()


def test_inspection_category_lt():
    pm = ProcessMaster()
    result = pm.categorize("SAMPLE_I01")
    assert result.category == "検査"
    assert result.standard_lt_business_days == 3
