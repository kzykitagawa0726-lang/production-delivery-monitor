from src.process_master import ProcessMaster


def test_known_sample_code_is_categorized():
    master = ProcessMaster()
    result = master.categorize("SAMPLE_G02")
    assert result.category == "歯切り系"
    assert result.is_bottleneck is True
    assert result.is_unknown_code is False


def test_unknown_code_falls_back_to_other_and_is_tracked():
    master = ProcessMaster()
    result = master.categorize("CR")  # 実データで確認された未登録コード
    assert result.category == "その他"
    assert result.is_bottleneck is False
    assert result.is_unknown_code is True
    assert "CR" in master.get_unknown_codes()
