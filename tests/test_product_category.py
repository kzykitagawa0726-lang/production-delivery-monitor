from src.product_category import ProductCategoryClassifier


def test_classifies_known_prefixes():
    c = ProductCategoryClassifier()
    assert c.classify("ＧＥＡＲ　Ｍ６ｘ６０　Ｓ") == "GEAR"
    assert c.classify("ＳＰＢ　Ｍ３Ｘ２４Ｔ　Ｌ") == "BEVEL"
    assert c.classify("複ＷＯＲＭ　Ｍ６Ｘ１口　Ｒ") == "WORM"


def test_longer_keyword_wins_over_shorter_prefix():
    # 「複ＷＯＲＭ」は「ＷＯＲＭ」で始まらないので通常のstartswithでも問題ないが、
    # 長い方を優先するルールで確実にWORMと判定されることを確認する。
    c = ProductCategoryClassifier()
    assert c.classify("複リードウォームホイル") == "WORM"


def test_unmatched_product_name_returns_none():
    c = ProductCategoryClassifier()
    assert c.classify("フランジ") is None
    assert c.classify("スプラインナット") is None


def test_none_and_empty_input():
    c = ProductCategoryClassifier()
    assert c.classify(None) is None
    assert c.classify("") is None
