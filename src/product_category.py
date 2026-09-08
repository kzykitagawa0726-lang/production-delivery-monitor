"""品名から品種(GEAR/BEVEL/WORM)を分類する。

工程コード対応表と同じ考え方で、キーワード→品種の対応をプログラムに埋め込まず、
config/product_category_keywords.csv に分離している(2026-09-08 kii-san要望: 将来
品名パターンが増えても自分で編集できるように)。

品名の「前方一致」で判定する(実データで、品種を表す語が品名の先頭に来る命名規則と確認済み)。
一致しない品名(フランジ/シャフト/スプライン等の付随部品、型番のみの表記など)はNoneを返し、
GEAR/BEVEL/WORMの集計からは除外する(誤って いずれかの品種に含めてしまうのを避けるため)。
"""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Optional

DEFAULT_KEYWORDS_PATH = Path("config/product_category_keywords.csv")


class ProductCategoryClassifier:
    def __init__(self, keywords_path: Path = DEFAULT_KEYWORDS_PATH) -> None:
        self._rules: list[tuple[str, str]] = []
        with open(keywords_path, encoding="utf-8", newline="") as f:
            for row in csv.DictReader(_skip_comments(f)):
                keyword = row["keyword"].strip()
                category = row["category"].strip()
                if keyword and category:
                    self._rules.append((keyword, category))
        # 長いキーワードを先に判定する(例:「複ＷＯＲＭ」を「ＷＯＲＭ」より先にマッチさせるため)。
        self._rules.sort(key=lambda kv: len(kv[0]), reverse=True)

    def classify(self, product_name: Optional[str]) -> Optional[str]:
        if not product_name:
            return None
        normalized = str(product_name).replace("　", " ").strip()
        for keyword, category in self._rules:
            if normalized.startswith(keyword):
                return category
        return None


def _skip_comments(f):
    for line in f:
        if not line.lstrip().startswith("#"):
            yield line
