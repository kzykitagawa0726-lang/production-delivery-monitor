"""工程コード → カテゴリ の対応表を外部設定(CSV/JSON)から読み込む。

将来コードが追加された場合に自分で編集できるよう、コードとカテゴリの対応は
プログラム内に埋め込まず、config/process_code_master.csv と
config/process_categories.json に分離している(kii-san要件)。

工程コード対応表(PDF由来)が届くまでは、実データの工程コードは全て
config/process_code_master.csv 上のサンプル行にしか一致しないため、
すべて「その他」(未知コード)として扱われる。未知コードは一覧にまとめ、
レポート上で警告表示する。
"""
from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from pathlib import Path

DEFAULT_CATEGORIES_PATH = Path("config/process_categories.json")
DEFAULT_MASTER_CSV_PATH = Path("config/process_code_master.csv")


@dataclass
class CategoryResult:
    category: str
    standard_lt_business_days: int
    is_bottleneck: bool
    is_unknown_code: bool


class ProcessMaster:
    def __init__(
        self,
        categories_path: Path = DEFAULT_CATEGORIES_PATH,
        master_csv_path: Path = DEFAULT_MASTER_CSV_PATH,
    ) -> None:
        with open(categories_path, encoding="utf-8") as f:
            categories_config = json.load(f)
        self._categories: dict[str, int] = {
            name: cfg["standard_lt_business_days"] for name, cfg in categories_config["categories"].items()
        }
        self._unknown_category = categories_config["unknown_code_category"]

        self._code_map: dict[str, tuple[str, bool]] = {}
        with open(master_csv_path, encoding="utf-8", newline="") as f:
            for row in csv.DictReader(_skip_comments(f)):
                code = row["process_code"].strip()
                category = row["category"].strip()
                is_bottleneck = row["is_bottleneck"].strip().lower() == "true"
                self._code_map[code] = (category, is_bottleneck)

        self._unknown_codes: set[str] = set()

    def categorize(self, process_code: str) -> CategoryResult:
        entry = self._code_map.get(process_code)
        if entry is None:
            self._unknown_codes.add(process_code)
            category = self._unknown_category
            is_bottleneck = False
            is_unknown = True
        else:
            category, is_bottleneck = entry
            is_unknown = False

        standard_lt = self._categories.get(category, self._categories[self._unknown_category])
        return CategoryResult(
            category=category,
            standard_lt_business_days=standard_lt,
            is_bottleneck=is_bottleneck,
            is_unknown_code=is_unknown,
        )

    def get_unknown_codes(self) -> list[str]:
        return sorted(self._unknown_codes)


def _skip_comments(f):
    for line in f:
        if not line.lstrip().startswith("#"):
            yield line
