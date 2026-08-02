"""工程コード対応表(外部CSV/JSON)のロードとカテゴリ判定。

コードとカテゴリの対応はプログラムに埋め込まず、
config/process_categories.json (カテゴリ別標準LT) と
config/process_code_master.csv (工程コード→カテゴリ対応) から読み込む。
未知の工程コードは「その他」として扱い、警告として収集する。
"""
from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from pathlib import Path

DEFAULT_CATEGORIES_PATH = Path(__file__).resolve().parent.parent / "config" / "process_categories.json"
DEFAULT_CODE_MASTER_PATH = Path(__file__).resolve().parent.parent / "config" / "process_code_master.csv"


@dataclass(frozen=True)
class ProcessCategoryResult:
    process_code: str
    category: str
    standard_lt_business_days: int
    is_bottleneck: bool
    is_unknown_code: bool


class ProcessMaster:
    def __init__(self, categories_path: Path = DEFAULT_CATEGORIES_PATH, code_master_path: Path = DEFAULT_CODE_MASTER_PATH):
        self._categories, self._unknown_category = self._load_categories(categories_path)
        self._code_mapping = self._load_code_mapping(code_master_path)
        self._unknown_codes_seen: set[str] = set()

    @staticmethod
    def _load_categories(path: Path) -> tuple[dict[str, int], str]:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        categories = {
            name: info["standard_lt_business_days"]
            for name, info in data["categories"].items()
        }
        unknown_category = data.get("unknown_code_category", "その他")
        if unknown_category not in categories:
            raise ValueError(
                f"unknown_code_category '{unknown_category}' is not defined in categories: {list(categories)}"
            )
        return categories, unknown_category

    @staticmethod
    def _load_code_mapping(path: Path) -> dict[str, dict]:
        mapping: dict[str, dict] = {}
        with open(path, encoding="utf-8") as f:
            reader = csv.DictReader(row for row in f if not row.lstrip().startswith("#"))
            for row in reader:
                code = row["process_code"].strip()
                if not code:
                    continue
                mapping[code] = {
                    "category": row["category"].strip(),
                    "is_bottleneck": row.get("is_bottleneck", "").strip().lower() == "true",
                }
        return mapping

    def categorize(self, process_code: str) -> ProcessCategoryResult:
        entry = self._code_mapping.get(process_code)
        if entry is None:
            self._unknown_codes_seen.add(process_code)
            category = self._unknown_category
            is_bottleneck = False
            is_unknown = True
        else:
            category = entry["category"]
            if category not in self._categories:
                # マスタに未定義のカテゴリが指定されている場合もフォールバックし警告対象とする
                self._unknown_codes_seen.add(process_code)
                category = self._unknown_category
                is_unknown = True
            else:
                is_unknown = False
            is_bottleneck = entry["is_bottleneck"]

        return ProcessCategoryResult(
            process_code=process_code,
            category=category,
            standard_lt_business_days=self._categories[category],
            is_bottleneck=is_bottleneck,
            is_unknown_code=is_unknown,
        )

    def get_unknown_codes(self) -> list[str]:
        return sorted(self._unknown_codes_seen)
