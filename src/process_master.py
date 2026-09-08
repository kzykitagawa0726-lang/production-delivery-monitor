"""工程コード → カテゴリ の対応表を外部設定(CSV/JSON)から読み込む。

将来コードが追加された場合に自分で編集できるよう、コードとカテゴリの対応は
プログラム内に埋め込まず、config/process_code_master.csv と
config/process_categories.json に分離している(kii-san要件)。

2026-09-08、kii-sanより実際の工程コード対応表(218件、コード別LT付き)を
提供いただいたため、config/process_code_master.csv に反映済み。
カテゴリ(熱処理系/歯切り系/研磨系/検査/その他)とボトルネック区分
(ウォーム・スプライン・歯研系)は、対応表に列がなかったため工程名詳細の
キーワードから自動推定した一次案。LT(標準リードタイム)は対応表記載の
コード別の値をそのまま使用する(カテゴリ単位の固定値は、CSV側にLTが
無い将来の未知コード用のフォールバックとしてのみ残す)。
"""
from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from pathlib import Path

DEFAULT_CATEGORIES_PATH = Path("config/process_categories.json")
DEFAULT_MASTER_CSV_PATH = Path("config/process_code_master.csv")


def normalize_process_code(process_code: str) -> str:
    """1桁の数値コードは対応表側がゼロ埋め表記(例:"05")のため、参照側もゼロ埋めして揃える。

    仕入データはExcelの数値セルとして "5" のように読めてしまうため、
    工程コード対応表と突き合わせる前に正規化する(process_master.py・
    supplier_history.py共通で使用)。
    """
    code = str(process_code).strip()
    if code.isdigit() and len(code) == 1:
        return code.zfill(2)
    return code


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

        self._code_map: dict[str, tuple[str, bool, int | None]] = {}
        with open(master_csv_path, encoding="utf-8", newline="") as f:
            for row in csv.DictReader(_skip_comments(f)):
                code = row["process_code"].strip()
                category = row["category"].strip()
                is_bottleneck = row["is_bottleneck"].strip().lower() == "true"
                raw_lt = (row.get("standard_lt_business_days") or "").strip()
                lt = int(raw_lt) if raw_lt else None  # 未記入の場合はカテゴリ既定値にフォールバック
                self._code_map[code] = (category, is_bottleneck, lt)

        self._unknown_codes: set[str] = set()

    def categorize(self, process_code: str) -> CategoryResult:
        process_code = normalize_process_code(process_code)
        entry = self._code_map.get(process_code)
        if entry is None:
            self._unknown_codes.add(process_code)
            category = self._unknown_category
            is_bottleneck = False
            is_unknown = True
            lt = None
        else:
            category, is_bottleneck, lt = entry
            is_unknown = False

        standard_lt = lt if lt is not None else self._categories.get(category, self._categories[self._unknown_category])
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
