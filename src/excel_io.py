"""基幹システムからエクスポートされた受注データExcelの読み込み。

実データの列名(不規則な空白等)・データ型・欠損状況を確認していないため、
現時点では未実装。Step 1(実データ確認)完了後にここへ実装する。
"""
from __future__ import annotations

from pathlib import Path

from src.models import OrderRecord


def load_orders(path: Path) -> list[OrderRecord]:
    raise NotImplementedError(
        "受注データExcelの列マッピングは実データ確認後に実装します(Step 1未完了)。 "
        f"入力ファイル: {path}"
    )
