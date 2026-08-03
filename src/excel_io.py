"""基幹システムからエクスポートされた受注データExcelの読み込み。

実データ確認の結果、列位置は固定(A〜H)で以下の通り(2026-08-03 kii-san確認):
  A: 図番/品名(改行区切り)
  B: 製造オーダー№(現行基幹システムキー)
  C: 客先注番
  D: 数量
  E: 自社納期\n受注日\nLT=...(1行目9文字が自社納期。2〜3行目は移管期間中は信頼できないため未使用)
  F: 残日(基幹システム側で計算済みの残営業日数。そのまま使用)
  G: 取引先NO
  H: 製番(旧基幹システムキー)

I列以降(工程進捗の状況欄など)は移管期間中でデータ品質が不安定なため使用しない
(kii-san了承済み)。データ品質が改善されたら再検討する。
"""
from __future__ import annotations

import datetime
import re
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook

from src.models import OrderRecord

COL_DRAWING_PRODUCT = 0  # A
COL_ORDER_NO = 1  # B
COL_CUSTOMER_ORDER_NO = 2  # C
COL_QTY = 3  # D
COL_DEADLINE_RAW = 4  # E
COL_REMAINING_DAYS = 5  # F
COL_CUSTOMER_NO = 6  # G
COL_OLD_SYSTEM_NO = 7  # H

# 将来「目標納期」列が追加された場合の拡張フック(現行エクスポートには存在しない)。
TARGET_DEADLINE_HEADER = "目標納期"

DEADLINE_LINE_PATTERN = re.compile(r"(\d{2})年(\d{1,2})月(\d{1,2})日")


def _clean_str(value) -> Optional[str]:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _parse_deadline(raw_e_value) -> Optional[datetime.date]:
    if raw_e_value is None:
        return None
    first_line = str(raw_e_value).split("\n", 1)[0]
    m = DEADLINE_LINE_PATTERN.match(first_line)
    if not m:
        return None
    yy, mm, dd = m.groups()
    try:
        return datetime.date(2000 + int(yy), int(mm), int(dd))
    except ValueError:
        return None


def _parse_qty(value) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _parse_remaining_business_days(value) -> Optional[int]:
    if value is None or value == "":
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _split_drawing_product(value) -> tuple[str, Optional[str]]:
    text = "" if value is None else str(value)
    parts = text.split("\n", 1)
    drawing_no = parts[0].strip()
    product_name = parts[1].strip() if len(parts) > 1 else None
    return drawing_no, (product_name or None)


def load_orders(path: Path) -> list[OrderRecord]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    rows = ws.iter_rows(min_row=1, values_only=True)
    header = next(rows, None)
    target_deadline_col = None
    if header is not None and TARGET_DEADLINE_HEADER in header:
        target_deadline_col = header.index(TARGET_DEADLINE_HEADER)

    orders: list[OrderRecord] = []
    for row in rows:
        order_no = _clean_str(row[COL_ORDER_NO])
        if order_no is None:
            continue  # 末尾の空行など

        drawing_no, product_name = _split_drawing_product(row[COL_DRAWING_PRODUCT])

        target_deadline = None
        if target_deadline_col is not None:
            raw_target = row[target_deadline_col]
            if isinstance(raw_target, (datetime.date, datetime.datetime)):
                target_deadline = raw_target.date() if isinstance(raw_target, datetime.datetime) else raw_target

        orders.append(
            OrderRecord(
                order_no=order_no,
                drawing_no=drawing_no,
                product_name=product_name,
                qty=_parse_qty(row[COL_QTY]),
                company_deadline=_parse_deadline(row[COL_DEADLINE_RAW]),
                remaining_business_days=_parse_remaining_business_days(row[COL_REMAINING_DAYS]),
                customer_order_no=_clean_str(row[COL_CUSTOMER_ORDER_NO]),
                customer_no=_clean_str(row[COL_CUSTOMER_NO]),
                old_system_no=_clean_str(row[COL_OLD_SYSTEM_NO]),
                target_deadline=target_deadline,
            )
        )

    return orders
