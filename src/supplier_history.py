"""仕入(外注)累積データから、図番×工程コードごとの仕入先実績を集計する。

kii-san提供の複数年分の仕入データExcel(基幹システムからのエクスポート)を
読み込み、「この図番のこの工程なら過去にどの仕入先が対応してきたか」を
集計する。①②で停滞している工程の「代替先候補」提示に使う(2026-09-08 設計合意)。

- 図番と工程コードの組み合わせが紐付けの最優先(kii-san指摘: 2026-09-08。
  図番だけで一致させると、その図番の別の工程を担当しただけの仕入先まで
  「代替候補」として出てしまい、実態と異なる提案になるため)。
  一致がなければ工程コード単位の一般的な実績にフォールバックする。
- 仕入単価・金額はセンシティブな財務情報のため、代替候補仕入先の提示・工程別仕入先実績
  ランキングなど「個別の受注・仕入先」が見える集計には一切含めない(件数・最終利用日のみ)。
  【2026-09-08 kii-san追加合意】仕入先の週別予測(調整が必要な仕入先の把握)に限り、
  「仕入先ごとの週次合計金額」のみを内部集計に使う(個別受注・個別仕入先の金額は
  一切表示しない。src/report_data.py の build_supplier_capacity_forecast 参照)。
- 工程コードが空欄の行(材料そのものの仕入と推測、kii-san確認済み)は集計対象外。
- 入力ファイルは受注データと異なり更新頻度が低く、かつ合計で数十万行規模になるため、
  集計結果のみをローカルキャッシュ(.cache/supplier_history.json、gitには含めない)に
  保存し、入力ファイル一式が変わらない限り再読み込みを省略する。
"""
from __future__ import annotations

import datetime
import json
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook

from src.process_master import normalize_process_code

DEFAULT_CACHE_PATH = Path(".cache/supplier_history.json")
# 集計ロジックを変更したら上げる。ソースファイルが同じでもキャッシュを無効化し、
# 古いロジックで集計された結果を誤って使い続けないようにするためのガード。
CACHE_SCHEMA_VERSION = 3

# 2022年のみ列数が少ない簡易フォーマット。2023年以降は55列共通フォーマット。
HEADER_19 = [
    "得意先ＣＤ", "仕入先ＣＤ", "担当CD", "商品名", "数量", "仕入単価", "金額（支払額）",
    "受注№", "行№", "情報ＮＯ", "注文№", "図番", "材質", "寸法", "伝票日付", "工程コード",
    "明細１", "明細２", "明細４",
]
HEADER_55 = [
    "部門ＣＤ", "得意先ＣＤ", "仕入先ＣＤ", "伝票日付(未)", "伝票種類", "伝票№", "行", "取引区分",
    "FILLER", "担当CD", "請求発行区分", "台帳発行区分", "翌月区分", "翌締区分", "商品グループ",
    "科目", "予備", "商品名", "単位", "数量", "仕入単価", "金額（支払額）", "予備２", "備考",
    "受注№", "行№", "情報ＮＯ", "注文№", "予備３", "完了区分", "図番", "貴社注文番号", "製品区分",
    "材質", "寸法", "仕入区分", "予備", "伝票日付", "予備", "工程コード", "予備", "更新", "抽出",
    "受注", "明細１", "明細２", "明細３", "明細４", "明細５", "明細６", "窓口ＣＤ", "消費税",
    "予備", "予備", "予備",
]


@dataclass
class SupplierSuggestion:
    supplier_code: str
    process_code: Optional[str]
    match_type: str  # "drawing_process"(図番+工程一致) or "process_code"(工程コードのみ一致)
    count: int
    last_used: Optional[datetime.date]

    @property
    def label(self) -> str:
        kind = "図番+工程一致" if self.match_type == "drawing_process" else "工程一致"
        last = self.last_used.isoformat() if self.last_used else "不明"
        return f"仕入先{self.supplier_code}({kind} {self.count}回, 最終{last})"


def _parse_date(value) -> Optional[datetime.date]:
    if not isinstance(value, int):
        return None
    text = str(value)
    if len(text) != 8:
        return None
    try:
        return datetime.date(int(text[0:4]), int(text[4:6]), int(text[6:8]))
    except ValueError:
        return None


def _iter_rows(path: Path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    header = HEADER_19 if ws.max_column <= 20 else HEADER_55
    idx = {name: i for i, name in enumerate(header)}
    for row in ws.iter_rows(min_row=2, values_only=True):
        yield idx, row


def _week_start(d: datetime.date) -> datetime.date:
    return d - datetime.timedelta(days=d.weekday())  # 月曜始まり


class SupplierHistory:
    def __init__(self) -> None:
        # by_drawing_process[(図番, 工程コード)][仕入先CD] = {"count": int, "last_used": date|None}
        self._by_drawing_process: dict[tuple[str, str], dict[str, dict]] = defaultdict(lambda: defaultdict(
            lambda: {"count": 0, "last_used": None}
        ))
        # by_process[工程コード][仕入先CD] = {"count": int, "last_used": date|None, "amount": float}
        # amountは仕入先週別予測(相対比較)専用の内部集計で、代替候補・ランキング表示には使わない。
        self._by_process: dict[str, dict[str, dict]] = defaultdict(lambda: defaultdict(
            lambda: {"count": 0, "last_used": None, "amount": 0.0}
        ))
        # weekly_amount_by_supplier[(週初め, 仕入先CD)] = 週次合計金額。仕入先週別予測の「普段の実績水準」に使う。
        self._weekly_amount_by_supplier: dict[tuple[datetime.date, str], float] = defaultdict(float)

    @classmethod
    def load(cls, paths: list[Path], cache_path: Path = DEFAULT_CACHE_PATH) -> "SupplierHistory":
        signature = {"__schema__": CACHE_SCHEMA_VERSION, **{str(p): p.stat().st_mtime for p in paths}}
        cached = _try_load_cache(cache_path, signature)
        if cached is not None:
            return cached

        history = cls()
        for path in paths:
            history._ingest_file(path)
        _save_cache(cache_path, signature, history)
        return history

    def _ingest_file(self, path: Path) -> None:
        for idx, row in _iter_rows(path):
            raw_code = row[idx["工程コード"]]
            if raw_code in (None, "", " "):
                continue  # 工程コード空欄=材料そのものの仕入。集計対象外(kii-san確認済み)
            process_code = normalize_process_code(raw_code)

            supplier_code = row[idx["仕入先ＣＤ"]]
            if supplier_code in (None, ""):
                continue
            supplier_code = str(supplier_code).strip()

            drawing_no_raw = row[idx["図番"]]
            drawing_no = str(drawing_no_raw).strip() if drawing_no_raw not in (None, "") else None

            date = _parse_date(row[idx["伝票日付"]])
            amount_raw = row[idx["金額（支払額）"]]
            amount = float(amount_raw) if isinstance(amount_raw, (int, float)) else 0.0

            proc_entry = self._by_process[process_code][supplier_code]
            proc_entry["count"] += 1
            proc_entry["amount"] += amount
            if date and (proc_entry["last_used"] is None or date > proc_entry["last_used"]):
                proc_entry["last_used"] = date

            if drawing_no:
                dp_entry = self._by_drawing_process[(drawing_no, process_code)][supplier_code]
                dp_entry["count"] += 1
                if date and (dp_entry["last_used"] is None or date > dp_entry["last_used"]):
                    dp_entry["last_used"] = date

            if date:
                self._weekly_amount_by_supplier[(_week_start(date), supplier_code)] += amount

    def suggest(self, drawing_no: Optional[str], process_code: Optional[str], top_n: int = 3) -> list[SupplierSuggestion]:
        if not process_code:
            return []
        normalized = normalize_process_code(process_code)

        if drawing_no:
            key = (drawing_no, normalized)
            if key in self._by_drawing_process:
                entries = self._by_drawing_process[key]
                ranked = sorted(entries.items(), key=lambda kv: kv[1]["count"], reverse=True)[:top_n]
                return [
                    SupplierSuggestion(
                        supplier_code=supplier, process_code=normalized,
                        match_type="drawing_process", count=data["count"], last_used=data["last_used"],
                    )
                    for supplier, data in ranked
                ]

        if normalized in self._by_process:
            entries = self._by_process[normalized]
            ranked = sorted(entries.items(), key=lambda kv: kv[1]["count"], reverse=True)[:top_n]
            return [
                SupplierSuggestion(
                    supplier_code=supplier, process_code=normalized,
                    match_type="process_code", count=data["count"], last_used=data["last_used"],
                )
                for supplier, data in ranked
            ]

        return []

    # --- 仕入先週別予測(report_data.build_supplier_capacity_forecastが使用) ---
    # 以下はいずれも「仕入先ごとの合計金額」のみを扱う(個別受注・個別行の金額は返さない)。

    def average_amount_for_process(self, process_code: str) -> Optional[float]:
        """工程コード1件あたりの典型的な仕入金額(平均)。仕入先週別予測の「大きさ」に使う。"""
        entries = self._by_process.get(normalize_process_code(process_code))
        if not entries:
            return None
        total_amount = sum(d["amount"] for d in entries.values())
        total_count = sum(d["count"] for d in entries.values())
        if not total_count:
            return None
        return total_amount / total_count

    def supplier_amount_shares_for_process(self, process_code: str) -> dict[str, float]:
        """工程コードの仕入先別シェア(実績金額比率、合計1.0)。仕入先週別予測の配分に使う。

        外注実績が一切ない工程コードは空辞書を返す(=按分先が無く、自然に予測対象外となる)。
        """
        entries = self._by_process.get(normalize_process_code(process_code))
        if not entries:
            return {}
        total = sum(d["amount"] for d in entries.values())
        if not total:
            return {}
        return {supplier: d["amount"] / total for supplier, d in entries.items() if d["amount"]}

    def weekly_amount_by_supplier(self) -> list[tuple[datetime.date, str, float]]:
        """(週初め, 仕入先CD, 実績金額合計) のリスト。週初め昇順。「普段の実績水準」の算出に使う。"""
        return sorted(
            ((week, supplier, amount) for (week, supplier), amount in self._weekly_amount_by_supplier.items()),
            key=lambda t: (t[0], t[1]),
        )

    def to_cache_dict(self) -> dict:
        return {
            "by_drawing_process": {
                f"{drawing}\x1f{code}": {
                    supplier: {
                        "count": data["count"],
                        "last_used": data["last_used"].isoformat() if data["last_used"] else None,
                    }
                    for supplier, data in suppliers.items()
                }
                for (drawing, code), suppliers in self._by_drawing_process.items()
            },
            "by_process": {
                code: {
                    supplier: {
                        "count": data["count"],
                        "last_used": data["last_used"].isoformat() if data["last_used"] else None,
                        "amount": data["amount"],
                    }
                    for supplier, data in suppliers.items()
                }
                for code, suppliers in self._by_process.items()
            },
            "weekly_amount_by_supplier": {
                f"{week.isoformat()}\x1f{supplier}": amount
                for (week, supplier), amount in self._weekly_amount_by_supplier.items()
            },
        }

    @classmethod
    def from_cache_dict(cls, payload: dict) -> "SupplierHistory":
        history = cls()
        for key, suppliers in payload["by_drawing_process"].items():
            drawing, code = key.split("\x1f", 1)
            for supplier, data in suppliers.items():
                entry = history._by_drawing_process[(drawing, code)][supplier]
                entry["count"] = data["count"]
                entry["last_used"] = datetime.date.fromisoformat(data["last_used"]) if data["last_used"] else None
        for code, suppliers in payload["by_process"].items():
            for supplier, data in suppliers.items():
                entry = history._by_process[code][supplier]
                entry["count"] = data["count"]
                entry["last_used"] = datetime.date.fromisoformat(data["last_used"]) if data["last_used"] else None
                entry["amount"] = data.get("amount", 0.0)
        for key, amount in payload.get("weekly_amount_by_supplier", {}).items():
            week_str, supplier = key.split("\x1f", 1)
            history._weekly_amount_by_supplier[(datetime.date.fromisoformat(week_str), supplier)] = amount
        return history


def _try_load_cache(cache_path: Path, signature: dict) -> Optional[SupplierHistory]:
    if not cache_path.exists():
        return None
    try:
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
    if payload.get("signature") != signature:
        return None
    return SupplierHistory.from_cache_dict(payload["data"])


def _save_cache(cache_path: Path, signature: dict, history: SupplierHistory) -> None:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"signature": signature, "data": history.to_cache_dict()}
    cache_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
