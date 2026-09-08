"""工程累積データ(社内実績)から、図番×工程コードごとの設備実績を集計する。

kii-san提供の複数年分の工程実績Excel(基幹システムからのエクスポート、
1行=1受注の1工程ステップ、内外作区分は全行1=社内、作業完了は全行1=完了済み)
を読み込む。仕入累積データ(src/supplier_history.py)が外注実績を担当するのに
対し、こちらは社内設備の代替候補提示と、実績リードタイムの参考表示を担当する
(2026-09-08 設計合意)。

- 図番と工程コードの組み合わせが紐付けの最優先(kii-san指摘: 2026-09-08。
  図番だけで一致させると、その図番の別の工程を担当しただけの設備まで
  「代替候補」として出てしまい、実態と異なる提案になるため)。
  一致がなければ工程コード単位の一般的な実績にフォールバックする。
- 「予定機械」列が週次WIPデータの「加工先」と94%一致することを確認済みのため、
  設備の識別子としてこの列を使う(「加工先」列は数値の部門/得意先寄りのコードで別物)。
- 標準LT(config/process_code_master.csv記載の値)は変更しない(kii-san指示)。
  実績LT(着手日→完成日の実日数)は、標準LTだけでは説明しきれない場合の
  「参考情報」としてのみ提供する(判定ロジックには使わない)。
- 入力ファイルは仕入累積データと同様に更新頻度が低く数十万行規模になるため、
  集計結果のみをローカルキャッシュ(.cache/process_history.json、gitには含めない)に
  保存し、入力ファイル一式が変わらない限り再読み込みを省略する。
"""
from __future__ import annotations

import datetime
import json
import statistics
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook

from src.process_master import normalize_process_code

DEFAULT_CACHE_PATH = Path(".cache/process_history.json")
# 集計ロジックを変更したら上げる。ソースファイルが同じでもキャッシュを無効化し、
# 古いロジックで集計された結果を誤って使い続けないようにするためのガード。
CACHE_SCHEMA_VERSION = 3

IDX_ORDER_NO = 0
IDX_PROC_SEQ = 2
IDX_DRAWING = 3
IDX_PROC_CODE = 5
IDX_MACHINE = 11  # 予定機械。週次WIPデータの「加工先」と94%一致(確認済み)
IDX_START = 90  # 着手日(実作業開始日、kii-san確認済み)
IDX_ACTUAL_COMPLETE = 92  # 完成日


@dataclass
class MachineSuggestion:
    machine_code: str
    process_code: Optional[str]
    match_type: str  # "drawing_process"(図番+工程一致) or "process_code"(工程コードのみ一致)
    count: int
    last_used: Optional[datetime.date]

    @property
    def label(self) -> str:
        kind = "図番+工程一致" if self.match_type == "drawing_process" else "工程一致"
        last = self.last_used.isoformat() if self.last_used else "不明"
        return f"設備{self.machine_code}(社内, {kind} {self.count}回, 最終{last})"


def _parse_date(value) -> Optional[datetime.date]:
    if not isinstance(value, int) or value == 0:
        return None
    text = str(value)
    if len(text) != 8:
        return None
    try:
        return datetime.date(int(text[0:4]), int(text[4:6]), int(text[6:8]))
    except ValueError:
        return None


class ProcessHistory:
    def __init__(self) -> None:
        # by_drawing_process[(図番, 工程コード)][設備コード] = {"count": int, "last_used": date|None}
        self._by_drawing_process: dict[tuple[str, str], dict[str, dict]] = defaultdict(lambda: defaultdict(
            lambda: {"count": 0, "last_used": None}
        ))
        # by_process[工程コード][設備コード] = {"count": int, "last_used": date|None}
        self._by_process: dict[str, dict[str, dict]] = defaultdict(lambda: defaultdict(
            lambda: {"count": 0, "last_used": None}
        ))
        # actual_durations[工程コード] = [実日数, ...](着手日→完成日。参考情報専用)
        self._actual_durations: dict[str, list[int]] = defaultdict(list)

    @classmethod
    def load(cls, paths: list[Path], cache_path: Path = DEFAULT_CACHE_PATH) -> "ProcessHistory":
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
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            raw_code = row[IDX_PROC_CODE]
            if raw_code in (None, "", " "):
                continue
            process_code = normalize_process_code(raw_code)

            machine = str(row[IDX_MACHINE]).strip() if row[IDX_MACHINE] is not None else ""
            if not machine:
                continue

            drawing_no_raw = row[IDX_DRAWING]
            drawing_no = str(drawing_no_raw).strip() if drawing_no_raw not in (None, "") else None

            complete_date = _parse_date(row[IDX_ACTUAL_COMPLETE])

            proc_entry = self._by_process[process_code][machine]
            proc_entry["count"] += 1
            if complete_date and (proc_entry["last_used"] is None or complete_date > proc_entry["last_used"]):
                proc_entry["last_used"] = complete_date

            if drawing_no:
                dp_entry = self._by_drawing_process[(drawing_no, process_code)][machine]
                dp_entry["count"] += 1
                if complete_date and (dp_entry["last_used"] is None or complete_date > dp_entry["last_used"]):
                    dp_entry["last_used"] = complete_date

            start_date = _parse_date(row[IDX_START])
            if start_date and complete_date:
                duration = (complete_date - start_date).days
                if duration >= 0:
                    self._actual_durations[process_code].append(duration)

    def suggest_machines(
        self, drawing_no: Optional[str], process_code: Optional[str], top_n: int = 3
    ) -> list[MachineSuggestion]:
        if not process_code:
            return []
        normalized = normalize_process_code(process_code)

        if drawing_no:
            key = (drawing_no, normalized)
            if key in self._by_drawing_process:
                entries = self._by_drawing_process[key]
                ranked = sorted(entries.items(), key=lambda kv: kv[1]["count"], reverse=True)[:top_n]
                return [
                    MachineSuggestion(
                        machine_code=machine, process_code=normalized,
                        match_type="drawing_process", count=data["count"], last_used=data["last_used"],
                    )
                    for machine, data in ranked
                ]

        if normalized in self._by_process:
            entries = self._by_process[normalized]
            ranked = sorted(entries.items(), key=lambda kv: kv[1]["count"], reverse=True)[:top_n]
            return [
                MachineSuggestion(
                    machine_code=machine, process_code=normalized,
                    match_type="process_code", count=data["count"], last_used=data["last_used"],
                )
                for machine, data in ranked
            ]

        return []

    def actual_lt_calendar_days_median(self, process_code: str) -> Optional[float]:
        """工程コード別の実績リードタイム中央値(暦日、着手日〜完成日)。参考情報専用。"""
        durations = self._actual_durations.get(normalize_process_code(process_code))
        if not durations:
            return None
        return statistics.median(durations)

    def to_cache_dict(self) -> dict:
        return {
            "by_drawing_process": {
                f"{drawing}\x1f{code}": {
                    machine: {"count": d["count"], "last_used": d["last_used"].isoformat() if d["last_used"] else None}
                    for machine, d in machines.items()
                }
                for (drawing, code), machines in self._by_drawing_process.items()
            },
            "by_process": {
                code: {
                    machine: {"count": d["count"], "last_used": d["last_used"].isoformat() if d["last_used"] else None}
                    for machine, d in machines.items()
                }
                for code, machines in self._by_process.items()
            },
            "actual_durations": dict(self._actual_durations),
        }

    @classmethod
    def from_cache_dict(cls, payload: dict) -> "ProcessHistory":
        history = cls()
        for key, machines in payload["by_drawing_process"].items():
            drawing, code = key.split("\x1f", 1)
            for machine, d in machines.items():
                entry = history._by_drawing_process[(drawing, code)][machine]
                entry["count"] = d["count"]
                entry["last_used"] = datetime.date.fromisoformat(d["last_used"]) if d["last_used"] else None
        for code, machines in payload["by_process"].items():
            for machine, d in machines.items():
                entry = history._by_process[code][machine]
                entry["count"] = d["count"]
                entry["last_used"] = datetime.date.fromisoformat(d["last_used"]) if d["last_used"] else None
        for code, durations in payload["actual_durations"].items():
            history._actual_durations[code] = list(durations)
        return history


def _try_load_cache(cache_path: Path, signature: dict) -> Optional[ProcessHistory]:
    if not cache_path.exists():
        return None
    try:
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
    if payload.get("signature") != signature:
        return None
    return ProcessHistory.from_cache_dict(payload["data"])


def _save_cache(cache_path: Path, signature: dict, history: ProcessHistory) -> None:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"signature": signature, "data": history.to_cache_dict()}
    cache_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
