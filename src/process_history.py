"""工程累積データ(社内実績)から、図番×工程コードごとの設備実績・負荷を集計する。

kii-san提供の複数年分の工程実績Excel(基幹システムからのエクスポート、
1行=1受注の1工程ステップ、内外作区分は全行1=社内、作業完了は全行1=完了済み)
を読み込む。仕入累積データ(src/supplier_history.py)が外注実績を担当するのに
対し、こちらは以下を担当する(2026-09-08〜 設計合意)。

  1. 代替候補設備の提示(社内)
  2. 実績LTの参考表示(標準LTは変更しない)
  3. 図番を指定した際の実績ベースLT予測(工程順の中央値でルートを推定し、
     各工程の実績日数を合計する)
  4. 進行中の受注について、残り工程の実績LTを積み上げて客先納期に
     間に合いそうかを予測する(report_data.pyのfeasibility計算が使用)
  5. 週別・部署別/設備別の負荷(実績総工数)の可視化用データ

- 図番と工程コードの組み合わせが紐付けの最優先(kii-san指摘: 2026-09-08。
  図番だけで一致させると、その図番の別の工程を担当しただけの設備まで
  「代替候補」として出てしまい、実態と異なる提案になるため)。
  一致がなければ工程コード単位の一般的な実績にフォールバックする。
- 「予定機械」列が週次WIPデータの「加工先」と94%一致することを確認済みのため、
  設備の識別子としてこの列を使う。「加工先」列(idx8)は数値で種類数が13と少なく、
  部署/コストセンター的なコードと推測される(kii-san確認: 部署別負荷の軸として使用)。
- 標準LT(config/process_code_master.csv記載の値)は変更しない(kii-san指示)。
  実績LT(着手日→完成日の実日数)は、標準LTだけでは説明しきれない場合の
  「参考情報」としてのみ提供する(判定ロジックには使わない)。
- 負荷(週別工数)は「実績総工数」列を使う。週は完成日が属するISO週(月曜始まり)。
  着手日〜完成日にまたがる稼働を日別に按分してはいない(MVP。必要なら精緻化する)。
- 入力ファイルは仕入累積データと同様に更新頻度が低く数十万行規模になるため、
  集計結果のみをローカルキャッシュ(.cache/process_history.json、gitには含めない)に
  保存し、入力ファイル一式が変わらない限り再読み込みを省略する。
"""
from __future__ import annotations

import datetime
import json
import statistics
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook

from src.process_master import normalize_process_code

DEFAULT_CACHE_PATH = Path(".cache/process_history.json")
# 集計ロジックを変更したら上げる。ソースファイルが同じでもキャッシュを無効化し、
# 古いロジックで集計された結果を誤って使い続けないようにするためのガード。
CACHE_SCHEMA_VERSION = 4

IDX_ORDER_NO = 0
IDX_PROC_SEQ = 2  # 工程順
IDX_DRAWING = 3
IDX_PROC_CODE = 5
IDX_DEPARTMENT = 8  # 加工先。数値・13種類のみのため部署/コストセンターコードと推測
IDX_MACHINE = 11  # 予定機械。週次WIPデータの「加工先」と94%一致(確認済み)
IDX_ACTUAL_MANHOURS = 24  # 実績側の総工数
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


@dataclass
class RouteStepEstimate:
    process_code: str
    typical_seq: float  # 工程順の中央値(ルートの並び順に使用)
    actual_lt_calendar_days: Optional[float]
    sample_count: int
    is_drawing_specific: bool  # True=この図番自体の実績、False=工程コード単位の全図番平均にフォールバック


@dataclass
class DrawingLtEstimate:
    drawing_no: str
    steps: list[RouteStepEstimate] = field(default_factory=list)
    total_calendar_days: Optional[float] = None  # 実績が分かる工程だけの合計(下限値)
    missing_process_codes: list[str] = field(default_factory=list)

    @property
    def has_full_data(self) -> bool:
        return bool(self.steps) and not self.missing_process_codes


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


def _week_start(d: datetime.date) -> datetime.date:
    return d - datetime.timedelta(days=d.weekday())  # 月曜始まり


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
        # actual_durations[工程コード] = [実日数, ...](着手日→完成日。工程コード単位、フォールバック用)
        self._actual_durations: dict[str, list[int]] = defaultdict(list)
        # actual_durations_by_drawing_process[(図番, 工程コード)] = [実日数, ...](図番固有、優先use)
        self._actual_durations_by_drawing_process: dict[tuple[str, str], list[int]] = defaultdict(list)
        # seq_by_drawing_process[(図番, 工程コード)] = [工程順, ...](ルート推定の並び順用)
        self._seq_by_drawing_process: dict[tuple[str, str], list[int]] = defaultdict(list)
        # drawing_process_codes[図番] = {工程コード, ...}(estimate_route_for_drawingの高速化用インデックス)
        self._drawing_process_codes: dict[str, set[str]] = defaultdict(set)
        # weekly_load_by_department[(週初め, 部署コード)] = 実績工数合計
        self._weekly_load_by_department: dict[tuple[datetime.date, str], float] = defaultdict(float)
        # weekly_load_by_machine[(週初め, 設備コード)] = 実績工数合計
        self._weekly_load_by_machine: dict[tuple[datetime.date, str], float] = defaultdict(float)

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
            department = str(row[IDX_DEPARTMENT]).strip() if row[IDX_DEPARTMENT] is not None else ""

            drawing_no_raw = row[IDX_DRAWING]
            drawing_no = str(drawing_no_raw).strip() if drawing_no_raw not in (None, "") else None

            complete_date = _parse_date(row[IDX_ACTUAL_COMPLETE])

            if machine:
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
                    if drawing_no:
                        self._actual_durations_by_drawing_process[(drawing_no, process_code)].append(duration)

            if drawing_no:
                seq = row[IDX_PROC_SEQ]
                if isinstance(seq, (int, float)):
                    self._seq_by_drawing_process[(drawing_no, process_code)].append(int(seq))
                self._drawing_process_codes[drawing_no].add(process_code)

            if complete_date:
                week = _week_start(complete_date)
                manhours = row[IDX_ACTUAL_MANHOURS]
                if isinstance(manhours, (int, float)) and manhours > 0:
                    if department:
                        self._weekly_load_by_department[(week, department)] += manhours
                    if machine:
                        self._weekly_load_by_machine[(week, machine)] += manhours

    # --- 代替候補設備(社内) ---

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

    # --- 実績LT(参考情報) ---

    def actual_lt_calendar_days_median(
        self, process_code: str, drawing_no: Optional[str] = None
    ) -> Optional[float]:
        """実績リードタイム中央値(暦日、着手日〜完成日)。参考情報専用。

        drawing_noを指定し、その図番自体の実績があればそちらを優先する
        (代替候補と同じ優先順位の考え方)。無ければ工程コード単位にフォールバック。
        """
        normalized = normalize_process_code(process_code)
        if drawing_no:
            specific = self._actual_durations_by_drawing_process.get((drawing_no, normalized))
            if specific:
                return statistics.median(specific)

        durations = self._actual_durations.get(normalized)
        if not durations:
            return None
        return statistics.median(durations)

    def estimate_route_for_drawing(self, drawing_no: str) -> Optional[DrawingLtEstimate]:
        """図番の典型的な工程ルート(工程順の中央値でソート)と、各工程の実績LTを推定する。"""
        codes = self._drawing_process_codes.get(drawing_no)
        if not codes:
            return None

        steps: list[RouteStepEstimate] = []
        missing: list[str] = []
        total = 0.0
        any_known = False

        for code in codes:
            seqs = self._seq_by_drawing_process.get((drawing_no, code))
            typical_seq = statistics.median(seqs) if seqs else float("inf")

            specific = self._actual_durations_by_drawing_process.get((drawing_no, code))
            if specific:
                lt = statistics.median(specific)
                is_specific = True
                sample_count = len(specific)
            else:
                lt = self.actual_lt_calendar_days_median(code)
                is_specific = False
                sample_count = len(self._actual_durations.get(code, []))

            if lt is None:
                missing.append(code)
            else:
                total += lt
                any_known = True

            steps.append(RouteStepEstimate(code, typical_seq, lt, sample_count, is_specific))

        steps.sort(key=lambda s: s.typical_seq)
        return DrawingLtEstimate(
            drawing_no=drawing_no,
            steps=steps,
            total_calendar_days=total if any_known else None,
            missing_process_codes=missing,
        )

    # --- 週別負荷(可視化用) ---

    def weekly_load_by_department(self) -> list[tuple[datetime.date, str, float]]:
        """(週初め, 部署コード, 実績工数合計) のリスト。週初め昇順。"""
        return sorted(
            ((week, dept, hours) for (week, dept), hours in self._weekly_load_by_department.items()),
            key=lambda t: (t[0], t[1]),
        )

    def weekly_load_by_machine(self) -> list[tuple[datetime.date, str, float]]:
        """(週初め, 設備コード, 実績工数合計) のリスト。週初め昇順。"""
        return sorted(
            ((week, machine, hours) for (week, machine), hours in self._weekly_load_by_machine.items()),
            key=lambda t: (t[0], t[1]),
        )

    # --- キャッシュ ---

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
            "actual_durations_by_drawing_process": {
                f"{drawing}\x1f{code}": durations
                for (drawing, code), durations in self._actual_durations_by_drawing_process.items()
            },
            "seq_by_drawing_process": {
                f"{drawing}\x1f{code}": seqs
                for (drawing, code), seqs in self._seq_by_drawing_process.items()
            },
            "weekly_load_by_department": {
                f"{week.isoformat()}\x1f{dept}": hours
                for (week, dept), hours in self._weekly_load_by_department.items()
            },
            "weekly_load_by_machine": {
                f"{week.isoformat()}\x1f{machine}": hours
                for (week, machine), hours in self._weekly_load_by_machine.items()
            },
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
        for key, durations in payload["actual_durations_by_drawing_process"].items():
            drawing, code = key.split("\x1f", 1)
            history._actual_durations_by_drawing_process[(drawing, code)] = list(durations)
            history._drawing_process_codes[drawing].add(code)
        for key, seqs in payload["seq_by_drawing_process"].items():
            drawing, code = key.split("\x1f", 1)
            history._seq_by_drawing_process[(drawing, code)] = list(seqs)
            history._drawing_process_codes[drawing].add(code)
        for key, hours in payload["weekly_load_by_department"].items():
            week_iso, dept = key.split("\x1f", 1)
            history._weekly_load_by_department[(datetime.date.fromisoformat(week_iso), dept)] = hours
        for key, hours in payload["weekly_load_by_machine"].items():
            week_iso, machine = key.split("\x1f", 1)
            history._weekly_load_by_machine[(datetime.date.fromisoformat(week_iso), machine)] = hours
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
