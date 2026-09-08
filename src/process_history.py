"""工程累積データ(社内実績)から、図番×工程コードごとの設備実績・負荷を集計する。

kii-san提供の複数年分の工程実績Excel(基幹システムからのエクスポート、
1行=1受注の1工程ステップ、内外作区分は全行1=社内、作業完了は全行1=完了済み)
を読み込む。仕入累積データ(src/supplier_history.py)が外注実績を担当するのに
対し、こちらは以下を担当する(2026-09-08〜 設計合意)。

  1. 代替候補設備の提示(社内)
  2. 実績LTの参考表示(標準LTは変更しない)
  3. 図番を指定した際の実績ベースLT予測
  4. 進行中の受注について、客先納期に間に合いそうかを予測する
     (report_data.pyのfeasibility計算が使用)
  5. 週別・部署別/設備別の負荷(実績総工数)の可視化用データ

  【2026-09-08 実データ検証で発覚した設計変更】3・4は当初、工程コードごとの
  実績LT(着手日〜完成日)を残り工程数分「単純合計」していたが、実データでは
  残り工程が10件超の受注で合計が数千日規模になる異常値が頻発した。各工程の
  実績日数には他案件との順番待ち時間が相当含まれており、工程コード単位の
  中央値を寄せ集めても、その受注が実際に要する通しの期間を表さないため。
  → kii-san合意により、図番単位で「受注ごとのインスタンス(受注№+行が
  同じ工程群)の最初の着手日〜最後の完成日」を1つの実績値とし、その中央値
  (order_level_lt_calendar_days_median)を使う設計に変更した。工程別の内訳
  (estimate_route_for_drawingのsteps)は参考情報として残すが、合計値の主役は
  図番単位の通し実績に置き換えている。

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
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook

from src.process_master import normalize_process_code

DEFAULT_CACHE_PATH = Path(".cache/process_history.json")
# 集計ロジックを変更したら上げる。ソースファイルが同じでもキャッシュを無効化し、
# 古いロジックで集計された結果を誤って使い続けないようにするためのガード。
CACHE_SCHEMA_VERSION = 6

IDX_ORDER_NO = 0
IDX_LINE = 1  # 行(受注№と組み合わせて1つの工程ルート=インスタンスを識別)
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
    steps: list[RouteStepEstimate] = field(default_factory=list)  # 工程別の内訳(参考情報)
    total_calendar_days: Optional[float] = None
    # True: 図番単位の「受注(最初の着手)〜完成(最後の完成)」実績の中央値(信頼度が高い)。
    # False: その実績が1件も無いためのフォールバックで、工程別実績LTの単純合計(残り工程が
    #        多い受注では過大評価になりやすいので注意)。
    total_is_order_level: bool = False
    total_sample_count: int = 0  # 合計値の根拠になった実績件数
    missing_process_codes: list[str] = field(default_factory=list)  # 工程別内訳で実績が無かった工程(参考)

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
        # order_level_durations_by_drawing[図番] = [実日数, ...]
        # (受注№+行が同じ工程群を1インスタンスとし、最初の着手日〜最後の完成日を1件の実績とする)
        self._order_level_durations_by_drawing: dict[str, list[int]] = defaultdict(list)
        # _ingest_file中の一時集計。全ファイル読み込み後にfinalizeして上記に変換し、破棄する。
        self._instance_bounds: dict[tuple[str, str], dict] = {}
        # department_counts_by_process[工程コード] = Counter(部署コード)(将来予測負荷で、工程コードの
        # 実績部署シェアを算出するために使う。設備別は既存のby_process[工程コード][設備]のcountを流用する)
        self._department_counts_by_process: dict[str, Counter[str]] = defaultdict(Counter)
        # manhours_by_process[工程コード] = [実績工数, ...](将来予測負荷で、工程1件あたりの
        # 典型的な工数を推定するために使う。平均を採用)
        self._manhours_by_process: dict[str, list[float]] = defaultdict(list)

    @classmethod
    def load(cls, paths: list[Path], cache_path: Path = DEFAULT_CACHE_PATH) -> "ProcessHistory":
        signature = {"__schema__": CACHE_SCHEMA_VERSION, **{str(p): p.stat().st_mtime for p in paths}}
        cached = _try_load_cache(cache_path, signature)
        if cached is not None:
            return cached

        history = cls()
        for path in paths:
            history._ingest_file(path)
        history._finalize_order_level_durations()
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

            manhours = row[IDX_ACTUAL_MANHOURS]
            has_manhours = isinstance(manhours, (int, float)) and manhours > 0

            if complete_date and has_manhours:
                week = _week_start(complete_date)
                if department:
                    self._weekly_load_by_department[(week, department)] += manhours
                if machine:
                    self._weekly_load_by_machine[(week, machine)] += manhours

            if department:
                self._department_counts_by_process[process_code][department] += 1
            if has_manhours:
                self._manhours_by_process[process_code].append(manhours)

            order_no = str(row[IDX_ORDER_NO]).strip() if row[IDX_ORDER_NO] not in (None, "") else None
            line = str(row[IDX_LINE]).strip() if row[IDX_LINE] not in (None, "") else None
            if order_no and line and (drawing_no or start_date or complete_date):
                key = (order_no, line)
                bounds = self._instance_bounds.setdefault(
                    key, {"drawing": None, "min_start": None, "max_complete": None}
                )
                if drawing_no and bounds["drawing"] is None:
                    bounds["drawing"] = drawing_no
                if start_date and (bounds["min_start"] is None or start_date < bounds["min_start"]):
                    bounds["min_start"] = start_date
                if complete_date and (bounds["max_complete"] is None or complete_date > bounds["max_complete"]):
                    bounds["max_complete"] = complete_date

    def _finalize_order_level_durations(self) -> None:
        """全ファイル読み込み後に1回呼ぶ。インスタンス(受注№+行)ごとの最初の着手日〜
        最後の完成日を、その図番の「受注〜完成」実績1件として確定する。"""
        for bounds in self._instance_bounds.values():
            drawing = bounds["drawing"]
            start = bounds["min_start"]
            complete = bounds["max_complete"]
            if drawing and start and complete and complete >= start:
                self._order_level_durations_by_drawing[drawing].append((complete - start).days)
        self._instance_bounds = {}  # 中間データは不要(キャッシュにも含めない)

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

    def order_level_lt_calendar_days_median(self, drawing_no: str) -> Optional[float]:
        """図番単位の「受注(最初の着手)〜完成(最後の完成)」実績日数の中央値(暦日)。

        工程別実績LTの単純合計とは違い、1受注インスタンスの通しの実績値そのものなので、
        工程数の多寡による過大評価が起きない。予測の主役として使う(kii-san合意: 2026-09-08)。
        """
        durations = self._order_level_durations_by_drawing.get(drawing_no)
        if not durations:
            return None
        return statistics.median(durations)

    def order_level_lt_sample_count(self, drawing_no: str) -> int:
        return len(self._order_level_durations_by_drawing.get(drawing_no, []))

    def estimate_route_for_drawing(self, drawing_no: str) -> Optional[DrawingLtEstimate]:
        """図番の典型的な工程ルート(工程順の中央値でソート)と、各工程の実績LTを推定する。

        合計値(total_calendar_days)は、図番単位の受注〜完成実績があればその中央値を
        優先して使う(total_is_order_level=True)。無ければ工程別実績LTの単純合計に
        フォールバックする(total_is_order_level=False。残り工程が多いほど過大評価になりやすい)。
        """
        codes = self._drawing_process_codes.get(drawing_no)
        order_level_durations = self._order_level_durations_by_drawing.get(drawing_no)
        if not codes and not order_level_durations:
            return None

        steps: list[RouteStepEstimate] = []
        missing: list[str] = []
        summed = 0.0
        any_known = False

        for code in codes or ():
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
                summed += lt
                any_known = True

            steps.append(RouteStepEstimate(code, typical_seq, lt, sample_count, is_specific))

        steps.sort(key=lambda s: s.typical_seq)

        if order_level_durations:
            total = statistics.median(order_level_durations)
            is_order_level = True
            total_sample_count = len(order_level_durations)
        else:
            total = summed if any_known else None
            is_order_level = False
            total_sample_count = 0

        return DrawingLtEstimate(
            drawing_no=drawing_no,
            steps=steps,
            total_calendar_days=total,
            total_is_order_level=is_order_level,
            total_sample_count=total_sample_count,
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

    # --- 将来予測負荷(report_data.build_capacity_forecastが使用) ---

    def average_manhours_for_process(self, process_code: str) -> Optional[float]:
        """工程コード1件あたりの典型的な実績工数(平均)。予測負荷の「大きさ」に使う。"""
        values = self._manhours_by_process.get(normalize_process_code(process_code))
        if not values:
            return None
        return statistics.mean(values)

    def department_shares_for_process(self, process_code: str) -> dict[str, float]:
        """工程コードの部署別シェア(実績件数比率、合計1.0)。理想は設備別だが(kii-san指摘)、
        設備は146種と多く1つの工程コードが多数の設備に分散するため、まず部署単位(13種)を主軸にする。
        """
        counts = self._department_counts_by_process.get(normalize_process_code(process_code))
        if not counts:
            return {}
        total = sum(counts.values())
        return {dept: n / total for dept, n in counts.items()}

    def machine_shares_for_process(self, process_code: str) -> dict[str, float]:
        """工程コードの設備別シェア(実績件数比率、合計1.0)。kii-san要望の設備単位の予測に使う。

        将来どの設備が空くかは予測できないため、過去の実績シェアに比例して配分する
        近似(1台に決め打ちしない)。
        """
        entries = self._by_process.get(normalize_process_code(process_code))
        if not entries:
            return {}
        total = sum(d["count"] for d in entries.values())
        if total == 0:
            return {}
        return {machine: d["count"] / total for machine, d in entries.items()}

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
            "order_level_durations_by_drawing": dict(self._order_level_durations_by_drawing),
            "department_counts_by_process": {
                code: dict(counts) for code, counts in self._department_counts_by_process.items()
            },
            "manhours_by_process": dict(self._manhours_by_process),
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
        for drawing, durations in payload.get("order_level_durations_by_drawing", {}).items():
            history._order_level_durations_by_drawing[drawing] = list(durations)
        for code, counts in payload.get("department_counts_by_process", {}).items():
            history._department_counts_by_process[code] = Counter(counts)
        for code, values in payload.get("manhours_by_process", {}).items():
            history._manhours_by_process[code] = list(values)
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
