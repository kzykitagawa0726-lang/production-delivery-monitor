"""オフライン単体HTMLレポート出力。

外部CDN・外部通信は一切使用しない。CSSはインラインで埋め込み、
工程別仕掛中ランキングの棒グラフはライブラリを使わずSVGで手描きする。
"""
from __future__ import annotations

import html
from pathlib import Path

from src.models import OrderRecord
from src.report_data import ReportData

TOP_N = 10
CHART_TOP_N = 15


def _esc(value) -> str:
    return html.escape("" if value is None else str(value))


def _top_list_rows(records: list[OrderRecord], kind: str) -> str:
    rows = []
    for r in records[:TOP_N]:
        if kind == "delayed":
            metric = f"{-r.remaining_business_days}営業日超過"
        else:
            metric = f"残{r.remaining_business_days}営業日"
        current = r.current_process
        current_label = f"{current.process_code}/{current.status}" if current else "(全工程完了)"
        alternatives = "; ".join(s.label for s in (*r.machine_suggestions, *r.supplier_suggestions))
        rows.append(
            "<tr>"
            f"<td>{_esc(r.order_no)}</td>"
            f"<td>{_esc(r.drawing_no)}</td>"
            f"<td>{_esc(r.product_name)}</td>"
            f"<td>{_esc(r.customer_order_no)}</td>"
            f"<td>{_esc(r.company_deadline)}</td>"
            f"<td>{_esc(metric)}</td>"
            f"<td>{_esc(current_label)}</td>"
            f"<td>{_esc(alternatives)}</td>"
            "</tr>"
        )
    if not rows:
        return '<tr><td colspan="8">該当なし</td></tr>'
    return "".join(rows)


WEEKLY_CHART_WEEKS = 16
WEEKLY_CHART_TOP_DEPARTMENTS = 8
LINE_COLORS = [
    "#4f7cff", "#e35d5d", "#f0ad4e", "#2ea043", "#9a6fd8",
    "#17a2b8", "#d6336c", "#6c757d",
]


def _weekly_load_chart_svg(data: ReportData) -> str:
    """直近WEEKLY_CHART_WEEKS週分の、部署別 週次実績工数の折れ線グラフ(上位のみ)。

    部署数(13種類)が多いため、直近期間内の合計工数が多い上位のみを表示し、
    全期間・全設備の詳細はExcelの「週別負荷」シートを参照する形にする。
    """
    points = data.weekly_load_by_department
    if not points:
        return "<p>工程累積データが指定されていません。</p>"

    weeks = sorted({week for week, _, _ in points})[-WEEKLY_CHART_WEEKS:]
    week_set = set(weeks)

    totals: dict[str, float] = {}
    grid: dict[tuple, float] = {}
    for week, dept, hours in points:
        if week in week_set:
            totals[dept] = totals.get(dept, 0.0) + hours
            grid[(week, dept)] = hours

    top_departments = [d for d, _ in sorted(totals.items(), key=lambda kv: kv[1], reverse=True)[:WEEKLY_CHART_TOP_DEPARTMENTS]]
    if not top_departments or not weeks:
        return "<p>直近期間の工程累積データがありません。</p>"

    chart_width, chart_height = 640, 260
    margin_left, margin_bottom, margin_top = 40, 30, 10
    plot_w = chart_width - margin_left - 10
    plot_h = chart_height - margin_top - margin_bottom
    max_hours = max((grid.get((w, d), 0.0) for w in weeks for d in top_departments), default=0.0) or 1.0
    step_x = plot_w / max(len(weeks) - 1, 1)

    def xy(i: int, hours: float) -> tuple[float, float]:
        x = margin_left + i * step_x
        y = margin_top + plot_h - (hours / max_hours) * plot_h
        return x, y

    lines = []
    legend = []
    for ci, dept in enumerate(top_departments):
        color = LINE_COLORS[ci % len(LINE_COLORS)]
        coords = [xy(i, grid.get((w, dept), 0.0)) for i, w in enumerate(weeks)]
        path = " ".join(f"{x:.1f},{y:.1f}" for x, y in coords)
        lines.append(f'<polyline points="{path}" fill="none" stroke="{color}" stroke-width="2"></polyline>')
        legend.append(
            f'<span style="display:inline-flex;align-items:center;gap:4px;margin-right:12px;">'
            f'<span style="width:10px;height:10px;background:{color};border-radius:2px;display:inline-block;"></span>'
            f"部署{_esc(dept)}</span>"
        )

    x_labels = []
    label_every = max(len(weeks) // 8, 1)
    for i, w in enumerate(weeks):
        if i % label_every == 0 or i == len(weeks) - 1:
            x, _ = xy(i, 0)
            x_labels.append(
                f'<text x="{x:.1f}" y="{chart_height - 8}" text-anchor="middle" class="chart-label">'
                f"{w.strftime('%m/%d')}</text>"
            )

    y_labels = []
    for frac in (0, 0.5, 1.0):
        y = margin_top + plot_h - frac * plot_h
        val = max_hours * frac
        y_labels.append(f'<text x="4" y="{y + 4:.1f}" class="chart-value">{val:.0f}</text>')

    svg = (
        f'<svg viewBox="0 0 {chart_width} {chart_height}" width="100%" height="{chart_height}" '
        'role="img" aria-label="週別部署別実績工数">'
        + "".join(y_labels) + "".join(x_labels) + "".join(lines)
        + "</svg>"
    )
    return svg + '<div style="margin-top:8px;">' + "".join(legend) + "</div>"


def _congestion_bar_chart_svg(data: ReportData) -> str:
    entries = data.congestion_ranking[:CHART_TOP_N]
    if not entries:
        return "<p>工程進捗データがありません。</p>"

    max_count = max(e.count for e in entries)
    bar_height = 24
    gap = 8
    label_width = 160
    chart_width = 480
    row_height = bar_height + gap
    svg_height = row_height * len(entries) + gap

    bars = []
    for i, e in enumerate(entries):
        y = gap + i * row_height
        bar_w = (e.count / max_count) * chart_width if max_count else 0
        color = "#e35d5d" if e.is_bottleneck else ("#9aa0a6" if e.is_unknown_code else "#4f7cff")
        label = f"{e.process_code} ({e.category})" if not e.is_unknown_code else f"{e.process_code} (未分類)"
        bars.append(
            f'<text x="{label_width - 8}" y="{y + bar_height * 0.7}" text-anchor="end" '
            f'class="chart-label">{_esc(label)}</text>'
            f'<rect x="{label_width}" y="{y}" width="{bar_w:.1f}" height="{bar_height}" fill="{color}" rx="3"></rect>'
            f'<text x="{label_width + bar_w + 6}" y="{y + bar_height * 0.7}" class="chart-value">{e.count}</text>'
        )

    total_width = label_width + chart_width + 60
    return (
        f'<svg viewBox="0 0 {total_width} {svg_height}" width="100%" height="{svg_height}" '
        'role="img" aria-label="工程別仕掛中件数">'
        + "".join(bars)
        + "</svg>"
    )


PAGE_TEMPLATE = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>生産管理支援ツール レポート({generated_at})</title>
<style>
  :root {{ color-scheme: light dark; }}
  body {{ font-family: "Segoe UI", "Hiragino Kaku Gothic ProN", Meiryo, sans-serif; margin: 0; padding: 24px;
          background: #f4f6f9; color: #1f2430; }}
  h1 {{ font-size: 1.4rem; margin-bottom: 4px; }}
  .subtitle {{ color: #6b7280; margin-bottom: 24px; }}
  .cards {{ display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 32px; }}
  .card {{ flex: 1; min-width: 180px; background: #fff; border-radius: 10px; padding: 18px 20px;
           box-shadow: 0 1px 3px rgba(0,0,0,0.08); border-top: 6px solid #ccc; }}
  .card.delayed {{ border-top-color: #e35d5d; }}
  .card.risk {{ border-top-color: #f0ad4e; }}
  .card.undetermined {{ border-top-color: #9aa0a6; }}
  .card.forecast {{ border-top-color: #4f7cff; }}
  .card.infeasible {{ border-top-color: #d6336c; }}
  .card .label {{ font-size: 0.9rem; color: #6b7280; }}
  .card .value {{ font-size: 2.2rem; font-weight: 700; margin-top: 4px; }}
  section {{ background: #fff; border-radius: 10px; padding: 20px 24px; margin-bottom: 24px;
             box-shadow: 0 1px 3px rgba(0,0,0,0.08); }}
  section h2 {{ margin-top: 0; font-size: 1.1rem; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 0.9rem; }}
  th, td {{ text-align: left; padding: 6px 10px; border-bottom: 1px solid #e5e7eb; }}
  th {{ background: #305496; color: #fff; }}
  .chart-label {{ font-size: 12px; fill: #1f2430; }}
  .chart-value {{ font-size: 12px; fill: #1f2430; }}
  .warning {{ color: #9c0006; font-size: 0.85rem; margin-top: 8px; }}
  .note {{ color: #6b7280; font-size: 0.85rem; margin-top: 8px; }}
  @media (prefers-color-scheme: dark) {{
    body {{ background: #14161a; color: #e5e7eb; }}
    .card, section {{ background: #1f2229; box-shadow: none; border: 1px solid #2c313a; }}
    th, td {{ border-bottom-color: #2c313a; }}
    .chart-label, .chart-value {{ fill: #e5e7eb; }}
  }}
</style>
</head>
<body>
  <h1>生産管理支援ツール レポート</h1>
  <p class="subtitle">出力日: {generated_at} / 対象件数: {total_count}件 ※本ファイルは外部通信を行いません</p>

  <div class="cards">
    <div class="card delayed"><div class="label">① 納期遅延(超過)</div><div class="value">{delayed_count}</div></div>
    <div class="card risk"><div class="label">② 納期遅延リスク(残5営業日以内)</div><div class="value">{risk_count}</div></div>
    <div class="card undetermined"><div class="label">判定不能・要確認</div><div class="value">{undetermined_count}</div></div>
    <div class="card forecast"><div class="label">(参考)内示・先行手配</div><div class="value">{forecast_count}</div></div>
    <div class="card infeasible"><div class="label">(参考)実績ベースで納期に間に合わない見込み</div><div class="value">{infeasible_count}</div></div>
  </div>

  <section>
    <h2>① 納期遅延 トップ{top_n}(超過日数順)</h2>
    <table>
      <thead><tr><th>製造オーダー№</th><th>図番</th><th>品名</th><th>客先注番</th><th>自社納期</th><th>超過</th><th>現在工程</th><th>代替候補</th></tr></thead>
      <tbody>{delayed_rows}</tbody>
    </table>
  </section>

  <section>
    <h2>② 納期遅延リスク トップ{top_n}(危険度順)</h2>
    <table>
      <thead><tr><th>製造オーダー№</th><th>図番</th><th>品名</th><th>客先注番</th><th>自社納期</th><th>残営業日</th><th>現在工程</th><th>代替候補</th></tr></thead>
      <tbody>{risk_rows}</tbody>
    </table>
  </section>

  <section>
    <h2>工程別 仕掛中件数(現在停滞している工程コード別)</h2>
    {congestion_chart}
    {unknown_warning}
  </section>

  <section>
    <h2>週別負荷 上位部署(直近{weekly_chart_weeks}週、実績工数)</h2>
    {weekly_load_chart}
    <p class="note">※ 全部署・全設備・全期間の詳細はExcelの「週別負荷」シートをご覧ください。</p>
  </section>
</body>
</html>
"""


def write_html_report(data: ReportData, output_path: Path) -> None:
    unknown_warning = ""
    if data.unknown_process_codes:
        unknown_warning = (
            f'<p class="warning">⚠ {len(data.unknown_process_codes)}件の工程コードが '
            "工程コード対応表(config/process_code_master.csv)に未登録のため「未分類」表示です。"
            "対応表が届き次第、正しいカテゴリ・ボトルネック区分に更新されます。</p>"
        )

    infeasible_count = sum(1 for e in data.feasibility_estimates if e.status == "間に合わない見込み")
    infeasible_display = infeasible_count if data.feasibility_estimates else "—"

    page = PAGE_TEMPLATE.format(
        generated_at=_esc(data.generated_at.isoformat()),
        total_count=data.total_count,
        delayed_count=len(data.delayed),
        risk_count=len(data.at_risk),
        undetermined_count=len(data.undetermined),
        forecast_count=data.forecast_order_count,
        infeasible_count=infeasible_display,
        top_n=TOP_N,
        delayed_rows=_top_list_rows(data.delayed, "delayed"),
        risk_rows=_top_list_rows(data.at_risk, "risk"),
        congestion_chart=_congestion_bar_chart_svg(data),
        unknown_warning=unknown_warning,
        weekly_chart_weeks=WEEKLY_CHART_WEEKS,
        weekly_load_chart=_weekly_load_chart_svg(data),
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(page, encoding="utf-8")
