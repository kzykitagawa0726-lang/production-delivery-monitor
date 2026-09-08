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

    page = PAGE_TEMPLATE.format(
        generated_at=_esc(data.generated_at.isoformat()),
        total_count=data.total_count,
        delayed_count=len(data.delayed),
        risk_count=len(data.at_risk),
        undetermined_count=len(data.undetermined),
        forecast_count=data.forecast_order_count,
        top_n=TOP_N,
        delayed_rows=_top_list_rows(data.delayed, "delayed"),
        risk_rows=_top_list_rows(data.at_risk, "risk"),
        congestion_chart=_congestion_bar_chart_svg(data),
        unknown_warning=unknown_warning,
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(page, encoding="utf-8")
