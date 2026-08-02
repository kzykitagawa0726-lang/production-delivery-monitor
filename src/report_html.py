"""オフライン単体HTMLレポート出力。

外部CDN・外部通信は一切使用しない。CSSはインラインで埋め込み、
グラフはSVGを直接描画する(Chart.js等の外部ライブラリ不使用)。
"""
from __future__ import annotations

import html
from pathlib import Path

from src.models import OrderRecord
from src.report_data import ReportData

TOP_N = 10


def _esc(value) -> str:
    return html.escape("" if value is None else str(value))


def _bar_chart_svg(congestion: list[dict], width: int = 720, bar_height: int = 26, gap: int = 8) -> str:
    if not congestion:
        return "<p>仕掛中データがありません。</p>"

    max_count = max(e["count"] for e in congestion)
    label_width = 160
    chart_width = width - label_width - 60
    height = len(congestion) * (bar_height + gap) + gap

    bars = []
    for i, entry in enumerate(congestion):
        y = gap + i * (bar_height + gap)
        bar_w = 0 if max_count == 0 else round((entry["count"] / max_count) * chart_width)
        color = "#e35d5d" if entry["is_bottleneck"] else "#4f7cff"
        label = _esc(entry["process_code"])
        category = _esc(entry.get("category") or "")
        bars.append(
            f'<text x="{label_width - 8}" y="{y + bar_height / 2 + 4}" text-anchor="end" '
            f'class="bar-label">{label}{" ★" if entry["is_bottleneck"] else ""}</text>'
            f'<rect x="{label_width}" y="{y}" width="{bar_w}" height="{bar_height}" fill="{color}" rx="3">'
            f'<title>{label}({category}): {entry["count"]}件</title></rect>'
            f'<text x="{label_width + bar_w + 6}" y="{y + bar_height / 2 + 4}" class="bar-value">{entry["count"]}</text>'
        )

    return (
        f'<svg viewBox="0 0 {width} {height}" width="100%" height="{height}" role="img" '
        f'aria-label="工程別仕掛件数の棒グラフ">{"".join(bars)}</svg>'
    )


def _top_list_rows(records: list[OrderRecord], kind: str) -> str:
    rows = []
    for r in records[:TOP_N]:
        if kind == "delayed":
            metric = f"{-r.remaining_business_days_to_deadline}営業日超過"
        else:
            metric = (
                f"残{r.remaining_business_days_to_deadline}営業日 / 必要{r.remaining_required_business_days}営業日"
            )
        rows.append(
            "<tr>"
            f"<td>{_esc(r.order_no)}</td>"
            f"<td>{_esc(r.drawing_no)}</td>"
            f"<td>{_esc(r.product_name)}</td>"
            f"<td>{_esc(r.assignee)}</td>"
            f"<td>{_esc(r.company_deadline)}</td>"
            f"<td>{_esc(metric)}</td>"
            f"<td>{_esc(r.process_code)}</td>"
            "</tr>"
        )
    if not rows:
        return '<tr><td colspan="7">該当なし</td></tr>'
    return "".join(rows)


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
  .card .label {{ font-size: 0.9rem; color: #6b7280; }}
  .card .value {{ font-size: 2.2rem; font-weight: 700; margin-top: 4px; }}
  section {{ background: #fff; border-radius: 10px; padding: 20px 24px; margin-bottom: 24px;
             box-shadow: 0 1px 3px rgba(0,0,0,0.08); }}
  section h2 {{ margin-top: 0; font-size: 1.1rem; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 0.9rem; }}
  th, td {{ text-align: left; padding: 6px 10px; border-bottom: 1px solid #e5e7eb; }}
  th {{ background: #305496; color: #fff; }}
  .bar-label {{ font-size: 12px; fill: #1f2430; }}
  .bar-value {{ font-size: 12px; fill: #1f2430; }}
  .warning {{ background: #fff3cd; border: 1px solid #ffe69c; padding: 10px 14px; border-radius: 6px;
              margin-bottom: 16px; }}
  @media (prefers-color-scheme: dark) {{
    body {{ background: #14161a; color: #e5e7eb; }}
    .card, section {{ background: #1f2229; box-shadow: none; border: 1px solid #2c313a; }}
    th, td {{ border-bottom-color: #2c313a; }}
    .bar-label, .bar-value {{ fill: #e5e7eb; }}
    .warning {{ background: #3a3320; border-color: #5c4f21; color: #f1e3b3; }}
  }}
</style>
</head>
<body>
  <h1>生産管理支援ツール レポート</h1>
  <p class="subtitle">出力日: {generated_at} / 対象件数(仕掛中): {total_count}件 ※本ファイルは外部通信を行いません</p>

  {warning_block}

  <div class="cards">
    <div class="card delayed"><div class="label">① 納期遅延(超過)</div><div class="value">{delayed_count}</div></div>
    <div class="card risk"><div class="label">② 納期遅延リスク</div><div class="value">{risk_count}</div></div>
    <div class="card undetermined"><div class="label">判定不能・要確認</div><div class="value">{undetermined_count}</div></div>
  </div>

  <section>
    <h2>工程別仕掛件数(混雑ランキング) ★=ボトルネック工程</h2>
    {bar_chart}
  </section>

  <section>
    <h2>① 納期遅延 トップ{top_n}</h2>
    <table>
      <thead><tr><th>受注No</th><th>図番</th><th>商品名</th><th>担当者</th><th>自社納期</th><th>超過</th><th>工程コード</th></tr></thead>
      <tbody>{delayed_rows}</tbody>
    </table>
  </section>

  <section>
    <h2>② 納期遅延リスク トップ{top_n}</h2>
    <table>
      <thead><tr><th>受注No</th><th>図番</th><th>商品名</th><th>担当者</th><th>自社納期</th><th>状況</th><th>工程コード</th></tr></thead>
      <tbody>{risk_rows}</tbody>
    </table>
  </section>
</body>
</html>
"""


def write_html_report(data: ReportData, output_path: Path) -> None:
    warning_block = ""
    if data.unknown_process_codes:
        codes = ", ".join(_esc(c) for c in data.unknown_process_codes)
        warning_block = f'<div class="warning">⚠ 未知の工程コードを検出しました(「その他」として集計): {codes}</div>'

    page = PAGE_TEMPLATE.format(
        generated_at=_esc(data.generated_at.isoformat()),
        total_count=data.total_count,
        warning_block=warning_block,
        delayed_count=len(data.delayed),
        risk_count=len(data.at_risk),
        undetermined_count=len(data.undetermined),
        bar_chart=_bar_chart_svg(data.congestion),
        top_n=TOP_N,
        delayed_rows=_top_list_rows(data.delayed, "delayed"),
        risk_rows=_top_list_rows(data.at_risk, "risk"),
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(page, encoding="utf-8")
