"""オフライン単体HTMLレポート出力。

外部CDN・外部通信は一切使用しない。CSSはインラインで埋め込む。
工程進捗データを使わない方針のため、工程別混雑グラフは本バージョンでは含まない。
"""
from __future__ import annotations

import html
from pathlib import Path

from src.models import OrderRecord
from src.report_data import ReportData

TOP_N = 10


def _esc(value) -> str:
    return html.escape("" if value is None else str(value))


def _top_list_rows(records: list[OrderRecord], kind: str) -> str:
    rows = []
    for r in records[:TOP_N]:
        if kind == "delayed":
            metric = f"{-r.remaining_business_days}営業日超過"
        else:
            metric = f"残{r.remaining_business_days}営業日"
        rows.append(
            "<tr>"
            f"<td>{_esc(r.order_no)}</td>"
            f"<td>{_esc(r.drawing_no)}</td>"
            f"<td>{_esc(r.product_name)}</td>"
            f"<td>{_esc(r.customer_order_no)}</td>"
            f"<td>{_esc(r.company_deadline)}</td>"
            f"<td>{_esc(metric)}</td>"
            "</tr>"
        )
    if not rows:
        return '<tr><td colspan="6">該当なし</td></tr>'
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
  @media (prefers-color-scheme: dark) {{
    body {{ background: #14161a; color: #e5e7eb; }}
    .card, section {{ background: #1f2229; box-shadow: none; border: 1px solid #2c313a; }}
    th, td {{ border-bottom-color: #2c313a; }}
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
  </div>

  <section>
    <h2>① 納期遅延 トップ{top_n}(超過日数順)</h2>
    <table>
      <thead><tr><th>製造オーダー№</th><th>図番</th><th>品名</th><th>客先注番</th><th>自社納期</th><th>超過</th></tr></thead>
      <tbody>{delayed_rows}</tbody>
    </table>
  </section>

  <section>
    <h2>② 納期遅延リスク トップ{top_n}(危険度順)</h2>
    <table>
      <thead><tr><th>製造オーダー№</th><th>図番</th><th>品名</th><th>客先注番</th><th>自社納期</th><th>残営業日</th></tr></thead>
      <tbody>{risk_rows}</tbody>
    </table>
  </section>
</body>
</html>
"""


def write_html_report(data: ReportData, output_path: Path) -> None:
    page = PAGE_TEMPLATE.format(
        generated_at=_esc(data.generated_at.isoformat()),
        total_count=data.total_count,
        delayed_count=len(data.delayed),
        risk_count=len(data.at_risk),
        undetermined_count=len(data.undetermined),
        top_n=TOP_N,
        delayed_rows=_top_list_rows(data.delayed, "delayed"),
        risk_rows=_top_list_rows(data.at_risk, "risk"),
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(page, encoding="utf-8")
