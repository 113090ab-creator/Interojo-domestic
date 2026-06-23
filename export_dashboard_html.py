from __future__ import annotations

import html
from pathlib import Path
from typing import Any

import pandas as pd

import dashboard


EXPORT_DIR = Path("exports")
EXPORT_FILE = EXPORT_DIR / "dashboard_static.html"


def fmt_int(value: Any) -> str:
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return "0"
    return f"{float(num):,.0f}"


def fmt_pct(value: Any) -> str:
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return "0.0%"
    return f"{float(num):,.1f}%"


def esc(value: Any) -> str:
    if value is None:
        return ""
    return html.escape(str(value), quote=True)


def table_html(df: pd.DataFrame, columns: list[str], max_rows: int = 20) -> str:
    view = df[[col for col in columns if col in df.columns]].head(max_rows).copy()
    if view.empty:
        return "<div class='empty'>표시할 데이터가 없습니다.</div>"

    numeric_like = {
        "요청 PACK",
        "용마입고 PACK",
        "미입고 PACK",
        "용마입고대기 PACK",
        "포장가능재고(PCS)",
        "생산부족 PCS",
        "재고수량",
    }
    percent_like = {"생산진도율", "용마입고율"}
    rows: list[str] = ["<table><thead><tr>"]
    for col in view.columns:
        rows.append(f"<th>{esc(col)}</th>")
    rows.append("</tr></thead><tbody>")
    for _, row in view.iterrows():
        rows.append("<tr>")
        for col in view.columns:
            value = row[col]
            klass = ""
            if col in percent_like:
                text = fmt_pct(value)
                klass = "num"
            elif col in numeric_like:
                text = fmt_int(value)
                klass = "num risk" if col in {"미입고 PACK", "생산부족 PCS"} and float(pd.to_numeric(value, errors="coerce") or 0) > 0 else "num"
            else:
                text = esc(value)
            rows.append(f"<td class='{klass}'>{text}</td>")
        rows.append("</tr>")
    rows.append("</tbody></table>")
    return "".join(rows)


def metric_card(label: str, value: str, tone: str = "") -> str:
    return f"""
    <div class="metric {tone}">
      <div class="metric-label">{esc(label)}</div>
      <div class="metric-value">{esc(value)}</div>
    </div>
    """


def status_board_html(kpi: dict[str, Any], operation_kpi: dict[str, Any], request_out_count: int, waiting_pcs: float) -> str:
    request_pack = float(pd.to_numeric(kpi.get("request_pack", 0), errors="coerce") or 0)
    yongma_in_pack = float(pd.to_numeric(kpi.get("yongma_in_pack", 0), errors="coerce") or 0)
    missing_pack = float(pd.to_numeric(kpi.get("shortage_pack", operation_kpi.get("packing_shortage_pack", 0)), errors="coerce") or 0)
    receipt_progress = float(pd.to_numeric(kpi.get("progress_pct", operation_kpi.get("packing_progress_pct", 0)), errors="coerce") or 0)
    production_progress = float(pd.to_numeric(operation_kpi.get("production_progress_pct", 0), errors="coerce") or 0)
    production_shortage = float(pd.to_numeric(operation_kpi.get("production_shortage_pcs", 0), errors="coerce") or 0)
    priority_products = int(pd.to_numeric(operation_kpi.get("priority_products", 0), errors="coerce") or 0)
    receipt_width = max(0.0, min(100.0, receipt_progress))
    missing_width = max(0.0, min(100.0 - receipt_width, 100.0))
    if request_out_count > 0 or priority_products > 0:
        tone = "risk"
        status = "우선 대응"
    elif missing_pack > 0 or production_shortage > 0:
        tone = "warn"
        status = "진행 관리"
    else:
        tone = "good"
        status = "정상"
    tiles = [
        ("용마입고율", fmt_pct(receipt_progress), ""),
        ("생산진도율", fmt_pct(production_progress), ""),
        ("미입고 PACK", fmt_int(missing_pack), "warn" if missing_pack > 0 else ""),
        ("생산부족 PCS", fmt_int(production_shortage), "warn" if production_shortage > 0 else ""),
        ("요청외 긴급", fmt_int(request_out_count), "warn" if request_out_count else ""),
        ("포장가능재고 PCS", fmt_int(waiting_pcs), "warn" if waiting_pcs else ""),
    ]
    tile_html = "".join(
        f"""
        <div class="status-tile {tile_tone}">
          <div class="metric-label">{esc(label)}</div>
          <div class="metric-value">{esc(value)}</div>
        </div>
        """
        for label, value, tile_tone in tiles
    )
    return f"""
    <div class="status-board {tone}">
      <div class="status-main">
        <div class="status-head"><span class="status-pill {tone}">{esc(status)}</span><strong>요청 대비 용마 입고 현황</strong></div>
        <div class="status-main-value">{receipt_progress:.1f}%</div>
        <div class="status-flow">
          <div class="status-flow-fill receipt" style="width:{receipt_width:.1f}%"></div>
          <div class="status-flow-fill shortage" style="width:{missing_width:.1f}%"></div>
        </div>
        <div class="status-flow-legend">
          <span>요청 {fmt_int(request_pack)} PACK</span>
          <span>입고 {fmt_int(yongma_in_pack)} PACK</span>
          <span>미입고 {fmt_int(missing_pack)} PACK</span>
        </div>
      </div>
      <div class="status-tile-grid">{tile_html}</div>
    </div>
    """


def family_cards_html(family_view: pd.DataFrame) -> str:
    if family_view.empty:
        return "<div class='empty'>분류 데이터가 없습니다.</div>"
    rows: list[str] = []
    for _, row in family_view.iterrows():
        request_pack = fmt_int(row.get("요청 PACK", 0))
        production = float(pd.to_numeric(row.get("생산진도율", 0), errors="coerce") or 0)
        receipt = float(pd.to_numeric(row.get("용마입고율", 0), errors="coerce") or 0)
        production_short = fmt_int(row.get("생산부족수량", 0))
        receipt_short = fmt_int(row.get("미입고수량", 0))
        rows.append(
            f"""
            <article class="family-card">
              <div class="family-head">
                <strong>{esc(row.get("본품분류", ""))}</strong>
                <span>요청 {request_pack} PACK</span>
              </div>
              <div class="bar-row"><span>생산</span><div class="bar"><i style="width:{max(0, min(production, 100)):.1f}%"></i></div><em>{production:.1f}%</em></div>
              <div class="bar-row"><span>입고</span><div class="bar"><i style="width:{max(0, min(receipt, 100)):.1f}%"></i></div><em>{receipt:.1f}%</em></div>
              <div class="family-foot"><span>생산부족 <b>{production_short}</b></span><span>미입고 <b>{receipt_short}</b></span></div>
            </article>
            """
        )
    return "".join(rows)


def build_html() -> str:
    files = dashboard.discover_source_files(Path("."))
    product_summary, code_summary, _lot_status_df, daily_inventory_df, sample_available_df = dashboard.load_dashboard_data(
        dashboard.file_fingerprint(files.request_file),
        dashboard.file_fingerprint(files.packing_file),
        dashboard.file_fingerprint(files.progress_file),
        dashboard.file_fingerprint(files.inventory_file),
        dashboard.file_fingerprint(files.daily_inventory_file),
    )

    kpi = dashboard.calc_kpi_from_code_summary(code_summary)
    operation_kpi = dashboard.calc_operation_kpis(
        product_summary,
        code_summary,
        dashboard.INVENTORY_STOCK_THRESHOLD_DEFAULT,
    )
    main_products, _sample_products = dashboard.split_main_sample(product_summary)
    family_view = dashboard.build_family_progress_view(main_products)
    daily_view = dashboard.build_daily_inventory_response_view(daily_inventory_df, code_summary, sample_available_df)
    exception_kpis, _exception_detail = dashboard.build_daily_exception_report_view(
        daily_inventory_df,
        code_summary,
        sample_available_df,
    )
    urgent = daily_view[daily_view["대응상태"].isin(["요청외 긴급", "요청내 긴급"])].copy()
    urgent = urgent.sort_values(["대응상태", "포장가능재고(PCS)", "생산부족 PCS"], ascending=[True, False, False], kind="stable")
    request_out_count = int(exception_kpis.get("request_out_count", 0.0))
    waiting_pcs = float(exception_kpis.get("waiting_pcs", 0.0))
    generated_at = pd.Timestamp.now(tz="Asia/Seoul").strftime("%Y-%m-%d %H:%M")

    source_list = [
        files.request_file,
        files.packing_file,
        files.progress_file,
        files.inventory_file,
        files.daily_inventory_file,
    ]
    source_html = "".join(f"<li>{esc(path.name if path else '(없음)')}</li>" for path in source_list)

    return f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>국내 생산·포장 현황 대시보드</title>
  <style>
    :root {{
      --bg:#e7e9ec; --panel:#ffffff; --line:rgba(62,74,88,0.16); --line-soft:rgba(62,74,88,0.08);
      --text:#202832; --muted:#6f7782; --muted2:#a5adb6;
      --blue:#24a6cf; --amber:#ffbf4a; --risk:#f5575c; --soft:#fff0f0; --section:#f2f4f6;
      --shadow:0 10px 22px rgba(38,52,67,0.08);
    }}
    * {{ box-sizing:border-box; }}
    body {{ margin:0; background:var(--bg); color:var(--text); font-family:"Noto Sans KR","Segoe UI",Arial,sans-serif; }}
    main {{ max-width:1440px; margin:0 auto; padding:28px 30px 48px; }}
    header {{ display:flex; justify-content:space-between; gap:20px; align-items:flex-end; margin-bottom:22px; }}
    h1 {{ margin:0; font-size:30px; line-height:1.2; letter-spacing:0; font-weight:800; }}
    h2 {{ margin:0 0 12px; font-size:18px; font-weight:800; }}
    .stamp {{ color:var(--muted); font-size:13px; }}
    .status-board {{ display:grid; grid-template-columns:minmax(320px,1.1fr) minmax(420px,1.9fr); gap:14px; margin:0 0 16px; }}
    .status-main, .status-tile {{ background:var(--panel); border:1px solid var(--line); border-radius:6px; box-shadow:var(--shadow); }}
    .status-main {{ padding:20px 22px; border-left:5px solid var(--blue); }}
    .status-board.warn .status-main {{ border-left-color:var(--amber); }}
    .status-board.risk .status-main {{ border-left-color:var(--risk); }}
    .status-head {{ display:flex; justify-content:space-between; gap:12px; align-items:center; margin-bottom:12px; }}
    .status-head strong {{ font-size:14px; font-weight:800; }}
    .status-pill {{ display:inline-flex; border-radius:999px; padding:5px 10px; font-size:11px; font-weight:800; background:#e7f8fc; color:#0c7c9a; }}
    .status-pill.warn {{ background:#fff5d8; color:#9d6700; }}
    .status-pill.risk {{ background:var(--soft); color:var(--risk); }}
    .status-main-value {{ font-size:38px; line-height:1; font-weight:900; color:var(--text); margin-bottom:16px; }}
    .status-flow {{ display:flex; height:14px; border-radius:999px; background:var(--section); overflow:hidden; }}
    .status-flow-fill.receipt {{ background:var(--blue); }}
    .status-flow-fill.shortage {{ background:var(--risk); }}
    .status-flow-legend {{ display:flex; justify-content:space-between; gap:10px; margin-top:10px; color:var(--muted); font-size:11px; }}
    .status-tile-grid {{ display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:12px; }}
    .status-tile {{ padding:14px 16px; min-height:86px; position:relative; overflow:hidden; }}
    .status-tile::before {{ content:""; position:absolute; top:0; left:0; right:0; height:5px; background:var(--blue); }}
    .status-tile::after {{
      content:""; position:absolute; right:14px; bottom:10px; width:76px; height:22px; opacity:.15; background:var(--blue);
      clip-path:polygon(0 70%,18% 46%,34% 58%,50% 28%,68% 44%,84% 16%,100% 34%,100% 100%,0 100%);
    }}
    .status-tile:nth-child(2)::before, .status-tile:nth-child(2)::after,
    .status-tile:nth-child(5)::before, .status-tile:nth-child(5)::after {{ background:var(--amber); }}
    .status-tile:nth-child(3)::before, .status-tile:nth-child(3)::after,
    .status-tile:nth-child(4)::before, .status-tile:nth-child(4)::after {{ background:var(--risk); }}
    .metrics {{ display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:12px; margin-bottom:18px; }}
    .metric {{ background:var(--panel); border:1px solid var(--line); border-radius:6px; padding:14px 16px; min-height:86px; box-shadow:var(--shadow); }}
    .metric.warn {{ background:var(--soft); border-color:#f3b8a8; }}
    .metric-label {{ color:var(--muted2); font-size:12px; margin-bottom:8px; font-weight:700; }}
    .metric-value {{ color:var(--blue); font-size:25px; font-weight:900; }}
    .metric.warn .metric-value, .status-tile.warn .metric-value, .risk {{ color:var(--risk); }}
    section {{ background:var(--panel); border:1px solid var(--line); border-radius:6px; padding:18px; margin-top:16px; box-shadow:var(--shadow); }}
    .family-grid {{ display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; }}
    .family-card {{ border:1px solid var(--line); border-radius:6px; padding:14px 12px 12px; background:#fff; position:relative; overflow:hidden; box-shadow:0 6px 14px rgba(38,52,67,.05); }}
    .family-card::before {{ content:""; position:absolute; top:0; left:0; right:0; height:4px; background:var(--blue); }}
    .family-head, .family-foot {{ display:flex; justify-content:space-between; gap:12px; align-items:center; }}
    .family-head span, .family-foot {{ color:var(--muted); font-size:12px; }}
    .family-foot b {{ color:var(--risk); }}
    .bar-row {{ display:grid; grid-template-columns:40px 1fr 52px; gap:8px; align-items:center; margin-top:12px; font-size:12px; }}
    .bar {{ height:8px; background:var(--section); border-radius:999px; overflow:hidden; }}
    .bar i {{ display:block; height:100%; background:var(--blue); }}
    .bar-row em {{ color:var(--muted); font-style:normal; text-align:right; }}
    table {{ width:100%; border-collapse:collapse; font-size:12px; }}
    th, td {{ border-bottom:1px solid #e5ebf1; padding:9px 8px; text-align:left; white-space:nowrap; }}
    th {{ color:var(--muted); font-weight:800; background:#f8fafc; position:sticky; top:0; }}
    td.num {{ text-align:right; font-variant-numeric:tabular-nums; }}
    .table-wrap {{ overflow:auto; border:1px solid var(--line); border-radius:6px; }}
    .empty {{ color:var(--muted); padding:20px; border:1px dashed var(--line); border-radius:6px; }}
    .sources {{ color:var(--muted); font-size:12px; }}
    @media (max-width:900px) {{
      main {{ padding:20px 14px 36px; }}
      header {{ display:block; }}
      .status-board {{ grid-template-columns:1fr; }}
      .status-tile-grid {{ grid-template-columns:repeat(2,minmax(0,1fr)); }}
      .metrics {{ grid-template-columns:repeat(2,minmax(0,1fr)); }}
      .family-grid {{ grid-template-columns:1fr; }}
      h1 {{ font-size:24px; }}
    }}
  </style>
</head>
<body>
  <main>
    <header>
      <div>
        <h1>국내 생산·포장 현황 대시보드</h1>
        <div class="stamp">정적 HTML 스냅샷 · 생성 {esc(generated_at)}</div>
      </div>
      <div class="stamp">기준: 요청물량, 수요정보, 용마입고, 일일 재고현황</div>
    </header>

    {status_board_html(kpi, operation_kpi, request_out_count, waiting_pcs)}

    <div class="metrics">
      {metric_card("요청 PACK", fmt_int(kpi.get("request_pack", 0)))}
      {metric_card("용마입고 PACK", fmt_int(kpi.get("yongma_in_pack", 0)))}
      {metric_card("미입고 PACK", fmt_int(operation_kpi.get("packing_shortage_pack", 0)), "warn")}
      {metric_card("생산부족 PCS", fmt_int(operation_kpi.get("production_shortage_pcs", 0)), "warn")}
      {metric_card("요청외 긴급 품목", fmt_int(request_out_count), "warn" if request_out_count else "")}
    </div>

    <section>
      <h2>본품 분류별 진도 현황</h2>
      <div class="family-grid">{family_cards_html(family_view)}</div>
    </section>

    <section>
      <h2>일일 재고 대응 우선 확인</h2>
      <div class="table-wrap">
        {table_html(urgent, ["대응상태", "품목코드", "제품명", "제품코드", "PACK", "POWER", "재고수량", "요청 PACK", "용마입고 PACK", "미입고 PACK", "용마입고대기 PACK", "포장가능재고(PCS)", "생산부족 PCS", "생산진도율", "최소 납기"], 30)}
      </div>
    </section>

    <section>
      <h2>소스</h2>
      <ul class="sources">{source_html}</ul>
    </section>
  </main>
</body>
</html>"""


def main() -> None:
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    EXPORT_FILE.write_text(build_html(), encoding="utf-8")
    print(EXPORT_FILE.resolve())


if __name__ == "__main__":
    main()
