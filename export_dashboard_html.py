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
        "현재 재고수량",
        "부족수량",
        "용마입고대기 PACK",
        "포장가능재고(PCS)",
        "생산부족 PCS",
        "재고수량",
        "순위",
        "GAP",
    }
    percent_like = {"생산진도율", "포장진도율", "용마입고율"}
    decimal_like = {"GAP"}
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
            elif col in decimal_like:
                num = pd.to_numeric(value, errors="coerce")
                text = "0.0" if pd.isna(num) else f"{float(num):,.1f}"
                klass = "num risk" if not pd.isna(num) and float(num) > 0 else "num"
            elif col in numeric_like:
                text = fmt_int(value)
                num = pd.to_numeric(value, errors="coerce")
                is_risk = (
                    (col in {"미입고 PACK", "생산부족 PCS", "부족수량"} and not pd.isna(num) and float(num) > 0)
                    or (col in {"현재 재고수량", "재고수량"} and not pd.isna(num) and float(num) < 0)
                )
                klass = "num risk" if is_risk else "num"
            else:
                text = esc(value)
            rows.append(f"<td class='{klass}'>{text}</td>")
        rows.append("</tr>")
    rows.append("</tbody></table>")
    return "".join(rows)


def status_board_html(kpi: dict[str, Any], operation_kpi: dict[str, Any], request_out_count: int) -> str:
    request_pack = float(pd.to_numeric(kpi.get("request_pack", 0), errors="coerce") or 0)
    packing_pack = float(pd.to_numeric(kpi.get("packing_pack", 0), errors="coerce") or 0)
    yongma_in_pack = float(pd.to_numeric(kpi.get("yongma_in_pack", 0), errors="coerce") or 0)
    missing_pack = float(pd.to_numeric(kpi.get("shortage_pack", operation_kpi.get("packing_shortage_pack", 0)), errors="coerce") or 0)
    packing_progress = float(pd.to_numeric(operation_kpi.get("packing_progress_pct", (packing_pack / request_pack * 100.0) if request_pack > 0 else 0), errors="coerce") or 0)
    receipt_progress = float(pd.to_numeric(operation_kpi.get("receipt_progress_pct", kpi.get("progress_pct", 0)), errors="coerce") or 0)
    production_progress = float(pd.to_numeric(operation_kpi.get("production_progress_pct", 0), errors="coerce") or 0)
    production_shortage = float(pd.to_numeric(operation_kpi.get("production_shortage_pcs", 0), errors="coerce") or 0)
    priority_products = int(pd.to_numeric(operation_kpi.get("priority_products", 0), errors="coerce") or 0)
    receipt_width = max(0.0, min(100.0, receipt_progress))
    missing_width = max(0.0, min(100.0 - receipt_width, 100.0))
    if request_out_count > 0 or priority_products > 0:
        tone = "risk"
    elif missing_pack > 0 or production_shortage > 0:
        tone = "warn"
    else:
        tone = "good"
    tiles = [
        ("국내 요청량 PACK", fmt_int(request_pack), ""),
        ("생산진도율", fmt_pct(production_progress), ""),
        ("포장진도율", fmt_pct(packing_progress), ""),
        ("용마입고율", fmt_pct(receipt_progress), ""),
        ("미입고 PACK", fmt_int(missing_pack), "warn" if missing_pack > 0 else ""),
        ("긴급 대응 품목 수", fmt_int(request_out_count), "warn" if request_out_count else ""),
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
        <div class="status-head"><span class="status-pill {tone}">용마입고율</span><strong>요청 대비 용마 입고 현황</strong></div>
        <div class="status-main-value">{receipt_progress:.1f}%</div>
        <div class="status-flow">
          <div class="status-flow-fill receipt" style="width:{receipt_width:.1f}%"></div>
          <div class="status-flow-fill shortage" style="width:{missing_width:.1f}%"></div>
        </div>
        <div class="status-flow-legend">
          <span>요청 {fmt_int(request_pack)} PACK</span>
          <span>용마입고 {fmt_int(yongma_in_pack)} PACK</span>
          <span>미입고 {fmt_int(missing_pack)} PACK</span>
          <span>용마입고율 {receipt_progress:.1f}%</span>
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
        rows.append(
            f"""
            <article class="family-card">
              <div class="family-head">
                <strong>{esc(row.get("본품분류", ""))}</strong>
                <span>요청 {request_pack} PACK</span>
              </div>
              <div class="bar-row"><span>생산</span><div class="bar"><i style="width:{max(0, min(production, 100)):.1f}%"></i></div><em>{production:.1f}%</em></div>
              <div class="bar-row"><span>입고</span><div class="bar"><i style="width:{max(0, min(receipt, 100)):.1f}%"></i></div><em>{receipt:.1f}%</em></div>
              <div class="family-foot"><span>생산부족 PCS <b>{production_short}</b></span></div>
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
    top_shortage_view = dashboard.build_top_shortage_view(product_summary, top_n=10)
    gap_top_view = dashboard.build_gap_top_view(product_summary, top_n=10)
    exception_kpis, exception_detail = dashboard.build_daily_exception_report_view(
        daily_inventory_df,
        code_summary,
        sample_available_df,
        max_rows=10,
    )
    request_out_count = int(exception_kpis.get("request_out_count", 0.0))
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
      --bg:#f4f7fa; --panel:#ffffff; --line:#d8e0e8; --text:#172437; --muted:#657386;
      --primary:#24496f; --accent:#ff4b4b; --risk:#d94b22; --soft:#fff1ed;
      --blue:#24496f; --amber:#ba7517; --section:#e9eff5; --muted2:#657386;
    }}
    * {{ box-sizing:border-box; }}
    body {{ margin:0; background:var(--bg); color:var(--text); font-family:"Noto Sans KR","Segoe UI",Arial,sans-serif; }}
    main {{ max-width:1440px; margin:0 auto; padding:30px 32px 48px; }}
    header {{ display:flex; justify-content:space-between; gap:20px; align-items:flex-end; margin-bottom:22px; }}
    h1 {{ margin:0; font-size:32px; line-height:1.2; letter-spacing:0; }}
    h2 {{ margin:0 0 12px; font-size:18px; }}
    .stamp {{ color:var(--muted); font-size:13px; }}
    .lead {{ color:var(--muted); font-size:14px; margin:8px 0 18px; }}
    .status-board {{ display:grid; grid-template-columns:minmax(320px,1.1fr) minmax(420px,1.9fr); gap:10px; margin:0 0 12px; }}
    .status-main, .status-tile {{ background:var(--panel); border:1px solid var(--line); border-radius:8px; }}
    .status-main {{ padding:16px 18px; border-left:4px solid #9e9d99; }}
    .status-board.warn .status-main {{ border-left-color:var(--amber); }}
    .status-board.risk .status-main {{ border-left-color:var(--risk); }}
    .status-head {{ display:flex; justify-content:space-between; gap:12px; align-items:center; margin-bottom:12px; }}
    .status-head strong {{ font-size:14px; font-weight:800; }}
    .status-pill {{ display:inline-flex; border-radius:999px; padding:5px 10px; font-size:11px; font-weight:700; background:#eaf6f1; color:#1d9e75; }}
    .status-pill.warn {{ background:#f7efe3; color:#ba7517; }}
    .status-pill.risk {{ background:var(--soft); color:var(--risk); }}
    .status-main-value {{ font-size:34px; line-height:1; font-weight:900; color:var(--text); margin-bottom:12px; }}
    .status-flow {{ display:flex; height:10px; border-radius:999px; background:#e9eff5; overflow:hidden; }}
    .status-flow-fill.receipt {{ background:#1d9e75; }}
    .status-flow-fill.shortage {{ background:var(--risk); }}
    .status-flow-legend {{ display:flex; justify-content:space-between; flex-wrap:wrap; gap:10px; margin-top:9px; color:var(--muted); font-size:11px; }}
    .status-tile-grid {{ display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:8px; }}
    .status-tile {{ padding:11px 13px; min-height:72px; }}
    .metric-label {{ color:var(--muted); font-size:12px; margin-bottom:8px; }}
    .metric-value {{ color:var(--primary); font-size:25px; font-weight:800; }}
    .status-tile.warn .metric-value, .risk {{ color:var(--risk); }}
    section {{ background:var(--panel); border:1px solid var(--line); border-radius:8px; padding:14px; margin-top:12px; }}
    .family-grid {{ display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:8px; }}
    .family-card {{ border:1px solid var(--line); border-radius:8px; padding:11px 13px; background:#fff; }}
    .family-head, .family-foot {{ display:flex; justify-content:space-between; gap:12px; align-items:center; }}
    .family-head span, .family-foot {{ color:var(--muted); font-size:12px; }}
    .family-foot b {{ color:var(--risk); }}
    .bar-row {{ display:grid; grid-template-columns:40px 1fr 52px; gap:8px; align-items:center; margin-top:12px; font-size:12px; }}
    .bar {{ height:8px; background:#e9eff5; border-radius:999px; overflow:hidden; }}
    .bar i {{ display:block; height:100%; background:var(--primary); }}
    .bar-row em {{ color:var(--muted); font-style:normal; text-align:right; }}
    table {{ width:100%; border-collapse:collapse; font-size:12px; }}
    th, td {{ border-bottom:1px solid #e5ebf1; padding:9px 8px; text-align:left; white-space:nowrap; }}
    th {{ color:var(--muted); font-weight:700; background:#f8fafc; position:sticky; top:0; }}
    td.num {{ text-align:right; font-variant-numeric:tabular-nums; }}
    .table-wrap {{ overflow:auto; border:1px solid var(--line); border-radius:8px; }}
    .empty {{ color:var(--muted); padding:20px; border:1px dashed var(--line); border-radius:8px; }}
    .sources {{ color:var(--muted); font-size:12px; }}
    @media (max-width:900px) {{
      main {{ padding:20px 14px 36px; }}
      header {{ display:block; }}
      .status-board {{ grid-template-columns:1fr; }}
      .status-tile-grid {{ grid-template-columns:repeat(2,minmax(0,1fr)); }}
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
    <p class="lead">국내 요청 물량이 생산 → 포장 → 용마 입고까지 정상적으로 진행되고 있는지 확인하고, 부족 및 지연 품목을 우선 대응하기 위한 화면입니다.</p>

    {status_board_html(kpi, operation_kpi, request_out_count)}

    <section>
      <h2>미입고 TOP10</h2>
      <div class="table-wrap">
        {table_html(top_shortage_view, ["순위", "제품명", "미입고 PACK", "생산진도율", "포장진도율", "용마입고율", "추정 원인"], 10)}
      </div>
    </section>

    <section>
      <h2>본품 분류별 진도현황</h2>
      <div class="family-grid">{family_cards_html(family_view)}</div>
    </section>

    <section>
      <h2>생산완료 후 미입고 TOP10</h2>
      <div class="table-wrap">
        {table_html(gap_top_view, ["순위", "제품명", "생산진도율", "용마입고율", "GAP"], 10)}
      </div>
    </section>

    <section>
      <h2>요청 긴급 요약</h2>
      <div class="table-wrap">
        {table_html(exception_detail, ["품목코드", "제품명", "현재 재고수량", "부족수량", "포장가능재고(PCS)", "대응가능 여부"], 10)}
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
