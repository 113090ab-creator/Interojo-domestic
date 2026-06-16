from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.offline import plot


def discover_request_and_packing(base_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    request_df = None
    packing_df = None

    for file in sorted(base_dir.glob("*.xlsx")):
        if file.name.startswith("~$"):
            continue
        xls = pd.ExcelFile(file)
        if len(xls.sheet_names) != 1:
            continue
        df = pd.read_excel(file)
        if len(df.columns) == 28:
            request_df = df
        elif len(df.columns) == 40:
            packing_df = df

    if request_df is None or packing_df is None:
        raise FileNotFoundError("요청/포장 원본 엑셀 파일을 찾지 못했습니다.")
    return request_df, packing_df


def normalize_frames(request_raw: pd.DataFrame, packing_raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    rq = list(request_raw.columns)
    pk = list(packing_raw.columns)

    request = pd.DataFrame(
        {
            "request_qty": pd.to_numeric(request_raw[rq[1]], errors="coerce").fillna(0.0),
            "sales_code": request_raw[rq[2]].fillna("").astype(str).str.strip(),
            "due_date": pd.to_datetime(request_raw[rq[4]], errors="coerce"),
            "product_code": request_raw[rq[9]].fillna("미기재").astype(str).str.strip(),
            "product_name": request_raw[rq[10]].fillna("제품명 미기재").astype(str).str.strip(),
            "prod_code": request_raw[rq[20]].fillna("").astype(str).str.strip(),
            "category": request_raw[rq[14]].fillna("미분류").astype(str).str.strip(),
        }
    )

    packing = pd.DataFrame(
        {
            "mark_date": pd.to_datetime(packing_raw[pk[29]], errors="coerce"),
            "prod_code": packing_raw[pk[11]].fillna("").astype(str).str.strip(),
            "pack_qty": pd.to_numeric(packing_raw[pk[19]], errors="coerce"),
            "piece_qty": pd.to_numeric(packing_raw[pk[20]], errors="coerce"),
            "pack_unit": pd.to_numeric(packing_raw[pk[16]], errors="coerce"),
        }
    )

    fill_mask = (
        packing["pack_qty"].isna()
        & packing["piece_qty"].notna()
        & packing["pack_unit"].notna()
        & (packing["pack_unit"] > 0)
    )
    packing["pack_qty_resolved"] = packing["pack_qty"]
    packing.loc[fill_mask, "pack_qty_resolved"] = packing.loc[fill_mask, "piece_qty"] / packing.loc[fill_mask, "pack_unit"]
    packing["pack_qty_resolved"] = packing["pack_qty_resolved"].fillna(0.0)

    return request, packing


def allocate_fifo(request_df: pd.DataFrame, packed_by_prod: dict[str, float]) -> pd.DataFrame:
    work = request_df.copy()
    work["allocated_qty"] = 0.0
    work = work.sort_values(by=["prod_code", "due_date", "sales_code"], na_position="last").reset_index(drop=True)

    for prod_code, idx_group in work.groupby("prod_code", dropna=False).groups.items():
        remain = float(packed_by_prod.get(prod_code, 0.0))
        if remain <= 0:
            continue
        for idx in idx_group:
            req_qty = float(work.at[idx, "request_qty"])
            if req_qty <= 0:
                continue
            alloc = min(req_qty, remain)
            work.at[idx, "allocated_qty"] = alloc
            remain -= alloc
            if remain <= 0:
                break

    work["remain_qty"] = (work["request_qty"] - work["allocated_qty"]).clip(lower=0.0)
    work["progress_pct"] = np.where(work["request_qty"] > 0, work["allocated_qty"] / work["request_qty"] * 100, 0.0)
    return work


def build_report(base_dir: Path) -> Path:
    request_raw, packing_raw = discover_request_and_packing(base_dir)
    request, packing = normalize_frames(request_raw, packing_raw)

    packed_by_prod = packing.groupby("prod_code", dropna=False)["pack_qty_resolved"].sum().to_dict()
    allocation = allocate_fifo(request, packed_by_prod)

    product_summary = (
        allocation.groupby(["product_code", "product_name"], dropna=False)[["request_qty", "allocated_qty", "remain_qty"]]
        .sum()
        .reset_index()
    )
    product_summary["progress_pct"] = np.where(
        product_summary["request_qty"] > 0,
        product_summary["allocated_qty"] / product_summary["request_qty"] * 100,
        0.0,
    )
    product_summary["product_label"] = product_summary["product_code"] + " | " + product_summary["product_name"]
    product_summary = product_summary.sort_values("request_qty", ascending=False)

    total_req = float(product_summary["request_qty"].sum())
    total_done = float(product_summary["allocated_qty"].sum())
    total_remain = float(product_summary["remain_qty"].sum())
    total_progress = (total_done / total_req * 100) if total_req > 0 else 0.0

    today = pd.Timestamp.today().normalize()
    due_req = float(allocation.loc[allocation["due_date"] <= today, "request_qty"].sum())
    due_done = float(allocation.loc[allocation["due_date"] <= today, "allocated_qty"].sum())
    due_progress = (due_done / due_req * 100) if due_req > 0 else 0.0

    focus_n = min(20, len(product_summary))
    focus = product_summary.head(focus_n).copy()
    focus = focus.sort_values("remain_qty", ascending=True)

    fig_stack = go.Figure()
    fig_stack.add_trace(
        go.Bar(
            y=focus["product_label"],
            x=focus["allocated_qty"],
            name="포장량",
            orientation="h",
            marker_color="#1f9d8b",
            hovertemplate="%{y}<br>포장량 %{x:,.0f} PACK<extra></extra>",
        )
    )
    fig_stack.add_trace(
        go.Bar(
            y=focus["product_label"],
            x=focus["remain_qty"],
            name="잔량",
            orientation="h",
            marker_color="#d74f3d",
            hovertemplate="%{y}<br>잔량 %{x:,.0f} PACK<extra></extra>",
        )
    )
    fig_stack.update_layout(
        title=f"제품별 총량 현황 TOP {focus_n} (잔량 중심 정렬)",
        barmode="stack",
        xaxis_title="PACK",
        yaxis_title="제품",
        height=820,
        margin=dict(l=20, r=20, t=60, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0),
        template="plotly_white",
    )

    risk_n = min(15, len(product_summary))
    risk = product_summary.sort_values("remain_qty", ascending=False).head(risk_n).copy()
    risk["progress_pct"] = risk["progress_pct"].round(1)
    risk_fig = px.bar(
        risk.sort_values("remain_qty", ascending=True),
        x="remain_qty",
        y="product_label",
        orientation="h",
        color="progress_pct",
        color_continuous_scale="RdYlGn",
        title=f"리스크 제품 TOP {risk_n} (잔량 큰 순)",
        labels={"remain_qty": "잔량(PACK)", "product_label": "제품", "progress_pct": "진도율(%)"},
    )
    risk_fig.update_layout(
        height=620,
        margin=dict(l=20, r=20, t=60, b=20),
        coloraxis_colorbar=dict(title="진도율(%)"),
        template="plotly_white",
    )

    table_df = product_summary.rename(
        columns={
            "product_code": "제품코드",
            "product_name": "제품명",
            "request_qty": "요청량(PACK)",
            "allocated_qty": "포장량(PACK)",
            "remain_qty": "잔량(PACK)",
            "progress_pct": "진도율(%)",
        }
    )[["제품코드", "제품명", "요청량(PACK)", "포장량(PACK)", "잔량(PACK)", "진도율(%)"]]
    table_df["진도율(%)"] = table_df["진도율(%)"].round(1)
    table_html = table_df.to_html(index=False, border=0, classes="tbl")

    stack_div = plot(fig_stack, include_plotlyjs="cdn", output_type="div")
    risk_div = plot(risk_fig, include_plotlyjs=False, output_type="div")

    html = f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <title>제품별 총량 현황 리포트</title>
  <style>
    :root {{
      --bg: #f4f7fb;
      --card: #ffffff;
      --line: #dbe4ef;
      --text: #14273b;
      --sub: #5c6f84;
      --accent: #17476a;
    }}
    body {{
      margin: 22px;
      background: var(--bg);
      color: var(--text);
      font-family: "Malgun Gothic", "Noto Sans KR", sans-serif;
    }}
    h1 {{
      margin: 0 0 6px 0;
      font-size: 30px;
      color: #0f2f4a;
    }}
    .sub {{
      margin-bottom: 14px;
      color: var(--sub);
      font-size: 14px;
    }}
    .kpi-wrap {{
      display: grid;
      grid-template-columns: repeat(5, minmax(150px, 1fr));
      gap: 10px;
      margin-bottom: 12px;
    }}
    .card {{
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 12px 14px;
      box-shadow: 0 3px 14px rgba(16, 31, 48, 0.05);
    }}
    .kpi-title {{
      font-size: 12px;
      color: var(--sub);
      margin-bottom: 6px;
    }}
    .kpi-val {{
      font-size: 28px;
      font-weight: 800;
      color: var(--accent);
      line-height: 1.1;
    }}
    .row {{
      display: grid;
      grid-template-columns: 1.6fr 1fr;
      gap: 12px;
      align-items: start;
      margin-bottom: 12px;
    }}
    .tbl {{
      width: 100%;
      border-collapse: collapse;
      background: #fff;
    }}
    .tbl th, .tbl td {{
      border: 1px solid #e4ebf3;
      padding: 6px 8px;
      font-size: 13px;
    }}
    .tbl th {{
      position: sticky;
      top: 0;
      background: #eef4fb;
    }}
    .table-wrap {{
      max-height: 620px;
      overflow: auto;
      border: 1px solid #e4ebf3;
      border-radius: 10px;
    }}
    @media (max-width: 1200px) {{
      .kpi-wrap {{
        grid-template-columns: repeat(2, minmax(140px, 1fr));
      }}
      .row {{
        grid-template-columns: 1fr;
      }}
    }}
  </style>
</head>
<body>
  <h1>제품별 총량 현황</h1>
  <div class="sub">
    생성시각: {pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")} · 기준일: {today.date()} · 데이터 폴더: {base_dir}
  </div>

  <div class="kpi-wrap">
    <div class="card"><div class="kpi-title">요청 총량(PACK)</div><div class="kpi-val">{total_req:,.0f}</div></div>
    <div class="card"><div class="kpi-title">포장 총량(PACK)</div><div class="kpi-val">{total_done:,.0f}</div></div>
    <div class="card"><div class="kpi-title">잔량(PACK)</div><div class="kpi-val">{total_remain:,.0f}</div></div>
    <div class="card"><div class="kpi-title">총량 진도율</div><div class="kpi-val">{total_progress:.1f}%</div></div>
    <div class="card"><div class="kpi-title">당일 납기 진도율</div><div class="kpi-val">{due_progress:.1f}%</div></div>
  </div>

  <div class="card">{stack_div}</div>

  <div class="row">
    <div class="card">{risk_div}</div>
    <div class="card">
      <h3 style="margin-top:0;">제품별 상세</h3>
      <div class="table-wrap">{table_html}</div>
    </div>
  </div>
</body>
</html>
"""

    output = base_dir / "product_total_report.html"
    output.write_text(html, encoding="utf-8")
    return output


if __name__ == "__main__":
    out = build_report(Path.cwd())
    print(f"SAVED: {out}")
