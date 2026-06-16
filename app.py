from __future__ import annotations

from io import BytesIO

import pandas as pd
import streamlit as st


st.set_page_config(page_title="포장 진도 요약", layout="wide")
st.title("포장 진도 요약 대시보드")


REQUIRED_COLUMN_ALIASES = {
    "product_name": ["제품명", "제품 명", "품명", "product_name"],
    "production_code": ["생산코드", "생산 코드", "prod_code", "production_code"],
    "request_qty": ["요청수량", "요청 수량", "request_qty", "요청량"],
    "packing_qty": ["포장수량", "포장 수량", "packing_qty", "포장량"],
}


def normalize_name(text: str) -> str:
    return "".join(str(text).strip().lower().split())


def resolve_columns(df: pd.DataFrame) -> tuple[dict[str, str], list[str]]:
    normalized_to_original = {normalize_name(col): col for col in df.columns}
    resolved: dict[str, str] = {}
    missing: list[str] = []

    for key, aliases in REQUIRED_COLUMN_ALIASES.items():
        found = None
        for alias in aliases:
            normalized_alias = normalize_name(alias)
            if normalized_alias in normalized_to_original:
                found = normalized_to_original[normalized_alias]
                break
        if found is None:
            missing.append(key)
        else:
            resolved[key] = found
    return resolved, missing


def to_number(series: pd.Series) -> pd.Series:
    cleaned = series.astype(str).str.replace(",", "", regex=False).str.strip()
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def read_excel(uploaded_file: BytesIO, sheet_name: str) -> pd.DataFrame:
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("엑셀 파일을 업로드해 주세요.")
    st.stop()

try:
    excel = pd.ExcelFile(uploaded_file)
except Exception as exc:
    st.error(f"엑셀 파일을 읽을 수 없습니다: {exc}")
    st.stop()

sheet_name = st.selectbox("시트 선택", excel.sheet_names, index=0)

try:
    raw_df = read_excel(uploaded_file, sheet_name=sheet_name)
except Exception as exc:
    st.error(f"시트 데이터를 읽는 중 오류가 발생했습니다: {exc}")
    st.stop()

st.subheader("업로드 데이터(DataFrame)")
st.dataframe(raw_df, use_container_width=True, height=320)

resolved_columns, missing_keys = resolve_columns(raw_df)
if missing_keys:
    missing_desc = ", ".join(missing_keys)
    st.error(
        "필수 컬럼이 없습니다. "
        f"누락 항목: {missing_desc}\n"
        "필요 기준: 제품명, 생산코드, 요청수량, 포장수량"
    )
    st.stop()

work = raw_df[
    [
        resolved_columns["product_name"],
        resolved_columns["production_code"],
        resolved_columns["request_qty"],
        resolved_columns["packing_qty"],
    ]
].copy()

work.columns = ["product_name", "production_code", "request_qty", "packing_qty"]
work["product_name"] = work["product_name"].fillna("").astype(str).str.strip()
work["production_code"] = work["production_code"].fillna("").astype(str).str.strip()
work["request_qty"] = to_number(work["request_qty"])
work["packing_qty"] = to_number(work["packing_qty"])

summary_df = (
    work.groupby(["product_name", "production_code"], dropna=False, as_index=False)[
        ["request_qty", "packing_qty"]
    ]
    .sum()
    .sort_values(["product_name", "production_code"], kind="stable")
)
summary_df["shortage_qty"] = summary_df["request_qty"] - summary_df["packing_qty"]

st.subheader("요약 결과")
st.dataframe(summary_df, use_container_width=True, height=360)

shortage_df = summary_df.loc[summary_df["shortage_qty"] > 0].copy()
shortage_df = shortage_df.sort_values("shortage_qty", ascending=False, kind="stable")

st.subheader("부족 품목 (부족수량 > 0)")
if shortage_df.empty:
    st.success("부족수량이 0보다 큰 품목이 없습니다.")
else:
    st.dataframe(shortage_df, use_container_width=True, height=320)
