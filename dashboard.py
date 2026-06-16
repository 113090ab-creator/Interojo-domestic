from __future__ import annotations

from dataclasses import dataclass
from html import escape
from io import BytesIO
from pathlib import Path
import re
from typing import Any

import numpy as np
import pandas as pd
import plotly.express as px
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt
import streamlit as st


st.set_page_config(page_title="국내 제품 포장 운영 대시보드", layout="wide")


class DashboardConfigError(Exception):
    def __init__(self, messages: list[str]):
        super().__init__("\n".join(messages))
        self.messages = messages


@dataclass
class SourceFiles:
    request_file: Path
    packing_file: Path
    progress_file: Path | None = None


STATUS_ORDER = ["미착수", "진행중", "완료"]
SAMPLE_KEYWORDS = ["샘플"]
GROUP_ORDER = ["전체", "본품", "샘플", "PIA", "Clalen", "Toric", "1Day", "Color", "Monthly", "기타"]
MAIN_PRODUCT_FAMILY_ORDER = [
    "전체",
    "Clalen 1Day",
    "O2O2 1Day",
    "O2O2 D 컬러",
    "O2O2 D Micelia",
    "O2O2 Toric",
    "O2O2 Monthly",
    "O2O2 M Micelia",
    "Clear",
    "PIA 1Day",
    "PIA Monthly",
    "Iris 컬러",
    "Iris Toric",
    "Toric",
    "부자재/기타",
    "기타",
]
DETAIL_FAMILY_PLACEHOLDER = "본품분류 선택"
STANDARD_PACK_BUCKETS = ["5P", "10P", "30P", "80P", "90P"]
POWER_RE = re.compile(r"(-?\d+(?:\.\d+)?)\s*$")
CODE_KEY_RE = re.compile(r"[^0-9A-Za-z가-힣]+")
BASE_P_CODE_RE = re.compile(r"^(P\d+)")
PIA_TOKEN_RE = re.compile(r"\bPIA\b", re.IGNORECASE)
PACK_UNIT_RE = re.compile(r"(?:_|\b|\()(\d+(?:\.\d+)?)\s*(?:P|팩)?\)?$", re.IGNORECASE)
PACK_PREFIX_RE = re.compile(r"^(\d+(?:\.\d+)?)\s*(?:P|팩)_", re.IGNORECASE)
PACK_ANY_RE = re.compile(r"(?:^|[^0-9A-Za-z가-힣])(\d+(?:\.\d+)?)\s*(?:P|팩)(?:$|[^0-9A-Za-z가-힣])", re.IGNORECASE)
PACK_SUFFIX_RE = re.compile(
    r"(?:_샘플\(\d+(?:\.\d+)?P\)|_\d+(?:\.\d+)?\s*(?:P|팩)|_\d+(?:\.\d+)?)$",
    re.IGNORECASE,
)
PACK_PREFIX_SUFFIX_RE = re.compile(r"^\d+(?:\.\d+)?\s*(?:P|팩)_", re.IGNORECASE)

NAVY = "#1f3556"
SOFT_NAVY = "#365276"
WHITE = "#ffffff"
LIGHT_GRAY = "#eef2f6"
MID_GRAY = "#d7dee7"
TEXT_DARK = "#1d2836"
TEXT_MUTED = "#5f6d80"
MUTED_ORANGE = "#c67a3d"
MUTED_RED = "#b55a4a"

REQUEST_COLS = {
    "sales_code": ["판매코드", "판매 코드", "품목코드", "sales_code"],
    "product_name": ["제품명", "제품 명", "품명", "product_name"],
    "request_qty": [
        "수량(PACK)",
        "수량 (PACK)",
        "요청 PACK",
        "요청수량",
        "수량",
        "request_qty",
    ],
    "request_pcs": ["수량(PCS)", "수량 (PCS)", "요청 PCS", "요청수량(PCS)", "요청물량 PCS"],
    "units_per_pack": ["입수(낱개)", "입수", "팩당수량", "pack_size"],
    "due_date": ["납기일자", "납기 일자", "납기일", "due_date"],
    "product_name_code": ["제품명코드", "제품명 코드", "제품규격", "품목코드"],
    "production_code": ["생산코드", "생산 코드", "제품코드", "제품 코드", "production_code"],
    "p_code": ["P코드(생산)", "P 코드(생산)", "P코드", "P 코드", "P code"],
    "q_code": ["Q코드(분리)", "Q 코드(분리)", "Q코드", "분리코드", "Q 코드"],
    "r_code": ["R코드(사출)", "R 코드(사출)", "R코드", "사출코드", "R 코드"],
    "market_type": ["국내/해외", "국내해외", "시장구분", "market_type"],
    "customer_name": ["거래처", "거래처명", "고객명", "고객 이름", "customer_name"],
}

PACKING_COLS = {
    "sales_code": ["판매코드", "판매 코드", "품목코드", "sales_code"],
    "packing_qty": ["팩수량", "포장수량", "포장완료수량", "수량", "packing_qty"],
}

PROCESS_STEPS = [
    {
        "id": "10",
        "header": "[10]사출조립",
        "label": "[10] 사출조립",
        "qty_col": "proc_10_qty",
        "due_col": "proc_10_due",
        "progress_pct": 20.0,
    },
    {
        "id": "20",
        "header": "[20]분리",
        "label": "[20] 분리",
        "qty_col": "proc_20_qty",
        "due_col": "proc_20_due",
        "progress_pct": 40.0,
    },
    {
        "id": "45",
        "header": "[45]하이드레이션/전면검사",
        "label": "[45] 하이드레이션/전면검사",
        "qty_col": "proc_45_qty",
        "due_col": "proc_45_due",
        "progress_pct": 60.0,
    },
    {
        "id": "55",
        "header": "[55]접착/멸균",
        "label": "[55] 접착/멸균",
        "qty_col": "proc_55_qty",
        "due_col": "proc_55_due",
        "progress_pct": 80.0,
    },
    {
        "id": "80",
        "header": "[80]누수/규격검사",
        "label": "[80] 누수/규격검사",
        "qty_col": "proc_80_qty",
        "due_col": "proc_80_due",
        "progress_pct": 100.0,
    },
]


def normalize_col(text: Any) -> str:
    return (
        str(text)
        .strip()
        .lower()
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
        .replace("(", "")
        .replace(")", "")
    )


def clean_str(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_match_key(value: Any) -> str:
    return CODE_KEY_RE.sub("", clean_str(value)).upper()


def extract_pack_unit(value: Any) -> float:
    text = clean_str(value)
    if not text:
        return np.nan
    match = PACK_UNIT_RE.search(text)
    if not match:
        match = PACK_PREFIX_RE.search(text)
    if not match:
        return np.nan
    try:
        return float(match.group(1))
    except ValueError:
        return np.nan


def strip_pack_unit_suffix(value: Any) -> str:
    text = clean_str(value)
    stripped = PACK_PREFIX_SUFFIX_RE.sub("", text)
    stripped = PACK_SUFFIX_RE.sub("", stripped).strip("_ -")
    return stripped or text


def format_pack_unit_label(unit: Any, product_name: Any = "") -> str:
    num = pd.to_numeric(unit, errors="coerce")
    if pd.isna(num) or float(num) <= 0:
        return "(미기재)"
    value = f"{float(num):g}P"
    if "샘플" in clean_str(product_name):
        return f"{value} 샘플"
    return value


def extract_base_p_code_key(value: Any) -> str:
    match = BASE_P_CODE_RE.match(normalize_match_key(value))
    return match.group(1) if match else ""


def build_first_value_map(df: pd.DataFrame, key_col: str, value_col: str) -> dict[str, str]:
    if key_col not in df.columns or value_col not in df.columns:
        return {}
    pairs = df[[key_col, value_col]].copy()
    pairs[key_col] = pairs[key_col].map(clean_str)
    pairs[value_col] = pairs[value_col].map(clean_str)
    pairs = pairs[(pairs[key_col] != "") & (pairs[value_col] != "")]
    if pairs.empty:
        return {}
    return pairs.drop_duplicates(key_col, keep="first").set_index(key_col)[value_col].to_dict()


def min_datetime(series: pd.Series) -> pd.Timestamp:
    dates = pd.to_datetime(series, errors="coerce")
    dates = dates.dropna()
    if dates.empty:
        return pd.NaT
    return dates.min()


def first_nonempty(series: pd.Series) -> str:
    for value in series:
        text = clean_str(value)
        if text:
            return text
    return ""


def join_unique(series: pd.Series, limit: int = 3) -> str:
    values = [clean_str(value) for value in series if clean_str(value)]
    unique = list(dict.fromkeys(values))
    if not unique:
        return ""
    if len(unique) <= limit:
        return ", ".join(unique)
    return f"{', '.join(unique[:limit])} 외 {len(unique) - limit}"


def to_number(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def find_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    normalized = {normalize_col(col): col for col in df.columns}
    for alias in aliases:
        key = normalize_col(alias)
        if key in normalized:
            return normalized[key]
    return None


def resolve_columns(
    df: pd.DataFrame,
    alias_map: dict[str, list[str]],
    required_keys: list[str],
    file_label: str,
) -> dict[str, str]:
    resolved: dict[str, str] = {}
    missing: list[str] = []
    for key, aliases in alias_map.items():
        col = find_column(df, aliases)
        if col is None:
            if key in required_keys:
                missing.append(f"{key} (후보: {', '.join(aliases)})")
        else:
            resolved[key] = col
    if missing:
        raise DashboardConfigError([f"[{file_label}] 필수 컬럼 누락: {'; '.join(missing)}"])
    return resolved


def list_excel_files(search_roots: list[Path]) -> list[Path]:
    seen: set[Path] = set()
    files: list[Path] = []
    for root in search_roots:
        if not root.exists():
            continue
        for pattern in ("*.xlsx", "*.xls"):
            for path in root.glob(pattern):
                if not path.is_file() or path.name.startswith("~$"):
                    continue
                real = path.resolve()
                if real in seen:
                    continue
                seen.add(real)
                files.append(path)
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return files


def pick_latest_by_name(files: list[Path], keywords: list[str]) -> Path | None:
    for file in files:
        stem = file.stem.lower()
        if any(keyword.lower() in stem for keyword in keywords):
            return file
    return None


def has_alias(columns: list[str], aliases: list[str]) -> bool:
    normalized_cols = {normalize_col(col) for col in columns}
    return any(normalize_col(alias) in normalized_cols for alias in aliases)


def discover_source_files(base_dir: Path) -> SourceFiles:
    files = list_excel_files([base_dir / "data", base_dir])
    if not files:
        raise DashboardConfigError(
            [
                "엑셀 파일을 찾지 못했습니다.",
                f"- 검색 위치: {base_dir / 'data'}",
                f"- 검색 위치: {base_dir}",
            ]
        )

    request_file = pick_latest_by_name(files, ["생산요청등록", "국내", "요청"])
    packing_file = pick_latest_by_name(files, ["포장설비투입현황", "포장설비투입", "포장실적", "포장"])
    progress_file = pick_latest_by_name(files, ["수요정보", "전공정"])

    if request_file is None or packing_file is None:
        for file in files:
            try:
                cols = pd.read_excel(file, nrows=0).columns.astype(str).tolist()
            except Exception:
                continue
            if request_file is None:
                if has_alias(cols, REQUEST_COLS["sales_code"]) and has_alias(cols, REQUEST_COLS["product_name"]) and has_alias(
                    cols, REQUEST_COLS["request_qty"]
                ):
                    request_file = file
            if packing_file is None:
                if has_alias(cols, PACKING_COLS["sales_code"]) and has_alias(cols, PACKING_COLS["packing_qty"]):
                    packing_file = file

    messages: list[str] = []
    if request_file is None:
        messages.append("생산요청등록(국내) 파일을 찾지 못했습니다.")
    if packing_file is None:
        messages.append("포장설비투입현황 파일을 찾지 못했습니다.")
    if messages:
        raise DashboardConfigError(messages)

    return SourceFiles(request_file=request_file, packing_file=packing_file, progress_file=progress_file)


def normalize_request(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path)
    cols = resolve_columns(
        raw,
        REQUEST_COLS,
        required_keys=["sales_code", "product_name", "request_qty"],
        file_label=path.name,
    )
    out = pd.DataFrame(
        {
            "sales_code": raw[cols["sales_code"]].map(clean_str),
            "product_name": raw[cols["product_name"]].map(clean_str).replace("", "(제품명 미기재)"),
            "request_pack": to_number(raw[cols["request_qty"]]),
        }
    )
    request_pcs = to_number(raw[cols["request_pcs"]]) if "request_pcs" in cols else pd.Series(0.0, index=raw.index)
    raw_units_per_pack = (
        to_number(raw[cols["units_per_pack"]])
        if "units_per_pack" in cols
        else pd.Series(np.nan, index=raw.index)
    )
    name_units_per_pack = out["product_name"].map(extract_pack_unit)
    units_per_pack = raw_units_per_pack.where(raw_units_per_pack > 0, name_units_per_pack)
    out["pack_unit"] = units_per_pack.where(units_per_pack > 0, np.nan)
    out["pack_unit_label"] = [
        format_pack_unit_label(unit, name)
        for unit, name in zip(out["pack_unit"], out["product_name"])
    ]
    out["base_product_name"] = out["product_name"].map(strip_pack_unit_suffix)
    fallback_pcs = out["request_pack"] * units_per_pack.where(units_per_pack > 0, 1.0)
    out["request_pcs"] = request_pcs.where(request_pcs > 0, fallback_pcs)
    optional_text_cols = {
        "product_name_code": "product_name_code",
        "production_code": "production_code",
        "p_code": "p_code",
        "q_code": "q_code",
        "r_code": "r_code",
        "market_type": "market_type",
        "customer_name": "customer_name",
    }
    for source_key, output_col in optional_text_cols.items():
        if source_key in cols:
            out[output_col] = raw[cols[source_key]].map(clean_str)
        else:
            out[output_col] = "(미기재)" if output_col == "customer_name" else ""
    if "due_date" in cols:
        out["request_due_date"] = pd.to_datetime(raw[cols["due_date"]], errors="coerce")
    else:
        out["request_due_date"] = pd.NaT

    if "market_type" in out.columns:
        overseas_mask = out["market_type"].astype(str).str.contains("해외", case=False, na=False)
        out = out[~overseas_mask].copy()

    for col in ["sales_code", "product_name", "product_name_code", "production_code", "p_code", "q_code", "r_code"]:
        out[f"{col}_key"] = out[col].map(normalize_match_key)
    return out


def normalize_packing(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path)
    cols = resolve_columns(
        raw,
        PACKING_COLS,
        required_keys=["sales_code", "packing_qty"],
        file_label=path.name,
    )
    return pd.DataFrame(
        {
            "sales_code": raw[cols["sales_code"]].map(clean_str),
            "packing_pack": to_number(raw[cols["packing_qty"]]),
        }
    )


def empty_progress_df() -> pd.DataFrame:
    columns = [
        "site_code",
        "customer_name",
        "order_no",
        "initial",
        "product_code",
        "demand_product_name",
        "demand_qty",
        "total_prod_qty",
        "total_due_date",
        "production_basis_qty",
        "product_code_key",
        "product_base_p_key",
        "demand_product_name_key",
        "linked_product_name",
        "match_source",
    ]
    for step in PROCESS_STEPS:
        columns.extend([step["qty_col"], step["due_col"]])
    return pd.DataFrame(columns=columns)

def find_progress_column_index(groups: pd.Series, fields: pd.Series, group_label: str, field_label: str) -> int | None:
    target_group = normalize_match_key(group_label)
    target_field = normalize_match_key(field_label)
    for idx in range(len(fields)):
        if normalize_match_key(groups.iloc[idx]) == target_group and normalize_match_key(fields.iloc[idx]) == target_field:
            return idx
    return None


def find_progress_base_column_index(fields: pd.Series, field_label: str) -> int | None:
    target_field = normalize_match_key(field_label)
    for idx in range(len(fields)):
        if normalize_match_key(fields.iloc[idx]) == target_field:
            return idx
    return None


def normalize_progress(path: Path | None, request_df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, int]]:
    if path is None:
        return empty_progress_df(), {"total_rows": 0, "domestic_rows": 0, "code_rows": 0, "name_rows": 0}

    raw = pd.read_excel(path, sheet_name="Sheet1", header=None)
    if raw.shape[0] < 3:
        raise DashboardConfigError([f"[{path.name}] Sheet1에 처리할 데이터가 없습니다."])

    groups = raw.iloc[0]
    fields = raw.iloc[1]
    data = raw.iloc[2:].copy()

    base_indices = {
        "site_code": find_progress_base_column_index(fields, "설비 사이트 코드"),
        "customer_name": find_progress_base_column_index(fields, "고객 이름"),
        "order_no": find_progress_base_column_index(fields, "수주번호"),
        "initial": find_progress_base_column_index(fields, "이니셜"),
        "product_code": find_progress_base_column_index(fields, "제품 코드"),
        "demand_product_name": find_progress_base_column_index(fields, "수요 제품 이름"),
        "demand_qty": find_progress_base_column_index(fields, "수요 수량"),
    }
    missing = [name for name, idx in base_indices.items() if idx is None and name in {"product_code", "demand_product_name"}]
    if missing:
        raise DashboardConfigError([f"[{path.name}] 수요정보 필수 컬럼 누락: {', '.join(missing)}"])

    out = pd.DataFrame(index=data.index)
    for name, idx in base_indices.items():
        if idx is None:
            out[name] = "" if name != "demand_qty" else 0.0
            continue
        if name == "demand_qty":
            out[name] = to_number(data.iloc[:, idx])
        else:
            out[name] = data.iloc[:, idx].map(clean_str)

    for step in PROCESS_STEPS:
        qty_idx = find_progress_column_index(groups, fields, str(step["header"]), "생산 수량")
        due_idx = find_progress_column_index(groups, fields, str(step["header"]), "납기일")
        out[step["qty_col"]] = to_number(data.iloc[:, qty_idx]) if qty_idx is not None else 0.0
        out[step["due_col"]] = pd.to_datetime(data.iloc[:, due_idx], errors="coerce") if due_idx is not None else pd.NaT

    total_qty_idx = find_progress_column_index(groups, fields, "총합계", "생산 수량")
    total_due_idx = find_progress_column_index(groups, fields, "총합계", "납기일")
    out["total_prod_qty"] = to_number(data.iloc[:, total_qty_idx]) if total_qty_idx is not None else 0.0
    out["total_due_date"] = pd.to_datetime(data.iloc[:, total_due_idx], errors="coerce") if total_due_idx is not None else pd.NaT

    inspection_step = next(step for step in PROCESS_STEPS if step["id"] == "80")
    out["production_basis_qty"] = out[str(inspection_step["qty_col"])]

    out["product_code_key"] = out["product_code"].map(normalize_match_key)
    out["product_base_p_key"] = out["product_code"].map(extract_base_p_code_key)
    out["demand_product_name_key"] = out["demand_product_name"].map(normalize_match_key)

    request_production_keys = set(request_df["production_code_key"].map(clean_str)) - {""}
    request_p_keys = set(request_df["p_code_key"].map(clean_str)) - {""}
    request_name_keys = set(request_df["product_name_key"].map(clean_str)) - {""}

    exact_production_match = out["product_code_key"].isin(request_production_keys)
    exact_p_match = out["product_code_key"].isin(request_p_keys)
    base_p_match = out["product_base_p_key"].isin(request_p_keys)
    name_match = out["demand_product_name_key"].isin(request_name_keys)
    code_match = exact_production_match | exact_p_match | base_p_match
    domestic_match = code_match | name_match

    production_name_map = build_first_value_map(request_df, "production_code_key", "product_name")
    p_name_map = build_first_value_map(request_df, "p_code_key", "product_name")
    request_name_map = build_first_value_map(request_df, "product_name_key", "product_name")

    out["linked_product_name"] = ""
    out.loc[exact_production_match, "linked_product_name"] = out.loc[exact_production_match, "product_code_key"].map(
        production_name_map
    )
    p_link = exact_p_match & (out["linked_product_name"] == "")
    out.loc[p_link, "linked_product_name"] = out.loc[p_link, "product_code_key"].map(p_name_map)
    base_p_link = base_p_match & (out["linked_product_name"] == "")
    out.loc[base_p_link, "linked_product_name"] = out.loc[base_p_link, "product_base_p_key"].map(p_name_map)
    name_link = name_match & (out["linked_product_name"] == "")
    out.loc[name_link, "linked_product_name"] = out.loc[name_link, "demand_product_name_key"].map(request_name_map)
    out["linked_product_name"] = out["linked_product_name"].fillna("")
    empty_link = out["linked_product_name"] == ""
    out.loc[empty_link, "linked_product_name"] = out.loc[empty_link, "demand_product_name"]

    out["match_source"] = ""
    out.loc[name_match, "match_source"] = "제품명"
    out.loc[base_p_match, "match_source"] = "P코드(생산)"
    out.loc[exact_p_match, "match_source"] = "P코드(생산)"
    out.loc[exact_production_match, "match_source"] = "생산코드"

    filtered = out[domestic_match].copy()
    info = {
        "total_rows": int(len(out)),
        "domestic_rows": int(len(filtered)),
        "code_rows": int(code_match.sum()),
        "name_rows": int((~code_match & name_match).sum()),
    }
    return filtered, info


def format_date(value: Any) -> str:
    date = pd.to_datetime(value, errors="coerce")
    if pd.isna(date):
        return ""
    return date.strftime("%Y-%m-%d")


def summarize_progress(progress_df: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
    if progress_df.empty:
        base_columns = group_cols + ["누수규격검사 생산수량"]
        return pd.DataFrame(columns=base_columns)

    agg_spec: dict[str, Any] = {
        "누수규격검사 생산수량": ("production_basis_qty", "sum"),
    }
    return progress_df.groupby(group_cols, dropna=False).agg(**agg_spec).reset_index()


def classify_status(packing_pack: float, packing_progress_pct: float) -> str:
    if packing_progress_pct >= 100.0:
        return "완료"
    if packing_pack > 0:
        return "진행중"
    return "미착수"


def finalize_summary(summary: pd.DataFrame) -> pd.DataFrame:
    out = summary.copy()
    out["포장부족수량"] = (out["요청 PACK"] - out["포장 PACK"]).clip(lower=0.0)
    raw_progress = np.where(
        out["요청 PACK"] > 0,
        out["포장 PACK"] / out["요청 PACK"] * 100.0,
        0.0,
    )
    out["포장진도율"] = np.clip(raw_progress, 0.0, 100.0)
    out["부족 PACK"] = out["포장부족수량"]
    out["진도율(%)"] = out["포장진도율"]
    out["상태"] = [
        classify_status(float(packing), float(progress))
        for packing, progress in zip(out["포장 PACK"], out["포장진도율"])
    ]
    return out


def calc_production_progress_pct(request_qty: Any, production_shortage_qty: Any) -> Any:
    request = pd.to_numeric(request_qty, errors="coerce").fillna(0.0)
    shortage = pd.to_numeric(production_shortage_qty, errors="coerce").fillna(0.0)
    produced = (request - shortage).clip(lower=0.0)
    produced = produced.where(produced <= request, request)
    return np.where(request > 0, produced / request * 100.0, 0.0)


def build_summaries(
    request_df: pd.DataFrame,
    packing_df: pd.DataFrame,
) -> tuple[pd.DataFrame, float, pd.DataFrame]:
    request_work = request_df.copy()
    optional_cols = [
        "product_name_code",
        "production_code",
        "p_code",
        "q_code",
        "r_code",
        "request_due_date",
        "request_pcs",
        "pack_unit",
        "pack_unit_label",
        "base_product_name",
        "customer_name",
        "sales_code_key",
        "product_name_key",
        "product_name_code_key",
        "production_code_key",
        "p_code_key",
        "q_code_key",
        "r_code_key",
    ]
    for col in optional_cols:
        if col not in request_work.columns:
            if col == "request_due_date":
                request_work[col] = pd.NaT
            elif col == "request_pcs":
                request_work[col] = request_work["request_pack"]
            elif col == "pack_unit":
                request_work[col] = np.nan
            elif col == "pack_unit_label":
                request_work[col] = "(미기재)"
            elif col == "base_product_name":
                request_work[col] = request_work["product_name"].map(strip_pack_unit_suffix)
            elif col == "customer_name":
                request_work[col] = "(미기재)"
            else:
                request_work[col] = ""

    group_cols = [
        "sales_code",
        "product_name",
        "product_name_code",
        "production_code",
        "p_code",
        "q_code",
        "r_code",
        "pack_unit",
        "pack_unit_label",
        "base_product_name",
        "customer_name",
        "sales_code_key",
        "product_name_key",
        "product_name_code_key",
        "production_code_key",
        "p_code_key",
        "q_code_key",
        "r_code_key",
    ]
    request_by_code = (
        request_work.groupby(group_cols, dropna=False)
        .agg(
            request_pack=("request_pack", "sum"),
            request_pcs=("request_pcs", "sum"),
            request_due_date=("request_due_date", min_datetime),
        )
        .reset_index()
    )
    packing_by_code = packing_df.groupby("sales_code", dropna=False)["packing_pack"].sum().reset_index()

    matched_code_summary = request_by_code.merge(packing_by_code, on="sales_code", how="left")
    matched_code_summary["packing_pack"] = matched_code_summary["packing_pack"].fillna(0.0)

    request_codes = set(request_by_code["sales_code"].astype(str))
    unmatched_pack = packing_by_code[~packing_by_code["sales_code"].astype(str).isin(request_codes)].copy()
    unmatched_packing_total = float(unmatched_pack["packing_pack"].sum()) if not unmatched_pack.empty else 0.0

    product_summary = (
        matched_code_summary.groupby("base_product_name", dropna=False)[["request_pack", "request_pcs", "packing_pack"]]
        .sum()
        .reset_index()
        .rename(
            columns={
                "base_product_name": "제품명",
                "request_pack": "요청 PACK",
                "request_pcs": "요청 PCS",
                "packing_pack": "포장 PACK",
            }
        )
    )
    product_summary = finalize_summary(product_summary)
    product_summary["제품분류"] = product_summary["제품명"].map(classify_product_group)
    product_summary["본품분류"] = product_summary["제품명"].map(classify_main_product_family)
    return product_summary, unmatched_packing_total, matched_code_summary


def enrich_product_summary(product_summary: pd.DataFrame, progress_df: pd.DataFrame) -> pd.DataFrame:
    progress_work = progress_df.copy()
    if "linked_product_name" not in progress_work.columns:
        progress_work["linked_product_name"] = ""
    progress_work["linked_base_product_name"] = progress_work["linked_product_name"].map(strip_pack_unit_suffix)
    progress_by_product = summarize_progress(progress_work, ["linked_base_product_name"]).rename(
        columns={
            "linked_base_product_name": "제품명",
        }
    )
    out = product_summary.merge(
        progress_by_product[["제품명", "누수규격검사 생산수량"]],
        on="제품명",
        how="left",
    )
    out["누수규격검사 생산수량"] = out["누수규격검사 생산수량"].fillna(0.0)
    out["생산부족수량"] = out["누수규격검사 생산수량"].clip(lower=0.0)
    out["생산진도율"] = calc_production_progress_pct(out["요청 PCS"], out["생산부족수량"])
    return out


def attach_progress_to_code_summary(code_summary: pd.DataFrame, progress_df: pd.DataFrame) -> pd.DataFrame:
    if progress_df.empty:
        progress_by_code = pd.DataFrame(columns=["production_code_key", "production_basis_qty", "production_due_date"])
    else:
        progress_work = progress_df.copy()
        due_source = pd.to_datetime(progress_work.get("total_due_date", pd.NaT), errors="coerce")
        inspection_step = next(step for step in PROCESS_STEPS if step["id"] == "80")
        inspection_due = pd.to_datetime(progress_work.get(str(inspection_step["due_col"]), pd.NaT), errors="coerce")
        progress_work["production_due_date"] = due_source.fillna(inspection_due)
        progress_by_code = (
            progress_work.groupby("product_code_key", dropna=False)
            .agg(
                production_basis_qty=("production_basis_qty", "sum"),
                production_due_date=("production_due_date", min_datetime),
            )
            .reset_index()
            .rename(columns={"product_code_key": "production_code_key"})
        )
    keep_cols = ["production_code_key", "production_basis_qty", "production_due_date"]
    out = code_summary.merge(progress_by_code[keep_cols], on="production_code_key", how="left")
    out["production_basis_qty"] = out["production_basis_qty"].fillna(0.0)
    out["production_due_date"] = pd.to_datetime(out["production_due_date"], errors="coerce")
    out["production_shortage_qty"] = out["production_basis_qty"].clip(lower=0.0)
    out["production_progress_pct"] = calc_production_progress_pct(
        out["request_pcs"],
        out["production_shortage_qty"],
    )
    return out


def build_production_code_view(code_summary: pd.DataFrame) -> pd.DataFrame:
    work = code_summary.copy()
    work["production_code"] = work["production_code"].replace("", "(생산코드 미기재)")
    grouped = (
        work.groupby("production_code", dropna=False)
        .agg(
            sales_code_count=("sales_code", "nunique"),
            product_name=("product_name", join_unique),
            request_pack=("request_pack", "sum"),
            request_pcs=("request_pcs", "sum"),
            packing_pack=("packing_pack", "sum"),
            production_basis_qty=("production_basis_qty", "max"),
            request_due_date=("request_due_date", min_datetime),
            production_due_date=("production_due_date", min_datetime),
        )
        .reset_index()
    )
    grouped["production_shortage_qty"] = grouped["production_basis_qty"].clip(lower=0.0)
    grouped["production_progress_pct"] = calc_production_progress_pct(
        grouped["request_pcs"],
        grouped["production_shortage_qty"],
    )
    grouped = grouped.rename(
        columns={
            "production_code": "생산코드",
            "sales_code_count": "연결 판매코드 수",
            "product_name": "제품명",
            "request_pack": "요청 PACK",
            "request_pcs": "요청 PCS",
            "packing_pack": "포장 PACK",
            "production_basis_qty": "누수규격검사 생산수량",
            "production_shortage_qty": "생산부족수량",
            "production_progress_pct": "생산진도율",
            "request_due_date": "납기일",
            "production_due_date": "생산완료예상일",
        }
    )
    grouped = finalize_summary(grouped)
    return grouped


def is_sample_name(name: str) -> bool:
    text = str(name)
    return any(keyword in text for keyword in SAMPLE_KEYWORDS)


def split_main_sample(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    sample_mask = df["제품명"].astype(str).map(is_sample_name)
    sample_df = df[sample_mask].copy()
    main_df = df[~sample_mask].copy()
    return main_df, sample_df


def parse_power_from_sales_code(value: Any) -> float:
    text = clean_str(value)
    if not text:
        return np.nan
    tail = text.split("-", 1)[1] if "-" in text else text
    match = re.search(r"(-?\d+(?:\.\d+)?)", tail)
    if match is None:
        match = POWER_RE.search(text)
    if not match:
        return np.nan
    try:
        number = float(match.group(1))
    except ValueError:
        return np.nan
    # Sales code values are typically encoded as 00.50, 01.00 ... and represent minus diopters.
    return round(-abs(number), 2)


def format_power(value: Any) -> str:
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return "(미기재)"
    return f"-{abs(float(num)):.2f}"


def classify_product_group(product_name: str) -> str:
    text = clean_str(product_name)
    upper = text.upper()
    if is_sample_name(text):
        return "샘플"
    if upper.startswith("PIA_") or upper.startswith("PIA ") or PIA_TOKEN_RE.search(upper):
        return "PIA"
    if "CLALEN" in upper:
        return "Clalen"
    if "TORIC" in upper or "사축" in text:
        return "Toric"
    if "1DAY" in upper or "원데이" in text:
        return "1Day"
    if "COLOR" in upper or "컬러" in text:
        return "Color"
    if "M_" in upper or " M " in upper or "먼슬리" in text or "MONTHLY" in upper:
        return "Monthly"
    return "기타"


def classify_main_product_family(product_name: str) -> str:
    text = clean_str(product_name)
    upper = text.upper()
    if is_sample_name(text):
        return "샘플"
    if upper.startswith("PIA_") or upper.startswith("PIA ") or PIA_TOKEN_RE.search(upper):
        if "_1D" in upper or " 1D" in upper:
            return "PIA 1Day"
        if "_1M" in upper or " 1M" in upper:
            return "PIA Monthly"
        return "PIA 기타"
    if "O2O2 D TORIC" in upper or "O2O2 D_TORIC" in upper or "O2O2 D TORIC_" in upper:
        return "O2O2 Toric"
    if "IRIS TORIC" in upper:
        return "Iris Toric"
    if "TORIC" in upper or "사축" in text or "정축" in text or upper.startswith("T38"):
        return "Toric"
    if "O2O2 1DAY" in upper:
        return "O2O2 1Day"
    if "O2O2 D_MICELIA" in upper or "O2O2 D MICELIA" in upper:
        return "O2O2 D Micelia"
    if "O2O2 D_" in upper or "O2O2 D " in upper:
        return "O2O2 D 컬러"
    if "O2O2 M_MICELIA" in upper or "O2O2 M MICELIA" in upper:
        return "O2O2 M Micelia"
    if "O2O2 M" in upper:
        return "O2O2 Monthly"
    if "CLALEN 1DAY" in upper or "CLALEN1DAY" in upper.replace(" ", ""):
        return "Clalen 1Day"
    if "CLEAR" in upper:
        return "Clear"
    if "IRIS" in upper:
        return "Iris 컬러"
    if upper.startswith("S38") or upper.startswith("S45") or upper.startswith("US38") or "BANDAGE" in upper:
        return "부자재/기타"
    return "기타"


def build_power_detail(code_summary: pd.DataFrame) -> pd.DataFrame:
    work = code_summary.copy()
    work["power_value"] = work["sales_code"].map(parse_power_from_sales_code)
    work = work[work["power_value"].notna()].copy()
    if work.empty:
        return pd.DataFrame(
            columns=[
                "제품분류",
                "제품명",
                "POWER",
                "요청수량",
                "요청PCS",
                "포장수량",
                "부족수량",
                "진도율",
                "상태",
                "생산부족수량",
                "생산진도율",
                "power_value",
            ]
        )

    work["제품분류"] = work["product_name"].map(classify_product_group)
    work["POWER"] = work["power_value"].map(format_power)

    grouped = (
        work.groupby(["제품분류", "product_name", "power_value", "POWER"], dropna=False)[
            ["request_pack", "request_pcs", "packing_pack"]
        ]
        .sum()
        .reset_index()
        .rename(
            columns={
                "product_name": "제품명",
                "request_pack": "요청수량",
                "request_pcs": "요청PCS",
                "packing_pack": "포장수량",
            }
        )
    )
    grouped["부족수량"] = (grouped["요청수량"] - grouped["포장수량"]).clip(lower=0.0)
    grouped["진도율"] = np.where(
        grouped["요청수량"] > 0,
        grouped["포장수량"] / grouped["요청수량"] * 100.0,
        0.0,
    )
    grouped["진도율"] = np.clip(grouped["진도율"], 0.0, 100.0)
    grouped["상태"] = [
        classify_status(float(packing), float(progress))
        for packing, progress in zip(grouped["포장수량"], grouped["진도율"])
    ]

    progress_source = work.copy()
    progress_source["_progress_dedupe_key"] = np.where(
        progress_source["production_code_key"].map(clean_str) != "",
        progress_source["production_code_key"],
        progress_source["sales_code_key"],
    )
    progress_source = progress_source.drop_duplicates(["product_name", "power_value", "_progress_dedupe_key"])
    progress_grouped = (
        progress_source.groupby(["product_name", "power_value"], dropna=False)
        .agg(
            production_basis_qty=("production_basis_qty", "sum"),
        )
        .reset_index()
        .rename(
            columns={
                "product_name": "제품명",
                "production_basis_qty": "누수규격검사 생산수량",
            }
        )
    )
    grouped = grouped.merge(progress_grouped, on=["제품명", "power_value"], how="left")
    grouped["누수규격검사 생산수량"] = grouped["누수규격검사 생산수량"].fillna(0.0)
    grouped["생산부족수량"] = grouped["누수규격검사 생산수량"].clip(lower=0.0)
    grouped["생산진도율"] = calc_production_progress_pct(grouped["요청PCS"], grouped["생산부족수량"])
    return grouped


def get_group_options(power_df: pd.DataFrame) -> list[str]:
    options = GROUP_ORDER.copy()
    if power_df.empty:
        return options
    available = set(power_df["제품분류"].astype(str))
    extras = sorted(available - set(options))
    return options + extras


def filter_power_detail(
    power_df: pd.DataFrame,
    group_name: str,
    product_name: str,
    high_power_only: bool,
    shortage_only: bool,
    not_started_only: bool,
) -> pd.DataFrame:
    out = power_df.copy()
    if group_name == "본품":
        out = out[~out["제품명"].astype(str).map(is_sample_name)]
    elif group_name == "샘플":
        out = out[out["제품명"].astype(str).map(is_sample_name)]
    elif group_name != "전체":
        out = out[out["제품분류"] == group_name]
    if product_name != "전체":
        out = out[out["제품명"] == product_name]
    if high_power_only:
        out = out[out["power_value"] <= -5.0]
    if shortage_only:
        out = out[out["부족수량"] > 0]
    if not_started_only:
        out = out[out["상태"] == "미착수"]
    return out


def calc_power_ops_kpi(power_df: pd.DataFrame) -> dict[str, float]:
    if power_df.empty:
        return {
            "rows": 0,
            "shortage_rows": 0,
            "not_started_rows": 0,
            "high_power_shortage_rows": 0,
            "shortage_qty": 0.0,
        }
    return {
        "rows": int(len(power_df)),
        "shortage_rows": int((power_df["부족수량"] > 0).sum()),
        "not_started_rows": int((power_df["상태"] == "미착수").sum()),
        "high_power_shortage_rows": int(((power_df["power_value"] <= -5.0) & (power_df["부족수량"] > 0)).sum()),
        "shortage_qty": float(power_df["부족수량"].sum()),
    }


def render_power_ops_table(power_df: pd.DataFrame, max_rows: int = 2000) -> None:
    if power_df.empty:
        st.warning("조건에 맞는 POWER 상세 데이터가 없습니다.")
        return

    ordered = power_df.sort_values(["power_value", "부족수량"], ascending=[True, False], kind="stable").head(max_rows).copy()
    rows: list[str] = []
    for _, row in ordered.iterrows():
        power_value = float(row["power_value"])
        power_label = escape(str(row["POWER"]))
        req = format_int(float(row["요청수량"]))
        packed = format_int(float(row["포장수량"]))
        shortage = float(row["부족수량"])
        shortage_txt = format_int(shortage)
        progress = float(row["진도율"])
        prod_progress = float(row.get("생산진도율", 0.0))

        power_class = "power-cell high" if power_value <= -5.0 else "power-cell"
        shortage_class = "num shortage" if shortage > 0 else "num"
        progress_html = progress_cell_html(progress, "포장")
        prod_progress_html = progress_cell_html(prod_progress, "생산")

        rows.append(
            "<tr>"
            f"<td class='{power_class}'>{power_label}</td>"
            f"<td class='num'>{req}</td>"
            f"<td class='num'>{packed}</td>"
            f"<td class='{shortage_class}'>{shortage_txt}</td>"
            f"<td>{prod_progress_html}</td>"
            f"<td>{progress_html}</td>"
            "</tr>"
        )

    header = (
        "<tr>"
        "<th>POWER</th>"
        "<th class='num'>요청수량</th>"
        "<th class='num'>포장수량</th>"
        "<th class='num'>부족수량</th>"
        "<th>생산진도율</th>"
        "<th>포장진도율</th>"
        "</tr>"
    )
    table_html = (
        "<div class='table-wrap'>"
        "<table class='ops-table'>"
        f"<thead>{header}</thead>"
        f"<tbody>{''.join(rows)}</tbody>"
        "</table>"
        "</div>"
    )
    st.markdown(table_html, unsafe_allow_html=True)


def build_power_heatmap(power_df: pd.DataFrame) -> px.imshow | None:
    if power_df.empty:
        return None

    power_order = (
        power_df[["POWER", "power_value"]]
        .drop_duplicates()
        .sort_values("power_value", ascending=True)["POWER"]
        .tolist()
    )
    product_order = (
        power_df.groupby("제품명", dropna=False)["부족수량"]
        .sum()
        .sort_values(ascending=False)
        .index.tolist()
    )
    matrix = power_df.pivot_table(index="제품명", columns="POWER", values="진도율", aggfunc="mean")
    matrix = matrix.reindex(index=product_order, columns=power_order)
    if matrix.shape[0] > 35:
        matrix = matrix.iloc[:35]

    fig = px.imshow(
        matrix,
        aspect="auto",
        zmin=0,
        zmax=100,
        color_continuous_scale=[(0.0, "#f3f5f8"), (0.5, MID_GRAY), (1.0, NAVY)],
        labels={"x": "POWER", "y": "제품명", "color": "포장진도율(%)"},
        title="제품/POWER 포장진도율 Heatmap",
    )
    fig.update_layout(
        paper_bgcolor=WHITE,
        plot_bgcolor=WHITE,
        margin=dict(l=8, r=8, t=52, b=8),
    )
    fig.update_traces(
        hovertemplate="제품명: %{y}<br>POWER: %{x}<br>포장진도율: %{z:.1f}%<extra></extra>"
    )
    fig.update_xaxes(type="category")
    return fig


def apply_filters(df: pd.DataFrame, query: str, statuses: list[str]) -> pd.DataFrame:
    out = df.copy()
    q = query.strip()
    if q:
        out = out[out["제품명"].astype(str).str.contains(q, case=False, na=False)]
    if statuses:
        out = out[out["상태"].isin(statuses)]
    else:
        out = out.iloc[0:0]
    return out


def product_scope_options(df: pd.DataFrame) -> list[str]:
    base = ["전체", "본품", "샘플"]
    if "제품분류" not in df.columns:
        return base
    extras = [value for value in GROUP_ORDER if value not in base and value in set(df["제품분류"].astype(str))]
    remaining = sorted(set(df["제품분류"].astype(str)) - set(base) - set(extras))
    return base + extras + remaining


def product_family_options(df: pd.DataFrame) -> list[str]:
    base = ["전체"]
    if "본품분류" not in df.columns:
        return base
    available = set(df["본품분류"].dropna().astype(str))
    ordered = [value for value in MAIN_PRODUCT_FAMILY_ORDER if value not in base and value in available]
    remaining = sorted(available - set(base) - set(ordered))
    return base + ordered + remaining


def apply_product_scope_filter(df: pd.DataFrame, scope: str) -> pd.DataFrame:
    if scope == "전체" or "제품분류" not in df.columns:
        return df.copy()
    if scope == "본품":
        return df[~df["제품명"].astype(str).map(is_sample_name)].copy()
    if scope == "샘플":
        return df[df["제품명"].astype(str).map(is_sample_name)].copy()
    return df[df["제품분류"].astype(str) == scope].copy()


def apply_product_family_filter(df: pd.DataFrame, family: str) -> pd.DataFrame:
    if family == "전체" or "본품분류" not in df.columns:
        return df.copy()
    return df[df["본품분류"].astype(str) == family].copy()


def calc_kpi(df: pd.DataFrame) -> dict[str, float]:
    request_pack = float(df["요청 PACK"].sum()) if not df.empty else 0.0
    packing_pack = float(df["포장 PACK"].sum()) if not df.empty else 0.0
    shortage_pack = float(df["포장부족수량"].sum()) if "포장부족수량" in df.columns and not df.empty else 0.0
    progress = (packing_pack / request_pack * 100.0) if request_pack > 0 else 0.0
    request_pcs = float(df["요청 PCS"].sum()) if "요청 PCS" in df.columns and not df.empty else 0.0
    production_shortage_qty = (
        float(df["생산부족수량"].sum()) if "생산부족수량" in df.columns and not df.empty else 0.0
    )
    production_progress = (
        (request_pcs - production_shortage_qty) / request_pcs * 100.0
        if request_pcs > 0
        else 0.0
    )
    return {
        "request_pack": request_pack,
        "packing_pack": packing_pack,
        "shortage_pack": shortage_pack,
        "progress_pct": min(100.0, max(0.0, progress)),
        "production_progress_pct": min(100.0, max(0.0, production_progress)),
        "production_shortage_products": int((df["생산부족수량"] > 0).sum()) if "생산부족수량" in df.columns else 0,
        "packing_shortage_products": int((df["포장부족수량"] > 0).sum()) if "포장부족수량" in df.columns else 0,
        "not_started_products": int((df["상태"] == "미착수").sum()) if "상태" in df.columns else 0,
        "completed_products": int((df["상태"] == "완료").sum()) if "상태" in df.columns else 0,
    }


def code_summary_for_products(code_summary: pd.DataFrame, product_names: pd.Series) -> pd.DataFrame:
    if code_summary.empty:
        return code_summary.copy()
    names = set(product_names.dropna().astype(str))
    if not names:
        return code_summary.iloc[0:0].copy()

    work = code_summary.copy()
    if "base_product_name" in work.columns:
        base_names = work["base_product_name"].astype(str)
    else:
        base_names = work["product_name"].map(strip_pack_unit_suffix).astype(str)
    return work[base_names.isin(names)].copy()


def add_allocated_production_basis(code_summary: pd.DataFrame) -> pd.DataFrame:
    work = code_summary.copy()
    if work.empty:
        work["_allocated_production_shortage_qty"] = 0.0
        return work

    if "production_basis_qty" not in work.columns:
        work["production_basis_qty"] = 0.0
    if "request_pcs" not in work.columns:
        work["request_pcs"] = work.get("request_pack", 0.0)

    production_key = work.get("production_code_key", pd.Series("", index=work.index)).map(clean_str)
    sales_key = work.get("sales_code_key", pd.Series("", index=work.index)).map(clean_str)
    fallback_key = work.get("sales_code", pd.Series("", index=work.index)).map(clean_str)
    work["_production_alloc_key"] = production_key.where(production_key != "", sales_key)
    work["_production_alloc_key"] = work["_production_alloc_key"].where(work["_production_alloc_key"] != "", fallback_key)

    key_request = work.groupby("_production_alloc_key", dropna=False)["request_pcs"].transform("sum")
    key_shortage = work.groupby("_production_alloc_key", dropna=False)["production_basis_qty"].transform("max")
    allocation_ratio = np.where(key_request > 0, work["request_pcs"] / key_request, 0.0)
    work["_allocated_production_shortage_qty"] = (key_shortage * allocation_ratio).clip(lower=0.0)
    return work


def calc_kpi_from_code_summary(code_summary: pd.DataFrame) -> dict[str, float]:
    if code_summary.empty:
        return {
            "request_pack": 0.0,
            "packing_pack": 0.0,
            "shortage_pack": 0.0,
            "progress_pct": 0.0,
            "production_progress_pct": 0.0,
        }

    request_pack = float(code_summary["request_pack"].sum())
    packing_pack = float(code_summary["packing_pack"].sum())
    shortage_pack = max(0.0, request_pack - packing_pack)
    packing_progress = (packing_pack / request_pack * 100.0) if request_pack > 0 else 0.0

    work = (
        code_summary.copy()
        if "_allocated_production_shortage_qty" in code_summary.columns
        else add_allocated_production_basis(code_summary)
    )
    request_pcs = float(work["request_pcs"].sum())
    shortage_pcs = float(work["_allocated_production_shortage_qty"].sum())
    production_progress = ((request_pcs - shortage_pcs) / request_pcs * 100.0) if request_pcs > 0 else 0.0

    return {
        "request_pack": request_pack,
        "packing_pack": packing_pack,
        "shortage_pack": shortage_pack,
        "progress_pct": min(100.0, max(0.0, packing_progress)),
        "production_progress_pct": min(100.0, max(0.0, production_progress)),
    }


def format_int(value: float) -> str:
    return f"{value:,.0f}"


def progress_tone(progress: float) -> str:
    if progress >= 100:
        return "done"
    if progress >= 80:
        return "active"
    if progress > 0:
        return "warn"
    return "risk"


def status_class(status: str) -> str:
    if status == "완료":
        return "done"
    if status == "진행중":
        return "active"
    if status == "부족":
        return "warn"
    return "risk"


def progress_cell_html(progress: float, label: str = "") -> str:
    width = max(0.0, min(100.0, float(progress)))
    tone = progress_tone(float(progress))
    prefix = f"<span class='progress-name'>{escape(label)}</span>" if label else ""
    return (
        "<div class='progress-cell'>"
        f"{prefix}"
        "<div class='progress-track'>"
        f"<div class='progress-fill {tone}' style='width:{width:.1f}%'></div>"
        "</div>"
        f"<span class='progress-text'>{progress:.1f}%</span>"
        "</div>"
    )


def render_ops_table(
    df: pd.DataFrame,
    compact: bool = False,
    max_rows: int = 500,
    show_family: bool = False,
) -> None:
    if df.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
        return

    ordered = df.sort_values(
        ["포장부족수량", "생산부족수량", "요청 PACK"],
        ascending=[False, False, False],
        kind="stable",
    ).head(max_rows).copy()

    rows: list[str] = []
    for _, row in ordered.iterrows():
        product = escape(str(row["제품명"]))
        family = escape(str(row.get("본품분류", ""))) if show_family else ""
        req = format_int(float(row["요청 PACK"]))
        packing_shortage = float(row["포장부족수량"])
        packing_shortage_txt = format_int(packing_shortage)
        packing_progress = float(row["포장진도율"])
        production_shortage = float(row.get("생산부족수량", 0.0))
        production_shortage_txt = format_int(production_shortage)
        prod_progress = float(row.get("생산진도율", 0.0))
        status = escape(str(row["상태"]))
        badge = f"<span class='status-badge {status_class(str(row['상태']))}'>{status}</span>"

        packing_shortage_class = "num shortage" if packing_shortage > 0 else "num"
        production_shortage_class = "num shortage" if production_shortage > 0 else "num"
        packing_progress_html = progress_cell_html(packing_progress, "포장")
        production_progress_html = progress_cell_html(prod_progress, "생산")

        if compact:
            rows.append(
                "<tr>"
                f"<td class='left'>{product}</td>"
                f"<td>{production_progress_html}</td>"
                f"<td>{packing_progress_html}</td>"
                f"<td class='{packing_shortage_class}'>{packing_shortage_txt}</td>"
                f"<td>{badge}</td>"
                "</tr>"
            )
        else:
            rows.append(
                "<tr>"
                f"<td class='left'>{product}</td>"
                f"{f'<td>{family}</td>' if show_family else ''}"
                f"<td class='num'>{req}</td>"
                f"<td>{production_progress_html}</td>"
                f"<td>{packing_progress_html}</td>"
                f"<td class='{production_shortage_class}'>{production_shortage_txt}</td>"
                f"<td class='{packing_shortage_class}'>{packing_shortage_txt}</td>"
                f"<td>{badge}</td>"
                "</tr>"
            )

    header = (
        "<tr>"
        "<th class='left'>제품명</th>"
        "<th>생산진도율</th>"
        "<th>포장진도율</th>"
        "<th class='num'>포장부족수량</th>"
        "<th>상태</th>"
        "</tr>"
        if compact
        else "<tr>"
        "<th class='left'>제품명</th>"
        f"{'<th>본품분류</th>' if show_family else ''}"
        "<th class='num'>요청 PACK</th>"
        "<th>생산진도율</th>"
        "<th>포장진도율</th>"
        "<th class='num'>생산부족수량</th>"
        "<th class='num'>포장부족수량</th>"
        "<th>상태</th>"
        "</tr>"
    )

    table_html = (
        "<div class='table-wrap'>"
        "<table class='ops-table'>"
        f"<thead>{header}</thead>"
        f"<tbody>{''.join(rows)}</tbody>"
        "</table>"
        "</div>"
    )
    st.markdown(table_html, unsafe_allow_html=True)


def build_family_progress_view(product_df: pd.DataFrame) -> pd.DataFrame:
    if product_df.empty or "본품분류" not in product_df.columns:
        return pd.DataFrame(
            columns=[
                "본품분류",
                "요청 PACK",
                "요청 PCS",
                "포장 PACK",
                "생산부족수량",
                "포장부족수량",
                "생산진도율",
                "포장진도율",
            ]
        )

    grouped = (
        product_df.groupby("본품분류", dropna=False)
        .agg(
            request_pack=("요청 PACK", "sum"),
            request_pcs=("요청 PCS", "sum"),
            packing_pack=("포장 PACK", "sum"),
            production_shortage_qty=("생산부족수량", "sum"),
            packing_shortage_qty=("포장부족수량", "sum"),
        )
        .reset_index()
        .rename(
            columns={
                "request_pack": "요청 PACK",
                "request_pcs": "요청 PCS",
                "packing_pack": "포장 PACK",
                "production_shortage_qty": "생산부족수량",
                "packing_shortage_qty": "포장부족수량",
            }
        )
    )
    grouped["생산진도율"] = calc_production_progress_pct(grouped["요청 PCS"], grouped["생산부족수량"])
    grouped["포장진도율"] = np.where(
        grouped["요청 PACK"] > 0,
        grouped["포장 PACK"] / grouped["요청 PACK"] * 100.0,
        0.0,
    )
    grouped["포장진도율"] = np.clip(grouped["포장진도율"], 0.0, 100.0)
    grouped["_order"] = grouped["본품분류"].map(
        {name: idx for idx, name in enumerate(MAIN_PRODUCT_FAMILY_ORDER)}
    ).fillna(999)
    return grouped.sort_values(
        ["_order", "포장부족수량", "요청 PACK"],
        ascending=[True, False, False],
        kind="stable",
    ).drop(columns=["_order"])


def render_family_progress_cards(family_df: pd.DataFrame, max_rows: int = 14) -> None:
    if family_df.empty:
        st.warning("본품 분류별 진도현황을 표시할 데이터가 없습니다.")
        return

    rows: list[str] = []
    for _, row in family_df.head(max_rows).iterrows():
        family = escape(str(row["본품분류"]))
        request_pack = format_int(float(row["요청 PACK"]))
        production_progress = float(row["생산진도율"])
        packing_progress = float(row["포장진도율"])
        production_shortage = format_int(float(row["생산부족수량"]))
        packing_shortage = format_int(float(row["포장부족수량"]))
        rows.append(
            "<div class='family-card'>"
            f"<div class='family-head'><span>{family}</span><b>요청 {request_pack} PACK</b></div>"
            f"{progress_cell_html(production_progress, '생산')}"
            f"{progress_cell_html(packing_progress, '포장')}"
            "<div class='family-shortages'>"
            f"<span>생산부족 <b>{production_shortage}</b></span>"
            f"<span>포장부족 <b>{packing_shortage}</b></span>"
            "</div>"
            "</div>"
        )

    st.markdown("<div class='family-grid'>" + "".join(rows) + "</div>", unsafe_allow_html=True)


def build_top_shortage_view(product_df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    if product_df.empty:
        return pd.DataFrame(columns=["제품명", "포장부족수량", "포장진도율"])
    return (
        product_df[product_df["포장부족수량"] > 0]
        .sort_values("포장부족수량", ascending=False, kind="stable")
        .head(top_n)[["제품명", "포장부족수량", "포장진도율"]]
        .copy()
    )


def render_top_shortage_list(top_df: pd.DataFrame) -> None:
    if top_df.empty:
        st.warning("포장부족 제품이 없습니다.")
        return

    rows: list[str] = []
    for idx, (_, row) in enumerate(top_df.iterrows(), start=1):
        product = escape(str(row["제품명"]))
        shortage = format_int(float(row["포장부족수량"]))
        progress = float(row["포장진도율"])
        rows.append(
            "<div class='top-row'>"
            f"<div class='top-rank'>{idx}</div>"
            f"<div class='top-name'>{product}</div>"
            f"<div class='top-shortage'>{shortage}</div>"
            f"<div class='top-progress'>{progress_cell_html(progress, '포장')}</div>"
            "</div>"
        )
    st.markdown("<div class='top-list'>" + "".join(rows) + "</div>", unsafe_allow_html=True)


def build_gap_top_view(product_df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    columns = ["제품명", "생산진도율", "포장진도율", "GAP", "포장부족수량"]
    if product_df.empty:
        return pd.DataFrame(columns=columns)

    source = product_df.copy()
    if "생산진도율" not in source.columns:
        source["생산진도율"] = 0.0
    if "포장진도율" not in source.columns:
        source["포장진도율"] = 0.0
    if "포장부족수량" not in source.columns:
        source["포장부족수량"] = 0.0
    if "요청 PACK" not in source.columns:
        source["요청 PACK"] = 0.0
    source["생산진도율"] = pd.to_numeric(source["생산진도율"], errors="coerce").fillna(0.0)
    source["포장진도율"] = pd.to_numeric(source["포장진도율"], errors="coerce").fillna(0.0)
    source["GAP"] = source["생산진도율"] - source["포장진도율"]
    source = source[source["GAP"] > 0].copy()
    if source.empty:
        return pd.DataFrame(columns=columns)

    return (
        source.sort_values(
            ["GAP", "포장부족수량", "요청 PACK"],
            ascending=[False, False, False],
            kind="stable",
        )
        .head(top_n)[columns]
        .copy()
    )


def render_gap_top_list(gap_df: pd.DataFrame) -> None:
    if gap_df.empty:
        st.warning("생산진도율이 포장진도율보다 높은 GAP 제품이 없습니다.")
        return

    rows: list[str] = []
    for idx, (_, row) in enumerate(gap_df.iterrows(), start=1):
        product = escape(str(row["제품명"]))
        production_progress = float(row["생산진도율"])
        packing_progress = float(row["포장진도율"])
        gap = float(row["GAP"])
        rows.append(
            "<div class='gap-row'>"
            f"<div class='top-rank'>{idx}</div>"
            f"<div class='top-name'>{product}</div>"
            f"<div class='gap-progress'>{progress_cell_html(production_progress, '생산')}</div>"
            f"<div class='gap-progress'>{progress_cell_html(packing_progress, '포장')}</div>"
            f"<div class='gap-value'>+{gap:.1f}</div>"
            "</div>"
        )
    st.markdown("<div class='gap-list'>" + "".join(rows) + "</div>", unsafe_allow_html=True)


def render_kpi_panel(title: str, kpi: dict[str, float]) -> None:
    progress = float(kpi["progress_pct"])
    production_progress = float(kpi.get("production_progress_pct", 0.0))
    shortage_class = "metric-value warn" if kpi["shortage_pack"] > 0 else "metric-value"

    panel_html = f"""
    <div class='kpi-panel scope-kpi'>
      <div class='kpi-title'>{escape(title)}</div>
      <div class='kpi-grid'>
        <div class='kpi-card'>
          <div class='metric-label'>요청 PACK</div>
          <div class='metric-value'>{format_int(kpi['request_pack'])}</div>
        </div>
        <div class='kpi-card'>
          <div class='metric-label'>포장 PACK</div>
          <div class='metric-value'>{format_int(kpi['packing_pack'])}</div>
        </div>
        <div class='kpi-card'>
          <div class='metric-label'>부족 PACK</div>
          <div class='{shortage_class}'>{format_int(kpi['shortage_pack'])}</div>
        </div>
        <div class='kpi-card'>
          <div class='metric-label'>포장진도율</div>
          <div class='metric-value'>{progress:.1f}%</div>
        </div>
        <div class='kpi-card'>
          <div class='metric-label'>생산진도율</div>
          <div class='metric-value'>{production_progress:.1f}%</div>
        </div>
      </div>
    </div>
    """
    st.markdown(panel_html, unsafe_allow_html=True)


def render_kpi_scope_panels(code_summary: pd.DataFrame, product_names: pd.Series | None = None) -> None:
    work = add_allocated_production_basis(code_summary)
    if product_names is not None:
        work = code_summary_for_products(work, product_names)

    kpi_cols = st.columns(3, gap="small")
    for col, title, (_, kpi) in zip(kpi_cols, ["전체 KPI", "본품 KPI", "샘플 KPI"], build_scope_kpis(work)):
        with col:
            render_kpi_panel(title, kpi)


def build_scope_kpis(code_summary: pd.DataFrame) -> list[tuple[str, dict[str, float]]]:
    sample_mask = (
        code_summary["product_name"].astype(str).map(is_sample_name)
        if "product_name" in code_summary.columns
        else pd.Series(False, index=code_summary.index)
    )
    return [
        ("전체", calc_kpi_from_code_summary(code_summary)),
        ("본품", calc_kpi_from_code_summary(code_summary[~sample_mask].copy())),
        ("샘플", calc_kpi_from_code_summary(code_summary[sample_mask].copy())),
    ]


def base_pack_label(value: Any) -> str:
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num) or float(num) <= 0:
        return "(미기재)"
    return f"{float(num):g}P"


def extract_pack_label_from_text(value: Any) -> str:
    unit = extract_pack_unit(value)
    if pd.notna(unit) and float(unit) > 0:
        return base_pack_label(unit)
    text = clean_str(value)
    if not text:
        return "(미기재)"
    match = PACK_ANY_RE.search(text)
    if not match:
        return "(미기재)"
    try:
        return base_pack_label(float(match.group(1)))
    except ValueError:
        return "(미기재)"


def standard_pack_bucket(label: Any) -> str:
    text = clean_str(label)
    return text if text in STANDARD_PACK_BUCKETS else "기타팩"


def row_pack_bucket(row: pd.Series) -> str:
    for value in [row.get("_pack_label", ""), row.get("product_name", ""), row.get("sales_code", "")]:
        label = extract_pack_label_from_text(value) if not str(value).endswith("P") else clean_str(value)
        bucket = standard_pack_bucket(label)
        if bucket != "기타팩":
            return bucket
    return "기타팩"


def pack_sort_key(label: Any) -> tuple[int, float, str]:
    text = clean_str(label)
    if text == "기타팩":
        return (1, 999998.0, text)
    match = re.match(r"^(\d+(?:\.\d+)?)P$", text)
    if match:
        return (0, float(match.group(1)), text)
    if text == "(미기재)":
        return (2, 999999.0, text)
    return (1, 999999.0, text)


def pack_sort_rank(label: Any) -> float:
    group, value, _ = pack_sort_key(label)
    return group * 1000000.0 + value


def sorted_pack_labels(labels: list[str]) -> list[str]:
    unique = list(dict.fromkeys([clean_str(label) for label in labels if clean_str(label)]))
    return sorted(unique, key=pack_sort_key)


def with_operational_columns(code_summary: pd.DataFrame) -> pd.DataFrame:
    work = code_summary.copy()
    if "base_product_name" not in work.columns:
        work["base_product_name"] = work["product_name"].map(strip_pack_unit_suffix)
    if "pack_unit" not in work.columns:
        work["pack_unit"] = work["product_name"].map(extract_pack_unit)
    work["_pack_label"] = work["pack_unit"].map(base_pack_label)
    work["_pack_sort"] = work["_pack_label"].map(pack_sort_key)
    work["제품분류"] = work["base_product_name"].map(classify_product_group)
    work["본품분류"] = work["base_product_name"].map(classify_main_product_family)
    work["본품/샘플"] = np.where(work["base_product_name"].astype(str).map(is_sample_name), "샘플", "본품")
    if "customer_name" not in work.columns:
        work["customer_name"] = "(미기재)"
    work["customer_name"] = work["customer_name"].replace("", "(미기재)").fillna("(미기재)")
    sales_power = work["sales_code"].map(parse_power_from_sales_code)
    production_power = work["production_code"].map(parse_power_from_sales_code)
    product_power = work["product_name"].map(parse_power_from_sales_code)
    work["power_value"] = sales_power.fillna(production_power).fillna(product_power)
    work["POWER"] = work["power_value"].map(format_power)
    work["production_code_display"] = work["production_code"].replace("", "(생산코드 미기재)")
    work["_pack_bucket"] = work.apply(row_pack_bucket, axis=1)
    work["_pack_bucket_sort"] = work["_pack_bucket"].map(pack_sort_rank)
    return work


def available_pack_options(code_summary: pd.DataFrame) -> list[str]:
    work = with_operational_columns(code_summary)
    labels = sorted_pack_labels(work["_pack_label"].dropna().astype(str).tolist())
    return ["전체"] + labels


def available_product_group_options(code_summary: pd.DataFrame) -> list[str]:
    work = with_operational_columns(code_summary)
    available = set(work["제품분류"].dropna().astype(str))
    ordered = [value for value in GROUP_ORDER if value not in {"본품", "샘플"} and value in available]
    remaining = sorted(available - set(ordered))
    return ["전체"] + ordered + remaining


def available_power_options(code_summary: pd.DataFrame) -> list[str]:
    work = with_operational_columns(code_summary)
    source = work[work["power_value"].notna()][["POWER", "power_value"]].drop_duplicates()
    source = source.sort_values("power_value", ascending=True, kind="stable")
    return ["전체"] + source["POWER"].astype(str).tolist()


def available_customer_options(code_summary: pd.DataFrame) -> list[str]:
    work = with_operational_columns(code_summary)
    values = sorted(work["customer_name"].dropna().astype(str).unique().tolist())
    return ["전체"] + values


def filter_operational_code_summary(
    code_summary: pd.DataFrame,
    product_query: str = "",
    production_query: str = "",
    sales_query: str = "",
    pack_label: str = "전체",
    product_group: str = "전체",
    sample_scope: str = "전체",
    power_label: str = "전체",
    customer_name: str = "전체",
) -> pd.DataFrame:
    out = with_operational_columns(code_summary)
    product_q = product_query.strip()
    if product_q:
        name_match = out["product_name"].astype(str).str.contains(product_q, case=False, na=False)
        base_match = out["base_product_name"].astype(str).str.contains(product_q, case=False, na=False)
        out = out[name_match | base_match]
    production_q = production_query.strip()
    if production_q:
        out = out[out["production_code_display"].astype(str).str.contains(production_q, case=False, na=False)]
    sales_q = sales_query.strip()
    if sales_q:
        out = out[out["sales_code"].astype(str).str.contains(sales_q, case=False, na=False)]
    if pack_label != "전체":
        out = out[out["_pack_label"] == pack_label]
    if product_group != "전체":
        out = out[out["제품분류"] == product_group]
    if sample_scope == "본품":
        out = out[out["본품/샘플"] == "본품"]
    elif sample_scope == "샘플":
        out = out[out["본품/샘플"] == "샘플"]
    if power_label != "전체":
        out = out[out["POWER"] == power_label]
    if customer_name != "전체":
        out = out[out["customer_name"] == customer_name]
    return out.copy()


def build_pack_pivot(
    code_summary: pd.DataFrame,
    index_cols: list[str],
    pack_labels: list[str],
) -> pd.DataFrame:
    if code_summary.empty:
        return pd.DataFrame(columns=index_cols + pack_labels)
    work = with_operational_columns(code_summary)
    work["_pivot_request_pack"] = pd.to_numeric(work["request_pack"], errors="coerce").fillna(0.0).astype(float)
    pivot = (
        work.pivot_table(
            index=index_cols,
            columns="_pack_label",
            values="_pivot_request_pack",
            aggfunc="sum",
            dropna=False,
        )
        .fillna(0.0)
        .reset_index()
        .rename_axis(None, axis=1)
    )
    for label in pack_labels:
        if label not in pivot.columns:
            pivot[label] = 0.0
    return pivot[index_cols + pack_labels]


def display_date_or_dash(value: Any) -> str:
    text = format_date(value)
    return text if text else "-"


def pct_from_parts(done: Any, total: Any) -> float:
    done_num = float(pd.to_numeric(done, errors="coerce") or 0.0)
    total_num = float(pd.to_numeric(total, errors="coerce") or 0.0)
    if total_num <= 0:
        return 0.0
    return min(100.0, max(0.0, done_num / total_num * 100.0))


def calc_operation_kpis(product_summary: pd.DataFrame, code_summary: pd.DataFrame) -> dict[str, float]:
    today = pd.Timestamp.now(tz="Asia/Seoul").tz_localize(None).normalize()
    due_limit = today + pd.Timedelta(days=7)
    work = with_operational_columns(code_summary)
    work["request_due_date"] = pd.to_datetime(work["request_due_date"], errors="coerce")
    shortage_mask = (work["request_pack"] - work["packing_pack"]).clip(lower=0.0) > 0
    due_mask = work["request_due_date"].notna() & (work["request_due_date"] <= due_limit)
    urgent_products = work.loc[shortage_mask & due_mask, "base_product_name"].dropna().astype(str).nunique()
    return {
        "urgent_products": float(urgent_products),
        "packing_shortage_pack": float(product_summary["포장부족수량"].sum()) if not product_summary.empty else 0.0,
        "production_shortage_pcs": float(product_summary["생산부족수량"].sum())
        if "생산부족수량" in product_summary.columns and not product_summary.empty
        else 0.0,
    }


def render_operation_kpis(product_summary: pd.DataFrame, code_summary: pd.DataFrame) -> None:
    kpi = calc_operation_kpis(product_summary, code_summary)
    c1, c2, c3 = st.columns(3, gap="small")
    c1.metric("긴급품목 수", f"{int(kpi['urgent_products']):,}")
    c2.metric("포장 부족 PACK", format_int(kpi["packing_shortage_pack"]))
    c3.metric("생산 부족 PCS", format_int(kpi["production_shortage_pcs"]))


def build_product_progress_main_view(
    product_summary: pd.DataFrame,
    code_summary: pd.DataFrame,
    pack_labels: list[str],
) -> pd.DataFrame:
    work = with_operational_columns(code_summary)
    due_by_product = (
        work.groupby("base_product_name", dropna=False)
        .agg(production_due_date=("production_due_date", min_datetime))
        .reset_index()
        .rename(columns={"base_product_name": "제품명"})
    )
    pivot = build_pack_pivot(work, ["base_product_name"], pack_labels).rename(columns={"base_product_name": "제품명"})

    out = product_summary.merge(pivot, on="제품명", how="left").merge(due_by_product, on="제품명", how="left")
    for label in pack_labels:
        out[label] = out[label].fillna(0.0)
    out["제품필요수량"] = out.get("생산부족수량", 0.0)
    out["진도율"] = out.get("생산진도율", 0.0)
    out["생산완료예상일"] = [
        display_date_or_dash(date) if float(shortage) > 0 else "-"
        for date, shortage in zip(out["production_due_date"], out["제품필요수량"])
    ]
    out["전체진도율"] = out["포장진도율"]
    ordered = [
        "제품명",
        *pack_labels,
        "요청 PACK",
        "요청 PCS",
        "제품필요수량",
        "진도율",
        "생산완료예상일",
        "포장부족수량",
        "전체진도율",
        "상태",
    ]
    out = out.rename(columns={"요청 PACK": "요청합계(PACK)", "요청 PCS": "요청합계(PCS)"})
    ordered = [
        "제품명",
        *pack_labels,
        "요청합계(PACK)",
        "요청합계(PCS)",
        "제품필요수량",
        "진도율",
        "생산완료예상일",
        "포장부족수량",
        "전체진도율",
        "상태",
    ]
    return out[ordered].copy()


def build_product_sku_detail_view(code_summary: pd.DataFrame, product_name: str) -> pd.DataFrame:
    work = with_operational_columns(code_summary)
    scope = work[work["base_product_name"] == product_name].copy()
    if scope.empty:
        scope = work[work["product_name"] == product_name].copy()
    if scope.empty:
        return pd.DataFrame(columns=["SKU", "생산코드", "판매코드 수", "요청 PACK", "부족 PACK", "진도율"])
    grouped = (
        scope.groupby(["product_name", "production_code_display"], dropna=False)
        .agg(
            sales_code_count=("sales_code", "nunique"),
            request_pack=("request_pack", "sum"),
            packing_pack=("packing_pack", "sum"),
        )
        .reset_index()
        .rename(
            columns={
                "product_name": "SKU",
                "production_code_display": "생산코드",
                "sales_code_count": "판매코드 수",
                "request_pack": "요청 PACK",
            }
        )
    )
    grouped["부족 PACK"] = (grouped["요청 PACK"] - grouped["packing_pack"]).clip(lower=0.0)
    grouped["진도율"] = np.where(grouped["요청 PACK"] > 0, grouped["packing_pack"] / grouped["요청 PACK"] * 100.0, 0.0)
    grouped["진도율"] = np.clip(grouped["진도율"], 0.0, 100.0)
    return grouped[["SKU", "생산코드", "판매코드 수", "요청 PACK", "부족 PACK", "진도율"]].sort_values(
        ["부족 PACK", "요청 PACK"], ascending=[False, False], kind="stable"
    )


def build_sales_pack_detail_view(code_summary: pd.DataFrame) -> pd.DataFrame:
    if code_summary.empty:
        return pd.DataFrame(columns=["판매코드", "PACK", "요청", "포장", "부족", "납기"])
    work = with_operational_columns(code_summary)
    out = (
        work.groupby(["sales_code", "_pack_label"], dropna=False)
        .agg(
            request_pack=("request_pack", "sum"),
            packing_pack=("packing_pack", "sum"),
            request_due_date=("request_due_date", min_datetime),
        )
        .reset_index()
        .rename(columns={"sales_code": "판매코드", "_pack_label": "PACK", "request_pack": "요청", "packing_pack": "포장"})
    )
    out["부족"] = (out["요청"] - out["포장"]).clip(lower=0.0)
    out["납기"] = out["request_due_date"].map(display_date_or_dash)
    out["_pack_sort"] = out["PACK"].map(pack_sort_key)
    return out.sort_values(["부족", "_pack_sort", "요청"], ascending=[False, True, False], kind="stable")[
        ["판매코드", "PACK", "요청", "포장", "부족", "납기"]
    ]


def build_production_progress_main_view(code_summary: pd.DataFrame, pack_labels: list[str]) -> pd.DataFrame:
    if code_summary.empty:
        return pd.DataFrame(
            columns=[
                "생산코드",
                "제품명",
                *pack_labels,
                "요청합계",
                "생산부족",
                "포장부족",
                "진도율",
                "판매코드수",
            ]
        )
    work = with_operational_columns(code_summary)
    base = build_production_code_view(work).rename(
        columns={
            "요청 PACK": "요청합계",
            "생산부족수량": "생산부족",
            "포장부족수량": "포장부족",
            "포장진도율": "진도율",
            "연결 판매코드 수": "판매코드수",
        }
    )
    pivot = build_pack_pivot(work, ["production_code_display"], pack_labels).rename(
        columns={"production_code_display": "생산코드"}
    )
    out = base.merge(pivot, on="생산코드", how="left")
    for label in pack_labels:
        out[label] = out[label].fillna(0.0)
    return out[
        ["생산코드", "제품명", *pack_labels, "요청합계", "생산부족", "포장부족", "진도율", "판매코드수", "상태"]
    ].sort_values(["포장부족", "생산부족", "요청합계"], ascending=[False, False, False], kind="stable")


def prepare_production_power_rows(code_summary: pd.DataFrame) -> pd.DataFrame:
    work = with_operational_columns(code_summary)
    work = add_allocated_production_basis(work)
    work["_production_shortage_pcs"] = pd.to_numeric(
        work["_allocated_production_shortage_qty"],
        errors="coerce",
    ).fillna(0.0)
    work["_packing_shortage_pack"] = (work["request_pack"] - work["packing_pack"]).clip(lower=0.0)
    work["_power_sort"] = pd.to_numeric(work["power_value"], errors="coerce").fillna(999999.0)
    return work


def available_production_power_options(code_summary: pd.DataFrame) -> list[str]:
    work = prepare_production_power_rows(code_summary)
    source = work[["POWER", "_power_sort"]].drop_duplicates().sort_values("_power_sort", ascending=True, kind="stable")
    values = source["POWER"].astype(str).tolist()
    return ["전체"] + values


def filter_production_power_rows(
    code_summary: pd.DataFrame,
    product_query: str,
    production_query: str,
    power_label: str,
    pack_bucket: str,
    sample_scope: str,
    product_group: str,
) -> pd.DataFrame:
    out = prepare_production_power_rows(code_summary)
    product_q = product_query.strip()
    if product_q:
        name_match = out["product_name"].astype(str).str.contains(product_q, case=False, na=False)
        base_match = out["base_product_name"].astype(str).str.contains(product_q, case=False, na=False)
        out = out[name_match | base_match]
    production_q = production_query.strip()
    if production_q:
        out = out[out["production_code_display"].astype(str).str.contains(production_q, case=False, na=False)]
    if power_label != "전체":
        out = out[out["POWER"] == power_label]
    if pack_bucket != "전체":
        out = out[out["_pack_bucket"] == pack_bucket]
    if sample_scope == "본품":
        out = out[out["본품/샘플"] == "본품"]
    elif sample_scope == "샘플":
        out = out[out["본품/샘플"] == "샘플"]
    if product_group != "전체":
        out = out[out["제품분류"] == product_group]
    return out.copy()


def bottleneck_status(production_progress: Any, packing_progress: Any) -> str:
    production = float(pd.to_numeric(production_progress, errors="coerce") or 0.0)
    packing = float(pd.to_numeric(packing_progress, errors="coerce") or 0.0)
    if production < 20.0 and packing < 20.0:
        return "미착수 ⚫"
    if production < packing - 20.0:
        return "생산 병목 🔴"
    if packing < production - 20.0:
        return "포장 병목 🟠"
    return "정상 🟢"


def status_from_progress(packing_pack: Any, packing_progress: Any) -> str:
    packing = float(pd.to_numeric(packing_pack, errors="coerce") or 0.0)
    progress = float(pd.to_numeric(packing_progress, errors="coerce") or 0.0)
    return classify_status(packing, progress)


def build_production_power_main_view(rows: pd.DataFrame, shortage_only: bool = False) -> pd.DataFrame:
    visible_columns = [
        "생산코드",
        "대표 제품명",
        "POWER",
        "5P 필요팩",
        "10P 필요팩",
        "30P 필요팩",
        "80P 필요팩",
        "90P 필요팩",
        "기타팩 필요팩",
        "요청합계(PACK)",
        "요청합계(PCS)",
        "생산부족수량",
        "포장부족수량",
        "생산진도율",
        "포장진도율",
        "최소 납기일",
        "병목 상태",
        "상태",
    ]
    if rows.empty:
        return pd.DataFrame(columns=visible_columns + ["_min_due_date_sort", "_power_sort"])

    group_cols = ["production_code_display", "POWER", "_power_sort"]
    base = (
        rows.groupby(group_cols, dropna=False)
        .agg(
            representative_product=("base_product_name", first_nonempty),
            request_pack=("request_pack", "sum"),
            request_pcs=("request_pcs", "sum"),
            packing_pack=("packing_pack", "sum"),
            production_shortage_pcs=("_production_shortage_pcs", "sum"),
            packing_shortage_pack=("_packing_shortage_pack", "sum"),
            min_due_date=("request_due_date", min_datetime),
        )
        .reset_index()
    )

    pack_pivot = (
        rows.pivot_table(
            index=group_cols,
            columns="_pack_bucket",
            values="request_pack",
            aggfunc="sum",
            dropna=False,
        )
        .fillna(0.0)
        .reset_index()
        .rename_axis(None, axis=1)
    )
    for bucket in [*STANDARD_PACK_BUCKETS, "기타팩"]:
        if bucket not in pack_pivot.columns:
            pack_pivot[bucket] = 0.0

    grouped = base.merge(pack_pivot[group_cols + [*STANDARD_PACK_BUCKETS, "기타팩"]], on=group_cols, how="left")
    grouped["생산진도율"] = calc_production_progress_pct(grouped["request_pcs"], grouped["production_shortage_pcs"])
    grouped["포장진도율"] = np.where(
        grouped["request_pack"] > 0,
        grouped["packing_pack"] / grouped["request_pack"] * 100.0,
        0.0,
    )
    grouped["포장진도율"] = np.clip(grouped["포장진도율"], 0.0, 100.0)
    grouped["병목 상태"] = [
        bottleneck_status(prod, pack)
        for prod, pack in zip(grouped["생산진도율"], grouped["포장진도율"])
    ]
    grouped["상태"] = [
        status_from_progress(packing, progress)
        for packing, progress in zip(grouped["packing_pack"], grouped["포장진도율"])
    ]
    grouped["_min_due_date_sort"] = pd.to_datetime(grouped["min_due_date"], errors="coerce")
    grouped["최소 납기일"] = grouped["min_due_date"].map(display_date_or_dash)

    out = grouped.rename(
        columns={
            "production_code_display": "생산코드",
            "representative_product": "대표 제품명",
            "5P": "5P 필요팩",
            "10P": "10P 필요팩",
            "30P": "30P 필요팩",
            "80P": "80P 필요팩",
            "90P": "90P 필요팩",
            "기타팩": "기타팩 필요팩",
            "request_pack": "요청합계(PACK)",
            "request_pcs": "요청합계(PCS)",
            "production_shortage_pcs": "생산부족수량",
            "packing_shortage_pack": "포장부족수량",
        }
    )
    if shortage_only:
        out = out[(out["생산부족수량"] > 0) | (out["포장부족수량"] > 0)].copy()

    out = out.sort_values(
        ["_min_due_date_sort", "포장부족수량", "생산부족수량"],
        ascending=[True, False, False],
        na_position="last",
        kind="stable",
    )
    return out[visible_columns + ["_min_due_date_sort", "_power_sort"]].copy()


def calc_production_power_kpis(view: pd.DataFrame) -> dict[str, float]:
    if view.empty:
        return {
            "production_code_count": 0.0,
            "request_pack": 0.0,
            "production_shortage_pcs": 0.0,
            "packing_shortage_pack": 0.0,
            "production_bottleneck_count": 0.0,
            "packing_bottleneck_count": 0.0,
        }
    return {
        "production_code_count": float(view["생산코드"].nunique()),
        "request_pack": float(view["요청합계(PACK)"].sum()),
        "production_shortage_pcs": float(view["생산부족수량"].sum()),
        "packing_shortage_pack": float(view["포장부족수량"].sum()),
        "production_bottleneck_count": float(view["병목 상태"].astype(str).str.contains("생산 병목", na=False).sum()),
        "packing_bottleneck_count": float(view["병목 상태"].astype(str).str.contains("포장 병목", na=False).sum()),
    }


def render_production_power_kpis(view: pd.DataFrame) -> None:
    kpi = calc_production_power_kpis(view)
    items = [
        ("생산코드 수", f"{int(kpi['production_code_count']):,}", "normal"),
        ("총 요청 PACK", format_int(kpi["request_pack"]), "normal"),
        ("총 생산부족 PCS", format_int(kpi["production_shortage_pcs"]), "risk"),
        ("총 포장부족 PACK", format_int(kpi["packing_shortage_pack"]), "warn"),
        ("생산 병목 수", f"{int(kpi['production_bottleneck_count']):,}", "risk"),
        ("포장 병목 수", f"{int(kpi['packing_bottleneck_count']):,}", "warn"),
    ]
    cards = "".join(
        "<div class='mini-kpi-card'>"
        f"<div class='metric-label'>{escape(label)}</div>"
        f"<div class='metric-value {tone}'>{value}</div>"
        "</div>"
        for label, value, tone in items
    )
    st.markdown(f"<div class='mini-kpi-grid'>{cards}</div>", unsafe_allow_html=True)


def due_d_day_label(value: Any) -> str:
    due = pd.to_datetime(value, errors="coerce")
    if pd.isna(due):
        return "-"
    today = pd.Timestamp.now(tz="Asia/Seoul").tz_localize(None).normalize()
    days = int((due.normalize() - today).days)
    if days <= 0:
        return "D-0 🔴"
    if days <= 1:
        return f"D-{days} 🟠"
    if days <= 3:
        return f"D-{days} 🟡"
    return f"D-{days} 🟢"


def render_pack_composition_chart(selected_row: pd.Series) -> None:
    chart_df = pd.DataFrame(
        {
            "PACK": STANDARD_PACK_BUCKETS,
            "필요팩": [float(selected_row.get(f"{bucket} 필요팩", 0.0)) for bucket in STANDARD_PACK_BUCKETS],
        }
    )
    fig = px.bar(
        chart_df,
        x="PACK",
        y="필요팩",
        text="필요팩",
        title="PACK 구성",
        color_discrete_sequence=[NAVY],
    )
    fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
    fig.update_layout(
        paper_bgcolor=WHITE,
        plot_bgcolor=WHITE,
        margin=dict(l=8, r=8, t=48, b=8),
        yaxis_title="필요팩",
        xaxis_title="",
        showlegend=False,
    )
    st.plotly_chart(fig, width="stretch")


def render_production_progress_panel(selected_row: pd.Series) -> None:
    production_progress = float(selected_row.get("생산진도율", 0.0))
    packing_progress = float(selected_row.get("포장진도율", 0.0))
    due_label = due_d_day_label(selected_row.get("_min_due_date_sort", pd.NaT))
    panel = (
        "<div class='progress-summary-panel'>"
        "<div>"
        "<div class='section-sub'>생산/포장 Progress Bar</div>"
        f"{progress_cell_html(production_progress, '생산')}"
        f"{progress_cell_html(packing_progress, '포장')}"
        "</div>"
        "<div class='dday-box'>"
        "<div class='metric-label'>납기 D-Day</div>"
        f"<div class='dday-value'>{escape(due_label)}</div>"
        "</div>"
        "</div>"
    )
    st.markdown(panel, unsafe_allow_html=True)


def build_production_sales_detail_view(rows: pd.DataFrame, production_code: str, power_label: str) -> pd.DataFrame:
    scope = rows[(rows["production_code_display"] == production_code) & (rows["POWER"] == power_label)].copy()
    columns = [
        "판매코드",
        "제품명",
        "PACK 단위",
        "필요팩",
        "요청PCS",
        "포장완료PACK",
        "포장부족PACK",
        "생산진도율",
        "포장진도율",
        "납기일자",
    ]
    if scope.empty:
        return pd.DataFrame(columns=columns + ["_pack_sort"])

    grouped = (
        scope.groupby(["sales_code", "product_name", "_pack_bucket", "_pack_bucket_sort"], dropna=False)
        .agg(
            request_pack=("request_pack", "sum"),
            request_pcs=("request_pcs", "sum"),
            packing_pack=("packing_pack", "sum"),
            production_shortage_pcs=("_production_shortage_pcs", "sum"),
            request_due_date=("request_due_date", min_datetime),
        )
        .reset_index()
        .rename(
            columns={
                "sales_code": "판매코드",
                "product_name": "제품명",
                "_pack_bucket": "PACK 단위",
                "_pack_bucket_sort": "_pack_sort",
                "request_pack": "필요팩",
                "request_pcs": "요청PCS",
                "packing_pack": "포장완료PACK",
            }
        )
    )
    grouped["포장부족PACK"] = (grouped["필요팩"] - grouped["포장완료PACK"]).clip(lower=0.0)
    grouped["생산진도율"] = calc_production_progress_pct(grouped["요청PCS"], grouped["production_shortage_pcs"])
    grouped["포장진도율"] = np.where(
        grouped["필요팩"] > 0,
        grouped["포장완료PACK"] / grouped["필요팩"] * 100.0,
        0.0,
    )
    grouped["포장진도율"] = np.clip(grouped["포장진도율"], 0.0, 100.0)
    grouped["납기일자"] = grouped["request_due_date"].map(display_date_or_dash)
    grouped = grouped.sort_values(
        ["_pack_sort", "request_due_date", "포장부족PACK", "필요팩"],
        ascending=[True, True, False, False],
        na_position="last",
        kind="stable",
    )
    return grouped[columns + ["_pack_sort"]].copy()


def production_scope_from_row(code_summary: pd.DataFrame, production_code: str) -> pd.DataFrame:
    work = with_operational_columns(code_summary)
    return work[work["production_code_display"] == production_code].copy()


def sales_status_label(row: pd.Series) -> str:
    shortage = float(row.get("포장부족", 0.0))
    due = pd.to_datetime(row.get("request_due_date", pd.NaT), errors="coerce")
    today = pd.Timestamp.now(tz="Asia/Seoul").tz_localize(None).normalize()
    if shortage <= 0:
        return "완료"
    if pd.notna(due) and due <= today + pd.Timedelta(days=7):
        return "긴급"
    return "부족"


def build_sales_order_main_view(code_summary: pd.DataFrame) -> pd.DataFrame:
    if code_summary.empty:
        return pd.DataFrame(
            columns=[
                "판매코드",
                "생산코드",
                "제품명",
                "PACK",
                "거래처",
                "POWER",
                "요청PACK",
                "요청PCS",
                "생산부족",
                "포장부족",
                "생산진도율",
                "포장진도율",
                "납기",
                "상태",
            ]
        )
    work = add_allocated_production_basis(with_operational_columns(code_summary))
    grouped = (
        work.groupby("sales_code", dropna=False)
        .agg(
            production_code=("production_code_display", join_unique),
            product_name=("product_name", join_unique),
            pack_label=("_pack_label", join_unique),
            customer_name=("customer_name", join_unique),
            power=("POWER", first_nonempty),
            power_value=("power_value", "min"),
            request_pack=("request_pack", "sum"),
            request_pcs=("request_pcs", "sum"),
            packing_pack=("packing_pack", "sum"),
            production_shortage=("_allocated_production_shortage_qty", "sum"),
            request_due_date=("request_due_date", min_datetime),
        )
        .reset_index()
        .rename(
            columns={
                "sales_code": "판매코드",
                "production_code": "생산코드",
                "product_name": "제품명",
                "pack_label": "PACK",
                "customer_name": "거래처",
                "power": "POWER",
                "request_pack": "요청PACK",
                "request_pcs": "요청PCS",
                "production_shortage": "생산부족",
            }
        )
    )
    grouped["포장부족"] = (grouped["요청PACK"] - grouped["packing_pack"]).clip(lower=0.0)
    grouped["생산진도율"] = calc_production_progress_pct(grouped["요청PCS"], grouped["생산부족"])
    grouped["포장진도율"] = np.where(grouped["요청PACK"] > 0, grouped["packing_pack"] / grouped["요청PACK"] * 100.0, 0.0)
    grouped["포장진도율"] = np.clip(grouped["포장진도율"], 0.0, 100.0)
    grouped["납기"] = grouped["request_due_date"].map(display_date_or_dash)
    grouped["상태"] = grouped.apply(sales_status_label, axis=1)
    return grouped[
        [
            "판매코드",
            "생산코드",
            "제품명",
            "PACK",
            "거래처",
            "POWER",
            "요청PACK",
            "요청PCS",
            "생산부족",
            "포장부족",
            "생산진도율",
            "포장진도율",
            "납기",
            "상태",
            "power_value",
        ]
    ].sort_values(["포장부족", "생산부족", "요청PACK"], ascending=[False, False, False], kind="stable")


def sales_scope_from_row(code_summary: pd.DataFrame, sales_code: str) -> pd.DataFrame:
    work = with_operational_columns(code_summary)
    return work[work["sales_code"] == sales_code].copy()


def production_shortage_pack_equivalent(work: pd.DataFrame) -> pd.Series:
    if "_allocated_production_shortage_qty" not in work.columns:
        work = add_allocated_production_basis(work)
    pack_unit = pd.to_numeric(work.get("pack_unit", pd.Series(np.nan, index=work.index)), errors="coerce")
    implied_unit = np.where(work["request_pack"] > 0, work["request_pcs"] / work["request_pack"], np.nan)
    unit = pack_unit.where(pack_unit > 0, implied_unit)
    unit = pd.Series(unit, index=work.index).replace([np.inf, -np.inf], np.nan).fillna(1.0)
    unit = unit.where(unit > 0, 1.0)
    return (work["_allocated_production_shortage_qty"] / unit).clip(lower=0.0)


def build_power_summary_view(code_summary: pd.DataFrame) -> pd.DataFrame:
    work = with_operational_columns(code_summary)
    work = work[work["power_value"].notna()].copy()
    if work.empty:
        return pd.DataFrame(columns=["POWER", "요청", "생산", "포장", "부족", "진도율", "power_value"])
    work = add_allocated_production_basis(work)
    work["_production_shortage_pack"] = production_shortage_pack_equivalent(work)
    work["_production_pack"] = (work["request_pack"] - work["_production_shortage_pack"]).clip(lower=0.0)
    grouped = (
        work.groupby(["power_value", "POWER"], dropna=False)
        .agg(
            request_pack=("request_pack", "sum"),
            production_pack=("_production_pack", "sum"),
            packing_pack=("packing_pack", "sum"),
        )
        .reset_index()
        .rename(
            columns={
                "request_pack": "요청",
                "production_pack": "생산",
                "packing_pack": "포장",
            }
        )
    )
    grouped["부족"] = (grouped["요청"] - grouped["포장"]).clip(lower=0.0)
    grouped["진도율"] = np.where(grouped["요청"] > 0, grouped["포장"] / grouped["요청"] * 100.0, 0.0)
    grouped["진도율"] = np.clip(grouped["진도율"], 0.0, 100.0)
    return grouped[["POWER", "요청", "생산", "포장", "부족", "진도율", "power_value"]].sort_values(
        "power_value", ascending=True, kind="stable"
    )


def build_power_sku_detail_view(code_summary: pd.DataFrame, power_label: str) -> pd.DataFrame:
    work = with_operational_columns(code_summary)
    scope = work[work["POWER"] == power_label].copy()
    if scope.empty:
        return pd.DataFrame(columns=["생산코드", "판매코드", "제품명", "PACK", "요청", "부족", "납기"])
    out = (
        scope.groupby(["production_code_display", "sales_code", "product_name", "_pack_label"], dropna=False)
        .agg(
            request_pack=("request_pack", "sum"),
            packing_pack=("packing_pack", "sum"),
            request_due_date=("request_due_date", min_datetime),
        )
        .reset_index()
        .rename(
            columns={
                "production_code_display": "생산코드",
                "sales_code": "판매코드",
                "product_name": "제품명",
                "_pack_label": "PACK",
                "request_pack": "요청",
            }
        )
    )
    out["부족"] = (out["요청"] - out["packing_pack"]).clip(lower=0.0)
    out["납기"] = out["request_due_date"].map(display_date_or_dash)
    return out[["생산코드", "판매코드", "제품명", "PACK", "요청", "부족", "납기"]].sort_values(
        ["부족", "요청"], ascending=[False, False], kind="stable"
    )


def empty_inventory_detail_view() -> pd.DataFrame:
    return pd.DataFrame(columns=["LOT", "멸균NO", "ERP재고", "재작업 가능", "샘플 전환 가능"])


def ppt_rgb(hex_color: str) -> RGBColor:
    return RGBColor.from_string(hex_color.replace("#", ""))


def set_cell_text(
    cell: Any,
    text: str,
    size: int = 9,
    bold: bool = False,
    color: str = TEXT_DARK,
    align: PP_ALIGN = PP_ALIGN.CENTER,
) -> None:
    cell.text = ""
    cell.margin_left = Inches(0.04)
    cell.margin_right = Inches(0.04)
    cell.margin_top = Inches(0.02)
    cell.margin_bottom = Inches(0.02)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    paragraph = cell.text_frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.name = "맑은 고딕"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = ppt_rgb(color)


def add_textbox(
    slide: Any,
    text: str,
    left: float,
    top: float,
    width: float,
    height: float,
    size: int,
    bold: bool = False,
    color: str = TEXT_DARK,
    align: PP_ALIGN = PP_ALIGN.LEFT,
) -> None:
    shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = shape.text_frame
    frame.clear()
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.name = "맑은 고딕"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = ppt_rgb(color)


def format_report_value(value: Any, is_percent: bool = False) -> str:
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return "0.0%" if is_percent else "0"
    return f"{float(num):.1f}%" if is_percent else format_int(float(num))


def build_priority_report_view(product_view: pd.DataFrame, max_rows: int = 8) -> pd.DataFrame:
    if product_view.empty:
        return pd.DataFrame(
            columns=["제품명", "요청 PACK", "생산진도율", "포장진도율", "생산부족수량", "포장부족수량", "상태"]
        )
    return (
        product_view.sort_values(
            ["포장부족수량", "생산부족수량", "요청 PACK"],
            ascending=[False, False, False],
            kind="stable",
        )
        .head(max_rows)
        .copy()
    )


def build_ppt_report(
    product_view: pd.DataFrame,
    code_summary: pd.DataFrame,
    product_names: pd.Series,
    scope_label: str,
) -> bytes:
    work = add_allocated_production_basis(code_summary)
    work = code_summary_for_products(work, product_names)
    scope_kpis = build_scope_kpis(work)
    priority_view = build_priority_report_view(product_view)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = ppt_rgb(WHITE)

    add_textbox(slide, "국내 제품 포장현황 운영 보고서", 0.35, 0.22, 8.8, 0.36, 20, True, NAVY)
    generated_at = pd.Timestamp.now(tz="Asia/Seoul").strftime("%Y-%m-%d %H:%M")
    add_textbox(
        slide,
        f"기준: {scope_label} | 산출시각: {generated_at} | 포장 기준 운영 우선순위",
        0.35,
        0.62,
        8.8,
        0.25,
        9,
        False,
        TEXT_MUTED,
    )

    kpi_table = slide.shapes.add_table(6, 4, Inches(0.35), Inches(1.05), Inches(12.65), Inches(1.65)).table
    for idx, width in enumerate([2.4, 3.4, 3.4, 3.4]):
        kpi_table.columns[idx].width = Inches(width)
    header_labels = ["지표", "전체 KPI", "본품 KPI", "샘플 KPI"]
    for col_idx, label in enumerate(header_labels):
        cell = kpi_table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = ppt_rgb(NAVY)
        set_cell_text(cell, label, size=9, bold=True, color=WHITE)

    metric_rows = [
        ("요청 PACK", "request_pack", False),
        ("포장 PACK", "packing_pack", False),
        ("부족 PACK", "shortage_pack", False),
        ("포장진도율", "progress_pct", True),
        ("생산진도율", "production_progress_pct", True),
    ]
    kpi_map = {name: kpi for name, kpi in scope_kpis}
    for row_idx, (label, key, is_percent) in enumerate(metric_rows, start=1):
        set_cell_text(kpi_table.cell(row_idx, 0), label, size=9, bold=True, color=NAVY, align=PP_ALIGN.LEFT)
        for col_idx, scope in enumerate(["전체", "본품", "샘플"], start=1):
            cell = kpi_table.cell(row_idx, col_idx)
            if key == "shortage_pack":
                cell.fill.solid()
                cell.fill.fore_color.rgb = ppt_rgb("#fff7ef")
                color = MUTED_ORANGE
            else:
                color = TEXT_DARK
            set_cell_text(cell, format_report_value(kpi_map[scope].get(key, 0.0), is_percent), size=10, bold=True, color=color)

    add_textbox(slide, "우선 대응 제품 TOP 8", 0.35, 3.0, 5.0, 0.25, 13, True, NAVY)
    add_textbox(
        slide,
        "포장부족수량, 생산부족수량, 요청 PACK 순 정렬",
        0.35,
        3.28,
        5.8,
        0.22,
        8,
        False,
        TEXT_MUTED,
    )

    rows = max(2, len(priority_view) + 1)
    table = slide.shapes.add_table(rows, 7, Inches(0.35), Inches(3.58), Inches(12.65), Inches(3.45)).table
    widths = [3.9, 1.25, 1.25, 1.25, 1.55, 1.55, 0.9]
    for idx, width in enumerate(widths):
        table.columns[idx].width = Inches(width)
    headers = ["제품명", "요청", "생산진도", "포장진도", "생산부족", "포장부족", "상태"]
    for col_idx, label in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = ppt_rgb(SOFT_NAVY)
        set_cell_text(cell, label, size=8, bold=True, color=WHITE)

    if priority_view.empty:
        set_cell_text(table.cell(1, 0), "조건에 맞는 제품 데이터가 없습니다.", size=9, color=TEXT_MUTED, align=PP_ALIGN.LEFT)
        for col_idx in range(1, 7):
            set_cell_text(table.cell(1, col_idx), "", size=8)
    else:
        for row_idx, (_, row) in enumerate(priority_view.iterrows(), start=1):
            values = [
                str(row["제품명"]),
                format_report_value(row["요청 PACK"]),
                format_report_value(row["생산진도율"], True),
                format_report_value(row["포장진도율"], True),
                format_report_value(row["생산부족수량"]),
                format_report_value(row["포장부족수량"]),
                str(row["상태"]),
            ]
            for col_idx, value in enumerate(values):
                color = MUTED_ORANGE if col_idx in {4, 5} else TEXT_DARK
                align = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.CENTER
                set_cell_text(table.cell(row_idx, col_idx), value, size=7 if col_idx == 0 else 8, bold=col_idx in {4, 5}, color=color, align=align)

    output = BytesIO()
    prs.save(output)
    return output.getvalue()


def build_status_count_chart(main_df: pd.DataFrame, sample_df: pd.DataFrame) -> px.bar:
    parts: list[pd.DataFrame] = []
    for category, frame in [("본품", main_df), ("샘플", sample_df)]:
        counts = frame["상태"].value_counts().reindex(STATUS_ORDER, fill_value=0).reset_index()
        counts.columns = ["상태", "제품 수"]
        counts["구분"] = category
        parts.append(counts)
    status_df = pd.concat(parts, ignore_index=True)

    fig = px.bar(
        status_df,
        x="상태",
        y="제품 수",
        color="구분",
        barmode="group",
        title="상태별 제품 수",
        text="제품 수",
        color_discrete_map={"본품": NAVY, "샘플": SOFT_NAVY},
    )
    fig.update_traces(texttemplate="%{text:,.0f}")
    fig.update_layout(
        paper_bgcolor=WHITE,
        plot_bgcolor=WHITE,
        margin=dict(l=8, r=8, t=52, b=8),
        legend_title_text="",
    )
    return fig


def build_progress_compare_chart(
    total_kpi: dict[str, float],
    main_kpi: dict[str, float],
    sample_kpi: dict[str, float],
) -> px.bar:
    compare_df = pd.DataFrame(
        {
            "구분": ["전체", "본품", "샘플"],
            "진도율(%)": [total_kpi["progress_pct"], main_kpi["progress_pct"], sample_kpi["progress_pct"]],
        }
    )
    fig = px.bar(
        compare_df,
        x="구분",
        y="진도율(%)",
        color="구분",
        title="전체/본품/샘플 진도율 비교",
        text="진도율(%)",
        color_discrete_map={"전체": NAVY, "본품": SOFT_NAVY, "샘플": "#7087a5"},
    )
    fig.update_traces(texttemplate="%{text:.1f}%")
    fig.update_layout(
        paper_bgcolor=WHITE,
        plot_bgcolor=WHITE,
        margin=dict(l=8, r=8, t=52, b=8),
        legend_title_text="",
        showlegend=False,
    )
    fig.update_yaxes(rangemode="tozero")
    return fig


def build_shortage_top_chart(main_df: pd.DataFrame, sample_df: pd.DataFrame, top_n: int = 10) -> px.bar | None:
    source = pd.concat(
        [
            main_df.assign(구분="본품"),
            sample_df.assign(구분="샘플"),
        ],
        ignore_index=True,
    )
    source = source[source["포장부족수량"] > 0].copy()
    if source.empty:
        return None
    source = source.nlargest(top_n, "포장부족수량").sort_values("포장부족수량", ascending=True)
    source["라벨"] = source.apply(lambda r: f"[{r['구분']}] {r['제품명']}", axis=1)

    fig = px.bar(
        source,
        x="포장부족수량",
        y="라벨",
        color="구분",
        orientation="h",
        title=f"포장부족 TOP {min(top_n, len(source))}",
        text="포장부족수량",
        color_discrete_map={"본품": MUTED_ORANGE, "샘플": "#d99a5f"},
    )
    fig.update_traces(texttemplate="%{text:,.0f}")
    fig.update_layout(
        paper_bgcolor=WHITE,
        plot_bgcolor=WHITE,
        margin=dict(l=8, r=8, t=52, b=8),
        legend_title_text="",
        yaxis_title="",
    )
    return fig


def build_product_progress_gap_chart(df: pd.DataFrame, top_n: int = 18) -> px.bar | None:
    if df.empty:
        return None
    source = df.copy()
    source = source.sort_values(
        ["포장부족수량", "생산부족수량", "요청 PACK"],
        ascending=[False, False, False],
        kind="stable",
    ).head(top_n)
    if source.empty:
        return None
    chart_df = source[["제품명", "생산진도율", "포장진도율"]].melt(
        id_vars="제품명",
        value_vars=["생산진도율", "포장진도율"],
        var_name="지표",
        value_name="진도율",
    )
    product_order = source["제품명"].tolist()[::-1]
    fig = px.bar(
        chart_df,
        x="진도율",
        y="제품명",
        color="지표",
        barmode="group",
        orientation="h",
        category_orders={"제품명": product_order, "지표": ["생산진도율", "포장진도율"]},
        title="제품별 생산진도율 vs 포장진도율",
        text="진도율",
        color_discrete_map={"생산진도율": SOFT_NAVY, "포장진도율": NAVY},
    )
    fig.update_traces(texttemplate="%{text:.1f}%")
    fig.update_layout(
        paper_bgcolor=WHITE,
        plot_bgcolor=WHITE,
        margin=dict(l=8, r=8, t=52, b=8),
        legend_title_text="",
        yaxis_title="",
    )
    fig.update_xaxes(range=[0, 100])
    return fig


def render_style() -> None:
    st.markdown(
        f"""
        <style>
        .stApp {{
            background: {LIGHT_GRAY};
            color: {TEXT_DARK};
        }}
        :root {{
            --ui-gap: 12px;
        }}
        .block-container {{
            padding-top: 1.2rem;
            padding-bottom: 2.0rem;
        }}
        .kpi-panel {{
            background: {WHITE};
            border: 1px solid {MID_GRAY};
            border-radius: 8px;
            padding: 12px;
            box-shadow: 0 1px 3px rgba(17, 34, 58, 0.05);
            margin-bottom: 0;
            height: 100%;
        }}
        .drill-kpi {{
            margin-bottom: 12px;
        }}
        .kpi-title {{
            font-size: 14px;
            font-weight: 700;
            color: {NAVY};
            margin-bottom: var(--ui-gap);
        }}
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
            gap: var(--ui-gap);
        }}
        .scope-kpi .kpi-grid {{
            grid-template-columns: repeat(5, minmax(0, 1fr));
            gap: 8px;
        }}
        .kpi-card {{
            border: 1px solid {MID_GRAY};
            border-radius: 8px;
            padding: 10px;
            background: {WHITE};
            min-height: 72px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }}
        .metric-label {{
            font-size: 11px;
            font-weight: 600;
            color: {TEXT_MUTED};
            margin-bottom: 5px;
        }}
        .metric-value {{
            font-size: 26px;
            line-height: 1.0;
            font-weight: 800;
            color: {NAVY};
        }}
        .scope-kpi .metric-value {{
            font-size: 22px;
        }}
        @media (max-width: 1100px) {{
            .scope-kpi .kpi-grid {{
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }}
        }}
        .metric-value.warn {{
            color: {MUTED_ORANGE};
        }}
        .metric-value.risk {{
            color: {MUTED_RED};
        }}
        .metric-value.normal {{
            color: {NAVY};
        }}
        .mini-kpi-grid {{
            display: grid;
            grid-template-columns: repeat(6, minmax(0, 1fr));
            gap: 10px;
            margin-bottom: 12px;
        }}
        .mini-kpi-card {{
            background: {WHITE};
            border: 1px solid {MID_GRAY};
            border-radius: 8px;
            padding: 12px;
            min-height: 76px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            box-shadow: 0 1px 3px rgba(17, 34, 58, 0.04);
        }}
        @media (max-width: 1100px) {{
            .mini-kpi-grid {{
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }}
        }}
        .shortage-card {{
            border-color: #efdcc8;
            background: #fffaf5;
        }}
        .panel-box {{
            background: {WHITE};
            border: 1px solid {MID_GRAY};
            border-radius: 8px;
            padding: 12px;
            box-shadow: 0 1px 3px rgba(17, 34, 58, 0.04);
        }}
        .family-grid {{
            display: grid;
            grid-template-columns: repeat(5, minmax(0, 1fr));
            gap: 10px;
        }}
        .family-card {{
            border: 1px solid {MID_GRAY};
            border-radius: 8px;
            background: {WHITE};
            padding: 12px;
            min-height: 128px;
            display: flex;
            flex-direction: column;
            gap: 10px;
        }}
        .family-head {{
            display: flex;
            align-items: baseline;
            justify-content: space-between;
            gap: 10px;
        }}
        .family-head span {{
            color: {NAVY};
            font-size: 14px;
            font-weight: 800;
        }}
        .family-head b {{
            color: {TEXT_MUTED};
            font-size: 12px;
            font-weight: 700;
            font-variant-numeric: tabular-nums;
            white-space: nowrap;
        }}
        .family-shortages {{
            display: flex;
            justify-content: space-between;
            gap: 8px;
            color: {TEXT_MUTED};
            font-size: 12px;
        }}
        .family-shortages b {{
            color: {MUTED_ORANGE};
            font-variant-numeric: tabular-nums;
        }}
        .top-list {{
            display: flex;
            flex-direction: column;
            gap: 7px;
        }}
        .top-row {{
            display: grid;
            grid-template-columns: 32px minmax(220px, 1fr) 120px minmax(240px, 0.9fr);
            gap: 10px;
            align-items: center;
            border: 1px solid #edf1f5;
            border-radius: 8px;
            padding: 8px 10px;
            background: {WHITE};
        }}
        .top-rank {{
            width: 24px;
            height: 24px;
            border-radius: 50%;
            background: #eef3f9;
            color: {NAVY};
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 12px;
            font-weight: 800;
        }}
        .top-name {{
            color: {TEXT_DARK};
            font-size: 12px;
            font-weight: 700;
            overflow-wrap: anywhere;
        }}
        .top-shortage {{
            color: {MUTED_ORANGE};
            font-size: 13px;
            font-weight: 800;
            text-align: right;
            font-variant-numeric: tabular-nums;
        }}
        .top-progress {{
            min-width: 0;
        }}
        .gap-list {{
            display: flex;
            flex-direction: column;
            gap: 7px;
        }}
        .gap-row {{
            display: grid;
            grid-template-columns: 32px minmax(220px, 1fr) minmax(220px, 0.78fr) minmax(220px, 0.78fr) 82px;
            gap: 10px;
            align-items: center;
            border: 1px solid #edf1f5;
            border-radius: 8px;
            padding: 8px 10px;
            background: {WHITE};
        }}
        .gap-progress {{
            min-width: 0;
        }}
        .gap-value {{
            color: {MUTED_ORANGE};
            font-size: 14px;
            font-weight: 800;
            text-align: right;
            font-variant-numeric: tabular-nums;
        }}
        @media (max-width: 900px) {{
            .family-grid {{
                grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            }}
            .top-row {{
                grid-template-columns: 28px 1fr;
            }}
            .top-shortage {{
                text-align: left;
            }}
            .top-progress {{
                grid-column: 2;
            }}
            .gap-row {{
                grid-template-columns: 28px 1fr;
            }}
            .gap-progress, .gap-value {{
                grid-column: 2;
            }}
            .gap-value {{
                text-align: left;
            }}
        }}
        .drill-panel {{
            margin-bottom: 12px;
        }}
        .table-wrap {{
            max-height: 640px;
            overflow: auto;
            border: 1px solid {MID_GRAY};
            border-radius: 8px;
            background: {WHITE};
        }}
        .ops-table {{
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
        }}
        .ops-table th {{
            position: sticky;
            top: 0;
            background: #f7f9fc;
            color: {NAVY};
            font-size: 12px;
            font-weight: 700;
            border-bottom: 1px solid {MID_GRAY};
            padding: 8px 8px;
            z-index: 1;
        }}
        .ops-table td {{
            border-bottom: 1px solid #edf1f5;
            padding: 8px 8px;
            font-size: 12px;
            color: {TEXT_DARK};
            vertical-align: middle;
            background: {WHITE};
        }}
        .ops-table td.left, .ops-table th.left {{
            text-align: left;
        }}
        .ops-table td.num, .ops-table th.num {{
            text-align: right;
            font-variant-numeric: tabular-nums;
        }}
        .ops-table td.num.shortage {{
            color: {MUTED_ORANGE};
            font-weight: 700;
        }}
        .ops-table td.power-cell {{
            text-align: center;
            font-variant-numeric: tabular-nums;
            font-weight: 700;
            color: {NAVY};
        }}
        .ops-table td.power-cell.high {{
            color: {NAVY};
        }}
        .progress-cell {{
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .progress-name {{
            min-width: 28px;
            color: {TEXT_MUTED};
            font-size: 11px;
            font-weight: 700;
        }}
        .progress-track {{
            flex: 1;
            min-width: 80px;
            height: 8px;
            border-radius: 999px;
            background: #edf2f7;
            overflow: hidden;
        }}
        .progress-fill {{
            height: 100%;
            border-radius: 999px;
        }}
        .progress-fill.done {{
            background: {NAVY};
        }}
        .progress-fill.active {{
            background: {SOFT_NAVY};
        }}
        .progress-fill.warn {{
            background: {SOFT_NAVY};
        }}
        .progress-fill.risk {{
            background: {MID_GRAY};
        }}
        .progress-text {{
            min-width: 52px;
            text-align: right;
            font-size: 11px;
            color: {TEXT_MUTED};
            font-variant-numeric: tabular-nums;
        }}
        .status-badge {{
            display: inline-block;
            padding: 3px 8px;
            border-radius: 999px;
            border: 1px solid transparent;
            font-size: 11px;
            font-weight: 700;
            line-height: 1.2;
        }}
        .status-badge.done {{
            background: #eef3f9;
            color: {NAVY};
            border-color: #dbe5f1;
        }}
        .status-badge.active {{
            background: #e8eef6;
            color: {SOFT_NAVY};
            border-color: #d2deeb;
        }}
        .status-badge.warn {{
            background: #f5f7fa;
            color: {TEXT_MUTED};
            border-color: #e1e7ef;
        }}
        .status-badge.risk {{
            background: #f5f7fa;
            color: {TEXT_MUTED};
            border-color: #e1e7ef;
        }}
        .section-title {{
            color: {NAVY};
            font-weight: 700;
            font-size: 16px;
            margin-bottom: 6px;
        }}
        .section-sub {{
            color: {TEXT_MUTED};
            font-size: 12px;
            margin-bottom: 10px;
        }}
        .breadcrumb {{
            display: flex;
            gap: 8px;
            align-items: center;
            color: {TEXT_MUTED};
            font-size: 12px;
            margin: 2px 0 10px 0;
        }}
        .breadcrumb span {{
            color: {NAVY};
            font-weight: 700;
        }}
        .breadcrumb b {{
            color: {TEXT_MUTED};
        }}
        .progress-summary-panel {{
            display: grid;
            grid-template-columns: minmax(0, 1fr) 160px;
            gap: 14px;
            align-items: stretch;
            background: {WHITE};
            border: 1px solid {MID_GRAY};
            border-radius: 8px;
            padding: 12px;
            margin-bottom: 12px;
        }}
        .progress-summary-panel .progress-cell {{
            margin: 10px 0;
        }}
        .dday-box {{
            border: 1px solid {MID_GRAY};
            border-radius: 8px;
            padding: 12px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            background: #fbfcfe;
        }}
        .dday-value {{
            color: {NAVY};
            font-size: 26px;
            font-weight: 800;
            font-variant-numeric: tabular-nums;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_panel_title(title: str, sub: str) -> None:
    st.markdown(
        f"<div class='section-title'>{escape(title)}</div>"
        f"<div class='section-sub'>{escape(sub)}</div>",
        unsafe_allow_html=True,
    )


def build_sales_code_view(code_summary: pd.DataFrame) -> pd.DataFrame:
    view = code_summary.rename(
        columns={
            "sales_code": "판매코드",
            "product_name": "제품명",
            "request_pack": "요청 PACK",
            "request_pcs": "요청 PCS",
            "packing_pack": "포장 PACK",
            "production_code": "생산코드",
            "q_code": "분리코드",
            "r_code": "사출코드",
            "production_basis_qty": "누수규격검사 생산수량",
            "production_shortage_qty": "생산부족수량",
            "production_progress_pct": "생산진도율",
        }
    )
    view = finalize_summary(view)
    return view


def format_sales_code_view(view: pd.DataFrame) -> pd.DataFrame:
    out = view.copy()
    for col in ["요청 PACK", "포장 PACK", "부족 PACK"]:
        out[col] = out[col].map(format_int)
    out["생산진도율"] = out["생산진도율"].map(lambda x: f"{float(x):.1f}%")
    out["포장진도율"] = out["포장진도율"].map(lambda x: f"{x:.1f}%")
    return out[
        [
            "판매코드",
            "요청 PACK",
            "포장 PACK",
            "부족 PACK",
            "생산진도율",
            "포장진도율",
        ]
    ]


def format_production_code_view(view: pd.DataFrame) -> pd.DataFrame:
    out = view.copy()
    for col in ["요청 PACK", "생산부족수량", "포장부족수량"]:
        out[col] = out[col].map(format_int)
    out["생산진도율"] = out["생산진도율"].map(lambda x: f"{float(x):.1f}%")
    out["포장진도율"] = out["포장진도율"].map(lambda x: f"{float(x):.1f}%")
    out["납기일"] = out["납기일"].map(format_date)
    return out[
        [
            "생산코드",
            "연결 판매코드 수",
            "제품명",
            "요청 PACK",
            "생산진도율",
            "생산부족수량",
            "포장진도율",
            "포장부족수량",
            "납기일",
        ]
    ]


def calc_drilldown_kpi(df: pd.DataFrame) -> dict[str, float]:
    if df.empty:
        return {
            "request_qty": 0.0,
            "request_pcs": 0.0,
            "production_shortage_qty": 0.0,
            "production_progress_pct": 0.0,
            "packing_qty": 0.0,
            "shortage_qty": 0.0,
        }
    request_qty = float(df["요청 PACK"].sum()) if "요청 PACK" in df.columns else float(df["요청수량"].sum())
    if "요청 PCS" in df.columns:
        request_pcs = float(df["요청 PCS"].sum())
    elif "요청PCS" in df.columns:
        request_pcs = float(df["요청PCS"].sum())
    else:
        request_pcs = request_qty
    production_shortage_qty = float(df["생산부족수량"].sum()) if "생산부족수량" in df.columns else 0.0
    production_progress_pct = (
        (request_pcs - production_shortage_qty) / request_pcs * 100.0
        if request_pcs > 0
        else 0.0
    )
    return {
        "request_qty": request_qty,
        "request_pcs": request_pcs,
        "production_shortage_qty": production_shortage_qty,
        "production_progress_pct": max(0.0, min(100.0, production_progress_pct)),
        "packing_qty": float(df["포장 PACK"].sum()) if "포장 PACK" in df.columns else float(df["포장수량"].sum()),
        "shortage_qty": float(df["부족 PACK"].sum()) if "부족 PACK" in df.columns else float(df["부족수량"].sum()),
    }


def render_drilldown_kpi(kpi: dict[str, float]) -> None:
    panel_html = f"""
    <div class='kpi-panel drill-kpi'>
      <div class='kpi-grid'>
        <div class='kpi-card'>
          <div class='metric-label'>요청 PACK</div>
          <div class='metric-value'>{format_int(kpi['request_qty'])}</div>
        </div>
        <div class='kpi-card'>
          <div class='metric-label'>생산부족수량</div>
          <div class='metric-value warn'>{format_int(kpi['production_shortage_qty'])}</div>
        </div>
        <div class='kpi-card'>
          <div class='metric-label'>생산진도율</div>
          <div class='metric-value'>{kpi['production_progress_pct']:.1f}%</div>
        </div>
        <div class='kpi-card'>
          <div class='metric-label'>포장수량</div>
          <div class='metric-value'>{format_int(kpi['packing_qty'])}</div>
        </div>
        <div class='kpi-card shortage-card'>
          <div class='metric-label'>포장부족수량</div>
          <div class='metric-value warn'>{format_int(kpi['shortage_qty'])}</div>
        </div>
      </div>
    </div>
    """
    st.markdown(panel_html, unsafe_allow_html=True)


def build_product_drilldown_view(product_summary: pd.DataFrame) -> pd.DataFrame:
    out = product_summary.copy()
    out = out.rename(
        columns={
            "요청 PACK": "요청수량",
            "포장 PACK": "포장수량",
            "포장부족수량": "포장부족수량",
        }
    )
    return out[
        ["제품명", "요청수량", "생산진도율", "포장진도율", "생산부족수량", "포장부족수량", "상태"]
    ]


def build_pack_unit_view(code_summary: pd.DataFrame, product_name: str) -> pd.DataFrame:
    if code_summary.empty:
        return pd.DataFrame(columns=["팩 단위", "요청수량", "포장수량", "부족수량", "진도율", "_sort"])

    selected_base = strip_pack_unit_suffix(product_name)
    work = code_summary.copy()
    if "base_product_name" not in work.columns:
        work["base_product_name"] = work["product_name"].map(strip_pack_unit_suffix)
    if "pack_unit" not in work.columns:
        work["pack_unit"] = work["product_name"].map(extract_pack_unit)
    if "pack_unit_label" not in work.columns:
        work["pack_unit_label"] = [
            format_pack_unit_label(unit, name)
            for unit, name in zip(work["pack_unit"], work["product_name"])
        ]

    scope = work[work["base_product_name"] == selected_base].copy()
    if scope.empty:
        scope = work[work["product_name"] == product_name].copy()
    if scope.empty:
        return pd.DataFrame(columns=["팩 단위", "요청수량", "포장수량", "부족수량", "진도율", "_sort"])

    grouped = (
        scope.groupby(["pack_unit", "pack_unit_label"], dropna=False)
        .agg(
            request_qty=("request_pack", "sum"),
            packing_qty=("packing_pack", "sum"),
        )
        .reset_index()
    )
    grouped["shortage_qty"] = (grouped["request_qty"] - grouped["packing_qty"]).clip(lower=0.0)
    grouped["progress_pct"] = np.where(
        grouped["request_qty"] > 0,
        grouped["packing_qty"] / grouped["request_qty"] * 100.0,
        0.0,
    )
    grouped["progress_pct"] = np.clip(grouped["progress_pct"], 0.0, 100.0)
    grouped["_sort"] = pd.to_numeric(grouped["pack_unit"], errors="coerce").fillna(999999.0)
    grouped = grouped.sort_values("_sort", kind="stable")

    out = grouped.rename(
        columns={
            "pack_unit_label": "팩 단위",
            "request_qty": "요청수량",
            "packing_qty": "포장수량",
            "shortage_qty": "부족수량",
            "progress_pct": "진도율",
        }
    )[["팩 단위", "요청수량", "포장수량", "부족수량", "진도율", "_sort"]]

    total_request = float(out["요청수량"].sum())
    total_packing = float(out["포장수량"].sum())
    total_shortage = max(0.0, total_request - total_packing)
    total_progress = (total_packing / total_request * 100.0) if total_request > 0 else 0.0
    total_row = pd.DataFrame(
        [
            {
                "팩 단위": "전체",
                "요청수량": total_request,
                "포장수량": total_packing,
                "부족수량": total_shortage,
                "진도율": min(100.0, max(0.0, total_progress)),
                "_sort": 1000000.0,
            }
        ]
    )
    return pd.concat([out, total_row], ignore_index=True)


def pack_unit_column_config() -> dict[str, Any]:
    numeric_format = "%,.0f"
    return {
        "요청수량": st.column_config.NumberColumn("요청수량", format=numeric_format),
        "포장수량": st.column_config.NumberColumn("포장수량", format=numeric_format),
        "부족수량": st.column_config.NumberColumn("부족수량", format=numeric_format),
        "진도율": st.column_config.ProgressColumn("진도율", min_value=0, max_value=100, format="%.2f%%"),
        "_sort": None,
    }


def build_production_drilldown_view(code_summary: pd.DataFrame) -> pd.DataFrame:
    out = build_production_code_view(code_summary).rename(
        columns={
            "요청 PACK": "요청수량",
            "포장 PACK": "포장수량",
        }
    )
    return out[
        [
            "생산코드",
            "요청수량",
            "생산부족수량",
            "생산진도율",
            "포장부족수량",
            "연결 판매코드 수",
            "포장진도율",
            "상태",
        ]
    ]


def build_sales_drilldown_view(code_summary: pd.DataFrame) -> pd.DataFrame:
    out = build_sales_code_view(code_summary).rename(
        columns={
            "요청 PACK": "요청수량",
            "포장 PACK": "포장수량",
            "부족 PACK": "부족수량",
        }
    )
    return out[
        [
            "판매코드",
            "요청수량",
            "포장수량",
            "부족수량",
            "생산진도율",
            "포장진도율",
            "생산코드",
            "상태",
        ]
    ]


def build_power_drilldown_view(code_summary: pd.DataFrame) -> pd.DataFrame:
    out = build_power_detail(code_summary).rename(
        columns={
            "요청수량": "요청수량",
            "포장수량": "포장수량",
            "부족수량": "부족수량",
            "진도율": "포장진도율",
        }
    )
    if out.empty:
        return pd.DataFrame(
            columns=["POWER", "요청수량", "포장수량", "부족수량", "생산진도율", "포장진도율", "power_value"]
        )
    out = out.sort_values("power_value", ascending=True, kind="stable")
    return out[
        ["POWER", "요청수량", "포장수량", "부족수량", "생산진도율", "포장진도율", "power_value"]
    ]


def drilldown_column_config() -> dict[str, Any]:
    numeric_format = "%,.0f"
    return {
        "요청합계(PACK)": st.column_config.NumberColumn("요청합계(PACK)", format=numeric_format),
        "요청합계(PCS)": st.column_config.NumberColumn("요청합계(PCS)", format=numeric_format),
        "제품필요수량": st.column_config.NumberColumn("제품필요수량", format=numeric_format),
        "5P 필요팩": st.column_config.NumberColumn("5P 필요팩", format=numeric_format),
        "10P 필요팩": st.column_config.NumberColumn("10P 필요팩", format=numeric_format),
        "30P 필요팩": st.column_config.NumberColumn("30P 필요팩", format=numeric_format),
        "80P 필요팩": st.column_config.NumberColumn("80P 필요팩", format=numeric_format),
        "90P 필요팩": st.column_config.NumberColumn("90P 필요팩", format=numeric_format),
        "기타팩 필요팩": st.column_config.NumberColumn("기타팩 필요팩", format=numeric_format),
        "진도율": st.column_config.ProgressColumn("진도율", min_value=0, max_value=100, format="%.1f%%"),
        "전체진도율": st.column_config.ProgressColumn("전체진도율", min_value=0, max_value=100, format="%.1f%%"),
        "요청합계": st.column_config.NumberColumn("요청합계", format=numeric_format),
        "생산부족": st.column_config.NumberColumn("생산부족", format=numeric_format),
        "포장부족": st.column_config.NumberColumn("포장부족", format=numeric_format),
        "판매코드수": st.column_config.NumberColumn("판매코드수", format=numeric_format),
        "판매코드 수": st.column_config.NumberColumn("판매코드 수", format=numeric_format),
        "요청 PACK": st.column_config.NumberColumn("요청 PACK", format=numeric_format),
        "부족 PACK": st.column_config.NumberColumn("부족 PACK", format=numeric_format),
        "요청": st.column_config.NumberColumn("요청", format=numeric_format),
        "포장": st.column_config.NumberColumn("포장", format=numeric_format),
        "부족": st.column_config.NumberColumn("부족", format=numeric_format),
        "요청PACK": st.column_config.NumberColumn("요청PACK", format=numeric_format),
        "요청PCS": st.column_config.NumberColumn("요청PCS", format=numeric_format),
        "생산": st.column_config.NumberColumn("생산", format=numeric_format),
        "필요팩": st.column_config.NumberColumn("필요팩", format=numeric_format),
        "포장완료PACK": st.column_config.NumberColumn("포장완료PACK", format=numeric_format),
        "포장부족PACK": st.column_config.NumberColumn("포장부족PACK", format=numeric_format),
        "요청수량": st.column_config.NumberColumn("요청수량", format=numeric_format),
        "생산부족수량": st.column_config.NumberColumn("생산부족수량", format=numeric_format),
        "포장부족수량": st.column_config.NumberColumn("포장부족수량", format=numeric_format),
        "생산진도율": st.column_config.ProgressColumn("생산진도율", min_value=0, max_value=100, format="%.1f%%"),
        "포장수량": st.column_config.NumberColumn("포장수량", format=numeric_format),
        "부족수량": st.column_config.NumberColumn("부족수량", format=numeric_format),
        "포장진도율": st.column_config.ProgressColumn("포장진도율", min_value=0, max_value=100, format="%.1f%%"),
        "power_value": None,
        "_power_sort": None,
        "_min_due_date_sort": None,
        "_pack_sort": None,
    }


def get_selected_row(selection_event: Any, df: pd.DataFrame) -> pd.Series | None:
    rows: list[int] = []
    if hasattr(selection_event, "selection"):
        rows = list(getattr(selection_event.selection, "rows", []) or [])
    elif isinstance(selection_event, dict):
        rows = list(selection_event.get("selection", {}).get("rows", []) or [])
    if not rows:
        return None
    row_idx = int(rows[0])
    if row_idx < 0 or row_idx >= len(df):
        return None
    return df.iloc[row_idx]


def render_selectable_table(
    title: str,
    sub: str,
    df: pd.DataFrame,
    key: str,
    height: int,
) -> pd.Series | None:
    render_panel_title(title, sub)
    st.markdown("<div class='panel-box drill-panel'>", unsafe_allow_html=True)
    if df.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
        st.markdown("</div>", unsafe_allow_html=True)
        return None
    column_config = drilldown_column_config()
    for col in df.columns:
        if re.match(r"^\d+(?:\.\d+)?P$", str(col)):
            column_config[col] = st.column_config.NumberColumn(str(col), format="%,.0f")
    event = st.dataframe(
        df,
        hide_index=True,
        height=height,
        width="stretch",
        column_config=column_config,
        on_select="rerun",
        selection_mode="single-row",
        key=key,
    )
    st.markdown("</div>", unsafe_allow_html=True)
    return get_selected_row(event, df)


def render_product_summary_tab(product_summary: pd.DataFrame, code_summary: pd.DataFrame) -> None:
    main_products, _ = split_main_sample(product_summary)
    pack_labels = available_pack_options(code_summary)[1:]

    render_kpi_scope_panels(code_summary)
    render_operation_kpis(product_summary, code_summary)

    family_view = build_family_progress_view(main_products)
    top_shortage_view = build_top_shortage_view(product_summary, top_n=10)
    gap_top_view = build_gap_top_view(product_summary, top_n=10)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    detail_col, report_col = st.columns([4.8, 1.2], gap="small", vertical_alignment="center")
    with detail_col:
        render_panel_title(
            "제품 진도 현황",
            "제품명 기준 PACK pivot, 생산 필요수량, 포장 전체진도율을 한 번에 확인합니다.",
        )
    with report_col:
        ppt_bytes = build_ppt_report(
            product_view=product_summary,
            code_summary=code_summary,
            product_names=product_summary["제품명"],
            scope_label="전체",
        )
        st.download_button(
            "PPT 보고서 다운로드",
            data=ppt_bytes,
            file_name=f"국내_제품_포장현황_운영보고서_{pd.Timestamp.now(tz='Asia/Seoul').strftime('%Y%m%d_%H%M')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            width="stretch",
            key="download_ppt_report",
        )

    pf1, pf2, pf3 = st.columns([1.4, 2.4, 1.5], gap="small")
    with pf1:
        product_scope = st.selectbox(
            "제품 범위",
            options=product_scope_options(product_summary),
            index=0,
            key="tab_summary_product_scope",
        )
    with pf2:
        product_query = st.text_input(
            "제품명/SKU 검색",
            value="",
            placeholder="제품명 또는 SKU 일부 입력",
            key="tab_summary_product_query",
        )
    with pf3:
        product_statuses = st.multiselect(
            "상태 필터",
            STATUS_ORDER,
            default=STATUS_ORDER,
            key="tab_summary_product_status",
        )

    product_filter_base = apply_product_scope_filter(product_summary, product_scope)
    product_filter_base = apply_filters(product_filter_base, query=product_query, statuses=product_statuses)
    product_view_all = build_product_progress_main_view(product_summary, code_summary, pack_labels)
    visible_products = set(product_filter_base["제품명"].astype(str))
    product_view = product_view_all[product_view_all["제품명"].astype(str).isin(visible_products)].copy()
    product_view = product_view.sort_values(
        ["포장부족수량", "제품필요수량", "요청합계(PACK)"],
        ascending=[False, False, False],
        kind="stable",
    )
    selected_product_row = render_selectable_table(
        "제품 진도 현황",
        f"제품 기준 요청/생산/포장 현황 | 표시 건수: {len(product_view):,}",
        product_view,
        key="product_progress_main_table",
        height=430,
    )

    if selected_product_row is not None:
        selected_product = str(selected_product_row["제품명"])
        st.markdown(f"<div class='breadcrumb'>제품 <span>{escape(selected_product)}</span></div>", unsafe_allow_html=True)
        sku_view = build_product_sku_detail_view(code_summary, selected_product)
        selected_sku_row = render_selectable_table(
            "SKU 상세",
            f"{selected_product} 기준 SKU/생산코드 현황 | 표시 건수: {len(sku_view):,}",
            sku_view,
            key="product_sku_detail_table",
            height=260,
        )
        if selected_sku_row is not None:
            selected_sku = str(selected_sku_row["SKU"])
            selected_production = str(selected_sku_row["생산코드"])
            work = with_operational_columns(code_summary)
            sales_scope = work[
                (work["base_product_name"] == selected_product)
                & (work["product_name"] == selected_sku)
                & (work["production_code_display"] == selected_production)
            ].copy()
            sales_detail = build_sales_pack_detail_view(sales_scope)
            st.markdown(
                "<div class='breadcrumb'>"
                f"제품 <span>{escape(selected_product)}</span>"
                f"<b>›</b> SKU <span>{escape(selected_sku)}</span>"
                f"<b>›</b> 생산코드 <span>{escape(selected_production)}</span>"
                "</div>",
                unsafe_allow_html=True,
            )
            render_selectable_table(
                "판매코드 상세",
                f"{selected_sku} / {selected_production} 기준 판매코드 상세 | 표시 건수: {len(sales_detail):,}",
                sales_detail,
                key="product_sales_detail_table",
                height=240,
            )

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    render_panel_title(
        "본품 분류별 진도현황",
        "제품보다 상위 제품군 기준으로 생산/포장 진도와 부족수량을 먼저 확인합니다.",
    )
    st.markdown("<div class='panel-box'>", unsafe_allow_html=True)
    render_family_progress_cards(family_view)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    render_panel_title(
        "포장부족 TOP10",
        "오늘 우선 대응해야 할 제품을 포장부족수량 기준으로 표시합니다.",
    )
    st.markdown("<div class='panel-box'>", unsafe_allow_html=True)
    render_top_shortage_list(top_shortage_view)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    render_panel_title(
        "생산/포장 GAP TOP10",
        "생산진도율 대비 포장진도율이 뒤처진 제품을 GAP 기준으로 표시합니다.",
    )
    st.markdown("<div class='panel-box'>", unsafe_allow_html=True)
    render_gap_top_list(gap_top_view)
    st.markdown("</div>", unsafe_allow_html=True)


def render_production_code_tab(code_summary: pd.DataFrame) -> None:
    render_panel_title(
        "생산코드 상세",
        "생산코드 + POWER 기준으로 부족, 병목, 납기 우선순위를 확인합니다.",
    )
    pack_options = ["전체", *STANDARD_PACK_BUCKETS, "기타팩"]
    power_options = available_production_power_options(code_summary)
    group_options = available_product_group_options(code_summary)

    pc1, pc2, pc3, pc4 = st.columns([1.9, 1.7, 1.2, 1.2], gap="small")
    with pc1:
        product_query = st.text_input(
            "제품명 검색",
            value="",
            placeholder="제품명/SKU 일부 입력",
            key="tab_production_product_query",
        )
    with pc2:
        production_query = st.text_input(
            "생산코드 검색",
            value="",
            placeholder="예: P3015",
            key="tab_production_code_query",
        )
    with pc3:
        selected_power = st.selectbox(
            "POWER 선택",
            options=power_options,
            index=0,
            key="tab_production_power",
        )
    with pc4:
        selected_pack = st.selectbox(
            "PACK 선택",
            options=pack_options,
            index=0,
            key="tab_production_pack",
        )

    pc5, pc6, pc7 = st.columns([1.2, 1.5, 1.2], gap="small")
    with pc5:
        sample_scope = st.selectbox(
            "본품/샘플 선택",
            options=["전체", "본품", "샘플"],
            index=0,
            key="tab_production_sample_scope",
        )
    with pc6:
        selected_group = st.selectbox(
            "분류 선택",
            options=group_options,
            index=0,
            key="tab_production_group",
        )
    with pc7:
        shortage_only = st.checkbox("부족품만 보기", value=False, key="tab_production_shortage_only")

    production_source = filter_production_power_rows(
        code_summary,
        product_query=product_query,
        production_query=production_query,
        power_label=selected_power,
        pack_bucket=selected_pack,
        sample_scope=sample_scope,
        product_group=selected_group,
    )
    production_view = build_production_power_main_view(production_source, shortage_only=shortage_only)
    render_production_power_kpis(production_view)

    selected_production_row = render_selectable_table(
        "생산코드 + POWER 메인 테이블",
        f"납기일, 포장부족, 생산부족 순 정렬 | 표시 건수: {len(production_view):,}",
        production_view,
        key="production_code_main_table",
        height=620,
    )
    if selected_production_row is None:
        return

    selected_production = str(selected_production_row["생산코드"])
    selected_power_for_detail = str(selected_production_row["POWER"])
    st.markdown(
        "<div class='breadcrumb'>"
        f"생산코드 <span>{escape(selected_production)}</span>"
        f"<b>›</b> POWER <span>{escape(selected_power_for_detail)}</span>"
        "</div>",
        unsafe_allow_html=True,
    )

    chart_col, progress_col = st.columns([1.4, 1.0], gap="small")
    with chart_col:
        render_pack_composition_chart(selected_production_row)
    with progress_col:
        render_production_progress_panel(selected_production_row)

    sales_detail = build_production_sales_detail_view(
        production_source,
        production_code=selected_production,
        power_label=selected_power_for_detail,
    )
    render_selectable_table(
        "판매코드 상세",
        f"{selected_production} / {selected_power_for_detail} 기준 판매코드 상세 | 표시 건수: {len(sales_detail):,}",
        sales_detail,
        key="production_sales_detail_table",
        height=360,
    )


def render_sales_code_tab(code_summary: pd.DataFrame) -> None:
    render_panel_title(
        "판매코드 상세",
        "출고/오더 관점에서 판매코드별 생산·포장 진도와 납기 상태를 확인합니다.",
    )
    pack_options = available_pack_options(code_summary)
    power_options = available_power_options(code_summary)
    customer_options = available_customer_options(code_summary)

    sf1, sf2, sf3, sf4, sf5, sf6 = st.columns([1.7, 1.5, 1.5, 1.1, 1.2, 1.5], gap="small")
    with sf1:
        product_query = st.text_input(
            "제품명 검색",
            value="",
            placeholder="제품명/SKU 일부 입력",
            key="tab_sales_product_query",
        )
    with sf2:
        sales_query = st.text_input(
            "판매코드 검색",
            value="",
            placeholder="예: S309",
            key="tab_sales_code_query",
        )
    with sf3:
        production_query = st.text_input(
            "생산코드 검색",
            value="",
            placeholder="예: P3015",
            key="tab_sales_production_query",
        )
    with sf4:
        selected_pack = st.selectbox("PACK 선택", options=pack_options, index=0, key="tab_sales_pack")
    with sf5:
        selected_power = st.selectbox("POWER 선택", options=power_options, index=0, key="tab_sales_power")
    with sf6:
        selected_customer = st.selectbox("거래처 선택", options=customer_options, index=0, key="tab_sales_customer")

    sales_source = filter_operational_code_summary(
        code_summary,
        product_query=product_query,
        production_query=production_query,
        sales_query=sales_query,
        pack_label=selected_pack,
        power_label=selected_power,
        customer_name=selected_customer,
    )
    sales_view = build_sales_order_main_view(sales_source)
    selected_sales_row = render_selectable_table(
        "판매코드",
        f"판매코드 기준 출고/오더 상세 | 표시 건수: {len(sales_view):,}",
        sales_view.drop(columns=["power_value"], errors="ignore"),
        key="sales_code_main_table",
        height=620,
    )
    if selected_sales_row is None:
        return

    selected_sales = str(selected_sales_row["판매코드"])
    st.markdown(f"<div class='breadcrumb'>판매코드 <span>{escape(selected_sales)}</span></div>", unsafe_allow_html=True)
    inventory_view = empty_inventory_detail_view()
    render_panel_title("SKU 상세", "LOT/멸균NO/ERP재고/재작업/샘플 전환 가능 여부")
    st.markdown("<div class='panel-box drill-panel'>", unsafe_allow_html=True)
    st.info("현재 연결된 엑셀 원천에는 LOT, 멸균NO, ERP재고 컬럼이 없어 상세 재고 항목은 빈 표로 표시합니다.")
    st.dataframe(inventory_view, hide_index=True, height=120, width="stretch")
    st.markdown("</div>", unsafe_allow_html=True)


def render_power_tab(code_summary: pd.DataFrame) -> None:
    render_panel_title(
        "POWER 상세",
        "렌즈 POWER 기준 요청/생산/포장/부족 현황과 하위 생산·판매코드를 확인합니다.",
    )
    power_options = available_power_options(code_summary)

    pf1, pf2, pf3, pf4 = st.columns([2.0, 1.7, 1.7, 1.2], gap="small")
    with pf1:
        product_query = st.text_input("제품명", value="", placeholder="제품명/SKU 일부 입력", key="tab_power_product_query")
    with pf2:
        production_query = st.text_input("생산코드", value="", placeholder="예: P3015", key="tab_power_production_query")
    with pf3:
        sales_query = st.text_input("판매코드", value="", placeholder="예: S309", key="tab_power_sales_query")
    with pf4:
        selected_power = st.selectbox("POWER", options=power_options, index=0, key="tab_power_power")

    power_source = filter_operational_code_summary(
        code_summary,
        product_query=product_query,
        production_query=production_query,
        sales_query=sales_query,
        power_label=selected_power,
    )
    power_detail_for_heatmap = build_power_detail(power_source)
    heatmap = build_power_heatmap(power_detail_for_heatmap)
    if heatmap is not None:
        st.plotly_chart(heatmap, width="stretch")

    power_summary = build_power_summary_view(power_source)
    ops_kpi = calc_power_ops_kpi(power_detail_for_heatmap)
    oc1, oc2, oc3, oc4, oc5 = st.columns(5, gap="small")
    oc1.metric("대상 도수", f"{ops_kpi['rows']:,}")
    oc2.metric("포장부족 도수", f"{ops_kpi['shortage_rows']:,}")
    oc3.metric("미착수 도수", f"{ops_kpi['not_started_rows']:,}")
    oc4.metric("하이파워 부족", f"{ops_kpi['high_power_shortage_rows']:,}")
    oc5.metric("포장부족수량 합계", format_int(ops_kpi["shortage_qty"]))

    selected_power_row = render_selectable_table(
        "POWER 히트맵 상세",
        f"POWER 기준 요청/생산/포장/부족 | 표시 건수: {len(power_summary):,}",
        power_summary.drop(columns=["power_value"], errors="ignore"),
        key="power_summary_table",
        height=430,
    )
    if selected_power_row is None:
        return

    selected_power_detail = str(selected_power_row["POWER"])
    st.markdown(f"<div class='breadcrumb'>POWER <span>{escape(selected_power_detail)}</span></div>", unsafe_allow_html=True)
    sku_detail = build_power_sku_detail_view(power_source, selected_power_detail)
    render_selectable_table(
        "SKU 상세",
        f"{selected_power_detail} 기준 생산코드/판매코드 상세 | 표시 건수: {len(sku_detail):,}",
        sku_detail,
        key="power_sku_detail_table",
        height=320,
    )


def render_drilldown_tab(product_summary: pd.DataFrame, code_summary: pd.DataFrame) -> None:
    render_drilldown_kpi(calc_drilldown_kpi(product_summary))

    filter_col1, filter_col2 = st.columns([3, 2], gap="small")
    with filter_col1:
        query = st.text_input("제품명 검색", value="", placeholder="제품명 일부 입력", key="drill_product_query")
    with filter_col2:
        statuses = st.multiselect("상태 필터", STATUS_ORDER, default=STATUS_ORDER, key="drill_product_status")

    product_filtered = apply_filters(product_summary, query=query, statuses=statuses)
    product_filtered = product_filtered.sort_values(
        ["포장부족수량", "생산부족수량", "요청 PACK"],
        ascending=[False, False, False],
        kind="stable",
    )
    product_view = build_product_drilldown_view(product_filtered)
    selected_product_row = render_selectable_table(
        "제품 진도현황",
        f"제품 기준 요청/생산부족/포장 현황 | 표시 건수: {len(product_view):,}",
        product_view,
        key="drill_product_table",
        height=430,
    )
    if selected_product_row is None:
        return

    selected_product = str(selected_product_row["제품명"])
    if "base_product_name" in code_summary.columns:
        product_scope = code_summary[code_summary["base_product_name"] == selected_product].copy()
    else:
        product_scope = code_summary[code_summary["product_name"].map(strip_pack_unit_suffix) == selected_product].copy()
    if product_scope.empty:
        product_scope = code_summary[code_summary["product_name"] == selected_product].copy()
    st.markdown(f"<div class='breadcrumb'>제품 <span>{escape(selected_product)}</span></div>", unsafe_allow_html=True)

    pack_unit_view = build_pack_unit_view(code_summary, selected_product)
    render_panel_title(
        "팩 단위 포장 진도",
        f"{strip_pack_unit_suffix(selected_product)} 기준 팩 단위 요청/포장/부족/진도율",
    )
    st.markdown("<div class='panel-box drill-panel'>", unsafe_allow_html=True)
    if pack_unit_view.empty:
        st.warning("팩 단위 상세 데이터가 없습니다.")
    else:
        st.dataframe(
            pack_unit_view,
            hide_index=True,
            height=min(260, 88 + 36 * len(pack_unit_view)),
            width="stretch",
            column_config=pack_unit_column_config(),
        )
    st.markdown("</div>", unsafe_allow_html=True)

    production_view = build_production_drilldown_view(product_scope)
    production_view = production_view.sort_values(
        ["포장부족수량", "생산부족수량", "요청수량"],
        ascending=[False, False, False],
        kind="stable",
    )
    selected_production_row = render_selectable_table(
        "생산코드",
        f"{selected_product} 기준 생산코드 현황 | 표시 건수: {len(production_view):,}",
        production_view,
        key="drill_production_table",
        height=300,
    )
    if selected_production_row is None:
        return

    selected_production = str(selected_production_row["생산코드"])
    production_key = selected_production if selected_production != "(생산코드 미기재)" else ""
    production_scope = product_scope[product_scope["production_code"].replace("", "(생산코드 미기재)") == selected_production].copy()
    if production_scope.empty and production_key:
        production_scope = product_scope[product_scope["production_code"] == production_key].copy()
    st.markdown(
        "<div class='breadcrumb'>"
        f"제품 <span>{escape(selected_product)}</span>"
        f"<b>›</b> 생산코드 <span>{escape(selected_production)}</span>"
        "</div>",
        unsafe_allow_html=True,
    )

    sales_view = build_sales_drilldown_view(production_scope)
    sales_view = sales_view.sort_values(["부족수량", "요청수량"], ascending=[False, False], kind="stable")
    selected_sales_row = render_selectable_table(
        "판매코드",
        f"{selected_production} 기준 판매코드 현황 | 표시 건수: {len(sales_view):,}",
        sales_view,
        key="drill_sales_table",
        height=280,
    )
    if selected_sales_row is None:
        return

    selected_sales = str(selected_sales_row["판매코드"])
    sales_scope = production_scope[production_scope["sales_code"] == selected_sales].copy()
    st.markdown(
        "<div class='breadcrumb'>"
        f"제품 <span>{escape(selected_product)}</span>"
        f"<b>›</b> 생산코드 <span>{escape(selected_production)}</span>"
        f"<b>›</b> 판매코드 <span>{escape(selected_sales)}</span>"
        "</div>",
        unsafe_allow_html=True,
    )

    power_view = build_power_drilldown_view(sales_scope)
    render_selectable_table(
        "POWER",
        f"{selected_sales} 기준 POWER 상세 운영현황 | 표시 건수: {len(power_view):,}",
        power_view,
        key="drill_power_table",
        height=240,
    )


def main() -> None:
    render_style()
    st.title("국내 제품 포장현황 대시보드")
    st.caption("국내 제품 전체 포장 진도현황 | 우선 대응 품목 판단용 통합 운영 화면")

    base_dir = Path.cwd()
    try:
        files = discover_source_files(base_dir)
        request_df = normalize_request(files.request_file)
        packing_df = normalize_packing(files.packing_file)
        product_summary, unmatched_packing_total, code_summary = build_summaries(request_df, packing_df)
        progress_df, progress_info = normalize_progress(files.progress_file, request_df)
        product_summary = enrich_product_summary(product_summary, progress_df)
        code_summary = attach_progress_to_code_summary(code_summary, progress_df)
    except DashboardConfigError as exc:
        st.error("데이터 설정 오류")
        for msg in exc.messages:
            st.write(f"- {msg}")
        st.stop()
    except Exception as exc:
        st.error(f"처리 중 오류가 발생했습니다: {exc}")
        st.stop()

    progress_file_label = files.progress_file.name if files.progress_file is not None else "찾지 못함"
    st.info(
        f"요청 파일: {files.request_file.name}\n"
        f"포장 파일: {files.packing_file.name}\n"
        f"수요정보 파일: {progress_file_label}"
    )
    if files.progress_file is None:
        st.warning("수요정보(전공정) 파일을 찾지 못해 생산부족수량은 0으로 처리됩니다.")
    else:
        st.caption(
            "누수/규격검사 생산수량을 생산부족수량으로 반영: "
            f"{progress_info['domestic_rows']:,}/{progress_info['total_rows']:,}행 "
            f"(코드 매칭 {progress_info['code_rows']:,}행, 제품명 보조 매칭 {progress_info['name_rows']:,}행)"
        )
    if unmatched_packing_total > 0:
        st.warning(
            f"요청 데이터에 없는 판매코드 포장량 {unmatched_packing_total:,.0f} PACK은 "
            "제품 진도 집계에서 제외되었습니다."
        )
    tab_summary, tab_production, tab_sales, tab_power = st.tabs(
        ["제품 진도 현황", "생산코드 상세", "판매코드 상세", "POWER 상세"]
    )

    with tab_summary:
        render_product_summary_tab(product_summary, code_summary)
    with tab_production:
        render_production_code_tab(code_summary)
    with tab_sales:
        render_sales_code_tab(code_summary)
    with tab_power:
        render_power_tab(code_summary)


if __name__ == "__main__":
    main()
