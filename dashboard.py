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
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt
import streamlit as st


st.set_page_config(
    page_title="국내 제품 생산, 포장 모니터링",
    layout="wide",
    initial_sidebar_state="collapsed",
)


class DashboardConfigError(Exception):
    def __init__(self, messages: list[str]):
        super().__init__("\n".join(messages))
        self.messages = messages


@dataclass
class SourceFiles:
    request_file: Path
    packing_file: Path
    progress_file: Path | None = None
    inventory_file: Path | None = None
    daily_inventory_file: Path | None = None


STATUS_ORDER = ["미착수", "진행중", "완료"]
UNIT_PACK = "PACK 기준"
UNIT_PCS = "PCS 기준"
UNIT_OPTIONS = [UNIT_PACK, UNIT_PCS]
DASHBOARD_TABS = ["제품 진도 현황", "일일 재고 대응", "생산코드 상세", "판매코드 상세", "POWER 상세", "포장 LOT 상세"]
DAILY_INVENTORY_FILE_STANDARD = "클라렌사업본부 재고현황_YYMMDD.xlsx"
DAILY_INVENTORY_FILE_KEYWORDS = ["클라렌사업본부 재고현황", "재고현황_"]
SAMPLE_KEYWORDS = ["샘플"]
GROUP_ORDER = ["전체", "본품", "샘플", "PIA", "Clalen", "Toric", "1Day", "Color", "Monthly", "기타"]
PRODUCTION_CODE_PACK_LABELS = ["1P", "2P", "5P", "6P", "10P", "30P", "40P", "80P", "90P"]
DATA_CACHE_VERSION = 10
PRODUCTION_PROGRESS_DUE_MONTH = "2026-06"
MAIN_PRODUCT_FAMILY_ORDER = [
    "전체",
    "Clalen 1Day",
    "O2O2 1Day",
    "O2O2 D 컬러",
    "O2O2 D Micelia",
    "O2O2 D Toric",
    "O2O2 Monthly",
    "O2O2 M Micelia",
    "Clear",
    "PIA 1Day",
    "PIA Monthly",
    "Iris 컬러",
    "Iris Toric",
    "T38 Toric",
    "기타 Toric",
    "부자재/기타",
    "기타",
]
DETAIL_FAMILY_PLACEHOLDER = "본품분류 선택"
FAMILY_CARD_SECTION_ORDER = ["1DAY", "FRP", "기타"]
FAMILY_CARD_1DAY_NAMES = {
    "O2O2 D 컬러",
    "O2O2 D Micelia",
    "O2O2 D Toric",
    "Iris 컬러",
    "Iris Toric",
}
FAMILY_CARD_MISC_NAMES = {"PIA 1Day", "PIA Monthly"}
STANDARD_PACK_BUCKETS = ["5P", "10P", "30P", "80P", "90P"]
PRODUCT_QUERY_ALIASES = {
    "딥블랙": ["Deep Black"],
    "레이크그레이": ["Lake Gray"],
    "뮤트브라운": ["Mute Brown"],
    "페일초코": ["Pale Choco"],
    "소울브라운": ["SoulBrown", "Soul Brown"],
    "수지그레이": ["Suzy Gray"],
    "수지브라운": ["Suzy Brown"],
    "알리샤브라운": ["Alicia Brown"],
    "페즈브라운": ["Fez Brown"],
    "블루문": ["Blue Moon", "Bluemoon"],
    "미셀리아": ["Micelia"],
}
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

COLOR_BLUE = "#185FA5"
COLOR_TEAL = "#1D9E75"
COLOR_ORANGE = "#D85A30"
COLOR_AMBER = "#BA7517"
COLOR_ALERT_BG = "#FAECE7"
COLOR_ALERT_BD = "#F0997B"
BG_PAGE = "#F4F4F2"
BG_CARD = "#FFFFFF"
BG_SECTION = "#EFEFEC"
TEXT_PRIMARY = "#1A1A1A"
TEXT_SECONDARY = "#6B6B68"
TEXT_TERTIARY = "#9E9D99"
BORDER_DEFAULT = "rgba(0,0,0,0.10)"
BORDER_LIGHT = "rgba(0,0,0,0.06)"

NAVY = COLOR_BLUE
SOFT_NAVY = COLOR_TEAL
WHITE = BG_CARD
LIGHT_GRAY = BG_PAGE
MID_GRAY = "#DEDED8"
TEXT_DARK = TEXT_PRIMARY
TEXT_MUTED = TEXT_SECONDARY
MUTED_ORANGE = COLOR_ORANGE
MUTED_RED = COLOR_ORANGE
PPT_FONT_NAME = "Noto Sans KR"
REPORT_BG = "#F6F7F9"
REPORT_PANEL = "#FFFFFF"
REPORT_PANEL_LINE = "#DDE3EA"
REPORT_HEADER = "#121A2A"
REPORT_MUTED = "#6B7280"
REPORT_FAINT = "#F1F4F8"
REPORT_ROW_ALT = "#F8FAFC"
REPORT_ACCENT = "#D85A30"
REPORT_NAVY = "#172033"
REPORT_BLUE_SOFT = "#EAF2F9"
REPORT_ACCENT_SOFT = "#FFF1E8"
REPORT_GREEN_SOFT = "#EAF7F1"

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
    "product_name": ["판매명", "생산명", "제품명", "제품 명", "품명", "product_name"],
    "lot_no": ["LOTNO", "LOT NO", "LOT", "lot_no"],
    "barcode_info": ["바코드정보", "바코드 정보", "barcode_info"],
    "packing_date": ["마킹일", "마킹시간", "포장일", "일자", "date"],
    "packing_qty": ["팩수량", "포장수량", "포장완료수량", "수량", "packing_qty"],
    "packing_pcs": ["낱개수량", "낱개 수량", "PCS수량", "PCS 수량", "pcs_qty", "packing_pcs"],
    "pack_unit": ["포장단위", "포장 단위", "입수", "pack_unit"],
}

YONGMA_COLS = {
    "sales_code": ["제품코드", "제품 코드", "판매코드", "판매 코드", "품목코드", "sales_code"],
    "product_name": ["품명", "제품명", "제품 명", "product_name"],
    "lot_no": ["LOTNO", "LOT NO", "LOT", "lot_no"],
    "receipt_qty": ["수량", "입고수량", "용마입고수량", "receipt_qty"],
}

SAMPLE_AVAILABLE_COLS = {
    "product_code": ["제품코드", "제품 코드", "품목코드", "생산코드", "production_code", "product_code"],
    "sample_available_qty": [
        "샘플 신청 가능 수량",
        "샘플신청가능수량",
        "샘플 신청가능수량",
        "sample_available_qty",
    ],
}

INVENTORY_STOCK_THRESHOLD_DEFAULT = 100

INVENTORY_COLS = {
    "sales_code": ["제품코드", "제품 코드", "판매코드", "판매 코드", "품목코드", "SKU", "sku"],
    "product_name": ["제품명", "제품 명", "품명", "product_name"],
    "available_stock_pack": [
        "실시간가용재고",
        "실시간 가용재고",
        "가용재고",
        "가용 재고",
        "재고수량",
        "재고 수량",
        "수량",
        "available_stock",
    ],
    "total_stock_pack": ["총수량", "총 수량", "총재고", "총 재고", "total_stock"],
    "product_spec": ["제품규격", "제품 규격", "규격", "product_spec"],
    "updated_at": ["전송일자", "전송 일자", "수집일자", "업데이트일자", "updated_at"],
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


def max_datetime(series: pd.Series) -> pd.Timestamp:
    dates = pd.to_datetime(series, errors="coerce")
    dates = dates.dropna()
    if dates.empty:
        return pd.NaT
    return dates.max()


def sum_numeric_or_nan(series: pd.Series) -> float:
    numbers = pd.to_numeric(series, errors="coerce").dropna()
    if numbers.empty:
        return np.nan
    return float(numbers.sum())


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
    inventory_file = pick_latest_by_name(files, ["용마WMS재고현황", "WMS재고현황", "WMS"])
    daily_inventory_file = pick_latest_by_name(files, DAILY_INVENTORY_FILE_KEYWORDS)

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
            if inventory_file is None:
                if has_alias(cols, INVENTORY_COLS["sales_code"]) and has_alias(
                    cols,
                    INVENTORY_COLS["available_stock_pack"],
                ):
                    inventory_file = file

    messages: list[str] = []
    if request_file is None:
        messages.append("생산요청등록(국내) 파일을 찾지 못했습니다.")
    if packing_file is None:
        messages.append("포장설비투입현황 파일을 찾지 못했습니다.")
    if messages:
        raise DashboardConfigError(messages)

    return SourceFiles(
        request_file=request_file,
        packing_file=packing_file,
        progress_file=progress_file,
        inventory_file=inventory_file,
        daily_inventory_file=daily_inventory_file,
    )


def read_excel_preferred_sheet(path: Path, preferred_sheet: str) -> pd.DataFrame:
    try:
        xl = pd.ExcelFile(path)
    except Exception:
        return pd.read_excel(path)
    sheet_name = preferred_sheet if preferred_sheet in xl.sheet_names else xl.sheet_names[0]
    return xl.parse(sheet_name=sheet_name)


def has_excel_sheet(path: Path, sheet_name: str) -> bool:
    try:
        return sheet_name in pd.ExcelFile(path).sheet_names
    except Exception:
        return False


def read_resolved_excel_sheet(
    xl: pd.ExcelFile,
    sheet_name: str,
    alias_map: dict[str, list[str]],
    required_keys: list[str],
    file_label: str,
) -> pd.DataFrame:
    header = xl.parse(sheet_name=sheet_name, nrows=0)
    cols = resolve_columns(
        header,
        alias_map,
        required_keys=required_keys,
        file_label=file_label,
    )
    usecols = list(dict.fromkeys(cols.values()))
    return xl.parse(sheet_name=sheet_name, usecols=usecols)


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


def normalize_packing_frame(raw: pd.DataFrame, file_label: str) -> pd.DataFrame:
    cols = resolve_columns(
        raw,
        PACKING_COLS,
        required_keys=["sales_code", "packing_qty"],
        file_label=file_label,
    )
    packing_pack = to_number(raw[cols["packing_qty"]])
    if "packing_pcs" in cols:
        packing_pcs = to_number(raw[cols["packing_pcs"]])
        if "pack_unit" in cols:
            pack_unit = to_number(raw[cols["pack_unit"]])
        elif "product_name" in cols:
            pack_unit = raw[cols["product_name"]].map(extract_pack_unit)
        else:
            pack_unit = pd.Series(0.0, index=raw.index)
        pack_unit = pd.to_numeric(pack_unit, errors="coerce").fillna(0.0)
        # Some exported pack counts are formatted as Excel dates; recover PACK from PCS when needed.
        derived_pack = (packing_pcs / pack_unit.where(pack_unit > 0, np.nan)).replace([np.inf, -np.inf], np.nan).fillna(0.0)
        packing_pack = packing_pack.where((packing_pack > 0) | (derived_pack <= 0), derived_pack)

    out = pd.DataFrame(
        {
            "sales_code": raw[cols["sales_code"]].map(clean_str),
            "packing_pack": packing_pack,
        }
    )
    out["sales_code_key"] = out["sales_code"].map(normalize_match_key)
    out["packing_product_name"] = raw[cols["product_name"]].map(clean_str) if "product_name" in cols else ""
    out["packing_lot"] = raw[cols["lot_no"]].map(clean_str) if "lot_no" in cols else ""
    out["packing_lot_key"] = out["packing_lot"].map(normalize_match_key)
    out["packing_barcode"] = raw[cols["barcode_info"]].map(clean_str) if "barcode_info" in cols else ""
    out["packing_barcode_key"] = out["packing_barcode"].map(normalize_match_key)
    out["packing_date"] = (
        parse_datetime_series(raw[cols["packing_date"]]) if "packing_date" in cols else pd.NaT
    )
    return out


def normalize_packing(path: Path) -> pd.DataFrame:
    try:
        xl = pd.ExcelFile(path)
        sheet_name = "포장실적" if "포장실적" in xl.sheet_names else xl.sheet_names[0]
        raw = read_resolved_excel_sheet(
            xl,
            sheet_name,
            PACKING_COLS,
            required_keys=["sales_code", "packing_qty"],
            file_label=path.name,
        )
    except DashboardConfigError:
        raise
    except Exception:
        raw = read_excel_preferred_sheet(path, "포장실적")
    return normalize_packing_frame(raw, path.name)


def empty_yongma_movement_df() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "sales_code",
            "sales_code_key",
            "yongma_product_name",
            "yongma_lot",
            "yongma_lot_key",
            "yongma_in_pack",
        ]
    )


def normalize_yongma_movement_frame(raw: pd.DataFrame, file_label: str) -> pd.DataFrame:
    cols = resolve_columns(
        raw,
        YONGMA_COLS,
        required_keys=["sales_code", "lot_no", "receipt_qty"],
        file_label=file_label,
    )
    out = pd.DataFrame(
        {
            "sales_code": raw[cols["sales_code"]].map(clean_str),
            "yongma_lot": raw[cols["lot_no"]].map(clean_str),
            "yongma_in_pack": to_number(raw[cols["receipt_qty"]]),
        }
    )
    out["sales_code_key"] = out["sales_code"].map(normalize_match_key)
    out["yongma_product_name"] = raw[cols["product_name"]].map(clean_str) if "product_name" in cols else ""
    out["yongma_lot_key"] = out["yongma_lot"].map(normalize_match_key)
    return out[(out["sales_code_key"] != "") & (out["yongma_in_pack"] > 0)].copy()


def normalize_yongma_movement(path: Path) -> pd.DataFrame:
    sheet_name = "용마이동현황"
    if not has_excel_sheet(path, sheet_name):
        return empty_yongma_movement_df()

    try:
        xl = pd.ExcelFile(path)
        raw = read_resolved_excel_sheet(
            xl,
            sheet_name,
            YONGMA_COLS,
            required_keys=["sales_code", "lot_no", "receipt_qty"],
            file_label=f"{path.name}:{sheet_name}",
        )
    except DashboardConfigError:
        raise
    except Exception:
        raw = pd.read_excel(path, sheet_name=sheet_name)
    return normalize_yongma_movement_frame(raw, f"{path.name}:{sheet_name}")


def empty_sample_available_df() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "product_code",
            "production_code_key",
            "sample_available_pcs",
        ]
    )


def normalize_sample_available_frame(raw: pd.DataFrame, file_label: str) -> pd.DataFrame:
    cols = resolve_columns(
        raw,
        SAMPLE_AVAILABLE_COLS,
        required_keys=["product_code", "sample_available_qty"],
        file_label=file_label,
    )
    out = pd.DataFrame(
        {
            "product_code": raw[cols["product_code"]].map(clean_str),
            "sample_available_pcs": to_number(raw[cols["sample_available_qty"]]),
        }
    )
    out["production_code_key"] = out["product_code"].map(normalize_match_key)
    return out[(out["production_code_key"] != "") & (out["sample_available_pcs"] > 0)].copy()


def normalize_sample_available(path: Path) -> pd.DataFrame:
    sheet_name = "샘플신청가능수량"
    if not has_excel_sheet(path, sheet_name):
        return empty_sample_available_df()

    try:
        xl = pd.ExcelFile(path)
        raw = read_resolved_excel_sheet(
            xl,
            sheet_name,
            SAMPLE_AVAILABLE_COLS,
            required_keys=["product_code", "sample_available_qty"],
            file_label=f"{path.name}:{sheet_name}",
        )
    except DashboardConfigError:
        raise
    except Exception:
        raw = pd.read_excel(path, sheet_name=sheet_name)
    return normalize_sample_available_frame(raw, f"{path.name}:{sheet_name}")


def normalize_packing_workbook(path: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    try:
        xl = pd.ExcelFile(path)
    except Exception:
        return (
            normalize_packing(path),
            normalize_yongma_movement(path),
            normalize_sample_available(path),
        )

    packing_sheet = "포장실적" if "포장실적" in xl.sheet_names else xl.sheet_names[0]
    packing_raw = read_resolved_excel_sheet(
        xl,
        packing_sheet,
        PACKING_COLS,
        required_keys=["sales_code", "packing_qty"],
        file_label=path.name,
    )
    packing_df = normalize_packing_frame(packing_raw, path.name)

    yongma_sheet = "용마이동현황"
    if yongma_sheet in xl.sheet_names:
        yongma_raw = read_resolved_excel_sheet(
            xl,
            yongma_sheet,
            YONGMA_COLS,
            required_keys=["sales_code", "lot_no", "receipt_qty"],
            file_label=f"{path.name}:{yongma_sheet}",
        )
        yongma_df = normalize_yongma_movement_frame(yongma_raw, f"{path.name}:{yongma_sheet}")
    else:
        yongma_df = empty_yongma_movement_df()

    sample_sheet = "샘플신청가능수량"
    if sample_sheet in xl.sheet_names:
        sample_raw = read_resolved_excel_sheet(
            xl,
            sample_sheet,
            SAMPLE_AVAILABLE_COLS,
            required_keys=["product_code", "sample_available_qty"],
            file_label=f"{path.name}:{sample_sheet}",
        )
        sample_available_df = normalize_sample_available_frame(sample_raw, f"{path.name}:{sample_sheet}")
    else:
        sample_available_df = empty_sample_available_df()

    return packing_df, yongma_df, sample_available_df


def empty_inventory_df() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "sales_code",
            "sales_code_key",
            "inventory_product_name",
            "available_stock_pack",
            "inventory_total_stock_pack",
            "inventory_product_spec",
            "inventory_updated_at",
        ]
    )


def normalize_inventory(path: Path | None) -> pd.DataFrame:
    if path is None:
        return empty_inventory_df()

    raw = pd.read_excel(path)
    cols = resolve_columns(
        raw,
        INVENTORY_COLS,
        required_keys=["sales_code", "available_stock_pack"],
        file_label=path.name,
    )
    out = pd.DataFrame(
        {
            "sales_code": raw[cols["sales_code"]].map(clean_str),
            "sales_code_key": raw[cols["sales_code"]].map(normalize_match_key),
            "available_stock_pack": to_number(raw[cols["available_stock_pack"]]),
        }
    )
    out["inventory_product_name"] = (
        raw[cols["product_name"]].map(clean_str) if "product_name" in cols else ""
    )
    out["inventory_total_stock_pack"] = (
        to_number(raw[cols["total_stock_pack"]]) if "total_stock_pack" in cols else np.nan
    )
    out["inventory_product_spec"] = (
        raw[cols["product_spec"]].map(clean_str) if "product_spec" in cols else ""
    )
    out["inventory_updated_at"] = (
        parse_datetime_series(raw[cols["updated_at"]]) if "updated_at" in cols else pd.NaT
    )
    return out[out["sales_code_key"] != ""].copy()


DAILY_INVENTORY_COLUMNS = [
    "제품명",
    "제품코드",
    "PACK",
    "POWER",
    "재고수량",
    "전일재고",
    "재고증감",
    "긴급요청",
    "대상품목",
]


def empty_daily_inventory_df() -> pd.DataFrame:
    return pd.DataFrame(columns=DAILY_INVENTORY_COLUMNS)


def numeric_scalar(value: Any, default: float = 0.0) -> float:
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number):
        return default
    return float(number)


def daily_power_label(value: Any) -> str:
    text = clean_str(value).replace("−", "-").replace("–", "-").replace("—", "-")
    if not text:
        return ""
    if text.upper() == "PL":
        return format_power(0.0)
    number = pd.to_numeric(text, errors="coerce")
    if pd.isna(number):
        return ""
    return format_power(float(number))


def parse_daily_power_tokens(value: Any) -> list[str]:
    text = clean_str(value).replace("−", "-").replace("–", "-").replace("—", "-")
    if not text:
        return []
    tokens = re.findall(r"PL|[+-]?\d+(?:\.\d+)?", text, flags=re.IGNORECASE)
    powers = [daily_power_label(token) for token in tokens]
    return list(dict.fromkeys([power for power in powers if power]))


def extract_daily_pack_label(value: Any) -> str:
    text = clean_str(value)
    unit = extract_pack_unit(text)
    if pd.notna(unit) and float(unit) > 0:
        return base_pack_label(unit)
    match = re.search(r"(\d+(?:\.\d+)?)\s*(?:P|p|팩|개입)", text)
    if not match:
        return ""
    return base_pack_label(float(match.group(1)))


def daily_inventory_key(product_name: Any, pack_label: Any, power_label: Any) -> str:
    product = compact_query_text(product_name)
    return f"{product}|{clean_str(pack_label).upper()}|{clean_str(power_label)}"


def normalize_daily_emergency_requests(xl: pd.ExcelFile) -> pd.DataFrame:
    sheet_name = "긴급요청"
    if sheet_name not in xl.sheet_names:
        return empty_daily_inventory_df()

    raw = xl.parse(sheet_name=sheet_name, header=None)
    header_idx = None
    for idx, row in raw.iterrows():
        values = [clean_str(value) for value in row.tolist()]
        if "제품명" in values and "제품코드" in values:
            header_idx = idx
            break
    if header_idx is None:
        return empty_daily_inventory_df()

    header_values = [clean_str(value) for value in raw.iloc[header_idx].tolist()]
    product_col = header_values.index("제품명")
    code_col = header_values.index("제품코드")
    target_start_col = code_col + 1

    rows: list[dict[str, Any]] = []
    for row_idx in range(header_idx + 1, len(raw)):
        row = raw.iloc[row_idx]
        product_name = clean_str(row.iloc[product_col] if product_col < len(row) else "")
        product_code = clean_str(row.iloc[code_col] if code_col < len(row) else "")
        if not product_name and not product_code:
            continue
        target_text = " ".join(clean_str(value) for value in row.iloc[target_start_col:].tolist() if clean_str(value))
        powers = parse_daily_power_tokens(target_text)
        pack_label = extract_daily_pack_label(product_name)
        for power in powers:
            rows.append(
                {
                    "제품명": product_name,
                    "제품코드": product_code,
                    "PACK": pack_label,
                    "POWER": power,
                    "재고수량": np.nan,
                    "전일재고": np.nan,
                    "재고증감": np.nan,
                    "긴급요청": True,
                    "대상품목": target_text,
                }
            )
    if not rows:
        return empty_daily_inventory_df()
    return pd.DataFrame(rows, columns=DAILY_INVENTORY_COLUMNS)


def normalize_daily_support_inventory(xl: pd.ExcelFile) -> pd.DataFrame:
    sheet_name = "지원파트 재고표"
    if sheet_name not in xl.sheet_names:
        return empty_daily_inventory_df()

    raw = xl.parse(sheet_name=sheet_name, header=None)
    rows: list[dict[str, Any]] = []
    current_product = ""
    current_pack = ""

    for _, row in raw.iterrows():
        product_candidate = clean_str(row.iloc[1] if len(row) > 1 else "")
        pack_candidate = extract_daily_pack_label(product_candidate)
        if product_candidate and pack_candidate:
            current_product = product_candidate
            current_pack = pack_candidate
        if not current_product or not current_pack:
            continue

        for offset in range(18):
            power_col = 2 + offset
            current_col = 21 + offset
            previous_col = 39 + offset
            if power_col >= len(row):
                continue
            power = daily_power_label(row.iloc[power_col])
            if not power:
                continue
            current_stock = numeric_scalar(row.iloc[current_col] if current_col < len(row) else np.nan, np.nan)
            previous_stock = numeric_scalar(row.iloc[previous_col] if previous_col < len(row) else np.nan, np.nan)
            if pd.isna(current_stock) and pd.isna(previous_stock):
                continue
            rows.append(
                {
                    "제품명": current_product,
                    "제품코드": "",
                    "PACK": current_pack,
                    "POWER": power,
                    "재고수량": current_stock,
                    "전일재고": previous_stock,
                    "재고증감": current_stock - previous_stock if pd.notna(current_stock) and pd.notna(previous_stock) else np.nan,
                    "긴급요청": False,
                    "대상품목": "",
                }
            )

    if not rows:
        return empty_daily_inventory_df()
    return pd.DataFrame(rows, columns=DAILY_INVENTORY_COLUMNS)


def normalize_daily_inventory_file(path: Path | None) -> pd.DataFrame:
    if path is None:
        return empty_daily_inventory_df()
    try:
        xl = pd.ExcelFile(path)
    except Exception:
        return empty_daily_inventory_df()

    support = normalize_daily_support_inventory(xl)
    emergency = normalize_daily_emergency_requests(xl)
    if support.empty and emergency.empty:
        return empty_daily_inventory_df()

    support_work = support.copy()
    emergency_work = emergency.copy()
    support_work["_daily_key"] = [
        daily_inventory_key(product, pack, power)
        for product, pack, power in zip(support_work["제품명"], support_work["PACK"], support_work["POWER"])
    ]
    emergency_work["_daily_key"] = [
        daily_inventory_key(product, pack, power)
        for product, pack, power in zip(emergency_work["제품명"], emergency_work["PACK"], emergency_work["POWER"])
    ]

    merged = support_work.merge(
        emergency_work[["_daily_key", "제품명", "제품코드", "PACK", "POWER", "긴급요청", "대상품목"]].rename(
            columns={
                "제품명": "_emergency_product_name",
                "제품코드": "_emergency_product_code",
                "PACK": "_emergency_pack",
                "POWER": "_emergency_power",
                "긴급요청": "_emergency_flag",
                "대상품목": "_emergency_target",
            }
        ),
        on="_daily_key",
        how="outer",
    )

    for col in DAILY_INVENTORY_COLUMNS:
        if col not in merged.columns:
            merged[col] = np.nan if col in {"재고수량", "전일재고", "재고증감"} else ""

    for base_col, emergency_col in [
        ("제품명", "_emergency_product_name"),
        ("제품코드", "_emergency_product_code"),
        ("PACK", "_emergency_pack"),
        ("POWER", "_emergency_power"),
    ]:
        if emergency_col in merged.columns:
            merged[base_col] = merged[base_col].where(
                merged[base_col].map(clean_str) != "",
                merged[emergency_col],
            )
    if "_emergency_product_code" in merged.columns:
        merged["제품코드"] = merged["제품코드"].where(merged["제품코드"].map(clean_str) != "", merged["_emergency_product_code"])
    if "_emergency_flag" in merged.columns:
        base_flag = merged["긴급요청"].apply(lambda value: bool(value) if not pd.isna(value) else False)
        emergency_flag = merged["_emergency_flag"].apply(lambda value: bool(value) if not pd.isna(value) else False)
        merged["긴급요청"] = base_flag | emergency_flag
    if "_emergency_target" in merged.columns:
        merged["대상품목"] = merged["대상품목"].where(merged["대상품목"].map(clean_str) != "", merged["_emergency_target"])

    negative_stock = pd.to_numeric(merged["재고수량"], errors="coerce") < 0
    merged["긴급요청"] = merged["긴급요청"].apply(lambda value: bool(value) if not pd.isna(value) else False) | negative_stock
    merged.loc[negative_stock & (merged["대상품목"].map(clean_str) == ""), "대상품목"] = "재고표 음수 재고"

    merged["재고증감"] = pd.to_numeric(merged["재고증감"], errors="coerce")
    missing_delta = merged["재고증감"].isna()
    merged.loc[missing_delta, "재고증감"] = (
        pd.to_numeric(merged.loc[missing_delta, "재고수량"], errors="coerce")
        - pd.to_numeric(merged.loc[missing_delta, "전일재고"], errors="coerce")
    )
    return merged[DAILY_INVENTORY_COLUMNS].copy()


def clean_inventory_product_name(value: Any) -> str:
    text = clean_str(value)
    if not text:
        return ""
    text = re.sub(r"[/\\]+$", "", text).strip()
    text = re.sub(r"[/_ -]*[+-]?\d{1,2}\.\d{2}$", "", text).strip("_-/ ")
    return text or clean_str(value)


def inventory_power_label_from_sales_code(value: Any) -> str:
    power_value = parse_power_from_sales_code(value)
    if pd.isna(power_value):
        return ""
    return format_power(power_value)


def pack_label_from_inventory_name(product_name: Any, product_spec: Any = "") -> str:
    product_text = clean_str(product_name)
    inline_pack = re.search(
        r"(?:^|[_\s])(\d+(?:\.\d+)?)\s*(?:P|팩)(?=$|[_\s])",
        product_text,
        flags=re.IGNORECASE,
    )
    if inline_pack:
        return base_pack_label(float(inline_pack.group(1)))
    product_text = re.sub(r"[/\\]+$", "", product_text).strip()
    product_text = re.sub(r"[/_ -]*[+-]?\d{1,2}\.\d{2}$", "", product_text).strip("_-/ ")
    unit = extract_pack_unit(product_text)
    if pd.isna(unit):
        unit = extract_pack_unit(product_spec)
    return base_pack_label(unit) if pd.notna(unit) and float(unit) > 0 else ""


def build_daily_wms_catalog(inventory_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    exact_columns = ["제품코드", "POWER", "_wms_product_name", "_wms_pack", "_wms_stock"]
    prefix_columns = ["제품코드", "_wms_product_name", "_wms_pack"]
    if inventory_df is None or inventory_df.empty:
        return pd.DataFrame(columns=exact_columns), pd.DataFrame(columns=prefix_columns)

    work = inventory_df.copy()
    work["제품코드"] = work["sales_code"].map(extract_sales_prefix)
    work["POWER"] = work["sales_code"].map(inventory_power_label_from_sales_code)
    work["_wms_product_name"] = work["inventory_product_name"].map(clean_inventory_product_name)
    work["_wms_pack"] = [
        pack_label_from_inventory_name(product_name, product_spec)
        for product_name, product_spec in zip(
            work.get("inventory_product_name", pd.Series("", index=work.index)),
            work.get("inventory_product_spec", pd.Series("", index=work.index)),
        )
    ]
    work["_wms_stock"] = pd.to_numeric(work.get("available_stock_pack", 0.0), errors="coerce")
    work = work[work["제품코드"].map(clean_str) != ""].copy()
    if work.empty:
        return pd.DataFrame(columns=exact_columns), pd.DataFrame(columns=prefix_columns)

    exact = (
        work[work["POWER"].map(clean_str) != ""]
        .groupby(["제품코드", "POWER"], dropna=False)
        .agg(
            _wms_product_name=("_wms_product_name", first_nonempty),
            _wms_pack=("_wms_pack", first_nonempty),
            _wms_stock=("_wms_stock", "sum"),
        )
        .reset_index()
    )
    prefix = (
        work.groupby("제품코드", dropna=False)
        .agg(
            _wms_product_name=("_wms_product_name", first_nonempty),
            _wms_pack=("_wms_pack", first_nonempty),
        )
        .reset_index()
    )
    return exact[exact_columns].copy(), prefix[prefix_columns].copy()


def fill_daily_product_code_from_wms(out: pd.DataFrame, inventory_df: pd.DataFrame) -> pd.DataFrame:
    if out.empty or inventory_df is None or inventory_df.empty:
        return out

    work = inventory_df.copy()
    work["제품코드"] = work["sales_code"].map(extract_sales_prefix)
    work["POWER"] = work["sales_code"].map(inventory_power_label_from_sales_code)
    work["PACK"] = [
        pack_label_from_inventory_name(product_name, product_spec)
        for product_name, product_spec in zip(
            work.get("inventory_product_name", pd.Series("", index=work.index)),
            work.get("inventory_product_spec", pd.Series("", index=work.index)),
        )
    ]
    work["_wms_product_name"] = work["inventory_product_name"].map(clean_inventory_product_name)
    work = work[
        (work["제품코드"].map(clean_str) != "")
        & (work["POWER"].map(clean_str) != "")
        & (work["PACK"].map(clean_str) != "")
        & (work["_wms_product_name"].map(clean_str) != "")
    ].copy()
    if work.empty:
        return out

    catalog = (
        work.groupby(["제품코드", "PACK", "POWER"], dropna=False)
        .agg(_wms_product_name=("_wms_product_name", first_nonempty))
        .reset_index()
    )
    filled = out.copy()
    needs_code = filled["제품코드"].map(clean_str) == ""
    for idx, row in filled[needs_code].iterrows():
        pack = clean_str(row.get("PACK", ""))
        power = clean_str(row.get("POWER", ""))
        terms = expand_product_query_terms(row.get("제품명", ""))
        if not pack or not power or not terms:
            continue
        candidates = catalog[(catalog["PACK"] == pack) & (catalog["POWER"] == power)].copy()
        if candidates.empty:
            continue
        candidates = candidates[contains_any_query_term(candidates["_wms_product_name"], terms)]
        product_codes = [code for code in candidates["제품코드"].dropna().astype(str).unique().tolist() if clean_str(code)]
        if len(product_codes) != 1:
            continue
        filled.at[idx, "제품코드"] = product_codes[0]
        product_name = first_nonempty(candidates["_wms_product_name"])
        if product_name and clean_str(filled.at[idx, "제품명"]) == "":
            filled.at[idx, "제품명"] = product_name
    return filled


def enrich_daily_inventory_from_wms(daily_inventory_df: pd.DataFrame, inventory_df: pd.DataFrame) -> pd.DataFrame:
    if daily_inventory_df is None or daily_inventory_df.empty or inventory_df is None or inventory_df.empty:
        return daily_inventory_df

    out = daily_inventory_df.copy()
    out["제품코드"] = out["제품코드"].map(clean_str).str.upper()
    out["POWER"] = out["POWER"].map(clean_str)
    exact_catalog, prefix_catalog = build_daily_wms_catalog(inventory_df)

    if not exact_catalog.empty:
        out = out.merge(exact_catalog, on=["제품코드", "POWER"], how="left")
        out["제품명"] = out["제품명"].where(out["제품명"].map(clean_str) != "", out["_wms_product_name"])
        out["PACK"] = out["PACK"].where(out["PACK"].map(clean_str) != "", out["_wms_pack"])
        out["재고수량"] = pd.to_numeric(out["재고수량"], errors="coerce")
        out["재고수량"] = out["재고수량"].where(out["재고수량"].notna(), out["_wms_stock"])
        out = out.drop(columns=["_wms_product_name", "_wms_pack", "_wms_stock"], errors="ignore")

    if not prefix_catalog.empty:
        out = out.merge(prefix_catalog, on="제품코드", how="left")
        out["제품명"] = out["제품명"].where(out["제품명"].map(clean_str) != "", out["_wms_product_name"])
        out["PACK"] = out["PACK"].where(out["PACK"].map(clean_str) != "", out["_wms_pack"])
        out = out.drop(columns=["_wms_product_name", "_wms_pack"], errors="ignore")

    out = fill_daily_product_code_from_wms(out, inventory_df)
    return out[DAILY_INVENTORY_COLUMNS].copy()


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


def parse_datetime_series(series: pd.Series) -> pd.Series:
    text = series.astype(str).str.strip()
    korean_ampm = text.str.replace("오전", "AM", regex=False).str.replace("오후", "PM", regex=False)
    parsed = pd.to_datetime(korean_ampm, format="%Y-%m-%d %p %I:%M:%S", errors="coerce")
    missing = parsed.isna()
    if missing.any():
        parsed.loc[missing] = pd.to_datetime(series.loc[missing], errors="coerce")
    return parsed


def summarize_progress(progress_df: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
    if progress_df.empty:
        base_columns = group_cols + ["누수규격검사 생산수량"]
        return pd.DataFrame(columns=base_columns)

    agg_spec: dict[str, Any] = {
        "누수규격검사 생산수량": ("production_basis_qty", "sum"),
    }
    return progress_df.groupby(group_cols, dropna=False).agg(**agg_spec).reset_index()


def filter_progress_for_production_month(
    progress_df: pd.DataFrame,
    target_month: str = PRODUCTION_PROGRESS_DUE_MONTH,
) -> pd.DataFrame:
    if progress_df.empty:
        return progress_df.copy()

    total_due_source = (
        progress_df["total_due_date"]
        if "total_due_date" in progress_df.columns
        else pd.Series(pd.NaT, index=progress_df.index)
    )
    total_due = pd.to_datetime(
        total_due_source,
        errors="coerce",
    )
    inspection_step = next(step for step in PROCESS_STEPS if step["id"] == "80")
    inspection_due_col = str(inspection_step["due_col"])
    inspection_due_source = (
        progress_df[inspection_due_col]
        if inspection_due_col in progress_df.columns
        else pd.Series(pd.NaT, index=progress_df.index)
    )
    inspection_due = pd.to_datetime(
        inspection_due_source,
        errors="coerce",
    )
    production_due = total_due.fillna(inspection_due)
    target_period = pd.Period(target_month, freq="M")
    return progress_df.loc[production_due.dt.to_period("M") == target_period].copy()


def classify_status(packing_pack: float, packing_progress_pct: float) -> str:
    if packing_progress_pct >= 100.0:
        return "완료"
    if packing_pack > 0:
        return "진행중"
    return "미착수"


def finalize_summary(summary: pd.DataFrame) -> pd.DataFrame:
    out = summary.copy()
    if "용마입고 PACK" not in out.columns:
        out["용마입고 PACK"] = out["포장 PACK"] if "포장 PACK" in out.columns else 0.0
    out["포장부족수량"] = (out["요청 PACK"] - out["포장 PACK"]).clip(lower=0.0)
    out["미입고수량"] = (out["요청 PACK"] - out["용마입고 PACK"]).clip(lower=0.0)
    out["입고대기수량"] = (out["포장 PACK"] - out["용마입고 PACK"]).clip(lower=0.0)
    raw_progress = np.where(
        out["요청 PACK"] > 0,
        out["용마입고 PACK"] / out["요청 PACK"] * 100.0,
        0.0,
    )
    packing_progress = np.where(
        out["요청 PACK"] > 0,
        out["포장 PACK"] / out["요청 PACK"] * 100.0,
        0.0,
    )
    out["용마입고율"] = np.clip(raw_progress, 0.0, 100.0)
    out["포장진도율"] = np.clip(packing_progress, 0.0, 100.0)
    out["부족 PACK"] = out["미입고수량"]
    out["진도율(%)"] = out["용마입고율"]
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
    yongma_df: pd.DataFrame | None = None,
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
    if packing_df.empty:
        packing_by_key = pd.DataFrame(
            columns=["sales_code_key", "packing_sales_code", "packing_product_name", "packing_pack"]
        )
    else:
        packing_by_key = (
            packing_df.copy()
            .assign(sales_code_key=lambda df: df["sales_code_key"].map(clean_str))
            .groupby("sales_code_key", dropna=False)
            .agg(
                packing_sales_code=("sales_code", first_nonempty),
                packing_product_name=("packing_product_name", first_nonempty),
                packing_pack=("packing_pack", "sum"),
            )
            .reset_index()
        )
    if yongma_df is None or yongma_df.empty:
        if packing_df.empty or "sales_code_key" not in packing_df.columns:
            yongma_by_key = pd.DataFrame(
                columns=["sales_code_key", "yongma_sales_code", "yongma_product_name", "yongma_in_pack"]
            )
        else:
            yongma_by_key = (
                packing_df.groupby("sales_code_key", dropna=False)["packing_pack"]
                .sum()
                .reset_index()
                .rename(columns={"packing_pack": "yongma_in_pack"})
            )
            yongma_by_key["yongma_sales_code"] = ""
            yongma_by_key["yongma_product_name"] = ""
    else:
        yongma_by_key = (
            yongma_df.copy()
            .assign(sales_code_key=lambda df: df["sales_code_key"].map(clean_str))
            .groupby("sales_code_key", dropna=False)
            .agg(
                yongma_sales_code=("sales_code", first_nonempty),
                yongma_product_name=("yongma_product_name", first_nonempty),
                yongma_in_pack=("yongma_in_pack", "sum"),
            )
            .reset_index()
        )

    matched_code_summary = request_by_code.merge(
        packing_by_key[["sales_code_key", "packing_pack"]],
        on="sales_code_key",
        how="left",
    )
    matched_code_summary["packing_pack"] = matched_code_summary["packing_pack"].fillna(0.0)
    matched_code_summary = matched_code_summary.merge(
        yongma_by_key[["sales_code_key", "yongma_in_pack"]],
        on="sales_code_key",
        how="left",
    )
    matched_code_summary["yongma_in_pack"] = matched_code_summary["yongma_in_pack"].fillna(0.0)

    request_keys = set(request_by_code["sales_code_key"].map(clean_str)) - {""}
    supply_by_key = packing_by_key.merge(yongma_by_key, on="sales_code_key", how="outer")
    for col in ["packing_pack", "yongma_in_pack"]:
        if col not in supply_by_key.columns:
            supply_by_key[col] = 0.0
        supply_by_key[col] = pd.to_numeric(supply_by_key[col], errors="coerce").fillna(0.0)
    unmatched_supply = supply_by_key[
        (supply_by_key["sales_code_key"].map(clean_str) != "")
        & ~supply_by_key["sales_code_key"].map(clean_str).isin(request_keys)
        & ((supply_by_key["packing_pack"] > 0) | (supply_by_key["yongma_in_pack"] > 0))
    ].copy()
    unmatched_packing_total = float(unmatched_supply["packing_pack"].sum()) if not unmatched_supply.empty else 0.0

    if not unmatched_supply.empty:
        unmatched_rows: list[dict[str, Any]] = []
        for _, row in unmatched_supply.iterrows():
            sales_code = clean_str(row.get("packing_sales_code", "")) or clean_str(row.get("yongma_sales_code", ""))
            product_name = (
                clean_str(row.get("packing_product_name", ""))
                or clean_str(row.get("yongma_product_name", ""))
                or sales_code
            )
            pack_unit = extract_pack_unit(product_name)
            unmatched_rows.append(
                {
                    "sales_code": sales_code,
                    "product_name": product_name,
                    "product_name_code": product_name,
                    "production_code": "",
                    "p_code": "",
                    "q_code": "",
                    "r_code": "",
                    "pack_unit": pack_unit,
                    "pack_unit_label": format_pack_unit_label(pack_unit, product_name),
                    "base_product_name": strip_pack_unit_suffix(product_name),
                    "customer_name": "(포장실적)",
                    "sales_code_key": clean_str(row.get("sales_code_key", "")),
                    "product_name_key": normalize_match_key(product_name),
                    "product_name_code_key": normalize_match_key(product_name),
                    "production_code_key": "",
                    "p_code_key": "",
                    "q_code_key": "",
                    "r_code_key": "",
                    "request_pack": 0.0,
                    "request_pcs": 0.0,
                    "request_due_date": pd.NaT,
                    "packing_pack": float(row.get("packing_pack", 0.0)),
                    "yongma_in_pack": float(row.get("yongma_in_pack", 0.0)),
                }
            )
        matched_code_summary = pd.concat([matched_code_summary, pd.DataFrame(unmatched_rows)], ignore_index=True)

    product_summary = (
        matched_code_summary.groupby("base_product_name", dropna=False)[
            ["request_pack", "request_pcs", "packing_pack", "yongma_in_pack"]
        ]
        .sum()
        .reset_index()
        .rename(
            columns={
                "base_product_name": "제품명",
                "request_pack": "요청 PACK",
                "request_pcs": "요청 PCS",
                "packing_pack": "포장 PACK",
                "yongma_in_pack": "용마입고 PACK",
            }
        )
    )
    product_summary = finalize_summary(product_summary)
    product_summary["제품분류"] = product_summary["제품명"].map(classify_product_group)
    product_summary["본품분류"] = product_summary["제품명"].map(classify_main_product_family)
    return product_summary, unmatched_packing_total, matched_code_summary


def attach_inventory_to_code_summary(code_summary: pd.DataFrame, inventory_df: pd.DataFrame) -> pd.DataFrame:
    out = code_summary.copy()
    if inventory_df.empty:
        out["available_stock_pack"] = np.nan
        out["inventory_total_stock_pack"] = np.nan
        out["inventory_product_name"] = ""
        out["inventory_product_spec"] = ""
        out["inventory_updated_at"] = pd.NaT
        out["inventory_matched"] = False
        return out

    inventory_by_code = (
        inventory_df.groupby("sales_code_key", dropna=False)
        .agg(
            available_stock_pack=("available_stock_pack", sum_numeric_or_nan),
            inventory_total_stock_pack=("inventory_total_stock_pack", sum_numeric_or_nan),
            inventory_product_name=("inventory_product_name", join_unique),
            inventory_product_spec=("inventory_product_spec", first_nonempty),
            inventory_updated_at=("inventory_updated_at", max_datetime),
        )
        .reset_index()
    )
    out = out.merge(inventory_by_code, on="sales_code_key", how="left")
    out["inventory_matched"] = out["available_stock_pack"].notna()
    return out


def attach_inventory_to_product_summary(product_summary: pd.DataFrame, code_summary: pd.DataFrame) -> pd.DataFrame:
    out = product_summary.copy()
    if code_summary.empty or "available_stock_pack" not in code_summary.columns:
        out["용마창고재고 (PACK)"] = np.nan
        out["재고매칭SKU수"] = 0
        return out

    work = with_operational_columns(code_summary)
    stock_by_product = (
        work.groupby("base_product_name", dropna=False)
        .agg(
            current_stock_pack=("available_stock_pack", sum_numeric_or_nan),
            inventory_matched_count=("inventory_matched", "sum"),
        )
        .reset_index()
        .rename(
            columns={
                "base_product_name": "제품명",
                "current_stock_pack": "용마창고재고 (PACK)",
                "inventory_matched_count": "재고매칭SKU수",
            }
        )
    )
    out = out.merge(stock_by_product, on="제품명", how="left")
    out["재고매칭SKU수"] = out["재고매칭SKU수"].fillna(0).astype(int)
    return out


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


def attach_sample_available_to_code_summary(code_summary: pd.DataFrame, sample_available_df: pd.DataFrame) -> pd.DataFrame:
    out = code_summary.copy()
    if "production_code_key" not in out.columns:
        out["production_code_key"] = ""
    out["sample_available_pcs"] = 0.0
    if sample_available_df.empty:
        return out

    sample_by_code = (
        sample_available_df.groupby("production_code_key", dropna=False)["sample_available_pcs"]
        .sum()
        .reset_index()
        .rename(columns={"sample_available_pcs": "_sample_available_pcs"})
    )
    out = out.merge(sample_by_code, on="production_code_key", how="left")
    out["sample_available_pcs"] = pd.to_numeric(out["_sample_available_pcs"], errors="coerce").fillna(0.0)
    return out.drop(columns=["_sample_available_pcs"])


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
    return f"-{abs(float(num)):05.2f}"


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
        return "O2O2 D Toric"
    if "IRIS TORIC" in upper:
        return "Iris Toric"
    if upper.startswith("T38") or "T38" in upper or "사축" in text or "정축" in text:
        return "T38 Toric"
    if "TORIC" in upper:
        return "기타 Toric"
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
        color_continuous_scale=[(0.0, BG_PAGE), (0.5, "#D9E7F3"), (1.0, COLOR_BLUE)],
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
    yongma_in_pack = (
        float(df["용마입고 PACK"].sum()) if "용마입고 PACK" in df.columns and not df.empty else packing_pack
    )
    shortage_pack = float(df["미입고수량"].sum()) if "미입고수량" in df.columns and not df.empty else max(0.0, request_pack - yongma_in_pack)
    progress = (yongma_in_pack / request_pack * 100.0) if request_pack > 0 else 0.0
    packing_progress = (packing_pack / request_pack * 100.0) if request_pack > 0 else 0.0
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
        "request_pcs": request_pcs,
        "packing_pack": packing_pack,
        "yongma_in_pack": yongma_in_pack,
        "shortage_pack": shortage_pack,
        "production_shortage_pcs": production_shortage_qty,
        "progress_pct": min(100.0, max(0.0, progress)),
        "packing_progress_pct": min(100.0, max(0.0, packing_progress)),
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
        work["_allocated_sample_available_pcs"] = 0.0
        return work

    if "production_basis_qty" not in work.columns:
        work["production_basis_qty"] = 0.0
    if "sample_available_pcs" not in work.columns:
        work["sample_available_pcs"] = 0.0
    work["production_basis_qty"] = pd.to_numeric(work["production_basis_qty"], errors="coerce").fillna(0.0)
    work["sample_available_pcs"] = pd.to_numeric(work["sample_available_pcs"], errors="coerce").fillna(0.0)

    production_key = work.get("production_code_key", pd.Series("", index=work.index)).map(clean_str)
    sales_key = work.get("sales_code_key", pd.Series("", index=work.index)).map(clean_str)
    fallback_key = work.get("sales_code", pd.Series("", index=work.index)).map(clean_str)
    work["_production_alloc_key"] = production_key.where(production_key != "", sales_key)
    work["_production_alloc_key"] = work["_production_alloc_key"].where(work["_production_alloc_key"] != "", fallback_key)

    work["_allocated_production_shortage_qty"] = work["production_basis_qty"].clip(lower=0.0).round(0).astype("int64")
    work["_allocated_sample_available_pcs"] = work["sample_available_pcs"].clip(lower=0.0).round(0).astype("int64")
    return work


def calc_kpi_from_code_summary(code_summary: pd.DataFrame) -> dict[str, float]:
    if code_summary.empty:
        return {
            "request_pack": 0.0,
            "request_pcs": 0.0,
            "packing_pack": 0.0,
            "yongma_in_pack": 0.0,
            "shortage_pack": 0.0,
            "production_shortage_pcs": 0.0,
            "packable_pcs": 0.0,
            "progress_pct": 0.0,
            "packing_progress_pct": 0.0,
            "production_progress_pct": 0.0,
        }

    request_pack = float(code_summary["request_pack"].sum())
    packing_pack = float(code_summary["packing_pack"].sum())
    yongma_in_pack = (
        float(code_summary["yongma_in_pack"].sum()) if "yongma_in_pack" in code_summary.columns else packing_pack
    )
    shortage_pack = max(0.0, request_pack - yongma_in_pack)
    receipt_progress = (yongma_in_pack / request_pack * 100.0) if request_pack > 0 else 0.0
    packing_progress = (packing_pack / request_pack * 100.0) if request_pack > 0 else 0.0

    work = (
        code_summary.copy()
        if "_allocated_production_shortage_qty" in code_summary.columns
        else add_allocated_production_basis(code_summary)
    )
    request_pcs = float(work["request_pcs"].sum())
    shortage_pcs = float(work["_allocated_production_shortage_qty"].sum())
    packable_pcs = max(0.0, request_pcs - shortage_pcs)
    production_progress = ((request_pcs - shortage_pcs) / request_pcs * 100.0) if request_pcs > 0 else 0.0

    return {
        "request_pack": request_pack,
        "request_pcs": request_pcs,
        "packing_pack": packing_pack,
        "yongma_in_pack": yongma_in_pack,
        "shortage_pack": shortage_pack,
        "production_shortage_pcs": shortage_pcs,
        "packable_pcs": packable_pcs,
        "progress_pct": min(100.0, max(0.0, receipt_progress)),
        "packing_progress_pct": min(100.0, max(0.0, packing_progress)),
        "production_progress_pct": min(100.0, max(0.0, production_progress)),
    }


def format_int(value: float) -> str:
    return f"{value:,.0f}"


def progress_tone(progress: float) -> str:
    if progress >= 100:
        return "done"
    if progress >= 80:
        return "active"
    if progress >= 50:
        return "warn"
    return "risk"


def status_class(status: str) -> str:
    if status in {"완료", "입고완료"}:
        return "done"
    if status in {"진행중"}:
        return "active"
    if status in {"부족"}:
        return "warn"
    if status in {"입고대기"}:
        return "waiting"
    return "risk"


def progress_cell_html(progress: float, label: str = "") -> str:
    width = max(0.0, min(100.0, float(progress)))
    tone = progress_tone(float(progress))
    semantic = ""
    if label in {"생산"}:
        semantic = " production"
    elif label in {"입고", "용마입고"}:
        semantic = " receipt"
    prefix = f"<span class='progress-name'>{escape(label)}</span>" if label else ""
    return (
        "<div class='progress-cell'>"
        f"{prefix}"
        "<div class='progress-track'>"
        f"<div class='progress-fill {tone}{semantic}' style='width:{width:.1f}%'></div>"
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

    receipt_shortage_col = "미입고수량" if "미입고수량" in df.columns else "포장부족수량"
    receipt_progress_col = "용마입고율" if "용마입고율" in df.columns else "포장진도율"

    ordered = df.sort_values(
        [receipt_shortage_col, "생산부족수량", "요청 PACK"],
        ascending=[False, False, False],
        kind="stable",
    ).head(max_rows).copy()

    rows: list[str] = []
    for _, row in ordered.iterrows():
        product = escape(str(row["제품명"]))
        family = escape(str(row.get("본품분류", ""))) if show_family else ""
        req = format_int(float(row["요청 PACK"]))
        receipt_shortage = float(row.get(receipt_shortage_col, 0.0))
        receipt_shortage_txt = format_int(receipt_shortage)
        receipt_progress = float(row.get(receipt_progress_col, 0.0))
        production_shortage = float(row.get("생산부족수량", 0.0))
        production_shortage_txt = format_int(production_shortage)
        prod_progress = float(row.get("생산진도율", 0.0))
        status = escape(str(row["상태"]))
        badge = f"<span class='status-badge {status_class(str(row['상태']))}'>{status}</span>"

        receipt_shortage_class = "num shortage" if receipt_shortage > 0 else "num"
        production_shortage_class = "num shortage" if production_shortage > 0 else "num"
        receipt_progress_html = progress_cell_html(receipt_progress, "입고")
        production_progress_html = progress_cell_html(prod_progress, "생산")

        if compact:
            rows.append(
                "<tr>"
                f"<td class='left'>{product}</td>"
                f"<td>{production_progress_html}</td>"
                f"<td>{receipt_progress_html}</td>"
                f"<td class='{receipt_shortage_class}'>{receipt_shortage_txt}</td>"
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
                f"<td>{receipt_progress_html}</td>"
                f"<td class='{production_shortage_class}'>{production_shortage_txt}</td>"
                f"<td class='{receipt_shortage_class}'>{receipt_shortage_txt}</td>"
                f"<td>{badge}</td>"
                "</tr>"
            )

    header = (
        "<tr>"
        "<th class='left'>제품명</th>"
        "<th>생산진도율</th>"
        "<th>용마입고율</th>"
        "<th class='num'>미입고수량</th>"
        "<th>상태</th>"
        "</tr>"
        if compact
        else "<tr>"
        "<th class='left'>제품명</th>"
        f"{'<th>본품분류</th>' if show_family else ''}"
        "<th class='num'>요청 PACK</th>"
        "<th>생산진도율</th>"
        "<th>용마입고율</th>"
        "<th class='num'>생산부족수량</th>"
        "<th class='num'>미입고수량</th>"
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
                "용마입고 PACK",
                "생산부족수량",
                "포장부족수량",
                "미입고수량",
                "생산진도율",
                "포장진도율",
                "용마입고율",
            ]
        )

    grouped = (
        product_df.groupby("본품분류", dropna=False)
        .agg(
            request_pack=("요청 PACK", "sum"),
            request_pcs=("요청 PCS", "sum"),
            packing_pack=("포장 PACK", "sum"),
            yongma_in_pack=("용마입고 PACK", "sum"),
            production_shortage_qty=("생산부족수량", "sum"),
            packing_shortage_qty=("포장부족수량", "sum"),
            receipt_shortage_qty=("미입고수량", "sum"),
        )
        .reset_index()
        .rename(
            columns={
                "request_pack": "요청 PACK",
                "request_pcs": "요청 PCS",
                "packing_pack": "포장 PACK",
                "yongma_in_pack": "용마입고 PACK",
                "production_shortage_qty": "생산부족수량",
                "packing_shortage_qty": "포장부족수량",
                "receipt_shortage_qty": "미입고수량",
            }
        )
    )
    grouped["생산진도율"] = calc_production_progress_pct(grouped["요청 PCS"], grouped["생산부족수량"])
    grouped["용마입고율"] = np.where(
        grouped["요청 PACK"] > 0,
        grouped["용마입고 PACK"] / grouped["요청 PACK"] * 100.0,
        0.0,
    )
    grouped["용마입고율"] = np.clip(grouped["용마입고율"], 0.0, 100.0)
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


def family_card_section(family: Any) -> str:
    text = clean_str(family)
    upper = text.upper()
    if text in FAMILY_CARD_MISC_NAMES:
        return "기타"
    if text in FAMILY_CARD_1DAY_NAMES:
        return "1DAY"
    if "1DAY" in upper or "1 DAY" in upper:
        return "1DAY"
    if text in {"부자재/기타", "기타", "샘플"}:
        return "기타"
    return "FRP"


def family_card_html(row: pd.Series) -> str:
    family = escape(str(row["본품분류"]))
    request_pack = format_int(float(row["요청 PACK"]))
    production_progress = float(row["생산진도율"])
    receipt_progress = float(row.get("용마입고율", row["포장진도율"]))
    production_shortage = format_int(float(row["생산부족수량"]))
    return (
        "<div class='family-card'>"
        f"<div class='family-head'><span>{family}</span><b>요청 {request_pack} PACK</b></div>"
        f"{progress_cell_html(production_progress, '생산')}"
        f"{progress_cell_html(receipt_progress, '입고')}"
        "<div class='family-shortages'>"
        f"<span>생산부족 PCS <b>{production_shortage}</b></span>"
        "</div>"
        "</div>"
    )


def render_family_progress_cards(family_df: pd.DataFrame, max_rows: int = 14) -> None:
    if family_df.empty:
        st.warning("본품 분류별 진도현황을 표시할 데이터가 없습니다.")
        return

    view = family_df.head(max_rows).copy()
    view["_section"] = view["본품분류"].map(family_card_section)
    sections: list[str] = []
    for section in FAMILY_CARD_SECTION_ORDER:
        scoped = view[view["_section"] == section]
        if scoped.empty:
            continue
        cards = "".join(family_card_html(row) for _, row in scoped.iterrows())
        sections.append(
            "<section class='family-section'>"
            f"<div class='family-section-title'>{escape(section)}</div>"
            f"<div class='family-grid'>{cards}</div>"
            "</section>"
        )
    st.markdown("".join(sections), unsafe_allow_html=True)


def build_top_shortage_view(product_df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    columns = ["순위", "제품명", "미입고 PACK", "생산진도율", "포장진도율", "용마입고율"]
    if product_df.empty:
        return pd.DataFrame(columns=columns)

    view = product_df.copy()
    for col in ["미입고수량", "생산진도율", "포장진도율", "용마입고율"]:
        if col not in view.columns:
            view[col] = 0.0
        view[col] = pd.to_numeric(view[col], errors="coerce").fillna(0.0)
    view = view[view["미입고수량"] > 0].sort_values("미입고수량", ascending=False, kind="stable").head(top_n).copy()
    if view.empty:
        return pd.DataFrame(columns=columns)
    view["미입고 PACK"] = view["미입고수량"]
    view["순위"] = range(1, len(view) + 1)
    return view[columns].copy()


def render_top_shortage_list(top_df: pd.DataFrame) -> None:
    if top_df.empty:
        st.warning("미입고 제품이 없습니다.")
        return
    rows: list[str] = []
    for _, row in top_df.iterrows():
        rows.append(
            "<tr>"
            f"<td class='num muted'>{format_int(float(row.get('순위', 0.0)))}</td>"
            f"<td class='left'>{escape(str(row.get('제품명', '')))}</td>"
            f"<td class='num shortage'>{format_int(float(row.get('미입고 PACK', 0.0)))}</td>"
            f"<td>{progress_cell_html(float(row.get('생산진도율', 0.0)))}</td>"
            f"<td>{progress_cell_html(float(row.get('포장진도율', 0.0)))}</td>"
            f"<td>{progress_cell_html(float(row.get('용마입고율', 0.0)))}</td>"
            "</tr>"
        )
    st.markdown(
        "<div class='table-wrap compact-table'>"
        "<table class='ops-table progress-summary-table'>"
        "<thead><tr>"
        "<th class='num'>순위</th>"
        "<th class='left'>제품명</th>"
        "<th class='num'>미입고 PACK</th>"
        "<th>생산진도율</th>"
        "<th>포장진도율</th>"
        "<th>용마입고율</th>"
        "</tr></thead>"
        f"<tbody>{''.join(rows)}</tbody>"
        "</table></div>",
        unsafe_allow_html=True,
    )


def build_gap_top_view(product_df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    columns = ["순위", "제품명", "생산진도율", "용마입고율", "GAP"]
    if product_df.empty:
        return pd.DataFrame(columns=columns)

    source = product_df.copy()
    if "생산진도율" not in source.columns:
        source["생산진도율"] = 0.0
    if "용마입고율" not in source.columns:
        source["용마입고율"] = source.get("포장진도율", 0.0)
    if "미입고수량" not in source.columns:
        source["미입고수량"] = 0.0
    if "요청 PACK" not in source.columns:
        source["요청 PACK"] = 0.0
    source["생산진도율"] = pd.to_numeric(source["생산진도율"], errors="coerce").fillna(0.0)
    source["용마입고율"] = pd.to_numeric(source["용마입고율"], errors="coerce").fillna(0.0)
    source["GAP"] = source["생산진도율"] - source["용마입고율"]
    source = source[source["GAP"] > 0].copy()
    if source.empty:
        return pd.DataFrame(columns=columns)

    out = (
        source.sort_values(
            ["GAP", "미입고수량", "요청 PACK"],
            ascending=[False, False, False],
            kind="stable",
        )
        .head(top_n)
        .copy()
    )
    out["순위"] = range(1, len(out) + 1)
    return out[columns].copy()


def render_gap_top_list(gap_df: pd.DataFrame) -> None:
    if gap_df.empty:
        st.warning("생산진도율이 용마입고율보다 높은 GAP 제품이 없습니다.")
        return
    rows: list[str] = []
    for _, row in gap_df.iterrows():
        rows.append(
            "<tr>"
            f"<td class='num muted'>{format_int(float(row.get('순위', 0.0)))}</td>"
            f"<td class='left'>{escape(str(row.get('제품명', '')))}</td>"
            f"<td>{progress_cell_html(float(row.get('생산진도율', 0.0)))}</td>"
            f"<td>{progress_cell_html(float(row.get('용마입고율', 0.0)))}</td>"
            f"<td class='num shortage'>+{float(row.get('GAP', 0.0)):.1f}</td>"
            "</tr>"
        )
    st.markdown(
        "<div class='table-wrap compact-table'>"
        "<table class='ops-table progress-summary-table'>"
        "<thead><tr>"
        "<th class='num'>순위</th>"
        "<th class='left'>제품명</th>"
        "<th>생산진도율</th>"
        "<th>용마입고율</th>"
        "<th class='num'>GAP</th>"
        "</tr></thead>"
        f"<tbody>{''.join(rows)}</tbody>"
        "</table></div>",
        unsafe_allow_html=True,
    )


def render_kpi_panel(title: str, kpi: dict[str, float], unit_mode: str = UNIT_PACK) -> None:
    progress = float(kpi["progress_pct"])
    production_progress = float(kpi.get("production_progress_pct", 0.0))
    packing_progress = float(kpi.get("packing_progress_pct", 0.0))
    shortage_class = "metric-value warn" if kpi["shortage_pack"] > 0 else "metric-value"
    production_shortage_class = "metric-value warn" if kpi.get("production_shortage_pcs", 0.0) > 0 else "metric-value"

    if unit_mode == UNIT_PCS:
        panel_html = f"""
        <div class='kpi-panel scope-kpi'>
          <div class='kpi-title'>{escape(title)}</div>
          <div class='kpi-grid'>
            <div class='kpi-card'>
              <div class='metric-label'>요청 PCS</div>
              <div class='metric-value'>{format_int(kpi.get('request_pcs', 0.0))}</div>
            </div>
            <div class='kpi-card'>
              <div class='metric-label'>생산부족 PCS</div>
              <div class='{production_shortage_class}'>{format_int(kpi.get('production_shortage_pcs', 0.0))}</div>
            </div>
            <div class='kpi-card'>
              <div class='metric-label'>생산진도율</div>
              <div class='metric-value'>{production_progress:.1f}%</div>
            </div>
            <div class='kpi-card'>
              <div class='metric-label'>용마입고 PACK</div>
              <div class='metric-value'>{format_int(kpi.get('yongma_in_pack', kpi['packing_pack']))}</div>
            </div>
            <div class='kpi-card'>
              <div class='metric-label'>용마입고율</div>
              <div class='metric-value'>{progress:.1f}%</div>
            </div>
          </div>
        </div>
        """
        st.markdown(panel_html, unsafe_allow_html=True)
        return

    panel_html = f"""
    <div class='kpi-panel scope-kpi'>
      <div class='kpi-title'>{escape(title)}</div>
      <div class='kpi-grid'>
        <div class='kpi-card'>
          <div class='metric-label'>요청 PACK</div>
          <div class='metric-value'>{format_int(kpi['request_pack'])}</div>
        </div>
        <div class='kpi-card'>
          <div class='metric-label'>생산진도율</div>
          <div class='metric-value'>{production_progress:.1f}%</div>
        </div>
        <div class='kpi-card'>
          <div class='metric-label'>포장진도율</div>
          <div class='metric-value'>{packing_progress:.1f}%</div>
        </div>
        <div class='kpi-card'>
          <div class='metric-label'>용마입고율</div>
          <div class='metric-value'>{progress:.1f}%</div>
        </div>
        <div class='kpi-card'>
          <div class='metric-label'>미입고 PACK</div>
          <div class='{shortage_class}'>{format_int(kpi['shortage_pack'])}</div>
        </div>
        <div class='kpi-card'>
          <div class='metric-label'>생산부족 PCS</div>
          <div class='{production_shortage_class}'>{format_int(kpi.get('production_shortage_pcs', 0.0))}</div>
        </div>
      </div>
    </div>
    """
    st.markdown(panel_html, unsafe_allow_html=True)


def render_kpi_scope_panels(
    code_summary: pd.DataFrame,
    product_names: pd.Series | None = None,
    unit_mode: str = UNIT_PACK,
) -> None:
    work = add_allocated_production_basis(code_summary)
    if product_names is not None:
        work = code_summary_for_products(work, product_names)

    scope_kpis = [
        (f"{name} KPI", kpi)
        for name, kpi in build_scope_kpis(work)
        if name in {"본품", "샘플"}
    ]
    kpi_cols = st.columns(len(scope_kpis), gap="small")
    for col, (title, kpi) in zip(kpi_cols, scope_kpis):
        with col:
            render_kpi_panel(title, kpi, unit_mode=unit_mode)


def render_product_scope_kpi_panels(product_summary: pd.DataFrame, unit_mode: str = UNIT_PACK) -> None:
    main_products, sample_products = split_main_sample(product_summary)
    scopes = [
        ("본품 KPI", main_products),
        ("샘플 KPI", sample_products),
    ]
    kpi_cols = st.columns(2, gap="small")
    for col, (title, scope_df) in zip(kpi_cols, scopes):
        with col:
            render_kpi_panel(title, calc_kpi(scope_df), unit_mode=unit_mode)


def metric_progress_tone(progress: float) -> str:
    if progress >= 80:
        return "good"
    if progress >= 50:
        return "mid"
    if progress > 0:
        return "warn"
    return "muted"


def render_metric_card_grid(items: list[tuple[str, str, str] | tuple[str, str, str, str]]) -> None:
    def card_html(item: tuple[str, str, str] | tuple[str, str, str, str]) -> str:
        label, value, tone = item[:3]
        note = item[3] if len(item) > 3 else ""
        note_html = f"<div class='metric-note'>{escape(note)}</div>" if note else ""
        return (
            "<div class='mini-kpi-card'>"
            f"<div class='metric-label'>{escape(label)}</div>"
            f"<div class='metric-value {tone}'>{escape(value)}</div>"
            f"{note_html}"
            "</div>"
        )

    cards = "".join(
        card_html(item)
        for item in items
    )
    st.markdown(f"<div class='mini-kpi-grid'>{cards}</div>", unsafe_allow_html=True)


def render_status_board(
    product_summary: pd.DataFrame,
    code_summary: pd.DataFrame,
    daily_inventory_df: pd.DataFrame | None,
    sample_available_df: pd.DataFrame | None,
    stock_threshold_pack: float,
) -> None:
    kpi = calc_operation_kpis(product_summary, code_summary, stock_threshold_pack)
    exception_kpis, _exception_detail = build_daily_exception_report_view(
        daily_inventory_df,
        code_summary,
        sample_available_df,
    )
    request_pack = float(kpi.get("request_pack", 0.0))
    yongma_in_pack = float(kpi.get("yongma_in_pack", 0.0))
    missing_pack = float(kpi.get("packing_shortage_pack", 0.0))
    production_shortage = float(kpi.get("production_shortage_pcs", 0.0))
    packing_progress = float(kpi.get("packing_progress_pct", 0.0))
    receipt_progress = float(kpi.get("receipt_progress_pct", kpi.get("packing_progress_pct", 0.0)))
    production_progress = float(kpi.get("production_progress_pct", 0.0))
    packing_todo_pack = float(kpi.get("packing_todo_pack", 0.0))
    receipt_wait_pack = float(kpi.get("receipt_wait_pack", 0.0))
    priority_products = int(kpi.get("priority_products", 0.0))
    request_out_count = int(exception_kpis.get("request_out_count", 0.0))

    # Guard the headline KPI against stale Streamlit/session state: the status
    # board must always use the current code-level receipt aggregate when present.
    direct_request_pack = (
        float(pd.to_numeric(code_summary.get("request_pack", pd.Series(dtype=float)), errors="coerce").fillna(0.0).sum())
        if code_summary is not None and not code_summary.empty
        else 0.0
    )
    direct_packing_pack = (
        float(pd.to_numeric(code_summary.get("packing_pack", pd.Series(dtype=float)), errors="coerce").fillna(0.0).sum())
        if code_summary is not None and not code_summary.empty
        else 0.0
    )
    direct_yongma_in_pack = (
        float(pd.to_numeric(code_summary.get("yongma_in_pack", pd.Series(dtype=float)), errors="coerce").fillna(0.0).sum())
        if code_summary is not None and not code_summary.empty
        else 0.0
    )
    if direct_request_pack > 0:
        request_pack = direct_request_pack
    if direct_packing_pack > 0:
        packing_progress = direct_packing_pack / request_pack * 100.0 if request_pack > 0 else 0.0
    if direct_yongma_in_pack > 0:
        yongma_in_pack = direct_yongma_in_pack
        missing_pack = max(0.0, request_pack - yongma_in_pack)
        receipt_progress = yongma_in_pack / request_pack * 100.0 if request_pack > 0 else 0.0

    receipt_width = max(0.0, min(100.0, receipt_progress))
    missing_width = max(0.0, min(100.0 - receipt_width, 100.0))
    emergency_count = request_out_count
    if emergency_count > 0 or priority_products > 0:
        board_tone = "risk"
    elif missing_pack > 0 or production_shortage > 0:
        board_tone = "warn"
    else:
        board_tone = "good"

    cards = [
        ("국내 요청량 PACK", format_int(request_pack), "normal"),
        ("생산진도율", f"{production_progress:.1f}%", metric_progress_tone(production_progress)),
        ("포장진도율", f"{packing_progress:.1f}%", metric_progress_tone(packing_progress)),
        ("용마입고율", f"{receipt_progress:.1f}%", metric_progress_tone(receipt_progress)),
        ("미입고 PACK", format_int(missing_pack), "warn" if missing_pack > 0 else "normal"),
        ("긴급 대응 품목 수", f"{emergency_count:,}", "warn" if emergency_count > 0 else "normal"),
    ]
    card_html = "".join(
        "<div class='status-tile'>"
        f"<div class='metric-label'>{escape(label)}</div>"
        f"<div class='metric-value {tone}'>{escape(value)}</div>"
        "</div>"
        for label, value, tone in cards
    )

    board_html = f"""
    <div class='status-board {board_tone}'>
      <div class='status-main'>
        <div class='status-head'>
          <span class='status-pill {board_tone}'>용마입고율</span>
          <strong>요청 대비 용마 입고 현황</strong>
        </div>
        <div class='status-main-value'>{receipt_progress:.1f}%</div>
        <div class='status-flow'>
          <div class='status-flow-fill receipt' style='width:{receipt_width:.1f}%'></div>
          <div class='status-flow-fill shortage' style='width:{missing_width:.1f}%'></div>
        </div>
        <div class='status-flow-legend'>
          <span>요청 {format_int(request_pack)} PACK</span>
          <span>용마입고 {format_int(yongma_in_pack)} PACK</span>
          <span>미입고 {format_int(missing_pack)} PACK</span>
          <span>용마입고율 {receipt_progress:.1f}%</span>
        </div>
      </div>
      <div class='status-tile-grid'>
        {card_html}
      </div>
    </div>
    """
    st.markdown(board_html, unsafe_allow_html=True)


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


def pack_pcs_label(label: Any) -> str:
    return f"{clean_str(label)}(PCS)"


def with_operational_columns(code_summary: pd.DataFrame) -> pd.DataFrame:
    work = code_summary.copy()
    if "base_product_name" not in work.columns:
        work["base_product_name"] = work["product_name"].map(strip_pack_unit_suffix)
    if "pack_unit" not in work.columns:
        work["pack_unit"] = work["product_name"].map(extract_pack_unit)
    work["_pack_label"] = work["pack_unit"].map(base_pack_label)
    work["_pack_sort"] = work["_pack_label"].map(pack_sort_rank)
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


def d_day_number(value: Any) -> float:
    due = pd.to_datetime(value, errors="coerce")
    if pd.isna(due):
        return np.nan
    today = pd.Timestamp.now(tz="Asia/Seoul").tz_localize(None).normalize()
    return float((due.normalize() - today).days)


def d_day_text(value: Any) -> str:
    days = d_day_number(value)
    if pd.isna(days):
        return "-"
    day_count = int(days)
    if day_count < 0:
        return f"D+{abs(day_count)}"
    if day_count == 0:
        return "D-Day"
    return f"D-{day_count}"


def priority_grade(shortage_pack: Any, due_date: Any, current_stock_pack: Any, stock_threshold_pack: float) -> str:
    shortage = float(pd.to_numeric(shortage_pack, errors="coerce") or 0.0)
    if shortage <= 0:
        return "완료"

    days = d_day_number(due_date)
    stock = pd.to_numeric(current_stock_pack, errors="coerce")
    stock_low = pd.notna(stock) and float(stock) <= float(stock_threshold_pack)
    due_over = pd.notna(days) and days <= 0
    due_very_close = pd.notna(days) and days <= 3
    due_close = pd.notna(days) and days <= 7

    if due_over or (due_very_close and stock_low):
        return "A 긴급"
    if due_close or stock_low:
        return "B 주의"
    return "C 일반"


def priority_sort_value(value: Any) -> int:
    order = {"A 긴급": 0, "B 주의": 1, "C 일반": 2, "완료": 3}
    return order.get(str(value), 9)


def add_priority_columns(
    df: pd.DataFrame,
    stock_threshold_pack: float,
    shortage_col: str = "포장부족수량",
    due_col: str = "request_due_date",
    stock_col: str = "용마창고재고 (PACK)",
    request_col: str | None = None,
) -> pd.DataFrame:
    out = df.copy()
    if due_col not in out.columns:
        out[due_col] = pd.NaT
    if stock_col not in out.columns:
        out[stock_col] = np.nan
    stock = pd.to_numeric(out[stock_col], errors="coerce")
    out["재고기준(PACK)"] = float(stock_threshold_pack)
    if request_col and request_col in out.columns:
        request = pd.to_numeric(out[request_col], errors="coerce").fillna(0.0)
        out["재고부족(PACK)"] = np.where(stock.notna(), (request - stock).clip(lower=0.0), np.nan)
    else:
        out["재고부족(PACK)"] = np.where(stock.notna(), (float(stock_threshold_pack) - stock).clip(lower=0.0), np.nan)
    out["D-Day"] = out[due_col].map(d_day_text)
    out["우선등급"] = [
        priority_grade(shortage, due, current_stock, stock_threshold_pack)
        for shortage, due, current_stock in zip(out[shortage_col], out[due_col], out[stock_col])
    ]
    out["_priority_sort"] = out["우선등급"].map(priority_sort_value)
    out["_request_due_date_sort"] = pd.to_datetime(out[due_col], errors="coerce")
    return out


def calc_operation_kpis(
    product_summary: pd.DataFrame,
    code_summary: pd.DataFrame,
    stock_threshold_pack: float,
) -> dict[str, float]:
    today = pd.Timestamp.now(tz="Asia/Seoul").tz_localize(None).normalize()
    due_limit = today + pd.Timedelta(days=7)
    work = add_allocated_production_basis(with_operational_columns(code_summary))
    work["request_due_date"] = pd.to_datetime(work["request_due_date"], errors="coerce")
    if "available_stock_pack" not in work.columns:
        work["available_stock_pack"] = np.nan
    if "yongma_in_pack" not in work.columns:
        work["yongma_in_pack"] = work["packing_pack"]
    work["_packing_shortage_pack"] = (work["request_pack"] - work["packing_pack"]).clip(lower=0.0)
    work["_receipt_shortage_pack"] = (work["request_pack"] - work["yongma_in_pack"]).clip(lower=0.0)
    product_priority = (
        work.groupby("base_product_name", dropna=False)
        .agg(
            packing_shortage_pack=("_receipt_shortage_pack", "sum"),
            request_due_date=("request_due_date", min_datetime),
            current_stock_pack=("available_stock_pack", sum_numeric_or_nan),
        )
        .reset_index()
    )
    product_priority["우선등급"] = [
        priority_grade(shortage, due, stock, stock_threshold_pack)
        for shortage, due, stock in zip(
            product_priority["packing_shortage_pack"],
            product_priority["request_due_date"],
            product_priority["current_stock_pack"],
        )
    ]
    shortage_mask = product_priority["packing_shortage_pack"] > 0
    due_mask = product_priority["request_due_date"].notna() & (product_priority["request_due_date"] <= due_limit)
    stock = pd.to_numeric(product_priority["current_stock_pack"], errors="coerce")
    stock_shortage_mask = stock.notna() & (stock <= float(stock_threshold_pack))
    priority_mask = product_priority["우선등급"].isin(["A 긴급", "B 주의"])
    code_request_pack = float(pd.to_numeric(work.get("request_pack", pd.Series(dtype=float)), errors="coerce").fillna(0.0).sum())
    code_request_pcs = float(pd.to_numeric(work.get("request_pcs", pd.Series(dtype=float)), errors="coerce").fillna(0.0).sum())
    code_packing_pack = float(pd.to_numeric(work.get("packing_pack", pd.Series(dtype=float)), errors="coerce").fillna(0.0).sum())
    request_pack = code_request_pack or (
        float(product_summary["요청 PACK"].sum()) if "요청 PACK" in product_summary.columns and not product_summary.empty else 0.0
    )
    request_pcs = code_request_pcs or (
        float(product_summary["요청 PCS"].sum()) if "요청 PCS" in product_summary.columns and not product_summary.empty else 0.0
    )
    packing_pack = code_packing_pack or (
        float(product_summary["포장 PACK"].sum()) if "포장 PACK" in product_summary.columns and not product_summary.empty else 0.0
    )
    product_yongma_in_pack = (
        float(product_summary["용마입고 PACK"].sum())
        if "용마입고 PACK" in product_summary.columns and not product_summary.empty
        else 0.0
    )
    code_yongma_in_pack = float(pd.to_numeric(work["yongma_in_pack"], errors="coerce").fillna(0.0).sum()) if not work.empty else 0.0
    yongma_in_pack = code_yongma_in_pack if code_yongma_in_pack > 0 else product_yongma_in_pack
    packing_shortage_pack = (
        float(product_summary["포장부족수량"].sum())
        if "포장부족수량" in product_summary.columns and not product_summary.empty
        else max(0.0, request_pack - packing_pack)
    )
    receipt_shortage_pack = max(0.0, request_pack - yongma_in_pack)
    receipt_wait_pack = (
        float(product_summary["입고대기수량"].sum())
        if "입고대기수량" in product_summary.columns and not product_summary.empty
        else max(0.0, packing_pack - yongma_in_pack)
    )
    code_production_shortage_pcs = float(
        pd.to_numeric(work.get("_allocated_production_shortage_qty", pd.Series(dtype=float)), errors="coerce")
        .fillna(0.0)
        .sum()
    )
    production_shortage_pcs = code_production_shortage_pcs or (
        float(product_summary["생산부족수량"].sum())
        if "생산부족수량" in product_summary.columns and not product_summary.empty
        else 0.0
    )
    packable_pcs = max(0.0, request_pcs - production_shortage_pcs)
    packing_progress = (packing_pack / request_pack * 100.0) if request_pack > 0 else 0.0
    receipt_progress = (yongma_in_pack / request_pack * 100.0) if request_pack > 0 else 0.0
    production_progress = (
        (request_pcs - production_shortage_pcs) / request_pcs * 100.0
        if request_pcs > 0
        else 0.0
    )
    return {
        "priority_products": float((shortage_mask & priority_mask).sum()),
        "urgent_products": float((shortage_mask & due_mask).sum()),
        "stock_shortage_products": float((shortage_mask & stock_shortage_mask).sum()),
        "request_pack": request_pack,
        "request_pcs": request_pcs,
        "yongma_in_pack": yongma_in_pack,
        "packing_done_pack": packing_pack,
        "packing_todo_pack": packing_shortage_pack,
        "receipt_shortage_pack": receipt_shortage_pack,
        "receipt_wait_pack": receipt_wait_pack,
        "packing_shortage_pack": receipt_shortage_pack,
        "production_shortage_pcs": production_shortage_pcs,
        "packable_pcs": packable_pcs,
        "packing_progress_pct": min(100.0, max(0.0, packing_progress)),
        "receipt_progress_pct": min(100.0, max(0.0, receipt_progress)),
        "production_progress_pct": min(100.0, max(0.0, production_progress)),
    }


def render_operation_kpis(
    product_summary: pd.DataFrame,
    code_summary: pd.DataFrame,
    stock_threshold_pack: float,
    unit_mode: str = UNIT_PACK,
) -> None:
    kpi = calc_operation_kpis(product_summary, code_summary, stock_threshold_pack)
    if unit_mode == UNIT_PCS:
        render_metric_card_grid(
            [
                ("요청 PCS", format_int(kpi["request_pcs"]), "normal"),
                ("생산부족 PCS", format_int(kpi["production_shortage_pcs"]), "warn"),
                ("생산진도율", f"{kpi['production_progress_pct']:.1f}%", metric_progress_tone(kpi["production_progress_pct"])),
                ("포장진도율", f"{kpi['packing_progress_pct']:.1f}%", metric_progress_tone(kpi["packing_progress_pct"])),
                ("용마입고율", f"{kpi['receipt_progress_pct']:.1f}%", metric_progress_tone(kpi["receipt_progress_pct"])),
            ]
        )
        return
    render_metric_card_grid(
        [
            ("긴급 대응 품목 수", f"{int(kpi['priority_products']):,}", "normal"),
            ("재고부족 품목 수", f"{int(kpi['stock_shortage_products']):,}", "normal"),
            ("미입고 PACK", format_int(kpi["packing_shortage_pack"]), "warn"),
            ("생산부족 PCS", format_int(kpi["production_shortage_pcs"]), "warn"),
        ]
    )


def render_unit_selector(key: str) -> str:
    unit_mode = st.radio(
        "조회 단위 선택",
        UNIT_OPTIONS,
        index=0,
        horizontal=True,
        key=key,
    )
    if unit_mode == UNIT_PCS:
        st.caption("포장가능재고·생산부족 기준 조회")
    else:
        st.caption("용마입고·포장부족 기준 조회")
    return unit_mode


def visible_columns(df: pd.DataFrame, columns: list[str]) -> list[str]:
    return [col for col in columns if col in df.columns]


def dataframe_for_streamlit(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    return df.loc[:, ~pd.Index(df.columns).duplicated()].copy()


def product_progress_column_order(df: pd.DataFrame, pack_labels: list[str], unit_mode: str) -> list[str]:
    if unit_mode == UNIT_PCS:
        columns = [
            "우선등급",
            "D-Day",
            "제품명",
            "요청합계(PCS)",
            "생산필요수량(PCS)",
            "생산부족수량(PCS)",
            "생산진도율",
            "용마입고율",
        ]
    else:
        columns = [
            "우선등급",
            "D-Day",
            "제품명",
            *pack_labels,
            "요청합계(PACK)",
            "용마입고 PACK",
            "재고부족(PACK)",
            "미입고(PACK)",
            "생산부족수량(PCS)",
            "용마입고율",
            "생산진도율",
        ]
    return visible_columns(df, columns)


def production_progress_column_order(df: pd.DataFrame, pack_labels: list[str], unit_mode: str) -> list[str]:
    columns = [
        "생산코드",
        "대표 제품명",
        *pack_labels,
        "요청합계(PACK)",
        "포장부족(PACK)",
        "생산부족수량(PCS)",
        "포장진도율",
        "생산진도율",
        "최소 납기일",
    ]
    return visible_columns(df, columns)


def production_power_detail_column_order(df: pd.DataFrame, pack_labels: list[str]) -> list[str]:
    columns = [
        "생산코드 전체",
        "대표 제품명",
        "POWER",
        *pack_labels,
        "요청합계(PACK)",
        "포장부족(PACK)",
        "생산부족수량(PCS)",
        "포장진도율",
        "생산진도율",
        "최소 납기일",
    ]
    return visible_columns(df, columns)


def sales_progress_column_order(df: pd.DataFrame, unit_mode: str) -> list[str]:
    if unit_mode == UNIT_PCS:
        columns = [
            "우선등급",
            "D-Day",
            "판매코드",
            "생산코드",
            "제품명",
            "POWER",
            "PACK",
            "생산요청물량(PCS)",
            "용마입고수량(PCS)",
            "용마입고대기수량(PCS)",
            "포장가능재고(PCS)",
            "포장부족(PCS)",
            "생산부족(PCS)",
            "생산진도율",
            "납기",
            "상태",
        ]
    else:
        columns = [
            "우선등급",
            "D-Day",
            "판매코드",
            "생산코드",
            "제품명",
            "POWER",
            "PACK",
            "생산요청물량(PACK)",
            "용마입고수량(PACK)",
            "용마입고대기수량(PACK)",
            "포장부족(PACK)",
            "용마입고율",
            "납기",
            "상태",
        ]
    return visible_columns(df, columns)


def power_progress_column_order(df: pd.DataFrame, unit_mode: str) -> list[str]:
    if unit_mode == UNIT_PCS:
        columns = [
            "POWER",
            "요청합계(PCS)",
            "생산필요수량(PCS)",
            "생산부족수량(PCS)",
            "생산진도율",
            "포장진도율",
        ]
    else:
        columns = [
            "POWER",
            "요청합계(PACK)",
            "포장 PACK",
            "포장부족(PACK)",
            "생산부족수량(PCS)",
            "포장진도율",
            "생산진도율",
        ]
    return visible_columns(df, columns)


def build_product_progress_main_view(
    product_summary: pd.DataFrame,
    code_summary: pd.DataFrame,
    pack_labels: list[str],
    stock_threshold_pack: float = INVENTORY_STOCK_THRESHOLD_DEFAULT,
) -> pd.DataFrame:
    work = with_operational_columns(code_summary)
    due_by_product = (
        work.groupby("base_product_name", dropna=False)
        .agg(
            production_due_date=("production_due_date", min_datetime),
            request_due_date=("request_due_date", min_datetime),
        )
        .reset_index()
        .rename(columns={"base_product_name": "제품명"})
    )
    pivot = build_pack_pivot(work, ["base_product_name"], pack_labels).rename(columns={"base_product_name": "제품명"})

    out = product_summary.merge(pivot, on="제품명", how="left").merge(due_by_product, on="제품명", how="left")
    for label in pack_labels:
        out[label] = out[label].fillna(0.0)
    if "용마창고재고 (PACK)" not in out.columns:
        out["용마창고재고 (PACK)"] = np.nan
    if "재고매칭SKU수" not in out.columns:
        out["재고매칭SKU수"] = 0
    out["제품필요수량"] = out.get("생산부족수량", 0.0)
    out["생산필요수량(PCS)"] = out["제품필요수량"]
    out["생산부족수량(PCS)"] = out.get("생산부족수량", 0.0)
    out["진도율"] = out.get("생산진도율", 0.0)
    out["전체진도율"] = out["용마입고율"]
    out = add_priority_columns(out, stock_threshold_pack, shortage_col="미입고수량", request_col="요청 PACK")
    out = out.rename(columns={"요청 PACK": "요청합계(PACK)", "요청 PCS": "요청합계(PCS)"})
    out["포장부족(PACK)"] = out["포장부족수량"]
    out["미입고(PACK)"] = out["미입고수량"]
    ordered = [
        "우선등급",
        "D-Day",
        "제품명",
        *pack_labels,
        "요청합계(PACK)",
        "요청합계(PCS)",
        "용마입고 PACK",
        "용마창고재고 (PACK)",
        "재고부족(PACK)",
        "미입고(PACK)",
        "포장부족(PACK)",
        "생산필요수량(PCS)",
        "생산부족수량(PCS)",
        "제품필요수량",
        "진도율",
        "생산진도율",
        "포장부족수량",
        "미입고수량",
        "포장진도율",
        "용마입고율",
        "전체진도율",
        "상태",
        "_priority_sort",
        "_request_due_date_sort",
    ]
    return out[ordered].copy()


def build_product_sku_detail_view(code_summary: pd.DataFrame, product_name: str) -> pd.DataFrame:
    work = with_operational_columns(code_summary)
    scope = work[work["base_product_name"] == product_name].copy()
    if scope.empty:
        scope = work[work["product_name"] == product_name].copy()
    if scope.empty:
        return pd.DataFrame(columns=["SKU", "생산코드", "판매코드 수", "요청 PACK", "용마입고 PACK", "미입고 PACK", "용마입고율"])
    grouped = (
        scope.groupby(["product_name", "production_code_display"], dropna=False)
        .agg(
            sales_code_count=("sales_code", "nunique"),
            request_pack=("request_pack", "sum"),
            packing_pack=("packing_pack", "sum"),
            yongma_in_pack=("yongma_in_pack", "sum"),
        )
        .reset_index()
        .rename(
            columns={
                "product_name": "SKU",
                "production_code_display": "생산코드",
                "sales_code_count": "판매코드 수",
                "request_pack": "요청 PACK",
                "yongma_in_pack": "용마입고 PACK",
            }
        )
    )
    grouped["미입고 PACK"] = (grouped["요청 PACK"] - grouped["용마입고 PACK"]).clip(lower=0.0)
    grouped["용마입고율"] = np.where(grouped["요청 PACK"] > 0, grouped["용마입고 PACK"] / grouped["요청 PACK"] * 100.0, 0.0)
    grouped["용마입고율"] = np.clip(grouped["용마입고율"], 0.0, 100.0)
    return grouped[["SKU", "생산코드", "판매코드 수", "요청 PACK", "용마입고 PACK", "미입고 PACK", "용마입고율"]].sort_values(
        ["미입고 PACK", "요청 PACK"], ascending=[False, False], kind="stable"
    )


def compact_query_text(value: Any) -> str:
    return re.sub(r"\s+", "", clean_str(value)).lower()


def expand_product_query_terms(query: str) -> list[str]:
    text = clean_str(query)
    if not text:
        return []
    terms = [text]
    compact = compact_query_text(text)
    for alias, values in PRODUCT_QUERY_ALIASES.items():
        alias_compact = compact_query_text(alias)
        if alias_compact and (alias_compact in compact or compact in alias_compact):
            terms.extend(values)
    return list(dict.fromkeys([term for term in terms if clean_str(term)]))


def contains_any_query_term(series: pd.Series, terms: list[str]) -> pd.Series:
    if not terms:
        return pd.Series(False, index=series.index)
    text = series.astype(str)
    mask = pd.Series(False, index=series.index)
    for term in terms:
        mask = mask | text.str.contains(term, case=False, na=False, regex=False)
    return mask


def split_lookup_query_terms(query: str) -> list[str]:
    text = clean_str(query)
    if not text:
        return []
    tokens = [clean_str(token) for token in re.split(r"[,，]+|\s+", text)]
    return [token for token in tokens if token]


def parse_quick_lookup_direct_query(
    query: str,
    pack_options: list[str],
    power_options: list[str],
) -> tuple[str, str, list[str]]:
    text = clean_str(query)
    if not text:
        return "", "전체", []

    if "," in text or "，" in text:
        tokens = [clean_str(token) for token in re.split(r"[,，]+", text)]
    else:
        tokens = [clean_str(token) for token in text.split()]
    tokens = [token for token in tokens if token]

    available_packs = set(pack_options)
    available_powers = set(power_options)
    product_terms: list[str] = []
    pack_label = "전체"
    power_labels: list[str] = []

    for token in tokens:
        normalized = token.replace("−", "-").replace("–", "-").replace("—", "-")
        pack_match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*(?:P|팩|개입)", normalized, flags=re.IGNORECASE)
        if pack_match:
            candidate_pack = base_pack_label(float(pack_match.group(1)))
            pack_label = candidate_pack if candidate_pack in available_packs else candidate_pack
            continue

        power_match = re.fullmatch(r"[+-]?\d+(?:\.\d+)?", normalized)
        if power_match and (normalized.startswith(("+", "-")) or "." in normalized):
            candidate_power = format_power(float(normalized))
            power_labels.append(candidate_power if candidate_power in available_powers else candidate_power)
            continue

        product_terms.append(token)

    return " ".join(product_terms).strip(), pack_label, list(dict.fromkeys(power_labels))


def build_product_pack_power_quick_view(
    code_summary: pd.DataFrame,
    product_query: str,
    pack_label: str,
    power_labels: list[str],
) -> pd.DataFrame:
    columns = [
        "제품명",
        "SKU",
        "PACK",
        "POWER",
        "판매코드 수",
        "요청 PACK",
        "포장 PACK",
        "용마입고 PACK",
        "입고대기 PACK",
        "미입고 PACK",
        "요청 PCS",
        "생산부족 PCS",
        "용마입고율",
        "생산진도율",
        "최소 납기",
        "생산완료예상일",
    ]
    if code_summary.empty:
        return pd.DataFrame(columns=columns)

    work = add_allocated_production_basis(with_operational_columns(code_summary))
    query = product_query.strip()
    if query:
        terms = expand_product_query_terms(query)
        mask = (
            contains_any_query_term(work["base_product_name"], terms)
            | contains_any_query_term(work["product_name"], terms)
            | contains_any_query_term(work["sales_code"], terms)
            | contains_any_query_term(work["production_code_display"], terms)
        )
        work = work[mask].copy()
    if pack_label != "전체":
        work = work[work["_pack_label"] == pack_label].copy()
    if power_labels:
        work = work[work["POWER"].isin(power_labels)].copy()
    if work.empty:
        return pd.DataFrame(columns=columns)
    if "yongma_in_pack" not in work.columns:
        work["yongma_in_pack"] = 0.0

    grouped = (
        work.groupby(["base_product_name", "product_name", "_pack_label", "POWER", "power_value"], dropna=False)
        .agg(
            sales_code_count=("sales_code", "nunique"),
            request_pack=("request_pack", "sum"),
            packing_pack=("packing_pack", "sum"),
            yongma_in_pack=("yongma_in_pack", "sum"),
            request_pcs=("request_pcs", "sum"),
            production_shortage_pcs=("_allocated_production_shortage_qty", "sum"),
            request_due_date=("request_due_date", min_datetime),
            production_due_date=("production_due_date", min_datetime),
        )
        .reset_index()
        .rename(
            columns={
                "base_product_name": "제품명",
                "product_name": "SKU",
                "_pack_label": "PACK",
                "sales_code_count": "판매코드 수",
                "request_pack": "요청 PACK",
                "packing_pack": "포장 PACK",
                "yongma_in_pack": "용마입고 PACK",
                "request_pcs": "요청 PCS",
                "production_shortage_pcs": "생산부족 PCS",
            }
        )
    )
    grouped["입고대기 PACK"] = (grouped["포장 PACK"] - grouped["용마입고 PACK"]).clip(lower=0.0)
    grouped["미입고 PACK"] = (grouped["요청 PACK"] - grouped["용마입고 PACK"]).clip(lower=0.0)
    grouped["용마입고율"] = np.where(
        grouped["요청 PACK"] > 0,
        grouped["용마입고 PACK"] / grouped["요청 PACK"] * 100.0,
        0.0,
    )
    grouped["용마입고율"] = np.clip(grouped["용마입고율"], 0.0, 100.0)
    grouped["생산부족 PCS"] = pd.to_numeric(grouped["생산부족 PCS"], errors="coerce").fillna(0.0).round(0)
    grouped["생산진도율"] = calc_production_progress_pct(grouped["요청 PCS"], grouped["생산부족 PCS"])
    grouped["최소 납기"] = grouped["request_due_date"].map(display_date_or_dash)
    grouped["생산완료예상일"] = grouped["production_due_date"].map(display_date_or_dash)
    grouped["_pack_sort"] = grouped["PACK"].map(pack_sort_key)
    return grouped.sort_values(
        ["power_value", "_pack_sort", "미입고 PACK", "요청 PACK"],
        ascending=[True, True, False, False],
        na_position="last",
        kind="stable",
    )[columns].copy()


def render_product_pack_power_quick_lookup(code_summary: pd.DataFrame) -> None:
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    render_panel_title(
        "제품·PACK·POWER 간편 조회",
        "제품명 일부, PACK, POWER 조합으로 포장·용마입고·생산 상태를 확인합니다.",
    )
    pack_options = available_pack_options(code_summary)
    power_options = available_power_options(code_summary)
    direct_query = st.text_input(
        "직접 검색",
        value="",
        placeholder="예: 소울브라운, 40P, -06.50",
        key="quick_lookup_direct_query",
    )
    q1, q2, q3 = st.columns([2.3, 1.1, 2.0], gap="small")
    with q1:
        product_query = st.text_input(
            "제품명/SKU/코드 검색",
            value="",
            placeholder="예: 소울브라운",
            key="quick_lookup_product_query",
        )
    with q2:
        pack_label = st.selectbox(
            "PACK 선택",
            options=pack_options,
            index=0,
            key="quick_lookup_pack",
        )
    with q3:
        power_labels = st.multiselect(
            "POWER 선택",
            options=power_options[1:],
            default=[],
            key="quick_lookup_power",
        )
    if direct_query.strip():
        product_query, pack_label, power_labels = parse_quick_lookup_direct_query(
            direct_query,
            pack_options=pack_options,
            power_options=power_options[1:],
        )

    if not product_query.strip() and pack_label == "전체" and not power_labels:
        return

    quick_view = build_product_pack_power_quick_view(code_summary, product_query, pack_label, power_labels)
    total_request = float(quick_view["요청 PACK"].sum()) if not quick_view.empty else 0.0
    total_yongma = float(quick_view["용마입고 PACK"].sum()) if not quick_view.empty else 0.0
    total_shortage = float(quick_view["미입고 PACK"].sum()) if not quick_view.empty else 0.0
    total_request_pcs = float(quick_view["요청 PCS"].sum()) if not quick_view.empty else 0.0
    total_production_shortage = float(quick_view["생산부족 PCS"].sum()) if not quick_view.empty else 0.0
    receipt_progress = (total_yongma / total_request * 100.0) if total_request > 0 else 0.0
    production_progress = (
        (total_request_pcs - total_production_shortage) / total_request_pcs * 100.0
        if total_request_pcs > 0
        else 0.0
    )
    render_metric_card_grid(
        [
            ("요청 PACK", format_int(total_request), "normal"),
            ("용마입고 PACK", format_int(total_yongma), "normal"),
            ("미입고 PACK", format_int(total_shortage), "warn" if total_shortage > 0 else "normal"),
            ("생산부족 PCS", format_int(total_production_shortage), "warn" if total_production_shortage > 0 else "normal"),
            ("용마입고율", f"{min(100.0, max(0.0, receipt_progress)):.1f}%", metric_progress_tone(receipt_progress)),
            ("생산진도율", f"{min(100.0, max(0.0, production_progress)):.1f}%", metric_progress_tone(production_progress)),
        ]
    )
    st.dataframe(
        quick_view,
        hide_index=True,
        height=min(360, 88 + 35 * max(len(quick_view), 1)),
        width="stretch",
        column_config=drilldown_column_config(),
    )


def extract_sales_prefix(value: Any) -> str:
    text = clean_str(value).upper()
    match = re.match(r"^([A-Z]\d+)", text)
    return match.group(1) if match else ""


def build_daily_request_match_view(code_summary: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "제품코드",
        "PACK",
        "POWER",
        "요청 PACK",
        "포장 PACK",
        "용마입고 PACK",
        "미입고 PACK",
        "요청 PCS",
        "생산부족 PCS",
        "용마입고대기 PACK",
        "포장가능재고(PCS)",
        "샘플신청가능수량",
        "생산진도율",
        "최소 납기",
        "요청제품명",
        "판매코드 수",
    ]
    if code_summary.empty:
        return pd.DataFrame(columns=columns)

    work = add_allocated_production_basis(with_operational_columns(code_summary))
    work["_sales_prefix"] = work["sales_code"].map(extract_sales_prefix)
    work = work[work["_sales_prefix"] != ""].copy()
    if work.empty:
        return pd.DataFrame(columns=columns)
    if "yongma_in_pack" not in work.columns:
        work["yongma_in_pack"] = 0.0

    grouped = (
        work.groupby(["_sales_prefix", "_pack_label", "POWER"], dropna=False)
        .agg(
            request_pack=("request_pack", "sum"),
            packing_pack=("packing_pack", "sum"),
            yongma_in_pack=("yongma_in_pack", "sum"),
            request_pcs=("request_pcs", "sum"),
            production_shortage_pcs=("_allocated_production_shortage_qty", "sum"),
            sample_available_pcs=("_allocated_sample_available_pcs", "sum"),
            request_due_date=("request_due_date", min_datetime),
            product_name=("product_name", join_unique),
            sales_code_count=("sales_code", "nunique"),
        )
        .reset_index()
        .rename(
            columns={
                "_sales_prefix": "제품코드",
                "_pack_label": "PACK",
                "request_pack": "요청 PACK",
                "packing_pack": "포장 PACK",
                "yongma_in_pack": "용마입고 PACK",
                "request_pcs": "요청 PCS",
                "production_shortage_pcs": "생산부족 PCS",
                "sample_available_pcs": "샘플신청가능수량",
                "product_name": "요청제품명",
                "sales_code_count": "판매코드 수",
            }
        )
    )
    grouped["미입고 PACK"] = (grouped["요청 PACK"] - grouped["용마입고 PACK"]).clip(lower=0.0)
    grouped["용마입고대기 PACK"] = (grouped["포장 PACK"] - grouped["용마입고 PACK"]).clip(lower=0.0)
    grouped["포장가능재고(PCS)"] = (
        grouped["요청 PCS"] - grouped["생산부족 PCS"] + grouped["샘플신청가능수량"]
    ).clip(lower=0.0)
    grouped["생산진도율"] = calc_production_progress_pct(grouped["요청 PCS"], grouped["생산부족 PCS"])
    grouped["최소 납기"] = grouped["request_due_date"].map(display_date_or_dash)
    return grouped[columns].copy()


def build_daily_production_power_catalog(code_summary: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "제품코드",
        "POWER",
        "_production_request_pack",
        "_production_request_pcs",
        "_production_shortage_pcs",
        "_production_sample_available_pcs",
        "_production_available_stock_pcs",
        "_production_progress_pct",
    ]
    if code_summary.empty:
        return pd.DataFrame(columns=columns)

    work = add_allocated_production_basis(with_operational_columns(code_summary))
    work["_sales_prefix"] = work["sales_code"].map(extract_sales_prefix)
    work["_production_key"] = work.get("production_code_key", pd.Series("", index=work.index)).map(clean_str)
    fallback_key = work.get("sales_code_key", pd.Series("", index=work.index)).map(clean_str)
    work["_production_key"] = work["_production_key"].where(work["_production_key"] != "", fallback_key)
    work = work[
        (work["_sales_prefix"].map(clean_str) != "")
        & (work["POWER"].map(clean_str) != "")
        & (work["_production_key"].map(clean_str) != "")
    ].copy()
    if work.empty:
        return pd.DataFrame(columns=columns)

    for col in ["request_pack", "request_pcs", "production_basis_qty", "sample_available_pcs"]:
        if col not in work.columns:
            work[col] = 0.0
        work[col] = pd.to_numeric(work[col], errors="coerce").fillna(0.0)

    by_production = (
        work.groupby(["_sales_prefix", "POWER", "_production_key"], dropna=False)
        .agg(
            request_pack=("request_pack", "sum"),
            request_pcs=("request_pcs", "sum"),
            production_shortage_pcs=("production_basis_qty", "max"),
            sample_available_pcs=("sample_available_pcs", "max"),
        )
        .reset_index()
    )
    grouped = (
        by_production.groupby(["_sales_prefix", "POWER"], dropna=False)
        .agg(
            _production_request_pack=("request_pack", "sum"),
            _production_request_pcs=("request_pcs", "sum"),
            _production_shortage_pcs=("production_shortage_pcs", "sum"),
            _production_sample_available_pcs=("sample_available_pcs", "sum"),
        )
        .reset_index()
        .rename(columns={"_sales_prefix": "제품코드"})
    )
    grouped["_production_available_stock_pcs"] = (
        grouped["_production_request_pcs"]
        - grouped["_production_shortage_pcs"]
        + grouped["_production_sample_available_pcs"]
    ).clip(lower=0.0)
    grouped["_production_progress_pct"] = calc_production_progress_pct(
        grouped["_production_request_pcs"],
        grouped["_production_shortage_pcs"],
    )
    for col in [
        "_production_request_pack",
        "_production_request_pcs",
        "_production_shortage_pcs",
        "_production_sample_available_pcs",
        "_production_available_stock_pcs",
    ]:
        grouped[col] = pd.to_numeric(grouped[col], errors="coerce").fillna(0.0).round(0)
    return grouped[columns].copy()


def build_daily_base_power_production_catalog(code_summary: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "_daily_base_product_name",
        "POWER",
        "_base_production_request_pack",
        "_base_production_request_pcs",
        "_base_production_shortage_pcs",
        "_base_production_sample_available_pcs",
        "_base_production_available_stock_pcs",
        "_base_production_progress_pct",
    ]
    if code_summary.empty:
        return pd.DataFrame(columns=columns)

    work = add_allocated_production_basis(with_operational_columns(code_summary))
    if "base_product_name" in work.columns:
        work["_daily_base_product_name"] = work["base_product_name"].map(clean_str)
    else:
        work["_daily_base_product_name"] = work["product_name"].map(strip_pack_unit_suffix).map(clean_str)
    work["_production_key"] = work.get("production_code_key", pd.Series("", index=work.index)).map(clean_str)
    fallback_key = work.get("sales_code_key", pd.Series("", index=work.index)).map(clean_str)
    work["_production_key"] = work["_production_key"].where(work["_production_key"] != "", fallback_key)
    work = work[
        (work["_daily_base_product_name"].map(clean_str) != "")
        & (work["POWER"].map(clean_str) != "")
        & (work["_production_key"].map(clean_str) != "")
    ].copy()
    if work.empty:
        return pd.DataFrame(columns=columns)

    for col in ["request_pack", "request_pcs", "production_basis_qty", "sample_available_pcs"]:
        if col not in work.columns:
            work[col] = 0.0
        work[col] = pd.to_numeric(work[col], errors="coerce").fillna(0.0)

    by_production = (
        work.groupby(["_daily_base_product_name", "POWER", "_production_key"], dropna=False)
        .agg(
            request_pack=("request_pack", "sum"),
            request_pcs=("request_pcs", "sum"),
            production_shortage_pcs=("production_basis_qty", "max"),
            sample_available_pcs=("sample_available_pcs", "max"),
        )
        .reset_index()
    )
    grouped = (
        by_production.groupby(["_daily_base_product_name", "POWER"], dropna=False)
        .agg(
            _base_production_request_pack=("request_pack", "sum"),
            _base_production_request_pcs=("request_pcs", "sum"),
            _base_production_shortage_pcs=("production_shortage_pcs", "sum"),
            _base_production_sample_available_pcs=("sample_available_pcs", "sum"),
        )
        .reset_index()
    )
    grouped["_base_production_available_stock_pcs"] = (
        grouped["_base_production_request_pcs"]
        - grouped["_base_production_shortage_pcs"]
        + grouped["_base_production_sample_available_pcs"]
    ).clip(lower=0.0)
    grouped["_base_production_progress_pct"] = calc_production_progress_pct(
        grouped["_base_production_request_pcs"],
        grouped["_base_production_shortage_pcs"],
    )
    for col in [
        "_base_production_request_pack",
        "_base_production_request_pcs",
        "_base_production_shortage_pcs",
        "_base_production_sample_available_pcs",
        "_base_production_available_stock_pcs",
    ]:
        grouped[col] = pd.to_numeric(grouped[col], errors="coerce").fillna(0.0).round(0)
    return grouped[columns].copy()


def pack_unit_from_label(value: Any) -> float:
    match = re.search(r"(\d+(?:\.\d+)?)", clean_str(value))
    if not match:
        return 1.0
    try:
        number = float(match.group(1))
    except ValueError:
        return 1.0
    return number if number > 0 else 1.0


def build_item_code_from_prefix_power(product_code: Any, power_label: Any) -> str:
    prefix = clean_str(product_code).upper()
    power = clean_str(power_label)
    if not prefix or not power:
        return ""
    return f"{prefix}{power}"


def replace_power_in_production_code(template: Any, source_power: Any, target_power: Any) -> str:
    text = clean_str(template)
    target = clean_str(target_power)
    source = clean_str(source_power)
    if not text or not target:
        return ""
    if source and source in text:
        return text.replace(source, target, 1)
    match = re.search(r"[+-]\d{1,2}\.\d{2}", text)
    if not match:
        return ""
    return f"{text[:match.start()]}{target}{text[match.end():]}"


def build_sample_available_lookup(sample_available_df: pd.DataFrame | None) -> dict[str, float]:
    if sample_available_df is None or sample_available_df.empty:
        return {}
    if "production_code_key" not in sample_available_df.columns or "sample_available_pcs" not in sample_available_df.columns:
        return {}
    grouped = (
        sample_available_df.copy()
        .assign(sample_available_pcs=lambda df: pd.to_numeric(df["sample_available_pcs"], errors="coerce").fillna(0.0))
        .groupby("production_code_key", dropna=False)["sample_available_pcs"]
        .sum()
    )
    return {clean_str(key): float(value) for key, value in grouped.items() if clean_str(key)}


def build_daily_product_catalog(code_summary: pd.DataFrame) -> pd.DataFrame:
    columns = ["제품코드", "PACK", "마스터제품명", "생산코드템플릿", "생산코드POWER"]
    if code_summary.empty:
        return pd.DataFrame(columns=columns)

    work = with_operational_columns(code_summary)
    work["_sales_prefix"] = work["sales_code"].map(extract_sales_prefix)
    work = work[(work["_sales_prefix"] != "") & (work["_pack_label"].map(clean_str) != "")].copy()
    if work.empty:
        return pd.DataFrame(columns=columns)

    grouped = (
        work.groupby(["_sales_prefix", "_pack_label"], dropna=False)
        .agg(
            product_name=("product_name", join_unique),
            production_code=("production_code", first_nonempty),
            power=("POWER", first_nonempty),
        )
        .reset_index()
        .rename(
            columns={
                "_sales_prefix": "제품코드",
                "_pack_label": "PACK",
                "product_name": "마스터제품명",
                "production_code": "생산코드템플릿",
                "power": "생산코드POWER",
            }
        )
    )
    return grouped[columns].copy()


def build_daily_code_power_catalog(code_summary: pd.DataFrame) -> pd.DataFrame:
    columns = ["제품코드", "POWER", "_code_pack", "_code_product_name"]
    if code_summary.empty:
        return pd.DataFrame(columns=columns)

    work = with_operational_columns(code_summary)
    work["_sales_prefix"] = work["sales_code"].map(extract_sales_prefix)
    work = work[(work["_sales_prefix"] != "") & (work["POWER"].map(clean_str) != "")].copy()
    if work.empty:
        return pd.DataFrame(columns=columns)

    grouped = (
        work.groupby(["_sales_prefix", "POWER"], dropna=False)
        .agg(
            pack=("_pack_label", first_nonempty),
            product_name=("product_name", first_nonempty),
        )
        .reset_index()
        .rename(
            columns={
                "_sales_prefix": "제품코드",
                "pack": "_code_pack",
                "product_name": "_code_product_name",
            }
        )
    )
    return grouped[columns].copy()


def build_daily_code_catalog(code_summary: pd.DataFrame) -> pd.DataFrame:
    columns = ["제품코드", "_code_default_pack", "_code_default_product_name"]
    if code_summary.empty:
        return pd.DataFrame(columns=columns)

    work = with_operational_columns(code_summary)
    work["_sales_prefix"] = work["sales_code"].map(extract_sales_prefix)
    work = work[work["_sales_prefix"] != ""].copy()
    if work.empty:
        return pd.DataFrame(columns=columns)

    grouped = (
        work.groupby("_sales_prefix", dropna=False)
        .agg(
            pack=("_pack_label", first_nonempty),
            product_name=("product_name", first_nonempty),
        )
        .reset_index()
        .rename(
            columns={
                "_sales_prefix": "제품코드",
                "pack": "_code_default_pack",
                "product_name": "_code_default_product_name",
            }
        )
    )
    return grouped[columns].copy()


def enrich_daily_inventory_from_code_summary(daily_inventory_df: pd.DataFrame, code_summary: pd.DataFrame) -> pd.DataFrame:
    if daily_inventory_df is None or daily_inventory_df.empty or code_summary.empty:
        return daily_inventory_df

    exact_catalog = build_daily_code_power_catalog(code_summary)
    code_catalog = build_daily_code_catalog(code_summary)
    if exact_catalog.empty and code_catalog.empty:
        return daily_inventory_df

    out = daily_inventory_df.copy()
    out["제품코드"] = out["제품코드"].map(clean_str).str.upper()
    out["POWER"] = out["POWER"].map(clean_str)
    if not exact_catalog.empty:
        out = out.merge(exact_catalog, on=["제품코드", "POWER"], how="left")
        catalog_pack = out["_code_pack"].map(clean_str)
        catalog_product = out["_code_product_name"].map(clean_str)
        out["PACK"] = out["_code_pack"].where(catalog_pack != "", out["PACK"])
        out["제품명"] = out["_code_product_name"].where(catalog_product != "", out["제품명"])
        out = out.drop(columns=["_code_pack", "_code_product_name"], errors="ignore")
    if not code_catalog.empty:
        out = out.merge(code_catalog, on="제품코드", how="left")
        current_pack = out["PACK"].map(clean_str)
        current_product = out["제품명"].map(clean_str)
        out["PACK"] = out["PACK"].where(current_pack != "", out["_code_default_pack"])
        out["제품명"] = out["제품명"].where(current_product != "", out["_code_default_product_name"])
        out = out.drop(columns=["_code_default_pack", "_code_default_product_name"], errors="ignore")
    return out[DAILY_INVENTORY_COLUMNS].copy()


def classify_daily_inventory_status(row: pd.Series) -> str:
    urgent = bool(row.get("긴급요청", False))
    stock = pd.to_numeric(row.get("재고수량", np.nan), errors="coerce")
    request_pack = pd.to_numeric(row.get("요청 PACK", 0.0), errors="coerce")
    has_request = pd.notna(request_pack) and float(request_pack) > 0
    stock_negative = pd.notna(stock) and float(stock) < 0
    if urgent and has_request:
        return "요청내 긴급"
    if urgent:
        return "요청외 긴급"
    if stock_negative and has_request:
        return "요청내 재고부족"
    if stock_negative:
        return "재고 음수"
    if has_request:
        return "요청내 재고확인"
    return "재고 모니터링"


def complete_daily_response_mask(df: pd.DataFrame) -> pd.Series:
    required_cols = ["품목코드", "제품명", "제품코드", "PACK", "POWER"]
    mask = pd.Series(True, index=df.index)
    for col in required_cols:
        if col not in df.columns:
            return pd.Series(False, index=df.index)
        mask &= df[col].map(clean_str) != ""
    return mask


def daily_item_code_base(value: Any) -> str:
    match = re.match(r"^(S\d{3})", clean_str(value).upper())
    return match.group(1) if match else ""


def daily_inventory_status_rank(value: Any) -> int:
    ranks = {
        "요청외 긴급": 0,
        "요청내 긴급": 1,
        "요청내 재고부족": 2,
        "재고 음수": 3,
        "요청내 재고확인": 4,
        "재고 모니터링": 5,
    }
    return ranks.get(clean_str(value), 99)


def first_daily_inventory_status(series: pd.Series) -> str:
    values = [clean_str(value) for value in series if clean_str(value)]
    if not values:
        return ""
    return min(values, key=daily_inventory_status_rank)


def build_daily_lot_wait_lookup(lot_status_df: pd.DataFrame | None) -> dict[str, float]:
    if lot_status_df is None or lot_status_df.empty:
        return {}
    required_cols = ["제품코드", "입고대기수량"]
    if any(col not in lot_status_df.columns for col in required_cols):
        return {}

    work = lot_status_df[required_cols].copy()
    work["_lot_item_key"] = work["제품코드"].map(normalize_match_key)
    work["입고대기수량"] = pd.to_numeric(work["입고대기수량"], errors="coerce").fillna(0.0)
    work = work[work["_lot_item_key"] != ""].copy()
    if work.empty:
        return {}

    grouped = work.groupby("_lot_item_key", dropna=False)["입고대기수량"].sum()
    return {clean_str(key): float(value) for key, value in grouped.items() if clean_str(key)}


def build_daily_inventory_response_view(
    daily_inventory_df: pd.DataFrame,
    code_summary: pd.DataFrame,
    sample_available_df: pd.DataFrame | None = None,
    lot_status_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    columns = [
        "대응상태",
        "품목코드",
        "제품명",
        "재고표 제품명",
        "제품코드",
        "PACK",
        "POWER",
        "재고수량",
        "전일재고",
        "재고증감",
        "재고부족수량",
        "긴급요청",
        "요청 PACK",
        "용마입고 PACK",
        "미입고 PACK",
        "포장 PACK",
        "요청 PCS",
        "생산부족 PCS",
        "용마입고대기 PACK",
        "포장부족(재고 PCS)",
        "포장가능재고(PCS)",
        "생산진도율",
        "최소 납기",
        "요청제품명",
        "판매코드 수",
        "대상품목",
    ]
    if daily_inventory_df.empty:
        return pd.DataFrame(columns=columns)

    daily = daily_inventory_df.copy()
    daily = enrich_daily_inventory_from_code_summary(daily, code_summary)
    daily["제품코드"] = daily["제품코드"].map(clean_str).str.upper()
    daily["PACK"] = daily["PACK"].map(clean_str)
    daily["POWER"] = daily["POWER"].map(clean_str)
    daily["_daily_base_product_name"] = daily["제품명"].map(strip_pack_unit_suffix).map(clean_str)
    daily["재고수량"] = pd.to_numeric(daily["재고수량"], errors="coerce")
    daily["전일재고"] = pd.to_numeric(daily["전일재고"], errors="coerce")
    daily["재고증감"] = pd.to_numeric(daily["재고증감"], errors="coerce")
    daily["긴급요청"] = daily["긴급요청"].apply(lambda value: bool(value) if not pd.isna(value) else False)

    request_match = build_daily_request_match_view(code_summary)
    out = daily.merge(
        request_match,
        on=["제품코드", "PACK", "POWER"],
        how="left",
    )
    production_catalog = build_daily_production_power_catalog(code_summary)
    out = out.merge(production_catalog, on=["제품코드", "POWER"], how="left")
    base_production_catalog = build_daily_base_power_production_catalog(code_summary)
    out = out.merge(base_production_catalog, on=["_daily_base_product_name", "POWER"], how="left")
    product_catalog = build_daily_product_catalog(code_summary)
    out = out.merge(product_catalog, on=["제품코드", "PACK"], how="left")
    out["재고표 제품명"] = out["제품명"].map(clean_str)
    out["품목코드"] = [
        build_item_code_from_prefix_power(product_code, power)
        for product_code, power in zip(out["제품코드"], out["POWER"])
    ]
    out["제품명"] = out["요청제품명"].where(out["요청제품명"].map(clean_str) != "", out["마스터제품명"])
    out["제품명"] = out["제품명"].where(out["제품명"].map(clean_str) != "", out["재고표 제품명"])
    for col in ["최소 납기", "요청제품명", "대상품목", "마스터제품명", "재고표 제품명", "품목코드"]:
        if col in out.columns:
            out[col] = out[col].fillna("")
    out = out[complete_daily_response_mask(out)].copy()
    if out.empty:
        return pd.DataFrame(columns=columns)
    sample_lookup = build_sample_available_lookup(sample_available_df)
    inferred_production_code = [
        replace_power_in_production_code(template, source_power, power)
        for template, source_power, power in zip(
            out.get("생산코드템플릿", pd.Series("", index=out.index)),
            out.get("생산코드POWER", pd.Series("", index=out.index)),
            out["POWER"],
        )
    ]
    inferred_sample_available = [
        sample_lookup.get(normalize_match_key(code), 0.0)
        for code in inferred_production_code
    ]
    out["_추정샘플신청가능수량"] = inferred_sample_available
    numeric_cols = [
        "요청 PACK",
        "포장 PACK",
        "용마입고 PACK",
        "미입고 PACK",
        "요청 PCS",
        "생산부족 PCS",
        "용마입고대기 PACK",
        "포장가능재고(PCS)",
        "샘플신청가능수량",
        "생산진도율",
        "판매코드 수",
        "_production_request_pack",
        "_production_request_pcs",
        "_production_shortage_pcs",
        "_production_sample_available_pcs",
        "_production_available_stock_pcs",
        "_production_progress_pct",
        "_base_production_request_pack",
        "_base_production_request_pcs",
        "_base_production_shortage_pcs",
        "_base_production_sample_available_pcs",
        "_base_production_available_stock_pcs",
        "_base_production_progress_pct",
    ]
    for col in numeric_cols:
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0)
    for col in [
        "용마입고대기 PACK",
        "포장가능재고(PCS)",
        "샘플신청가능수량",
        "_production_request_pcs",
        "_production_shortage_pcs",
        "_production_sample_available_pcs",
        "_production_available_stock_pcs",
        "_production_progress_pct",
        "_base_production_request_pcs",
        "_base_production_shortage_pcs",
        "_base_production_sample_available_pcs",
        "_base_production_available_stock_pcs",
        "_base_production_progress_pct",
    ]:
        if col not in out.columns:
            out[col] = 0.0
    has_request = out["요청 PACK"] > 0
    has_code_production_context = (
        (out["_production_request_pcs"] > 0)
        | (out["_production_shortage_pcs"] > 0)
    )
    has_base_production_context = (
        (out["_base_production_request_pcs"] > 0)
        | (out["_base_production_shortage_pcs"] > 0)
    )
    use_base_production = ~has_code_production_context & has_base_production_context
    matched_production_shortage_pcs = np.where(
        use_base_production,
        out["_base_production_shortage_pcs"],
        out["_production_shortage_pcs"],
    )
    matched_production_sample_pcs = np.where(
        use_base_production,
        out["_base_production_sample_available_pcs"],
        out["_production_sample_available_pcs"],
    )
    matched_production_available_pcs = np.where(
        use_base_production,
        out["_base_production_available_stock_pcs"],
        out["_production_available_stock_pcs"],
    )
    matched_production_progress_pct = np.where(
        use_base_production,
        out["_base_production_progress_pct"],
        out["_production_progress_pct"],
    )
    has_production_context = has_code_production_context | has_base_production_context
    out["샘플신청가능수량"] = np.where(
        has_production_context,
        matched_production_sample_pcs,
        out["샘플신청가능수량"],
    )
    out["샘플신청가능수량"] = out["샘플신청가능수량"].where(
        has_request | has_production_context,
        out["_추정샘플신청가능수량"],
    )
    lot_wait_lookup = build_daily_lot_wait_lookup(lot_status_df)
    if lot_wait_lookup:
        item_keys = out["품목코드"].map(normalize_match_key)
        out["용마입고대기 PACK"] = item_keys.map(lambda key: lot_wait_lookup.get(key, 0.0))
    else:
        out["용마입고대기 PACK"] = 0.0
    out["생산부족 PCS"] = np.where(
        has_production_context,
        matched_production_shortage_pcs,
        out["생산부족 PCS"],
    )
    out["생산진도율"] = np.where(
        has_production_context,
        matched_production_progress_pct,
        out["생산진도율"],
    )
    out["생산부족 PCS"] = pd.to_numeric(out["생산부족 PCS"], errors="coerce").fillna(0.0).round(0)
    exact_available_stock_pcs = np.where(
        has_request,
        (out["요청 PCS"] - out["생산부족 PCS"] + out["샘플신청가능수량"]).clip(lower=0.0),
        0.0,
    )
    available_stock_pcs = np.where(
        has_production_context,
        matched_production_available_pcs,
        exact_available_stock_pcs,
    )
    out["재고부족수량"] = (-out["재고수량"]).clip(lower=0.0).fillna(0.0)
    pack_units = out["PACK"].map(pack_unit_from_label)
    stock_shortage_pcs = (out["재고부족수량"] * pack_units).clip(lower=0.0)
    stock_available_pcs = (out["재고수량"].clip(lower=0.0).fillna(0.0) * pack_units).clip(lower=0.0)
    sample_available_pcs = pd.to_numeric(out["샘플신청가능수량"], errors="coerce").fillna(0.0)
    fallback_available_stock_pcs = sample_available_pcs.where(
        sample_available_pcs >= stock_available_pcs,
        stock_available_pcs,
    )
    has_supply_context = has_request | has_production_context
    out["포장부족(재고 PCS)"] = np.where(
        has_supply_context,
        (stock_shortage_pcs - available_stock_pcs).clip(lower=0.0),
        (stock_shortage_pcs - fallback_available_stock_pcs).clip(lower=0.0),
    )
    out["포장가능재고(PCS)"] = np.where(
        has_supply_context,
        available_stock_pcs,
        fallback_available_stock_pcs,
    )
    out["포장가능재고(PCS)"] = pd.to_numeric(out["포장가능재고(PCS)"], errors="coerce").fillna(0.0).round(0)
    out["대응상태"] = out.apply(classify_daily_inventory_status, axis=1)
    out["_urgent_sort"] = out["긴급요청"].astype(int)
    out["_negative_sort"] = (out["재고수량"] < 0).astype(int)
    out["_request_sort"] = (out["요청 PACK"] > 0).astype(int)
    out["_pack_sort"] = out["PACK"].map(pack_sort_key)
    out["_power_sort"] = pd.to_numeric(out["POWER"].str.replace("-00.00", "0", regex=False), errors="coerce").fillna(0.0)
    out = out.sort_values(
        ["_urgent_sort", "_negative_sort", "_request_sort", "재고부족수량", "제품명", "_pack_sort", "_power_sort"],
        ascending=[False, False, False, False, True, True, True],
        kind="stable",
    )
    return out[columns].copy()


def build_daily_inventory_main_view(response_view: pd.DataFrame) -> pd.DataFrame:
    visible_columns = [
        "대응상태",
        "품목코드",
        "대표 제품명",
        "상세 건수",
        "긴급요청 수",
        "재고수량",
        "재고부족수량",
        "요청 PACK",
        "용마입고 PACK",
        "용마입고대기 PACK",
        "포장가능재고(PCS)",
        "생산부족 PCS",
        "생산진도율",
        "최소 납기",
    ]
    if response_view.empty:
        return pd.DataFrame(columns=visible_columns + ["_daily_item_code_base", "_daily_min_due_date_sort"])

    work = response_view.copy()
    work["_daily_item_code_base"] = work["품목코드"].map(daily_item_code_base)
    work = work[work["_daily_item_code_base"] != ""].copy()
    if work.empty:
        return pd.DataFrame(columns=visible_columns + ["_daily_item_code_base", "_daily_min_due_date_sort"])

    numeric_cols = [
        "재고수량",
        "재고부족수량",
        "요청 PACK",
        "용마입고 PACK",
        "미입고 PACK",
        "포장 PACK",
        "요청 PCS",
        "생산부족 PCS",
        "용마입고대기 PACK",
        "포장가능재고(PCS)",
    ]
    for col in numeric_cols:
        if col not in work.columns:
            work[col] = 0.0
        work[col] = pd.to_numeric(work[col], errors="coerce").fillna(0.0)
    work["_min_due_date_sort"] = pd.to_datetime(work.get("최소 납기", pd.NaT), errors="coerce")

    grouped = (
        work.groupby("_daily_item_code_base", dropna=False)
        .agg(
            status=("대응상태", first_daily_inventory_status),
            product_name=("제품명", first_nonempty),
            detail_count=("품목코드", "count"),
            urgent_count=("긴급요청", "sum"),
            stock_qty=("재고수량", "sum"),
            stock_shortage_qty=("재고부족수량", "sum"),
            request_pack=("요청 PACK", "sum"),
            yongma_in_pack=("용마입고 PACK", "sum"),
            yongma_wait_pack=("용마입고대기 PACK", "sum"),
            packable_stock_pcs=("포장가능재고(PCS)", "sum"),
            production_shortage_pcs=("생산부족 PCS", "sum"),
            request_pcs=("요청 PCS", "sum"),
            min_due_date=("_min_due_date_sort", min_datetime),
        )
        .reset_index()
    )
    grouped["생산진도율"] = calc_production_progress_pct(grouped["request_pcs"], grouped["production_shortage_pcs"])
    grouped["_daily_status_sort"] = grouped["status"].map(daily_inventory_status_rank)
    grouped["_daily_negative_sort"] = (grouped["stock_qty"] < 0).astype(int)
    grouped["_daily_min_due_date_sort"] = pd.to_datetime(grouped["min_due_date"], errors="coerce")
    grouped["최소 납기"] = grouped["min_due_date"].map(display_date_or_dash)

    out = grouped.rename(
        columns={
            "_daily_item_code_base": "품목코드",
            "status": "대응상태",
            "product_name": "대표 제품명",
            "detail_count": "상세 건수",
            "urgent_count": "긴급요청 수",
            "stock_qty": "재고수량",
            "stock_shortage_qty": "재고부족수량",
            "request_pack": "요청 PACK",
            "yongma_in_pack": "용마입고 PACK",
            "yongma_wait_pack": "용마입고대기 PACK",
            "packable_stock_pcs": "포장가능재고(PCS)",
            "production_shortage_pcs": "생산부족 PCS",
        }
    )
    out["_daily_item_code_base"] = out["품목코드"]
    out = out.sort_values(
        [
            "_daily_status_sort",
            "_daily_negative_sort",
            "재고부족수량",
            "용마입고대기 PACK",
            "_daily_min_due_date_sort",
            "품목코드",
        ],
        ascending=[True, False, False, False, True, True],
        na_position="last",
        kind="stable",
    )
    return out[
        visible_columns
        + [
            "_daily_item_code_base",
            "_daily_status_sort",
            "_daily_negative_sort",
            "_daily_min_due_date_sort",
        ]
    ].copy()


def daily_inventory_detail_column_order(df: pd.DataFrame) -> list[str]:
    columns = [
        "대응상태",
        "품목코드",
        "제품명",
        "제품코드",
        "PACK",
        "POWER",
        "재고수량",
        "긴급요청",
        "요청 PACK",
        "용마입고 PACK",
        "용마입고대기 PACK",
        "포장가능재고(PCS)",
        "생산부족 PCS",
        "생산진도율",
        "최소 납기",
    ]
    return visible_columns(df, columns)


def daily_inventory_search_variants(token: str) -> list[str]:
    normalized = clean_str(token).replace("−", "-").replace("–", "-").replace("—", "-")
    variants = [normalized]
    variants.extend(expand_product_query_terms(normalized))

    pack_label = extract_daily_pack_label(normalized)
    if pack_label:
        variants.append(pack_label)

    power_label = daily_power_label(normalized)
    if power_label:
        variants.append(power_label)

    return list(dict.fromkeys([variant for variant in variants if clean_str(variant)]))


def is_power_query_token(token: str) -> bool:
    normalized = clean_str(token).replace("−", "-").replace("–", "-").replace("—", "-")
    if normalized.upper() == "PL":
        return True
    return bool(re.fullmatch(r"[+-]?\d+(?:\.\d+)?", normalized) and (normalized.startswith(("+", "-")) or "." in normalized))


def is_item_code_query_token(token: str) -> bool:
    normalized = clean_str(token).replace("−", "-").replace("–", "-").replace("—", "-").upper()
    return bool(re.fullmatch(r"[A-Z]\d+[+-]\d+(?:\.\d+)?", normalized))


def daily_inventory_query_mask(view: pd.DataFrame, query: str) -> pd.Series:
    tokens = split_lookup_query_terms(query)
    if not tokens:
        return pd.Series(True, index=view.index)

    text_cols = [
        col
        for col in ["품목코드", "제품명", "재고표 제품명", "제품코드", "요청제품명", "대상품목"]
        if col in view.columns
    ]
    mask = pd.Series(True, index=view.index)
    for token in tokens:
        normalized = clean_str(token).replace("−", "-").replace("–", "-").replace("—", "-")
        token_mask = pd.Series(False, index=view.index)
        power_label = daily_power_label(normalized) if is_power_query_token(normalized) else ""
        pack_label = extract_daily_pack_label(normalized)
        if is_item_code_query_token(normalized) and "품목코드" in view.columns:
            token_mask = contains_any_query_term(view["품목코드"].fillna(""), [normalized.upper()])
        elif power_label and "POWER" in view.columns:
            token_mask = contains_any_query_term(view["POWER"].fillna(""), [power_label])
        elif pack_label and "PACK" in view.columns:
            token_mask = contains_any_query_term(view["PACK"].fillna(""), [pack_label])
        else:
            variants = daily_inventory_search_variants(normalized)
            for col in text_cols:
                token_mask = token_mask | contains_any_query_term(view[col].fillna(""), variants)
        mask = mask & token_mask
    return mask


def build_sales_pack_detail_view(code_summary: pd.DataFrame) -> pd.DataFrame:
    if code_summary.empty:
        return pd.DataFrame(columns=["판매코드", "PACK", "요청", "용마입고", "미입고", "납기"])
    work = with_operational_columns(code_summary)
    out = (
        work.groupby(["sales_code", "_pack_label"], dropna=False)
        .agg(
            request_pack=("request_pack", "sum"),
            packing_pack=("packing_pack", "sum"),
            yongma_in_pack=("yongma_in_pack", "sum"),
            request_due_date=("request_due_date", min_datetime),
        )
        .reset_index()
        .rename(columns={"sales_code": "판매코드", "_pack_label": "PACK", "request_pack": "요청", "yongma_in_pack": "용마입고"})
    )
    out["미입고"] = (out["요청"] - out["용마입고"]).clip(lower=0.0)
    out["납기"] = out["request_due_date"].map(display_date_or_dash)
    out["_pack_sort"] = out["PACK"].map(pack_sort_key)
    return out.sort_values(["미입고", "_pack_sort", "요청"], ascending=[False, True, False], kind="stable")[
        ["판매코드", "PACK", "요청", "용마입고", "미입고", "납기"]
    ]


def build_lot_receipt_status_view(
    packing_df: pd.DataFrame,
    yongma_df: pd.DataFrame,
    code_summary: pd.DataFrame,
) -> pd.DataFrame:
    columns = [
        "제품코드",
        "제품명",
        "LOTNO",
        "포장일",
        "포장실적수량",
        "용마입고수량",
        "입고대기수량",
        "상태",
    ]
    if packing_df.empty or code_summary.empty:
        return pd.DataFrame(columns=columns)

    domestic_code_keys = set(code_summary.get("sales_code_key", pd.Series(dtype=str)).map(clean_str)) - {""}
    if not domestic_code_keys:
        return pd.DataFrame(columns=columns)

    pack = packing_df.copy()
    for col in ["sales_code_key", "packing_lot_key", "packing_barcode_key"]:
        if col not in pack.columns:
            pack[col] = ""
    pack["sales_code_key"] = pack["sales_code_key"].map(clean_str)
    pack = pack[pack["sales_code_key"].isin(domestic_code_keys)].copy()
    if pack.empty:
        return pd.DataFrame(columns=columns)

    for col in ["packing_product_name", "packing_lot", "packing_barcode"]:
        if col not in pack.columns:
            pack[col] = ""
    if "packing_date" not in pack.columns:
        pack["packing_date"] = pd.NaT

    grouped = (
        pack.groupby(
            [
                "sales_code_key",
                "sales_code",
                "packing_product_name",
                "packing_lot",
                "packing_lot_key",
                "packing_barcode",
                "packing_barcode_key",
            ],
            dropna=False,
        )
        .agg(
            packing_pack=("packing_pack", "sum"),
            packing_date=("packing_date", min_datetime),
        )
        .reset_index()
    )
    grouped["용마입고수량"] = 0.0

    if yongma_df is not None and not yongma_df.empty:
        yongma = yongma_df.copy()
        yongma["sales_code_key"] = yongma["sales_code_key"].map(clean_str)
        yongma = yongma[yongma["sales_code_key"].isin(domestic_code_keys)].copy()
        yongma["yongma_lot_key"] = yongma["yongma_lot_key"].map(clean_str)
        code_meta = code_summary.copy()
        code_meta["sales_code_key"] = code_meta.get("sales_code_key", pd.Series("", index=code_meta.index)).map(clean_str)
        code_sales_by_key = build_first_value_map(code_meta, "sales_code_key", "sales_code")
        code_product_by_key = build_first_value_map(code_meta, "sales_code_key", "product_name")
        receipt_only_rows: list[dict[str, Any]] = []

        def add_receipt_to_indices(indices: list[int], qty: float) -> None:
            remaining = float(qty)
            for idx in indices:
                if remaining <= 0:
                    break
                packed = pd.to_numeric(grouped.at[idx, "packing_pack"], errors="coerce")
                received = pd.to_numeric(grouped.at[idx, "용마입고수량"], errors="coerce")
                capacity = max(0.0, float(packed if not pd.isna(packed) else 0.0) - float(received if not pd.isna(received) else 0.0))
                add_qty = min(remaining, capacity) if capacity > 0 else 0.0
                if add_qty > 0:
                    grouped.at[idx, "용마입고수량"] += add_qty
                    remaining -= add_qty
            if remaining > 0 and indices:
                grouped.at[indices[0], "용마입고수량"] += remaining

        for _, receipt in yongma.iterrows():
            code_key = clean_str(receipt.get("sales_code_key", ""))
            lot_key = clean_str(receipt.get("yongma_lot_key", ""))
            qty_value = pd.to_numeric(receipt.get("yongma_in_pack", 0.0), errors="coerce")
            qty = 0.0 if pd.isna(qty_value) else float(qty_value)
            if not code_key or not lot_key or qty <= 0:
                continue

            candidates = grouped[grouped["sales_code_key"] == code_key]
            if candidates.empty:
                receipt_only_rows.append(
                    {
                        "sales_code_key": code_key,
                        "sales_code": clean_str(receipt.get("sales_code", "")) or code_sales_by_key.get(code_key, ""),
                        "packing_product_name": clean_str(receipt.get("yongma_product_name", ""))
                        or code_product_by_key.get(code_key, ""),
                        "packing_lot": clean_str(receipt.get("yongma_lot", "")) or "(용마 LOT 미기재)",
                        "packing_lot_key": lot_key,
                        "packing_barcode": "",
                        "packing_barcode_key": "",
                        "packing_pack": 0.0,
                        "packing_date": pd.NaT,
                        "용마입고수량": qty,
                    }
                )
                continue
            exact = candidates[candidates["packing_lot_key"] == lot_key]
            target = exact
            if target.empty:
                barcode_match = candidates[
                    candidates["packing_barcode_key"].astype(str).str.contains(lot_key, regex=False, na=False)
                ]
                target = barcode_match
            if target.empty:
                target = candidates.sort_values(["packing_date", "packing_lot"], na_position="last", kind="stable")
                add_receipt_to_indices(target.index.tolist(), qty)
                continue
            add_receipt_to_indices(target.index.tolist(), qty)

        if receipt_only_rows:
            receipt_only = (
                pd.DataFrame(receipt_only_rows)
                .groupby(
                    [
                        "sales_code_key",
                        "sales_code",
                        "packing_product_name",
                        "packing_lot",
                        "packing_lot_key",
                        "packing_barcode",
                        "packing_barcode_key",
                    ],
                    dropna=False,
                )
                .agg(
                    packing_pack=("packing_pack", "sum"),
                    packing_date=("packing_date", min_datetime),
                    용마입고수량=("용마입고수량", "sum"),
                )
                .reset_index()
            )
            grouped = pd.concat([grouped, receipt_only], ignore_index=True)

    grouped["포장실적수량"] = pd.to_numeric(grouped["packing_pack"], errors="coerce").fillna(0.0)
    grouped["용마입고수량"] = pd.to_numeric(grouped["용마입고수량"], errors="coerce").fillna(0.0)
    grouped["입고대기수량"] = (grouped["포장실적수량"] - grouped["용마입고수량"]).clip(lower=0.0)
    grouped["상태"] = np.select(
        [
            (grouped["포장실적수량"] <= 0) & (grouped["용마입고수량"] > 0),
            (grouped["입고대기수량"] > 0) & (grouped["용마입고수량"] > 0),
            grouped["입고대기수량"] > 0,
        ],
        ["용마입고만", "부분입고", "입고대기"],
        default="입고완료",
    )
    grouped["_status_sort"] = grouped["상태"].map({"입고대기": 0, "부분입고": 1, "용마입고만": 2}).fillna(3)
    grouped["포장일"] = grouped["packing_date"].map(display_date_or_dash)
    grouped = grouped.rename(
        columns={
            "sales_code": "제품코드",
            "packing_product_name": "제품명",
            "packing_lot": "LOTNO",
        }
    )
    grouped["제품명"] = grouped["제품명"].replace("", "(미기재)")
    return grouped.sort_values(
        ["_status_sort", "입고대기수량", "포장실적수량"],
        ascending=[True, False, False],
        kind="stable",
    )[columns].copy()


def render_packing_lot_tab(lot_status_df: pd.DataFrame) -> None:
    render_panel_title(
        "세부 포장 진도 현황",
        "LOT 기준 포장실적과 용마입고수량을 비교합니다.",
    )
    if lot_status_df.empty:
        st.warning("표시할 포장 LOT 데이터가 없습니다.")
        return

    f1, f2 = st.columns([2.4, 1.2], gap="small")
    with f1:
        query = st.text_input(
            "제품코드/제품명/LOTNO 검색",
            value="",
            key="packing_lot_query",
        )
    with f2:
        statuses = st.multiselect(
            "상태 필터",
            ["입고대기", "입고완료"],
            default=["입고대기", "입고완료"],
            key="packing_lot_status",
        )

    view = lot_status_df.copy()
    if query.strip():
        q = query.strip()
        mask = (
            view["제품코드"].astype(str).str.contains(q, case=False, na=False)
            | view["제품명"].astype(str).str.contains(q, case=False, na=False)
            | view["LOTNO"].astype(str).str.contains(q, case=False, na=False)
        )
        view = view[mask].copy()
    if statuses:
        view = view[view["상태"].isin(statuses)].copy()

    waiting_qty = float(view["입고대기수량"].sum()) if not view.empty else 0.0
    render_metric_card_grid(
        [
            ("포장실적수량", format_int(float(view["포장실적수량"].sum()) if not view.empty else 0.0), "normal"),
            ("용마입고수량", format_int(float(view["용마입고수량"].sum()) if not view.empty else 0.0), "normal"),
            ("입고대기수량", format_int(waiting_qty), "warn" if waiting_qty > 0 else "normal"),
        ]
    )
    dl_col, _ = st.columns([1.2, 4.8], gap="small")
    with dl_col:
        render_excel_download(
            "엑셀 다운로드",
            "포장_LOT_상세",
            {"포장 LOT 상세": view},
            key="download_packing_lot_excel",
        )

    st.dataframe(
        view,
        hide_index=True,
        height=620,
        width="stretch",
        column_config=drilldown_column_config(),
    )


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
    work = work[work["production_code_display"].map(is_p_production_code)].copy()
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
    pack_label: str,
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
    if pack_label != "전체":
        out = out[out["_pack_label"] == pack_label]
    if sample_scope == "본품":
        out = out[out["본품/샘플"] == "본품"]
    elif sample_scope == "샘플":
        out = out[out["본품/샘플"] == "샘플"]
    if product_group != "전체":
        out = out[out["제품분류"] == product_group]
    return out.copy()


def is_p_production_code(value: Any) -> bool:
    return clean_str(value).upper().startswith("P")


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


def production_code_prefix(value: Any) -> str:
    text = clean_str(value)
    if not text or text == "(생산코드 미기재)":
        return "(생산코드 미기재)"
    return text[:5].upper()


def build_production_power_main_view(
    rows: pd.DataFrame,
    pack_labels: list[str],
    shortage_only: bool = False,
) -> pd.DataFrame:
    visible_columns = [
        "생산코드",
        "대표 제품명",
        *pack_labels,
        "요청합계(PACK)",
        "포장부족(PACK)",
        "생산부족수량(PCS)",
        "포장진도율",
        "생산진도율",
        "최소 납기일",
    ]
    if rows.empty:
        return pd.DataFrame(
            columns=visible_columns
            + [
                "요청합계(PCS)",
                "생산부족수량",
                "포장부족수량",
                "병목 상태",
                "_production_code_prefix",
                "_min_due_date_sort",
            ]
        )

    work = rows.copy()
    work["_production_code_prefix"] = work["production_code_display"].map(production_code_prefix)
    group_cols = ["_production_code_prefix"]
    base = (
        work.groupby(group_cols, dropna=False)
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
        work.pivot_table(
            index=group_cols,
            columns="_pack_label",
            values="request_pack",
            aggfunc="sum",
            dropna=True,
        )
        .fillna(0.0)
        .reset_index()
        .rename_axis(None, axis=1)
    )
    for label in pack_labels:
        if label not in pack_pivot.columns:
            pack_pivot[label] = 0.0

    grouped = base.merge(pack_pivot[group_cols + pack_labels], on=group_cols, how="left")
    for label in pack_labels:
        grouped[label] = grouped[label].fillna(0.0)
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
            "_production_code_prefix": "생산코드",
            "representative_product": "대표 제품명",
            "request_pack": "요청합계(PACK)",
            "request_pcs": "요청합계(PCS)",
            "production_shortage_pcs": "생산부족수량",
            "packing_shortage_pack": "포장부족수량",
        }
    )
    out["생산부족수량(PCS)"] = out["생산부족수량"]
    out["포장부족(PACK)"] = out["포장부족수량"]
    out["_production_code_prefix"] = out["생산코드"]
    if shortage_only:
        out = out[(out["생산부족수량"] > 0) | (out["포장부족수량"] > 0)].copy()

    out = out.sort_values(
        ["_min_due_date_sort", "포장부족수량", "생산부족수량"],
        ascending=[True, False, False],
        na_position="last",
        kind="stable",
    )
    return_columns = list(
        dict.fromkeys(
            visible_columns
            + [
                "요청합계(PCS)",
                "생산부족수량",
                "포장부족수량",
                "병목 상태",
                "_production_code_prefix",
                "_min_due_date_sort",
            ]
        )
    )
    return out[return_columns].copy()


def build_production_power_detail_view(
    rows: pd.DataFrame,
    pack_labels: list[str],
    production_prefix: str | None = None,
) -> pd.DataFrame:
    visible_columns = [
        "생산코드 전체",
        "대표 제품명",
        "POWER",
        *pack_labels,
        "요청합계(PACK)",
        "포장부족(PACK)",
        "생산부족수량(PCS)",
        "포장진도율",
        "생산진도율",
        "최소 납기일",
    ]
    if rows.empty:
        return pd.DataFrame(
            columns=visible_columns
            + [
                "요청합계(PCS)",
                "생산부족수량",
                "포장부족수량",
                "_production_code_prefix",
                "_min_due_date_sort",
                "_power_sort",
            ]
        )

    work = rows.copy()
    work["_production_code_prefix"] = work["production_code_display"].map(production_code_prefix)
    if production_prefix is not None:
        work = work[work["_production_code_prefix"] == production_prefix].copy()
    if work.empty:
        return pd.DataFrame(
            columns=visible_columns
            + ["요청합계(PCS)", "생산부족수량", "포장부족수량", "_production_code_prefix", "_min_due_date_sort", "_power_sort"]
        )

    group_cols = ["_production_code_prefix", "production_code_display", "POWER", "_power_sort"]
    base = (
        work.groupby(group_cols, dropna=False)
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
        work.pivot_table(
            index=group_cols,
            columns="_pack_label",
            values="request_pack",
            aggfunc="sum",
            dropna=True,
        )
        .fillna(0.0)
        .reset_index()
        .rename_axis(None, axis=1)
    )
    for label in pack_labels:
        if label not in pack_pivot.columns:
            pack_pivot[label] = 0.0

    grouped = base.merge(pack_pivot[group_cols + pack_labels], on=group_cols, how="left")
    for label in pack_labels:
        grouped[label] = grouped[label].fillna(0.0)
    grouped["생산진도율"] = calc_production_progress_pct(grouped["request_pcs"], grouped["production_shortage_pcs"])
    grouped["포장진도율"] = np.where(
        grouped["request_pack"] > 0,
        grouped["packing_pack"] / grouped["request_pack"] * 100.0,
        0.0,
    )
    grouped["포장진도율"] = np.clip(grouped["포장진도율"], 0.0, 100.0)
    grouped["상태"] = [
        status_from_progress(packing, progress)
        for packing, progress in zip(grouped["packing_pack"], grouped["포장진도율"])
    ]
    grouped["_min_due_date_sort"] = pd.to_datetime(grouped["min_due_date"], errors="coerce")
    grouped["최소 납기일"] = grouped["min_due_date"].map(display_date_or_dash)

    out = grouped.rename(
        columns={
            "production_code_display": "생산코드 전체",
            "representative_product": "대표 제품명",
            "request_pack": "요청합계(PACK)",
            "request_pcs": "요청합계(PCS)",
            "production_shortage_pcs": "생산부족수량",
            "packing_shortage_pack": "포장부족수량",
        }
    )
    out["생산부족수량(PCS)"] = out["생산부족수량"]
    out["포장부족(PACK)"] = out["포장부족수량"]
    out = out.sort_values(
        ["_min_due_date_sort", "포장부족수량", "생산부족수량"],
        ascending=[True, False, False],
        na_position="last",
        kind="stable",
    )
    return_columns = list(
        dict.fromkeys(
            visible_columns
            + [
                "요청합계(PCS)",
                "생산부족수량",
                "포장부족수량",
                "_production_code_prefix",
                "_min_due_date_sort",
                "_power_sort",
            ]
        )
    )
    return out[return_columns].copy()


def calc_production_power_kpis(view: pd.DataFrame) -> dict[str, float]:
    if view.empty:
        return {
            "production_code_count": 0.0,
            "request_pack": 0.0,
            "request_pcs": 0.0,
            "production_shortage_pcs": 0.0,
            "packing_shortage_pack": 0.0,
            "production_progress_pct": 0.0,
            "packing_progress_pct": 0.0,
            "production_bottleneck_count": 0.0,
            "packing_bottleneck_count": 0.0,
        }
    request_pack = float(view["요청합계(PACK)"].sum())
    request_pcs = float(view["요청합계(PCS)"].sum())
    production_shortage_pcs = float(view["생산부족수량"].sum())
    packing_shortage_pack = float(view["포장부족수량"].sum())
    production_progress = (
        (request_pcs - production_shortage_pcs) / request_pcs * 100.0
        if request_pcs > 0
        else 0.0
    )
    packing_done_pack = max(0.0, request_pack - packing_shortage_pack)
    packing_progress = (packing_done_pack / request_pack * 100.0) if request_pack > 0 else 0.0
    return {
        "production_code_count": float(view["생산코드"].nunique()),
        "request_pack": request_pack,
        "request_pcs": request_pcs,
        "production_shortage_pcs": production_shortage_pcs,
        "packing_shortage_pack": packing_shortage_pack,
        "production_progress_pct": min(100.0, max(0.0, production_progress)),
        "packing_progress_pct": min(100.0, max(0.0, packing_progress)),
        "production_bottleneck_count": float(view["병목 상태"].astype(str).str.contains("생산 병목", na=False).sum()),
        "packing_bottleneck_count": float(view["병목 상태"].astype(str).str.contains("포장 병목", na=False).sum()),
    }


def render_production_power_kpis(view: pd.DataFrame, unit_mode: str = UNIT_PACK) -> None:
    kpi = calc_production_power_kpis(view)
    if unit_mode == UNIT_PCS:
        items = [
            ("생산코드 수", f"{int(kpi['production_code_count']):,}", "normal"),
            ("총 요청 PCS", format_int(kpi["request_pcs"]), "normal"),
            ("총 생산부족수량(PCS)", format_int(kpi["production_shortage_pcs"]), "risk"),
            ("생산진도율", f"{kpi['production_progress_pct']:.1f}%", "normal"),
            ("포장진도율", f"{kpi['packing_progress_pct']:.1f}%", "normal"),
        ]
    else:
        items = [
            ("생산코드 수", f"{int(kpi['production_code_count']):,}", "normal"),
            ("총 요청 PACK", format_int(kpi["request_pack"]), "normal"),
            ("총 포장부족(PACK)", format_int(kpi["packing_shortage_pack"]), "warn"),
            ("총 생산부족수량(PCS)", format_int(kpi["production_shortage_pcs"]), "risk"),
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


def render_pack_composition_chart(selected_row: pd.Series, pack_labels: list[str]) -> None:
    chart_df = pd.DataFrame(
        {
            "PACK": pack_labels,
            "필요팩": [float(selected_row.get(label, 0.0)) for label in pack_labels],
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
        "생산필요수량(PCS)",
        "생산부족수량(PCS)",
        "생산진도율",
        "포장진도율",
        "납기일자",
    ]
    if scope.empty:
        return pd.DataFrame(columns=columns + ["_pack_sort"])

    grouped = (
        scope.groupby(["sales_code", "product_name", "_pack_label", "_pack_sort"], dropna=False)
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
                "_pack_label": "PACK 단위",
                "request_pack": "필요팩",
                "request_pcs": "요청PCS",
                "packing_pack": "포장완료PACK",
            }
        )
    )
    grouped["포장부족PACK"] = (grouped["필요팩"] - grouped["포장완료PACK"]).clip(lower=0.0)
    grouped["생산필요수량(PCS)"] = grouped["production_shortage_pcs"]
    grouped["생산부족수량(PCS)"] = grouped["production_shortage_pcs"]
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


def build_sales_order_main_view(
    code_summary: pd.DataFrame,
    stock_threshold_pack: float = INVENTORY_STOCK_THRESHOLD_DEFAULT,
) -> pd.DataFrame:
    if code_summary.empty:
        return pd.DataFrame(
            columns=[
                "우선등급",
                "D-Day",
                "판매코드",
                "생산코드",
                "제품명",
                "PACK",
                "POWER",
                "요청PACK",
                "요청PCS",
                "생산요청물량",
                "생산요청물량(PACK)",
                "생산요청물량(PCS)",
                "용마입고수량",
                "용마입고수량(PACK)",
                "용마입고수량(PCS)",
                "용마입고대기수량",
                "용마입고대기수량(PACK)",
                "용마입고대기수량(PCS)",
                "포장가능재고(PCS)",
                "용마창고재고 (PACK)",
                "재고기준(PACK)",
                "재고부족(PACK)",
                "생산부족",
                "생산부족(PCS)",
                "포장부족",
                "포장부족(PCS)",
                "생산진도율",
                "용마입고율",
                "납기",
                "상태",
            ]
        )
    work = with_operational_columns(code_summary)
    pack_unit = pd.to_numeric(work.get("pack_unit", pd.Series(np.nan, index=work.index)), errors="coerce")
    request_pack = pd.to_numeric(work.get("request_pack", pd.Series(0.0, index=work.index)), errors="coerce").fillna(0.0)
    request_pcs = pd.to_numeric(work.get("request_pcs", pd.Series(0.0, index=work.index)), errors="coerce").fillna(0.0)
    implied_unit = np.where(request_pack > 0, request_pcs / request_pack, np.nan)
    pcs_per_pack = pack_unit.where(pack_unit > 0, implied_unit)
    pcs_per_pack = pd.Series(pcs_per_pack, index=work.index).replace([np.inf, -np.inf], np.nan).fillna(1.0)
    pcs_per_pack = pcs_per_pack.where(pcs_per_pack > 0, 1.0)
    yongma_in_pack = pd.to_numeric(work.get("yongma_in_pack", pd.Series(0.0, index=work.index)), errors="coerce").fillna(0.0)
    packing_pack = pd.to_numeric(work.get("packing_pack", pd.Series(0.0, index=work.index)), errors="coerce").fillna(0.0)
    work["_yongma_in_pcs"] = (yongma_in_pack * pcs_per_pack).clip(lower=0.0)
    work["_yongma_wait_pcs"] = ((packing_pack - yongma_in_pack).clip(lower=0.0) * pcs_per_pack).clip(lower=0.0)
    work["_packing_shortage_pcs"] = ((request_pack - yongma_in_pack).clip(lower=0.0) * pcs_per_pack).clip(lower=0.0)
    grouped = (
        work.groupby("sales_code", dropna=False)
        .agg(
            production_code=("production_code_display", join_unique),
            product_name=("product_name", join_unique),
            pack_label=("_pack_label", join_unique),
            power=("POWER", first_nonempty),
            power_value=("power_value", "min"),
            request_pack=("request_pack", "sum"),
            request_pcs=("request_pcs", "sum"),
            packing_pack=("packing_pack", "sum"),
            yongma_in_pack=("yongma_in_pack", "sum"),
            yongma_in_pcs=("_yongma_in_pcs", "sum"),
            yongma_wait_pcs=("_yongma_wait_pcs", "sum"),
            packing_shortage_pcs=("_packing_shortage_pcs", "sum"),
            available_stock_pack=("available_stock_pack", sum_numeric_or_nan),
            production_shortage=("_allocated_production_shortage_qty", "sum"),
            sample_available_pcs=("_allocated_sample_available_pcs", "sum"),
            request_due_date=("request_due_date", min_datetime),
        )
        .reset_index()
        .rename(
            columns={
                "sales_code": "판매코드",
                "production_code": "생산코드",
                "product_name": "제품명",
                "pack_label": "PACK",
                "power": "POWER",
                "request_pack": "요청PACK",
                "request_pcs": "요청PCS",
                "yongma_in_pack": "용마입고수량",
                "yongma_in_pcs": "용마입고수량(PCS)",
                "yongma_wait_pcs": "용마입고대기수량(PCS)",
                "packing_shortage_pcs": "포장부족(PCS)",
                "available_stock_pack": "용마창고재고 (PACK)",
                "production_shortage": "생산부족",
                "sample_available_pcs": "샘플신청가능수량",
            }
        )
    )
    grouped["용마입고대기수량"] = (grouped["packing_pack"] - grouped["용마입고수량"]).clip(lower=0.0)
    grouped["용마입고수량(PACK)"] = grouped["용마입고수량"]
    grouped["용마입고대기수량(PACK)"] = grouped["용마입고대기수량"]
    grouped["포장가능재고(PCS)"] = (
        grouped["요청PCS"] - grouped["생산부족"] + grouped["샘플신청가능수량"]
    ).clip(lower=0.0)
    grouped["포장부족"] = (grouped["요청PACK"] - grouped["용마입고수량"]).clip(lower=0.0)
    grouped["생산진도율"] = calc_production_progress_pct(grouped["요청PCS"], grouped["생산부족"])
    grouped["용마입고율"] = np.where(
        grouped["요청PACK"] > 0,
        grouped["용마입고수량"] / grouped["요청PACK"] * 100.0,
        0.0,
    )
    grouped["용마입고율"] = np.clip(grouped["용마입고율"], 0.0, 100.0)
    grouped["납기"] = grouped["request_due_date"].map(display_date_or_dash)
    grouped["상태"] = grouped.apply(sales_status_label, axis=1)
    grouped = add_priority_columns(
        grouped,
        stock_threshold_pack,
        shortage_col="포장부족",
        due_col="request_due_date",
        stock_col="용마창고재고 (PACK)",
        request_col="요청PACK",
    )
    grouped["요청합계(PACK)"] = grouped["요청PACK"]
    grouped["요청합계(PCS)"] = grouped["요청PCS"]
    grouped["생산요청물량"] = grouped["요청PCS"]
    grouped["생산요청물량(PACK)"] = grouped["요청PACK"]
    grouped["생산요청물량(PCS)"] = grouped["요청PCS"]
    grouped["생산필요수량(PCS)"] = grouped["생산부족"]
    grouped["생산부족수량(PCS)"] = grouped["생산부족"]
    grouped["생산부족(PCS)"] = grouped["생산부족"]
    grouped["포장부족(PACK)"] = grouped["포장부족"]
    return grouped[
        [
            "우선등급",
            "D-Day",
            "판매코드",
            "생산코드",
            "제품명",
            "PACK",
            "POWER",
            "요청PACK",
            "요청PCS",
            "요청합계(PACK)",
            "요청합계(PCS)",
            "생산요청물량",
            "생산요청물량(PACK)",
            "생산요청물량(PCS)",
            "용마입고수량",
            "용마입고수량(PACK)",
            "용마입고수량(PCS)",
            "용마입고대기수량",
            "용마입고대기수량(PACK)",
            "용마입고대기수량(PCS)",
            "포장가능재고(PCS)",
            "샘플신청가능수량",
            "용마창고재고 (PACK)",
            "재고기준(PACK)",
            "재고부족(PACK)",
            "생산부족",
            "생산필요수량(PCS)",
            "생산부족수량(PCS)",
            "생산부족(PCS)",
            "포장부족",
            "포장부족(PACK)",
            "포장부족(PCS)",
            "생산진도율",
            "용마입고율",
            "납기",
            "상태",
            "power_value",
            "_priority_sort",
            "_request_due_date_sort",
        ]
    ].sort_values(
        ["_priority_sort", "_request_due_date_sort", "재고부족(PACK)", "포장부족", "생산부족"],
        ascending=[True, True, False, False, False],
        na_position="last",
        kind="stable",
    )


def build_urgent_sales_packing_view(sales_view: pd.DataFrame, max_rows: int = 20) -> pd.DataFrame:
    columns = [
        "우선등급",
        "D-Day",
        "판매코드",
        "제품명",
        "POWER",
        "PACK",
        "생산요청물량(PACK)",
        "포장부족(PACK)",
        "납기",
    ]
    if sales_view.empty:
        return pd.DataFrame(columns=columns)

    out = sales_view[
        (pd.to_numeric(sales_view["포장부족"], errors="coerce").fillna(0.0) > 0)
        & (sales_view["우선등급"].isin(["A 긴급", "B 주의"]))
    ].copy()
    if out.empty:
        return pd.DataFrame(columns=columns)

    out = out.sort_values(
        ["_priority_sort", "_request_due_date_sort", "재고부족(PACK)", "포장부족", "생산부족"],
        ascending=[True, True, False, False, False],
        na_position="last",
        kind="stable",
    )
    return out[columns].head(max_rows).copy()


def render_urgent_sales_packing_list(sales_view: pd.DataFrame) -> None:
    urgent_view = build_urgent_sales_packing_view(sales_view)
    render_panel_title(
        "긴급 포장 리스트",
        "용마 보유 재고는 긴급도 판단에만 사용하고, 표에는 PACK 기준 요청·부족 수량만 표시합니다.",
    )
    st.markdown("<div class='panel-box drill-panel'>", unsafe_allow_html=True)
    if urgent_view.empty:
        st.info("현재 기준에 해당하는 긴급 포장 판매코드가 없습니다.")
    else:
        st.dataframe(
            urgent_view,
            hide_index=True,
            height=260,
            width="stretch",
            column_config=drilldown_column_config(),
        )
    st.markdown("</div>", unsafe_allow_html=True)


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
        return pd.DataFrame(
            columns=[
                "POWER",
                "요청합계(PACK)",
                "요청합계(PCS)",
                "포장 PACK",
                "포장부족(PACK)",
                "생산필요수량(PCS)",
                "생산부족수량(PCS)",
                "생산진도율",
                "포장진도율",
                "power_value",
            ]
        )
    work = add_allocated_production_basis(work)
    grouped = (
        work.groupby(["power_value", "POWER"], dropna=False)
        .agg(
            request_pack=("request_pack", "sum"),
            request_pcs=("request_pcs", "sum"),
            packing_pack=("packing_pack", "sum"),
            production_shortage_pcs=("_allocated_production_shortage_qty", "sum"),
        )
        .reset_index()
        .rename(
            columns={
                "request_pack": "요청합계(PACK)",
                "request_pcs": "요청합계(PCS)",
                "packing_pack": "포장 PACK",
                "production_shortage_pcs": "생산부족수량(PCS)",
            }
        )
    )
    grouped["포장부족(PACK)"] = (grouped["요청합계(PACK)"] - grouped["포장 PACK"]).clip(lower=0.0)
    grouped["생산필요수량(PCS)"] = grouped["생산부족수량(PCS)"]
    grouped["생산진도율"] = calc_production_progress_pct(grouped["요청합계(PCS)"], grouped["생산부족수량(PCS)"])
    grouped["포장진도율"] = np.where(
        grouped["요청합계(PACK)"] > 0,
        grouped["포장 PACK"] / grouped["요청합계(PACK)"] * 100.0,
        0.0,
    )
    grouped["포장진도율"] = np.clip(grouped["포장진도율"], 0.0, 100.0)
    return grouped[
        [
            "POWER",
            "요청합계(PACK)",
            "요청합계(PCS)",
            "포장 PACK",
            "포장부족(PACK)",
            "생산필요수량(PCS)",
            "생산부족수량(PCS)",
            "생산진도율",
            "포장진도율",
            "power_value",
        ]
    ].sort_values("power_value", ascending=True, kind="stable")


def build_power_sku_detail_view(code_summary: pd.DataFrame, power_label: str) -> pd.DataFrame:
    work = add_allocated_production_basis(with_operational_columns(code_summary))
    scope = work[work["POWER"] == power_label].copy()
    if scope.empty:
        return pd.DataFrame(
            columns=[
                "생산코드",
                "판매코드",
                "제품명",
                "PACK",
                "요청합계(PACK)",
                "요청합계(PCS)",
                "포장부족(PACK)",
                "생산필요수량(PCS)",
                "생산부족수량(PCS)",
                "납기",
            ]
        )
    out = (
        scope.groupby(["production_code_display", "sales_code", "product_name", "_pack_label"], dropna=False)
        .agg(
            request_pack=("request_pack", "sum"),
            request_pcs=("request_pcs", "sum"),
            packing_pack=("packing_pack", "sum"),
            production_shortage_pcs=("_allocated_production_shortage_qty", "sum"),
            request_due_date=("request_due_date", min_datetime),
        )
        .reset_index()
        .rename(
            columns={
                "production_code_display": "생산코드",
                "sales_code": "판매코드",
                "product_name": "제품명",
                "_pack_label": "PACK",
                "request_pack": "요청합계(PACK)",
                "request_pcs": "요청합계(PCS)",
                "production_shortage_pcs": "생산부족수량(PCS)",
            }
        )
    )
    out["포장부족(PACK)"] = (out["요청합계(PACK)"] - out["packing_pack"]).clip(lower=0.0)
    out["생산필요수량(PCS)"] = out["생산부족수량(PCS)"]
    out["납기"] = out["request_due_date"].map(display_date_or_dash)
    return out[
        [
            "생산코드",
            "판매코드",
            "제품명",
            "PACK",
            "요청합계(PACK)",
            "요청합계(PCS)",
            "포장부족(PACK)",
            "생산필요수량(PCS)",
            "생산부족수량(PCS)",
            "납기",
        ]
    ].sort_values(
        ["포장부족(PACK)", "요청합계(PACK)"], ascending=[False, False], kind="stable"
    )


def empty_inventory_detail_view() -> pd.DataFrame:
    return pd.DataFrame(columns=["판매코드", "WMS제품명", "용마창고재고 (PACK)", "총수량(PACK)", "제품규격", "전송일자", "매칭여부"])


def build_inventory_detail_view(code_summary: pd.DataFrame, sales_code: str) -> pd.DataFrame:
    if code_summary.empty:
        return empty_inventory_detail_view()
    work = with_operational_columns(code_summary)
    scope = work[work["sales_code"].astype(str) == str(sales_code)].copy()
    if scope.empty:
        return empty_inventory_detail_view()

    grouped = (
        scope.groupby("sales_code", dropna=False)
        .agg(
            inventory_product_name=("inventory_product_name", first_nonempty),
            available_stock_pack=("available_stock_pack", sum_numeric_or_nan),
            inventory_total_stock_pack=("inventory_total_stock_pack", sum_numeric_or_nan),
            inventory_product_spec=("inventory_product_spec", first_nonempty),
            inventory_updated_at=("inventory_updated_at", max_datetime),
            inventory_matched=("inventory_matched", "max"),
        )
        .reset_index()
        .rename(
            columns={
                "sales_code": "판매코드",
                "inventory_product_name": "WMS제품명",
                "available_stock_pack": "용마창고재고 (PACK)",
                "inventory_total_stock_pack": "총수량(PACK)",
                "inventory_product_spec": "제품규격",
                "inventory_updated_at": "전송일자",
                "inventory_matched": "매칭여부",
            }
        )
    )
    grouped["전송일자"] = grouped["전송일자"].map(display_date_or_dash)
    grouped["매칭여부"] = np.where(grouped["매칭여부"], "매칭", "미매칭")
    return grouped[["판매코드", "WMS제품명", "용마창고재고 (PACK)", "총수량(PACK)", "제품규격", "전송일자", "매칭여부"]]


def ppt_rgb(hex_color: str) -> RGBColor:
    return RGBColor.from_string(hex_color.replace("#", ""))


def apply_ppt_font(
    run: Any,
    size: int | float,
    bold: bool = False,
    color: str = TEXT_DARK,
) -> None:
    run.font.name = PPT_FONT_NAME
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = ppt_rgb(color)

    r_pr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        font_el = r_pr.find(qn(tag))
        if font_el is None:
            font_el = OxmlElement(tag)
            r_pr.append(font_el)
        font_el.set("typeface", PPT_FONT_NAME)


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
    cell.text_frame.word_wrap = True
    paragraph = cell.text_frame.paragraphs[0]
    paragraph.alignment = align
    paragraph.space_after = Pt(0)
    paragraph.space_before = Pt(0)
    run = paragraph.add_run()
    run.text = text
    apply_ppt_font(run, size=size, bold=bold, color=color)


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
    vertical_anchor: Any | None = None,
) -> None:
    shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = shape.text_frame
    frame.clear()
    frame.margin_left = 0
    frame.margin_right = 0
    frame.margin_top = 0
    frame.margin_bottom = 0
    frame.word_wrap = True
    if vertical_anchor is not None:
        frame.vertical_anchor = vertical_anchor
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    paragraph.space_after = Pt(0)
    paragraph.space_before = Pt(0)
    run = paragraph.add_run()
    run.text = text
    apply_ppt_font(run, size=size, bold=bold, color=color)


def truncate_report_text(value: Any, max_chars: int = 34) -> str:
    text = clean_str(value)
    if len(text) <= max_chars:
        return text
    return text[: max_chars - 1].rstrip() + "..."


def add_report_rule(slide: Any, left: float, top: float, width: float, color: str = MID_GRAY) -> None:
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(0.01))
    shape.fill.solid()
    shape.fill.fore_color.rgb = ppt_rgb(color)
    shape.line.fill.background()


def add_report_shape(
    slide: Any,
    shape_type: Any,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color: str,
    line_color: str | None = None,
    line_width: float = 0.5,
) -> Any:
    shape = slide.shapes.add_shape(shape_type, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = ppt_rgb(fill_color)
    if line_color:
        shape.line.color.rgb = ppt_rgb(line_color)
        shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape


def report_progress_color(value: Any) -> str:
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return REPORT_HEADER
    return REPORT_HEADER


def add_kpi_card(
    slide: Any,
    title: str,
    kpi: dict[str, float],
    dot_color: str,
    left: float,
    top: float,
    width: float,
    height: float,
) -> None:
    add_report_shape(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left,
        top,
        width,
        height,
        REPORT_PANEL,
        REPORT_PANEL_LINE,
        0.5,
    )
    add_report_shape(slide, MSO_SHAPE.RECTANGLE, left + 0.12, top + 0.08, 0.46, 0.045, dot_color, dot_color)
    add_textbox(
        slide,
        title,
        left + 0.14,
        top + 0.18,
        width - 0.28,
        0.18,
        8.4,
        True,
        REPORT_HEADER,
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )

    receipt_progress = kpi.get("progress_pct", 0.0)
    packing_progress = kpi.get("packing_progress_pct", 0.0)
    production_progress = kpi.get("production_progress_pct", 0.0)
    progress_width = max(0.0, min(1.0, to_report_float(receipt_progress) / 100.0)) * (width - 0.3)
    packing_width = max(0.0, min(1.0, to_report_float(packing_progress) / 100.0)) * (width - 0.3)
    add_textbox(
        slide,
        format_report_value(receipt_progress, True),
        left + 0.14,
        top + 0.4,
        1.28,
        0.36,
        20,
        True,
        REPORT_HEADER,
        PP_ALIGN.LEFT,
        MSO_ANCHOR.MIDDLE,
    )
    add_textbox(
        slide,
        "용마입고율",
        left + width - 1.3,
        top + 0.52,
        0.9,
        0.2,
        7,
        False,
        REPORT_MUTED,
        PP_ALIGN.RIGHT,
        MSO_ANCHOR.MIDDLE,
    )
    add_report_shape(slide, MSO_SHAPE.RECTANGLE, left + 0.14, top + 0.83, width - 0.3, 0.06, REPORT_FAINT)
    if packing_width > 0:
        add_report_shape(slide, MSO_SHAPE.RECTANGLE, left + 0.14, top + 0.83, packing_width, 0.06, REPORT_ACCENT_SOFT, REPORT_ACCENT_SOFT)
    if progress_width > 0:
        add_report_shape(slide, MSO_SHAPE.RECTANGLE, left + 0.14, top + 0.83, progress_width, 0.06, dot_color, dot_color)

    metrics = [
        ("요청", format_report_value(kpi.get("request_pack", 0.0)), REPORT_HEADER, True),
        ("생산율", format_report_value(production_progress, True), REPORT_HEADER, True),
        ("포장률", format_report_value(packing_progress, True), REPORT_HEADER, True),
        ("입고율", format_report_value(receipt_progress, True), REPORT_HEADER, False),
        ("미입고", format_report_value(kpi.get("shortage_pack", 0.0)), REPORT_ACCENT, False),
    ]
    col_width = (width - 0.28) / len(metrics)
    for idx, (label, value, value_color, big) in enumerate(metrics):
        metric_left = left + 0.14 + idx * col_width
        add_textbox(
            slide,
            label,
            metric_left,
            top + 1.0,
            col_width,
            0.14,
            6.6,
            False,
            REPORT_MUTED,
            PP_ALIGN.LEFT,
            MSO_ANCHOR.MIDDLE,
        )
        add_textbox(
            slide,
            value,
            metric_left,
            top + 1.15,
            col_width,
            0.2,
            8.9 if big else 8.3,
            True,
            value_color,
            PP_ALIGN.LEFT,
            MSO_ANCHOR.MIDDLE,
        )


def format_report_value(value: Any, is_percent: bool = False) -> str:
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return "0.0%" if is_percent else "0"
    return f"{float(num):.1f}%" if is_percent else format_int(float(num))


def sanitize_excel_sheet_name(name: str) -> str:
    cleaned = re.sub(r"[\[\]\:\*\?\/\\]", "_", clean_str(name))
    return (cleaned or "Sheet")[:31]


def dataframe_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    drop_cols = [
        col
        for col in out.columns
        if str(col).startswith("_") or str(col) in {"power_value"}
    ]
    if drop_cols:
        out = out.drop(columns=drop_cols, errors="ignore")
    for col in out.columns:
        if pd.api.types.is_datetime64tz_dtype(out[col]):
            out[col] = out[col].dt.tz_localize(None)
    return out


def excel_text_length(value: Any) -> int:
    if value is None:
        return 0
    try:
        if pd.isna(value):
            return 0
    except (TypeError, ValueError):
        pass
    return len(str(value))


def make_unique_excel_columns(columns: pd.Index) -> list[str]:
    used_counts: dict[str, int] = {}
    unique_columns: list[str] = []
    for idx, column in enumerate(columns, start=1):
        base_name = clean_str(column) or f"컬럼{idx}"
        current_count = used_counts.get(base_name, 0)
        used_counts[base_name] = current_count + 1
        if current_count:
            unique_columns.append(f"{base_name}_{current_count + 1}")
        else:
            unique_columns.append(base_name)
    return unique_columns


def build_excel_download_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    used_names: set[str] = set()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for raw_name, df in sheets.items():
            base_name = sanitize_excel_sheet_name(raw_name)
            sheet_name = base_name
            suffix = 1
            while sheet_name in used_names:
                suffix += 1
                sheet_name = f"{base_name[:28]}_{suffix}"
            used_names.add(sheet_name)

            excel_df = dataframe_for_excel(df)
            if excel_df.empty:
                excel_df = pd.DataFrame({"내용": ["조건에 맞는 데이터가 없습니다."]})
            excel_df = excel_df.copy()
            excel_df.columns = make_unique_excel_columns(excel_df.columns)
            excel_df.to_excel(writer, sheet_name=sheet_name, index=False)

            worksheet = writer.sheets[sheet_name]
            for col_idx, col_name in enumerate(excel_df.columns, start=1):
                column_values = excel_df.iloc[:, col_idx - 1].head(300).tolist()
                value_lengths = [excel_text_length(value) for value in column_values]
                max_len = max([len(str(col_name)), *(value_lengths if value_lengths else [0])])
                worksheet.column_dimensions[worksheet.cell(row=1, column=col_idx).column_letter].width = min(
                    max(max_len + 2, 10),
                    45,
                )
    return output.getvalue()


def render_excel_download(
    label: str,
    file_prefix: str,
    sheets: dict[str, pd.DataFrame],
    key: str,
    width: str = "stretch",
) -> None:
    timestamp = pd.Timestamp.now(tz="Asia/Seoul").strftime("%Y%m%d_%H%M")
    st.download_button(
        label,
        data=build_excel_download_bytes(sheets),
        file_name=f"{file_prefix}_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width=width,
        key=key,
    )


def build_priority_report_view(product_view: pd.DataFrame, max_rows: int = 6) -> pd.DataFrame:
    columns = ["제품명", "요청 PACK", "생산진도율", "용마입고율", "생산부족수량", "미입고수량", "상태"]
    if product_view.empty:
        return pd.DataFrame(columns=columns)

    view = product_view.copy()
    if "용마입고율" not in view.columns:
        view["용마입고율"] = view.get("포장진도율", 0.0)
    if "미입고수량" not in view.columns:
        view["미입고수량"] = view.get("포장부족수량", 0.0)
    for col in columns:
        if col not in view.columns:
            view[col] = ""
    return view.sort_values(
        ["미입고수량", "생산부족수량", "요청 PACK"],
        ascending=[False, False, False],
        kind="stable",
    ).head(max_rows)[columns].copy()


def build_daily_exception_report_view(
    daily_inventory_df: pd.DataFrame | None,
    code_summary: pd.DataFrame,
    sample_available_df: pd.DataFrame | None = None,
    max_rows: int = 5,
) -> tuple[dict[str, float], pd.DataFrame]:
    columns = ["품목코드", "제품명", "현재 재고수량", "부족수량", "포장가능재고(PCS)", "대응가능 여부"]
    empty_kpis = {"request_out_count": 0.0, "negative_count": 0.0, "waiting_pcs": 0.0}
    if daily_inventory_df is None or daily_inventory_df.empty:
        return empty_kpis, pd.DataFrame(columns=columns)

    view = build_daily_inventory_response_view(daily_inventory_df, code_summary, sample_available_df)
    if view.empty or "대응상태" not in view.columns:
        return empty_kpis, pd.DataFrame(columns=columns)

    out = view[view["대응상태"] == "요청외 긴급"].copy()
    if out.empty:
        return empty_kpis, pd.DataFrame(columns=columns)

    out["재고수량"] = pd.to_numeric(out["재고수량"], errors="coerce")
    out["재고부족수량"] = pd.to_numeric(out.get("재고부족수량", 0.0), errors="coerce").fillna(0.0)
    pack_units = out["PACK"].map(pack_unit_from_label)
    out["부족수량"] = (out["재고부족수량"] * pack_units).clip(lower=0.0).round(0)
    out["포장가능재고(PCS)"] = pd.to_numeric(out["포장가능재고(PCS)"], errors="coerce").fillna(0.0)
    out["대응가능 여부"] = [
        classify_exception_response(available, shortage)
        for available, shortage in zip(out["포장가능재고(PCS)"], out["부족수량"])
    ]
    out["현재 재고수량"] = out["재고수량"]
    out["_stock_missing_sort"] = out["재고수량"].isna().astype(int)
    out["_stock_sort"] = out["재고수량"].fillna(0.0)
    out["_response_sort"] = out["대응가능 여부"].map({"대응 필요": 0, "일부 가능": 1, "충당 가능": 2}).fillna(3)
    kpis = {
        "request_out_count": float(len(out)),
        "negative_count": float((out["재고수량"] < 0).sum()),
        "waiting_pcs": float(out["포장가능재고(PCS)"].sum()),
    }
    detail = out.sort_values(
        ["_response_sort", "_stock_missing_sort", "_stock_sort", "포장가능재고(PCS)", "품목코드"],
        ascending=[True, True, True, False, True],
        kind="stable",
    ).head(max_rows)
    for col in columns:
        if col not in detail.columns:
            detail[col] = ""
    return kpis, detail[columns].copy()


def build_urgent_request_summary_view(
    daily_inventory_df: pd.DataFrame | None,
    code_summary: pd.DataFrame,
    sample_available_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    columns = ["S코드", "요청구분", "제품명", "SKU 수"]
    if daily_inventory_df is None or daily_inventory_df.empty:
        return pd.DataFrame(columns=columns)

    view = build_daily_inventory_response_view(daily_inventory_df, code_summary, sample_available_df)
    if view.empty or "긴급요청" not in view.columns:
        return pd.DataFrame(columns=columns)

    urgent = view[view["긴급요청"].fillna(False).astype(bool)].copy()
    if urgent.empty:
        return pd.DataFrame(columns=columns)

    urgent["S코드"] = [
        extract_sales_prefix(product_code) or extract_sales_prefix(item_code)
        for product_code, item_code in zip(
            urgent.get("제품코드", pd.Series("", index=urgent.index)),
            urgent.get("품목코드", pd.Series("", index=urgent.index)),
        )
    ]
    urgent = urgent[urgent["S코드"].map(lambda value: bool(re.fullmatch(r"S\d{3}", clean_str(value))))].copy()
    if urgent.empty:
        return pd.DataFrame(columns=columns)

    urgent["_sku_key"] = [
        clean_str(item_code)
        or "|".join([clean_str(product_code), clean_str(pack), clean_str(power)])
        for item_code, product_code, pack, power in zip(
            urgent.get("품목코드", pd.Series("", index=urgent.index)),
            urgent["S코드"],
            urgent.get("PACK", pd.Series("", index=urgent.index)),
            urgent.get("POWER", pd.Series("", index=urgent.index)),
        )
    ]
    urgent["_request_pack"] = pd.to_numeric(
        urgent.get("요청 PACK", pd.Series(0.0, index=urgent.index)),
        errors="coerce",
    ).fillna(0.0)
    urgent["_request_scope"] = np.where(urgent["_request_pack"] > 0, "요청내", "요청외")
    urgent["_request_in_sku_key"] = urgent["_sku_key"].where(urgent["_request_scope"] == "요청내", "")
    urgent["_request_out_sku_key"] = urgent["_sku_key"].where(urgent["_request_scope"] == "요청외", "")
    grouped = (
        urgent.groupby("S코드", dropna=False)
        .agg(
            제품명=("제품명", join_unique),
            request_in_count=("_request_in_sku_key", lambda series: len({clean_str(value) for value in series if clean_str(value)})),
            request_out_count=("_request_out_sku_key", lambda series: len({clean_str(value) for value in series if clean_str(value)})),
            sku_count=("_sku_key", "nunique"),
        )
        .reset_index()
    )
    grouped = grouped.rename(
        columns={
            "request_in_count": "요청내 SKU",
            "request_out_count": "요청외 SKU",
            "sku_count": "SKU 수",
        }
    )
    for col in ["요청내 SKU", "요청외 SKU"]:
        grouped[col] = pd.to_numeric(grouped[col], errors="coerce").fillna(0).astype(int)
    grouped["요청구분"] = np.where(grouped["요청외 SKU"] > 0, "요청외", "요청내")
    grouped["SKU 수"] = pd.to_numeric(grouped["SKU 수"], errors="coerce").fillna(0).astype(int)
    return grouped.sort_values(["SKU 수", "S코드"], ascending=[False, True], kind="stable")[columns].copy()


def classify_exception_response(available_pcs: Any, shortage_pcs: Any) -> str:
    available_num = pd.to_numeric(available_pcs, errors="coerce")
    shortage_num = pd.to_numeric(shortage_pcs, errors="coerce")
    available = 0.0 if pd.isna(available_num) else float(available_num)
    shortage = 0.0 if pd.isna(shortage_num) else float(shortage_num)
    if available <= 0:
        return "대응 필요"
    if available >= shortage:
        return "충당 가능"
    return "일부 가능"


def render_exception_summary_table(exception_detail: pd.DataFrame) -> None:
    if exception_detail.empty:
        st.warning("요청외 긴급 품목이 없습니다.")
        return

    headers = ["품목코드", "제품명", "현재 재고수량", "부족수량", "포장가능재고(PCS)", "대응가능 여부"]
    rows: list[str] = []
    for _, row in exception_detail.iterrows():
        stock = pd.to_numeric(row.get("현재 재고수량", np.nan), errors="coerce")
        shortage = pd.to_numeric(row.get("부족수량", 0.0), errors="coerce")
        available = pd.to_numeric(row.get("포장가능재고(PCS)", 0.0), errors="coerce")
        response = clean_str(row.get("대응가능 여부", ""))
        response_class = {
            "충당 가능": "ok",
            "일부 가능": "partial",
            "대응 필요": "need",
        }.get(response, "need")
        stock_text = "-" if pd.isna(stock) else format_int(float(stock))
        stock_class = "num negative" if not pd.isna(stock) and float(stock) < 0 else "num"
        rows.append(
            "<tr>"
            f"<td>{escape(str(row.get('품목코드', '')))}</td>"
            f"<td class='left'>{escape(str(row.get('제품명', '')))}</td>"
            f"<td class='{stock_class}'>{stock_text}</td>"
            f"<td class='num shortage'>{format_int(float(shortage) if not pd.isna(shortage) else 0.0)}</td>"
            f"<td class='num'>{format_int(float(available) if not pd.isna(available) else 0.0)}</td>"
            f"<td><span class='response-badge {response_class}'>{escape(response)}</span></td>"
            "</tr>"
        )
    header_html = "".join(f"<th class='{'left' if header == '제품명' else ''}'>{escape(header)}</th>" for header in headers)
    st.markdown(
        "<div class='table-wrap compact-table'>"
        "<table class='ops-table urgent-summary-table'>"
        f"<thead><tr>{header_html}</tr></thead>"
        f"<tbody>{''.join(rows)}</tbody>"
        "</table>"
        "</div>",
        unsafe_allow_html=True,
    )


def render_urgent_request_summary_table(summary_view: pd.DataFrame) -> None:
    if summary_view.empty:
        st.warning("긴급요청 품목이 없습니다.")
        return

    rows: list[str] = []
    for _, row in summary_view.iterrows():
        sku_num = pd.to_numeric(row.get("SKU 수", 0), errors="coerce")
        sku_count = 0.0 if pd.isna(sku_num) else float(sku_num)
        scope = clean_str(row.get("요청구분", ""))
        scope_class = "in" if scope == "요청내" else "out"
        rows.append(
            "<tr>"
            f"<td class='left code-cell'>{escape(clean_str(row.get('S코드', '')))}</td>"
            f"<td class='left'><span class='request-scope-badge {scope_class}'>{escape(scope)}</span></td>"
            f"<td class='left'>{escape(clean_str(row.get('제품명', '')))}</td>"
            f"<td class='num shortage'>{format_int(sku_count)}</td>"
            "</tr>"
        )
    st.markdown(
        "<div class='table-wrap compact-table'>"
        "<table class='ops-table urgent-summary-table'>"
        "<thead><tr>"
        "<th class='left'>S코드</th>"
        "<th class='left'>요청구분</th>"
        "<th class='left'>제품명</th>"
        "<th class='num'>SKU 수</th>"
        "</tr></thead>"
        f"<tbody>{''.join(rows)}</tbody>"
        "</table>"
        "</div>",
        unsafe_allow_html=True,
    )


def to_report_float(value: Any) -> float:
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return 0.0
    return float(num)


def add_report_status_badge(slide: Any, status: str, left: float, top: float, width: float, height: float) -> None:
    status_text = clean_str(status) or "-"
    if status_text in {"완료", "입고완료"}:
        fill_color = "#E8F5F0"
        line_color = "#9ED8C5"
        text_color = COLOR_TEAL
    elif status_text in {"진행중"}:
        fill_color = "#FFF4DE"
        line_color = "#E4B968"
        text_color = COLOR_AMBER
    else:
        fill_color = COLOR_ALERT_BG
        line_color = COLOR_ALERT_BD
        text_color = COLOR_ORANGE

    add_report_shape(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left,
        top,
        width,
        height,
        fill_color,
        line_color,
        0.5,
    )
    add_textbox(
        slide,
        status_text,
        left,
        top,
        width,
        height,
        7.5,
        True,
        text_color,
        PP_ALIGN.CENTER,
        MSO_ANCHOR.MIDDLE,
    )


def add_priority_report_table(
    slide: Any,
    priority_view: pd.DataFrame,
    left: float = 0.45,
    top: float = 3.62,
    width: float = 8.35,
    height: float = 3.28,
) -> None:
    add_report_shape(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left,
        top,
        width,
        height,
        REPORT_PANEL,
        REPORT_PANEL_LINE,
        0.5,
    )
    add_report_shape(slide, MSO_SHAPE.RECTANGLE, left, top, width, 0.38, REPORT_NAVY, REPORT_NAVY, 0.5)

    headers = ["순위", "제품명", "요청 PACK", "용마입고율", "미입고 PACK", "생산진도율"]
    col_widths = [0.58, 3.18, 1.08, 1.14, 1.12, 1.08]
    col_lefts = [left + 0.16]
    for width in col_widths[:-1]:
        col_lefts.append(col_lefts[-1] + width)

    for idx, header in enumerate(headers):
        add_textbox(
            slide,
            header,
            col_lefts[idx],
            top + 0.04,
            col_widths[idx],
            0.3,
            8.2,
            True,
            "#FFFFFF",
            PP_ALIGN.LEFT if idx == 1 else PP_ALIGN.CENTER if idx == 0 else PP_ALIGN.RIGHT,
            MSO_ANCHOR.MIDDLE,
        )

    if priority_view.empty:
        add_textbox(
            slide,
            "조건에 맞는 제품 데이터가 없습니다.",
            left + 0.2,
            top + 0.55,
            width - 0.4,
            0.35,
            8.4,
            False,
            REPORT_MUTED,
            PP_ALIGN.LEFT,
            MSO_ANCHOR.MIDDLE,
        )
        return

    row_height = 0.48
    for row_idx, (_, row) in enumerate(priority_view.iterrows(), start=1):
        row_top = top + 0.38 + (row_idx - 1) * row_height
        production_progress = to_report_float(row["생산진도율"])
        receipt_progress = to_report_float(row["용마입고율"])
        receipt_shortage = to_report_float(row["미입고수량"])

        if row_idx % 2 == 0:
            add_report_shape(slide, MSO_SHAPE.RECTANGLE, left + 0.06, row_top, width - 0.12, row_height, REPORT_ROW_ALT)

        cell_top = row_top + 0.01
        cell_height = row_height - 0.02
        values = [
            str(row_idx),
            truncate_report_text(row["제품명"], max_chars=32),
            format_report_value(row["요청 PACK"]),
            format_report_value(receipt_progress, True),
            format_report_value(receipt_shortage),
            format_report_value(production_progress, True),
        ]
        colors = [
            REPORT_MUTED,
            REPORT_HEADER,
            REPORT_HEADER,
            REPORT_HEADER,
            REPORT_MUTED if receipt_shortage <= 0 else REPORT_ACCENT,
            REPORT_HEADER,
        ]
        bolds = [False, False, False, False, receipt_shortage > 0, False]
        aligns = [PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT]

        for col_idx, value in enumerate(values):
            add_textbox(
                slide,
                value,
                col_lefts[col_idx],
                cell_top,
                col_widths[col_idx],
                cell_height,
                8.6 if col_idx != 1 else 8.8,
                bolds[col_idx],
                colors[col_idx],
                aligns[col_idx],
                MSO_ANCHOR.MIDDLE,
            )

        add_report_rule(slide, left + 0.1, row_top + row_height, width - 0.2, REPORT_FAINT)


def add_daily_exception_report_panel(
    slide: Any,
    exception_kpis: dict[str, float],
    exception_view: pd.DataFrame,
    left: float = 8.95,
    top: float = 3.62,
    width: float = 3.95,
    height: float = 3.28,
) -> None:
    add_report_shape(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left,
        top,
        width,
        height,
        REPORT_PANEL,
        REPORT_PANEL_LINE,
        0.5,
    )
    add_report_shape(slide, MSO_SHAPE.RECTANGLE, left, top, width, 0.38, REPORT_NAVY, REPORT_NAVY, 0.5)

    headers = ["품목", "제품명", "재고", "가용 PCS"]
    col_widths = [0.72, 1.62, 0.56, 0.78]
    col_lefts = [left + 0.12]
    for col_width in col_widths[:-1]:
        col_lefts.append(col_lefts[-1] + col_width)
    for idx, header in enumerate(headers):
        add_textbox(
            slide,
            header,
            col_lefts[idx],
            top + 0.04,
            col_widths[idx],
            0.3,
            7.8,
            True,
            "#FFFFFF",
            PP_ALIGN.LEFT if idx in {0, 1} else PP_ALIGN.RIGHT,
            MSO_ANCHOR.MIDDLE,
        )

    if exception_view.empty:
        add_textbox(
            slide,
            "요청물량 외 긴급 대응 품목이 없습니다.",
            left + 0.18,
            top + 0.72,
            width - 0.36,
            0.28,
            8.2,
            False,
            REPORT_MUTED,
            vertical_anchor=MSO_ANCHOR.MIDDLE,
        )
        return

    row_height = 0.48
    for row_idx, (_, row) in enumerate(exception_view.iterrows(), start=1):
        row_top = top + 0.38 + (row_idx - 1) * row_height
        stock = pd.to_numeric(row.get("현재 재고수량", row.get("재고수량", np.nan)), errors="coerce")
        waiting_pcs = pd.to_numeric(row.get("포장가능재고(PCS)", np.nan), errors="coerce")
        stock_text = "-" if pd.isna(stock) else format_report_value(stock)
        waiting_text = "-" if pd.isna(waiting_pcs) else format_report_value(waiting_pcs)
        stock_color = REPORT_ACCENT if pd.notna(stock) and float(stock) < 0 else REPORT_HEADER
        waiting_color = REPORT_HEADER
        if row_idx % 2 == 0:
            add_report_shape(slide, MSO_SHAPE.RECTANGLE, left + 0.03, row_top, width - 0.06, row_height, REPORT_ROW_ALT)
        values = [
            truncate_report_text(row.get("품목코드", ""), 12),
            truncate_report_text(row.get("제품명", ""), 20),
            stock_text,
            waiting_text,
        ]
        colors = [REPORT_HEADER, REPORT_HEADER, stock_color, waiting_color]
        aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT]
        for col_idx, value in enumerate(values):
            add_textbox(
                slide,
                value,
                col_lefts[col_idx],
                row_top + 0.03,
                col_widths[col_idx],
                row_height - 0.06,
                7.7 if col_idx != 1 else 7.9,
                col_idx == 2 and stock_color == REPORT_ACCENT,
                colors[col_idx],
                aligns[col_idx],
                MSO_ANCHOR.MIDDLE,
            )
        add_report_rule(slide, left + 0.08, row_top + row_height, width - 0.16, REPORT_PANEL_LINE)


def add_urgent_request_summary_panel(
    slide: Any,
    urgent_summary_view: pd.DataFrame,
    left: float = 8.95,
    top: float = 3.56,
    width: float = 3.95,
    height: float = 3.47,
) -> None:
    add_report_shape(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left,
        top,
        width,
        height,
        REPORT_PANEL,
        REPORT_PANEL_LINE,
        0.5,
    )
    add_report_shape(slide, MSO_SHAPE.RECTANGLE, left, top, width, 0.36, REPORT_NAVY, REPORT_NAVY, 0.5)

    headers = ["S코드", "구분", "제품명", "SKU"]
    content_width = max(0.0, width - 0.24)
    if width >= 8.0:
        col_widths = [1.0, 1.1, max(1.0, content_width - 2.85), 0.75]
        product_max_chars = 68
        body_font_size = 8.0
        product_font_size = 7.8
    else:
        col_widths = [0.58, 0.58, 2.12, 0.38]
        product_max_chars = 22
        body_font_size = 6.5
        product_font_size = 6.1
    col_lefts = [left + 0.12]
    for col_width in col_widths[:-1]:
        col_lefts.append(col_lefts[-1] + col_width)

    for idx, header in enumerate(headers):
        add_textbox(
            slide,
            header,
            col_lefts[idx],
            top + 0.04,
            col_widths[idx],
            0.27,
            7.2,
            True,
            "#FFFFFF",
            PP_ALIGN.RIGHT if idx == 3 else PP_ALIGN.LEFT,
            MSO_ANCHOR.MIDDLE,
        )

    if urgent_summary_view.empty:
        add_textbox(
            slide,
            "긴급요청 품목이 없습니다.",
            left + 0.18,
            top + 0.7,
            width - 0.36,
            0.3,
            8.0,
            False,
            REPORT_MUTED,
            vertical_anchor=MSO_ANCHOR.MIDDLE,
        )
        return

    max_rows = 11
    rows = urgent_summary_view.head(max_rows).copy()
    hidden_count = max(0, len(urgent_summary_view) - len(rows))
    row_height = min(0.27, (height - 0.58) / max(len(rows), 1))
    for row_idx, (_, row) in enumerate(rows.iterrows(), start=1):
        row_top = top + 0.36 + (row_idx - 1) * row_height
        if row_idx % 2 == 0:
            add_report_shape(slide, MSO_SHAPE.RECTANGLE, left + 0.04, row_top, width - 0.08, row_height, REPORT_ROW_ALT)

        scope = clean_str(row.get("요청구분", ""))
        scope_color = REPORT_ACCENT if scope == "요청외" else COLOR_TEAL
        values = [
            clean_str(row.get("S코드", "")),
            scope,
            truncate_report_text(row.get("제품명", ""), product_max_chars),
            format_report_value(row.get("SKU 수", 0.0)),
        ]
        colors = [COLOR_BLUE, scope_color, REPORT_HEADER, REPORT_ACCENT]
        bolds = [True, True, False, True]
        aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT]
        for col_idx, value in enumerate(values):
            add_textbox(
                slide,
                value,
                col_lefts[col_idx],
                row_top + 0.01,
                col_widths[col_idx],
                row_height - 0.02,
                body_font_size if col_idx != 2 else product_font_size,
                bolds[col_idx],
                colors[col_idx],
                aligns[col_idx],
                MSO_ANCHOR.MIDDLE,
            )
        add_report_rule(slide, left + 0.08, row_top + row_height, width - 0.16, REPORT_FAINT)

    note = "요청외 SKU 포함 시 S코드 전체 요청외"
    if hidden_count:
        note = f"{note} / 외 {hidden_count:,}개"
    add_textbox(
        slide,
        note,
        left + 0.12,
        top + height - 0.19,
        width - 0.24,
        0.16,
        5.9,
        False,
        REPORT_MUTED,
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )


def add_urgent_request_summary_slide(
    prs: Presentation,
    urgent_summary_view: pd.DataFrame,
    scope_label: str,
    generated_at: str,
) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = ppt_rgb(REPORT_BG)

    add_report_shape(slide, MSO_SHAPE.RECTANGLE, 0.0, 0.0, 13.333, 0.88, REPORT_NAVY)
    add_report_shape(slide, MSO_SHAPE.RECTANGLE, 0.0, 0.86, 13.333, 0.03, REPORT_ACCENT, REPORT_ACCENT)
    add_textbox(
        slide,
        "요청 긴급 S코드 요약",
        0.45,
        0.16,
        6.8,
        0.32,
        18,
        True,
        "#FFFFFF",
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )
    add_textbox(
        slide,
        "일일 재고표 기준 긴급요청 품목을 S코드 단위로 집계",
        0.45,
        0.51,
        7.2,
        0.18,
        8.3,
        False,
        "#CBD5E1",
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )
    add_textbox(
        slide,
        f"기준: {scope_label} / 산출시각: {generated_at}",
        8.1,
        0.33,
        4.75,
        0.2,
        7.8,
        False,
        "#E5E7EB",
        PP_ALIGN.RIGHT,
        MSO_ANCHOR.MIDDLE,
    )

    total_sku = int(pd.to_numeric(urgent_summary_view.get("SKU 수", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
    request_out_count = int((urgent_summary_view.get("요청구분", pd.Series(dtype=str)).astype(str) == "요청외").sum())
    add_textbox(
        slide,
        f"S코드 {len(urgent_summary_view):,}개 / SKU {total_sku:,}개 / 요청외 S코드 {request_out_count:,}개",
        0.62,
        1.12,
        8.0,
        0.24,
        10,
        True,
        REPORT_HEADER,
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )
    add_textbox(
        slide,
        "요청외는 해당 S코드 안에 요청외 긴급 SKU가 1개 이상 포함된 경우로 분류합니다.",
        0.62,
        6.98,
        11.8,
        0.2,
        7.5,
        False,
        REPORT_MUTED,
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )

    left = 0.62
    top = 1.5
    width = 12.1
    height = 5.3
    add_report_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height, REPORT_PANEL, REPORT_PANEL_LINE, 0.5)
    add_report_shape(slide, MSO_SHAPE.RECTANGLE, left, top, width, 0.42, REPORT_NAVY, REPORT_NAVY, 0.5)

    headers = ["S코드", "요청구분", "제품명", "SKU 수"]
    col_widths = [1.05, 1.25, 8.55, 0.95]
    col_lefts = [left + 0.16]
    for col_width in col_widths[:-1]:
        col_lefts.append(col_lefts[-1] + col_width)

    for idx, header in enumerate(headers):
        add_textbox(
            slide,
            header,
            col_lefts[idx],
            top + 0.06,
            col_widths[idx],
            0.3,
            8.5,
            True,
            "#FFFFFF",
            PP_ALIGN.RIGHT if idx == 3 else PP_ALIGN.LEFT,
            MSO_ANCHOR.MIDDLE,
        )

    if urgent_summary_view.empty:
        add_textbox(
            slide,
            "긴급요청 품목이 없습니다.",
            left + 0.2,
            top + 0.72,
            width - 0.4,
            0.35,
            9,
            False,
            REPORT_MUTED,
            vertical_anchor=MSO_ANCHOR.MIDDLE,
        )
        return

    rows = urgent_summary_view.head(12).copy()
    row_height = 0.39
    for row_idx, (_, row) in enumerate(rows.iterrows(), start=1):
        row_top = top + 0.42 + (row_idx - 1) * row_height
        if row_idx % 2 == 0:
            add_report_shape(slide, MSO_SHAPE.RECTANGLE, left + 0.08, row_top, width - 0.16, row_height, REPORT_ROW_ALT)
        scope = clean_str(row.get("요청구분", ""))
        scope_color = REPORT_ACCENT if scope == "요청외" else COLOR_TEAL
        values = [
            clean_str(row.get("S코드", "")),
            scope,
            truncate_report_text(row.get("제품명", ""), 64),
            format_report_value(row.get("SKU 수", 0.0)),
        ]
        colors = [COLOR_BLUE, scope_color, REPORT_HEADER, REPORT_ACCENT]
        aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT]
        for col_idx, value in enumerate(values):
            add_textbox(
                slide,
                value,
                col_lefts[col_idx],
                row_top + 0.02,
                col_widths[col_idx],
                row_height - 0.04,
                8.7 if col_idx != 2 else 8.4,
                col_idx in {0, 1, 3},
                colors[col_idx],
                aligns[col_idx],
                MSO_ANCHOR.MIDDLE,
            )
        add_report_rule(slide, left + 0.1, row_top + row_height, width - 0.2, REPORT_FAINT)

    hidden_count = max(0, len(urgent_summary_view) - len(rows))
    if hidden_count:
        add_textbox(
            slide,
            f"외 {hidden_count:,}개 S코드는 대시보드와 엑셀 다운로드에서 확인 가능합니다.",
            left + 0.2,
            top + height - 0.36,
            width - 0.4,
            0.2,
            7.5,
            False,
            REPORT_MUTED,
            vertical_anchor=MSO_ANCHOR.MIDDLE,
        )


def add_report_legend(slide: Any) -> None:
    legend_items = [
        (COLOR_TEAL, "진도율 정상 (>=80%)"),
        (COLOR_ORANGE, "진도율 미달 / 미입고 / 경고"),
        (COLOR_AMBER, "생산부족 발생"),
        ("#AAAAAA", "부족 없음"),
    ]
    for idx, (color, label) in enumerate(legend_items):
        left = 0.3 + idx * 3.1
        add_report_shape(slide, MSO_SHAPE.RECTANGLE, left, 7.28, 0.18, 0.1, color, color)
        add_textbox(
            slide,
            label,
            left + 0.22,
            7.22,
            2.8,
            0.22,
            7.5,
            False,
            TEXT_SECONDARY,
            vertical_anchor=MSO_ANCHOR.MIDDLE,
        )


def build_ppt_report(
    product_view: pd.DataFrame,
    code_summary: pd.DataFrame,
    product_names: pd.Series,
    scope_label: str,
    daily_inventory_df: pd.DataFrame | None = None,
    sample_available_df: pd.DataFrame | None = None,
) -> bytes:
    work = add_allocated_production_basis(code_summary)
    work = code_summary_for_products(work, product_names)
    scope_kpis = build_scope_kpis(work)
    urgent_summary_view = build_urgent_request_summary_view(
        daily_inventory_df,
        code_summary,
        sample_available_df,
    )
    urgent_sku_count = int(
        pd.to_numeric(urgent_summary_view.get("SKU 수", pd.Series(dtype=float)), errors="coerce")
        .fillna(0)
        .sum()
    )

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = ppt_rgb(REPORT_BG)

    add_report_shape(slide, MSO_SHAPE.RECTANGLE, 0.0, 0.0, 13.333, 0.88, REPORT_NAVY)
    add_report_shape(slide, MSO_SHAPE.RECTANGLE, 0.0, 0.86, 13.333, 0.03, REPORT_ACCENT, REPORT_ACCENT)
    add_textbox(
        slide,
        "국내 제품 포장현황 운영 보고서",
        0.45,
        0.15,
        6.6,
        0.34,
        19,
        True,
        "#FFFFFF",
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )
    add_textbox(
        slide,
        "국내 요청 물량의 생산, 포장, 용마 입고 진행 현황",
        0.45,
        0.51,
        6.6,
        0.18,
        8.3,
        False,
        "#CBD5E1",
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )
    generated_at = pd.Timestamp.now(tz="Asia/Seoul").strftime("%Y-%m-%d %H:%M")
    add_textbox(
        slide,
        f"기준: {scope_label}",
        8.35,
        0.22,
        4.65,
        0.18,
        7.9,
        False,
        "#E5E7EB",
        PP_ALIGN.RIGHT,
        MSO_ANCHOR.MIDDLE,
    )
    add_textbox(
        slide,
        f"산출시각: {generated_at}",
        8.35,
        0.48,
        4.65,
        0.18,
        7.9,
        False,
        "#CBD5E1",
        PP_ALIGN.RIGHT,
        MSO_ANCHOR.MIDDLE,
    )

    kpi_map = {name: kpi for name, kpi in scope_kpis}
    total_kpi = kpi_map.get("전체", {})
    total_progress = to_report_float(total_kpi.get("progress_pct", 0.0))
    total_packing_progress = to_report_float(total_kpi.get("packing_progress_pct", 0.0))
    total_shortage = to_report_float(total_kpi.get("shortage_pack", 0.0))
    exception_count = float(urgent_sku_count)
    if total_shortage > 0 or exception_count > 0:
        banner_fill = "#FFFFFF"
        banner_color = REPORT_ACCENT
        status_label = "주의"
    else:
        banner_fill = "#FFFFFF"
        banner_color = COLOR_TEAL
        status_label = "정상"
    banner_text = (
        f"[{status_label}] 포장진도율 {total_packing_progress:.1f}%"
        f" / 용마입고율 {total_progress:.1f}%"
        f" / 미입고 {format_report_value(total_shortage)} PACK"
        f" / 긴급 SKU {format_report_value(exception_count)}개"
    )

    add_report_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.45, 1.04, 12.45, 0.48, banner_fill, REPORT_PANEL_LINE, 0.5)
    add_report_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.62, 1.15, 0.62, 0.24, banner_color, banner_color, 0.4)
    add_textbox(
        slide,
        status_label,
        0.62,
        1.16,
        0.62,
        0.2,
        7.2,
        True,
        "#FFFFFF",
        PP_ALIGN.CENTER,
        MSO_ANCHOR.MIDDLE,
    )
    add_textbox(
        slide,
        banner_text.replace(f"[{status_label}] ", ""),
        1.38,
        1.18,
        11.1,
        0.2,
        8.4,
        True,
        REPORT_HEADER,
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )

    add_textbox(slide, "공급 운영 KPI", 0.45, 1.66, 2.0, 0.22, 8.5, True, REPORT_HEADER)
    add_kpi_card(slide, "전체 KPI", total_kpi, COLOR_BLUE, 0.45, 1.9, 4.0, 1.38)
    add_kpi_card(slide, "본품 KPI", kpi_map.get("본품", {}), COLOR_TEAL, 4.68, 1.9, 4.0, 1.38)
    add_kpi_card(slide, "샘플 KPI", kpi_map.get("샘플", {}), COLOR_AMBER, 8.9, 1.9, 4.0, 1.38)

    add_textbox(slide, "요청 긴급 S코드 요약", 0.45, 3.34, 12.45, 0.22, 8.5, True, REPORT_HEADER)

    add_urgent_request_summary_panel(slide, urgent_summary_view, left=0.45, top=3.56, width=12.45)
    add_textbox(
        slide,
        "진도율은 생산요청 기준이며, 요청/긴급 대응 품목은 일일 재고표 기준으로 산출됩니다.",
        0.45,
        7.16,
        12.4,
        0.22,
        7.5,
        False,
        REPORT_MUTED,
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )

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
        color_discrete_map={"전체": COLOR_BLUE, "본품": COLOR_TEAL, "샘플": COLOR_AMBER},
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
        color_discrete_map={"본품": COLOR_ORANGE, "샘플": COLOR_AMBER},
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
    receipt_progress_col = "용마입고율" if "용마입고율" in source.columns else "포장진도율"
    chart_source = source[["제품명", "생산진도율", receipt_progress_col]].rename(
        columns={receipt_progress_col: "용마입고율"}
    )
    chart_df = chart_source.melt(
        id_vars="제품명",
        value_vars=["생산진도율", "용마입고율"],
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
        category_orders={"제품명": product_order, "지표": ["생산진도율", "용마입고율"]},
        title="제품별 생산진도율 vs 용마입고율",
        text="진도율",
        color_discrete_map={"생산진도율": COLOR_BLUE, "용마입고율": COLOR_TEAL},
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
        @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');

        html, body, [class*="css"], .stApp {{
            font-family: 'Pretendard', 'Apple SD Gothic Neo', sans-serif !important;
            color: {TEXT_PRIMARY};
        }}
        .stApp {{
            background: {BG_PAGE};
            color: {TEXT_PRIMARY};
        }}
        :root {{
            --ui-gap: 12px;
            --color-blue: {COLOR_BLUE};
            --color-teal: {COLOR_TEAL};
            --color-orange: {COLOR_ORANGE};
            --color-amber: {COLOR_AMBER};
            --bg-page: {BG_PAGE};
            --bg-card: {BG_CARD};
            --bg-section: {BG_SECTION};
            --text-primary: {TEXT_PRIMARY};
            --text-secondary: {TEXT_SECONDARY};
            --text-tertiary: {TEXT_TERTIARY};
            --border-default: {BORDER_DEFAULT};
            --border-light: {BORDER_LIGHT};
        }}
        .block-container {{
            padding: 24px 32px 32px !important;
            max-width: 100% !important;
        }}
        h1 {{
            font-size: 22px !important;
            font-weight: 500 !important;
            color: {TEXT_PRIMARY} !important;
            margin-bottom: 4px !important;
            letter-spacing: 0 !important;
        }}
        h2 {{
            font-size: 15px !important;
            font-weight: 500 !important;
            color: {TEXT_PRIMARY} !important;
            margin-bottom: 2px !important;
            letter-spacing: 0 !important;
        }}
        h3 {{
            font-size: 13px !important;
            font-weight: 500 !important;
            color: {TEXT_PRIMARY} !important;
            letter-spacing: 0 !important;
        }}
        [data-baseweb="tab-list"] {{
            gap: 0 !important;
            border-bottom: 1.5px solid {BORDER_DEFAULT} !important;
            background: transparent !important;
            margin-bottom: 20px !important;
        }}
        [data-baseweb="tab"] {{
            font-size: 13px !important;
            padding: 10px 20px !important;
            color: {TEXT_SECONDARY} !important;
            background: transparent !important;
            letter-spacing: 0 !important;
        }}
        [aria-selected="true"][data-baseweb="tab"] {{
            color: {COLOR_ORANGE} !important;
            font-weight: 500 !important;
        }}
        [data-baseweb="tab-highlight"] {{
            background-color: {COLOR_ORANGE} !important;
            height: 2px !important;
        }}
        [data-baseweb="tab-border"] {{
            display: none !important;
        }}
        [data-testid="stSegmentedControl"] {{
            margin: 2px 0 14px 0 !important;
        }}
        [data-testid="stSegmentedControl"] button {{
            border: 0 !important;
            border-radius: 0 !important;
            background: transparent !important;
            color: {TEXT_SECONDARY} !important;
            font-size: 13px !important;
            font-weight: 400 !important;
            padding: 10px 20px !important;
            box-shadow: none !important;
        }}
        [data-testid="stSegmentedControl"] button[aria-pressed="true"] {{
            color: {COLOR_ORANGE} !important;
            font-weight: 500 !important;
            border-bottom: 2px solid {COLOR_ORANGE} !important;
        }}
        .dashboard-nav-divider {{
            height: 1.5px;
            background: {BORDER_DEFAULT};
            margin: -15px 0 20px 0;
        }}
        [data-testid="stRadio"] label,
        [data-testid="stCheckbox"] label {{
            font-size: 13px !important;
            color: {TEXT_SECONDARY} !important;
        }}
        [data-testid="stRadio"] [data-testid="stMarkdownContainer"] p,
        [data-testid="stCheckbox"] [data-testid="stMarkdownContainer"] p {{
            font-size: 13px !important;
            color: {TEXT_SECONDARY} !important;
        }}
        [data-testid="metric-container"] {{
            background: {BG_CARD};
            border: 0.5px solid {BORDER_DEFAULT};
            border-radius: 12px;
            padding: 14px 18px;
            box-shadow: none;
        }}
        [data-testid="stMetricLabel"] {{
            font-size: 11px !important;
            color: {TEXT_TERTIARY} !important;
            font-weight: 400 !important;
        }}
        [data-testid="stMetricValue"] {{
            font-size: 20px !important;
            line-height: 1.1 !important;
            font-weight: 500 !important;
            color: {TEXT_PRIMARY} !important;
        }}
        [data-testid="stMetricDelta"] {{
            display: none !important;
        }}
        [data-testid="stDataFrame"] {{
            border: 0.5px solid {BORDER_DEFAULT} !important;
            border-radius: 12px !important;
            overflow: hidden;
            background: {BG_CARD};
        }}
        [data-testid="stDataFrame"] th {{
            background: {BG_PAGE} !important;
            font-size: 11px !important;
            font-weight: 500 !important;
            color: {TEXT_SECONDARY} !important;
            padding: 8px 12px !important;
            border-bottom: 1px solid {BORDER_DEFAULT} !important;
            white-space: nowrap;
        }}
        [data-testid="stDataFrame"] td {{
            font-size: 12px !important;
            color: {TEXT_PRIMARY} !important;
            padding: 7px 12px !important;
            border-bottom: 0.5px solid {BORDER_LIGHT} !important;
        }}
        [data-testid="stDataFrame"] tr:hover td {{
            background: {BG_PAGE} !important;
        }}
        [data-testid="stTextInput"] input,
        [data-testid="stNumberInput"] input,
        [data-testid="stSelectbox"] > div > div,
        [data-testid="stMultiSelect"] > div > div {{
            font-size: 13px !important;
            border-radius: 8px !important;
            border: 0.5px solid rgba(0,0,0,0.15) !important;
            background: {BG_CARD} !important;
            color: {TEXT_PRIMARY} !important;
        }}
        [data-testid="stTextInput"] label,
        [data-testid="stNumberInput"] label,
        [data-testid="stSelectbox"] label,
        [data-testid="stMultiSelect"] label {{
            font-size: 12px !important;
            font-weight: 500 !important;
            color: {TEXT_SECONDARY} !important;
        }}
        [data-testid="stMultiSelect"] [data-baseweb="tag"] {{
            background: {COLOR_ALERT_BG} !important;
            color: #993C1D !important;
            border-radius: 20px !important;
            border: 0 !important;
            font-size: 12px !important;
            font-weight: 500 !important;
        }}
        [data-testid="stButton"] button,
        [data-testid="stDownloadButton"] button {{
            border-radius: 8px !important;
            font-size: 13px !important;
            font-weight: 500 !important;
            border: 0.5px solid rgba(0,0,0,0.15) !important;
            background: {BG_CARD} !important;
            color: {TEXT_PRIMARY} !important;
            box-shadow: none !important;
        }}
        [data-testid="stButton"] button:hover,
        [data-testid="stDownloadButton"] button:hover {{
            background: {BG_PAGE} !important;
            border-color: rgba(0,0,0,0.25) !important;
            color: {TEXT_PRIMARY} !important;
        }}
        hr {{
            border-color: rgba(0,0,0,0.08) !important;
            margin: 20px 0 !important;
        }}
        ::-webkit-scrollbar {{
            width: 6px;
            height: 6px;
        }}
        ::-webkit-scrollbar-track {{
            background: transparent;
        }}
        ::-webkit-scrollbar-thumb {{
            background: rgba(0,0,0,0.15);
            border-radius: 3px;
        }}
        .kpi-panel {{
            background: {BG_CARD};
            border: 0.5px solid {BORDER_DEFAULT};
            border-radius: 12px;
            padding: 14px 18px;
            box-shadow: none;
            margin-bottom: 0;
            height: 100%;
        }}
        .drill-kpi {{
            margin-bottom: 12px;
        }}
        .kpi-title {{
            font-size: 15px;
            font-weight: 500;
            color: {TEXT_PRIMARY};
            margin-bottom: var(--ui-gap);
        }}
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
            gap: var(--ui-gap);
        }}
        .scope-kpi .kpi-grid {{
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 8px;
        }}
        .kpi-card {{
            border: 0.5px solid {BORDER_DEFAULT};
            border-radius: 12px;
            padding: 14px 18px;
            background: {BG_CARD};
            min-height: 72px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }}
        .metric-label {{
            font-size: 11px;
            font-weight: 400;
            color: {TEXT_TERTIARY};
            margin-bottom: 6px;
        }}
        .metric-value {{
            font-size: 20px;
            line-height: 1.1;
            font-weight: 500;
            color: {TEXT_PRIMARY};
            white-space: nowrap;
            overflow-wrap: normal;
            word-break: normal;
        }}
        .scope-kpi .metric-value {{
            font-size: 20px;
        }}
        @media (max-width: 1100px) {{
            .scope-kpi .kpi-grid {{
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }}
        }}
        .metric-value.warn {{
            color: {COLOR_ORANGE};
        }}
        .metric-value.risk {{
            color: {COLOR_ORANGE};
        }}
        .metric-value.normal {{
            color: {TEXT_PRIMARY};
        }}
        .metric-value.good {{
            color: {COLOR_TEAL};
        }}
        .metric-value.mid {{
            color: {COLOR_AMBER};
        }}
        .metric-value.muted {{
            color: {TEXT_TERTIARY};
        }}
        .metric-note {{
            color: {TEXT_TERTIARY};
            font-size: 11px;
            line-height: 1.2;
            margin-top: 4px;
        }}
        .mini-kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 10px;
            margin-bottom: 12px;
        }}
        .mini-kpi-card {{
            background: {BG_CARD};
            border: 0.5px solid {BORDER_DEFAULT};
            border-radius: 12px;
            padding: 14px 18px;
            min-height: 76px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            box-shadow: none;
        }}
        @media (max-width: 1100px) {{
            .mini-kpi-grid {{
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }}
        }}
        .shortage-card {{
            border-color: {COLOR_ALERT_BD};
            background: {COLOR_ALERT_BG};
        }}
        .status-board {{
            display: grid;
            grid-template-columns: minmax(320px, 1.1fr) minmax(420px, 1.9fr);
            gap: 10px;
            margin: 2px 0 10px;
            align-items: stretch;
        }}
        .status-main,
        .status-tile {{
            background: {BG_CARD};
            border: 0.5px solid {BORDER_DEFAULT};
            border-radius: 8px;
            box-shadow: none;
        }}
        .status-main {{
            padding: 16px 18px;
            border-left: 4px solid {TEXT_TERTIARY};
            box-shadow: none;
        }}
        .status-board.warn .status-main {{
            border-left-color: {COLOR_AMBER};
        }}
        .status-board.risk .status-main {{
            border-left-color: {COLOR_ORANGE};
        }}
        .status-head {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            margin-bottom: 10px;
        }}
        .status-head strong {{
            color: {TEXT_PRIMARY};
            font-size: 14px;
            font-weight: 800;
        }}
        .status-pill {{
            display: inline-flex;
            align-items: center;
            border-radius: 999px;
            padding: 4px 9px;
            font-size: 11px;
            font-weight: 700;
            color: {TEXT_SECONDARY};
            background: {BG_SECTION};
        }}
        .status-pill.warn {{
            color: {COLOR_AMBER};
            background: #F7EFE3;
        }}
        .status-pill.risk {{
            color: {COLOR_ORANGE};
            background: {COLOR_ALERT_BG};
        }}
        .status-main-value {{
            color: {TEXT_PRIMARY};
            font-size: 34px;
            line-height: 1;
            font-weight: 900;
            font-variant-numeric: tabular-nums;
            margin-bottom: 12px;
        }}
        .status-flow {{
            display: flex;
            width: 100%;
            height: 10px;
            border-radius: 999px;
            background: {BG_SECTION};
            overflow: hidden;
        }}
        .status-flow-fill.receipt {{
            background: {COLOR_TEAL};
        }}
        .status-flow-fill.shortage {{
            background: {COLOR_ORANGE};
        }}
        .status-flow-legend {{
            display: flex;
            justify-content: space-between;
            gap: 10px;
            flex-wrap: wrap;
            margin-top: 9px;
            color: {TEXT_SECONDARY};
            font-size: 11px;
            font-variant-numeric: tabular-nums;
        }}
        .status-tile-grid {{
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 8px;
        }}
        .status-tile {{
            min-height: 72px;
            padding: 11px 13px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }}
        .status-tile .metric-value {{
            font-size: 20px;
        }}
        @media (max-width: 1100px) {{
            .status-board {{
                grid-template-columns: 1fr;
            }}
            .status-tile-grid {{
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }}
        }}
        @media (max-width: 640px) {{
            .status-tile-grid {{
                grid-template-columns: 1fr;
            }}
            .status-flow-legend {{
                flex-direction: column;
                gap: 4px;
            }}
        }}
        .panel-box {{
            background: {BG_CARD};
            border: 0.5px solid {BORDER_DEFAULT};
            border-radius: 8px;
            padding: 12px 14px;
            box-shadow: none;
        }}
        .family-section {{
            display: flex;
            flex-direction: column;
            gap: 8px;
        }}
        .family-section + .family-section {{
            margin-top: 16px;
        }}
        .family-section-title {{
            display: inline-block;
            width: fit-content;
            color: #444441;
            background: {BG_SECTION};
            font-size: 13px;
            font-weight: 500;
            line-height: 1.2;
            padding: 5px 12px;
            border-radius: 8px;
        }}
        .family-grid {{
            display: grid;
            grid-template-columns: repeat(5, minmax(0, 1fr));
            gap: 8px;
        }}
        .family-card {{
            border: 0.5px solid {BORDER_DEFAULT};
            border-radius: 8px;
            background: {BG_CARD};
            padding: 11px 13px;
            min-height: 106px;
            display: flex;
            flex-direction: column;
            gap: 8px;
        }}
        .family-card:has(.progress-fill.production.risk),
        .family-card:has(.progress-fill.production.warn) {{
            border-color: {COLOR_ALERT_BD};
        }}
        .family-head {{
            display: flex;
            align-items: baseline;
            justify-content: space-between;
            gap: 10px;
        }}
        .family-head span {{
            color: {TEXT_PRIMARY};
            font-size: 12px;
            font-weight: 500;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }}
        .family-head b {{
            color: {TEXT_TERTIARY};
            font-size: 11px;
            font-weight: 400;
            font-variant-numeric: tabular-nums;
            white-space: nowrap;
        }}
        .family-shortages {{
            display: flex;
            justify-content: flex-start;
            gap: 8px;
            color: {TEXT_SECONDARY};
            font-size: 10px;
        }}
        .family-shortages b {{
            color: {COLOR_ORANGE};
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
            gap: 12px;
            align-items: center;
            border-bottom: 0.5px solid {BORDER_LIGHT};
            padding: 10px 16px;
            background: {BG_CARD};
        }}
        .top-rank {{
            font-size: 12px;
            font-weight: 500;
            color: {TEXT_TERTIARY};
            text-align: center;
        }}
        .top-name {{
            color: {TEXT_PRIMARY};
            font-size: 13px;
            font-weight: 400;
            overflow-wrap: anywhere;
        }}
        .top-shortage {{
            color: {COLOR_ORANGE};
            font-size: 13px;
            font-weight: 500;
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
            gap: 12px;
            align-items: center;
            border-bottom: 0.5px solid {BORDER_LIGHT};
            padding: 10px 16px;
            background: {BG_CARD};
        }}
        .gap-progress {{
            min-width: 0;
        }}
        .gap-value {{
            color: {COLOR_ORANGE};
            font-size: 13px;
            font-weight: 500;
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
            border: 0.5px solid {BORDER_DEFAULT};
            border-radius: 8px;
            background: {BG_CARD};
        }}
        .compact-table {{
            max-height: 360px;
        }}
        .ops-table {{
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
        }}
        .ops-table th {{
            position: sticky;
            top: 0;
            background: {BG_PAGE};
            color: {TEXT_SECONDARY};
            font-size: 11px;
            font-weight: 500;
            border-bottom: 1px solid {BORDER_DEFAULT};
            padding: 8px 12px;
            z-index: 1;
        }}
        .ops-table td {{
            border-bottom: 0.5px solid {BORDER_LIGHT};
            padding: 7px 12px;
            font-size: 12px;
            color: {TEXT_PRIMARY};
            vertical-align: middle;
            background: {BG_CARD};
        }}
        .ops-table tbody tr:hover td {{
            background: {BG_PAGE};
        }}
        .ops-table td.left, .ops-table th.left {{
            text-align: left;
        }}
        .ops-table td.num, .ops-table th.num {{
            text-align: right;
            font-variant-numeric: tabular-nums;
        }}
        .ops-table td.num.shortage {{
            color: {COLOR_ORANGE};
            font-weight: 500;
        }}
        .ops-table td.num.muted {{
            color: {TEXT_SECONDARY};
            font-weight: 500;
        }}
        .ops-table td.num.negative {{
            color: {COLOR_ORANGE};
            font-weight: 700;
        }}
        .ops-table td.code-cell {{
            color: {COLOR_BLUE};
            font-weight: 700;
            font-variant-numeric: tabular-nums;
        }}
        .request-scope-badge {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            border-radius: 999px;
            padding: 3px 8px;
            font-size: 11px;
            font-weight: 700;
            border: 0.5px solid {BORDER_DEFAULT};
            background: {BG_SECTION};
            color: {TEXT_SECONDARY};
            white-space: nowrap;
        }}
        .request-scope-badge.in {{
            color: {COLOR_TEAL};
            background: #E8F5F0;
            border-color: #B9E3D4;
        }}
        .request-scope-badge.out {{
            color: {COLOR_ORANGE};
            background: {COLOR_ALERT_BG};
            border-color: {COLOR_ALERT_BD};
        }}
        .request-scope-badge.mixed {{
            color: {COLOR_AMBER};
            background: #F7EFE3;
            border-color: #E4B968;
        }}
        .response-badge {{
            display: inline-flex;
            align-items: center;
            min-width: 72px;
            justify-content: center;
            border-radius: 999px;
            padding: 3px 8px;
            font-size: 11px;
            font-weight: 600;
            background: {BG_SECTION};
            color: {TEXT_SECONDARY};
        }}
        .response-badge.partial {{
            background: #F7EFE3;
            color: {COLOR_AMBER};
        }}
        .response-badge.need {{
            background: {COLOR_ALERT_BG};
            color: {COLOR_ORANGE};
        }}
        .ops-table td.power-cell {{
            text-align: center;
            font-variant-numeric: tabular-nums;
            font-weight: 500;
            color: {TEXT_PRIMARY};
        }}
        .ops-table td.power-cell.high {{
            color: {TEXT_PRIMARY};
        }}
        .progress-cell {{
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .progress-name {{
            min-width: 28px;
            color: {TEXT_SECONDARY};
            font-size: 11px;
            font-weight: 500;
        }}
        .progress-track {{
            flex: 1;
            min-width: 80px;
            height: 5px;
            border-radius: 3px;
            background: {BG_SECTION};
            overflow: hidden;
        }}
        .progress-fill {{
            height: 100%;
            border-radius: 3px;
        }}
        .progress-fill.done {{
            background: {COLOR_TEAL};
        }}
        .progress-fill.active {{
            background: {COLOR_TEAL};
        }}
        .progress-fill.warn {{
            background: {COLOR_AMBER};
        }}
        .progress-fill.risk {{
            background: {COLOR_ORANGE};
        }}
        .progress-fill.production {{
            background: {COLOR_BLUE};
        }}
        .progress-fill.receipt {{
            background: {COLOR_TEAL};
        }}
        .progress-fill.risk.receipt {{
            background: {TEXT_TERTIARY};
        }}
        .progress-text {{
            min-width: 52px;
            text-align: right;
            font-size: 11px;
            color: {TEXT_TERTIARY};
            font-variant-numeric: tabular-nums;
        }}
        .status-badge {{
            display: inline-block;
            padding: 3px 8px;
            border-radius: 6px;
            border: 1px solid transparent;
            font-size: 11px;
            font-weight: 400;
            line-height: 1.2;
        }}
        .status-badge.done {{
            background: #EAF3DE;
            color: #3B6D11;
        }}
        .status-badge.active {{
            background: #E6F1FB;
            color: {COLOR_BLUE};
        }}
        .status-badge.warn {{
            background: {COLOR_ALERT_BG};
            color: #993C1D;
            font-weight: 500;
        }}
        .status-badge.waiting {{
            background: #F1EFE8;
            color: #5F5E5A;
        }}
        .status-badge.risk {{
            background: {COLOR_ALERT_BG};
            color: #993C1D;
        }}
        .section-title {{
            color: {TEXT_PRIMARY};
            font-weight: 500;
            font-size: 15px;
            margin-bottom: 4px;
        }}
        .section-sub {{
            color: {TEXT_SECONDARY};
            font-size: 12px;
            font-weight: 400;
            margin-bottom: 10px;
        }}
        .breadcrumb {{
            display: flex;
            gap: 8px;
            align-items: center;
            color: {TEXT_SECONDARY};
            font-size: 12px;
            margin: 2px 0 10px 0;
        }}
        .breadcrumb span {{
            color: {COLOR_BLUE};
            font-weight: 500;
        }}
        .breadcrumb b {{
            color: {TEXT_SECONDARY};
        }}
        .progress-summary-panel {{
            display: grid;
            grid-template-columns: minmax(0, 1fr) 160px;
            gap: 14px;
            align-items: stretch;
            background: {BG_CARD};
            border: 0.5px solid {BORDER_DEFAULT};
            border-radius: 12px;
            padding: 16px 20px;
            margin-bottom: 12px;
        }}
        .progress-summary-panel .progress-cell {{
            margin: 10px 0;
        }}
        .dday-box {{
            border: 0.5px solid {BORDER_DEFAULT};
            border-radius: 12px;
            padding: 12px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            background: {BG_PAGE};
        }}
        .dday-value {{
            color: {COLOR_ORANGE};
            font-size: 20px;
            font-weight: 500;
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
    out["용마입고율"] = out["용마입고율"].map(lambda x: f"{float(x):.1f}%")
    return out[
        [
            "판매코드",
            "요청 PACK",
            "포장 PACK",
            "부족 PACK",
            "생산진도율",
            "용마입고율",
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
            "용마입고율",
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
        "생산요청물량": st.column_config.NumberColumn("생산요청물량", format=numeric_format),
        "생산요청물량(PACK)": st.column_config.NumberColumn("생산요청물량(PACK)", format=numeric_format),
        "생산요청물량(PCS)": st.column_config.NumberColumn("생산요청물량(PCS)", format=numeric_format),
        "용마창고재고 (PACK)": st.column_config.NumberColumn("용마창고재고 (PACK)", format=numeric_format),
        "총수량(PACK)": st.column_config.NumberColumn("총수량(PACK)", format=numeric_format),
        "재고기준(PACK)": st.column_config.NumberColumn("재고기준(PACK)", format=numeric_format),
        "재고부족(PACK)": st.column_config.NumberColumn("재고부족(PACK)", format=numeric_format),
        "용마입고 PACK": st.column_config.NumberColumn("용마입고 PACK", format=numeric_format),
        "용마입고": st.column_config.NumberColumn("용마입고", format=numeric_format),
        "미입고": st.column_config.NumberColumn("미입고", format=numeric_format),
        "미입고 PACK": st.column_config.NumberColumn("미입고 PACK", format=numeric_format),
        "용마입고수량": st.column_config.NumberColumn("용마입고수량", format=numeric_format),
        "용마입고수량(PACK)": st.column_config.NumberColumn("용마입고수량(PACK)", format=numeric_format),
        "용마입고수량(PCS)": st.column_config.NumberColumn("용마입고수량(PCS)", format=numeric_format),
        "용마입고대기 PACK": st.column_config.NumberColumn("용마입고대기 PACK", format=numeric_format),
        "용마입고대기수량": st.column_config.NumberColumn("용마입고대기수량", format=numeric_format),
        "용마입고대기수량(PACK)": st.column_config.NumberColumn("용마입고대기수량(PACK)", format=numeric_format),
        "용마입고대기수량(PCS)": st.column_config.NumberColumn("용마입고대기수량(PCS)", format=numeric_format),
        "포장가능재고(PCS)": st.column_config.NumberColumn("포장가능재고(PCS)", format=numeric_format),
        "샘플신청가능수량": st.column_config.NumberColumn("샘플신청가능수량", format=numeric_format),
        "순위": st.column_config.NumberColumn("순위", format="%d", width="small"),
        "현재 재고수량": st.column_config.NumberColumn("현재 재고수량", format=numeric_format),
        "부족수량": st.column_config.NumberColumn("부족수량", format=numeric_format),
        "상세 건수": st.column_config.NumberColumn("상세 건수", format=numeric_format),
        "긴급요청 수": st.column_config.NumberColumn("긴급요청 수", format=numeric_format),
        "미입고(PACK)": st.column_config.NumberColumn("미입고(PACK)", format=numeric_format),
        "미입고수량": st.column_config.NumberColumn("미입고수량", format=numeric_format),
        "입고대기수량": st.column_config.NumberColumn("입고대기수량", format=numeric_format),
        "제품필요수량": st.column_config.NumberColumn("제품필요수량", format=numeric_format),
        "생산필요수량(PCS)": st.column_config.NumberColumn("생산필요수량(PCS)", format=numeric_format),
        "생산부족 PCS": st.column_config.NumberColumn("생산부족 PCS", format=numeric_format),
        "생산부족수량(PCS)": st.column_config.NumberColumn("생산부족수량(PCS)", format=numeric_format),
        "포장부족(PACK)": st.column_config.NumberColumn("포장부족(PACK)", format=numeric_format),
        "포장부족(PCS)": st.column_config.NumberColumn("포장부족(PCS)", format=numeric_format),
        "포장부족(재고 PCS)": st.column_config.NumberColumn("포장부족(재고 PCS)", format=numeric_format),
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
        "포장 PACK": st.column_config.NumberColumn("포장 PACK", format=numeric_format),
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
        "생산부족(PCS)": st.column_config.NumberColumn("생산부족(PCS)", format=numeric_format),
        "생산부족수량": st.column_config.NumberColumn("생산부족수량", format=numeric_format),
        "포장부족수량": st.column_config.NumberColumn("포장부족수량", format=numeric_format),
        "생산진도율": st.column_config.ProgressColumn("생산진도율", min_value=0, max_value=100, format="%.1f%%"),
        "용마입고율": st.column_config.ProgressColumn("용마입고율", min_value=0, max_value=100, format="%.1f%%"),
        "포장수량": st.column_config.NumberColumn("포장수량", format=numeric_format),
        "부족수량": st.column_config.NumberColumn("부족수량", format=numeric_format),
        "포장진도율": st.column_config.ProgressColumn("포장진도율", min_value=0, max_value=100, format="%.1f%%"),
        "GAP": st.column_config.NumberColumn("GAP", format="%.1f"),
        "power_value": None,
        "_power_sort": None,
        "_production_code_prefix": None,
        "_min_due_date_sort": None,
        "_priority_sort": None,
        "_request_due_date_sort": None,
        "_pack_sort": None,
        "_daily_item_code_base": None,
        "_daily_status_sort": None,
        "_daily_negative_sort": None,
        "_daily_min_due_date_sort": None,
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
    column_order: list[str] | None = None,
) -> pd.Series | None:
    render_panel_title(title, sub)
    st.markdown("<div class='panel-box drill-panel'>", unsafe_allow_html=True)
    if df.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
        st.markdown("</div>", unsafe_allow_html=True)
        return None
    display_df = dataframe_for_streamlit(df)
    column_config = drilldown_column_config()
    for col in display_df.columns:
        if re.match(r"^\d+(?:\.\d+)?P(?:\(PCS\))?$", str(col)):
            column_config[col] = st.column_config.NumberColumn(str(col), format="%,.0f")
    event = st.dataframe(
        display_df,
        hide_index=True,
        height=height,
        width="stretch",
        column_config=column_config,
        column_order=visible_columns(display_df, column_order) if column_order is not None else None,
        on_select="rerun",
        selection_mode="single-row",
        key=key,
    )
    st.markdown("</div>", unsafe_allow_html=True)
    return get_selected_row(event, df)


def render_production_power_detail_dialog(
    selected_row: pd.Series,
    detail_view: pd.DataFrame,
    pack_labels: list[str],
    table_nonce_key: str,
) -> None:
    production_code = clean_str(selected_row.get("생산코드", ""))
    product_name = clean_str(selected_row.get("대표 제품명", ""))
    title = f"생산코드 {production_code} 상세 - {product_name}"

    @st.dialog(title, width="large")
    def _dialog() -> None:
        st.caption(
            f"{production_code}에 해당하는 POWER별 PACK 단위 수량, 부족수량, 진도율, 납기 현황 | 표시 건수: {len(detail_view):,}"
        )
        if detail_view.empty:
            st.warning("상세 데이터가 없습니다.")
        else:
            st.dataframe(
                dataframe_for_streamlit(detail_view),
                hide_index=True,
                height=520,
                width="stretch",
                column_config=drilldown_column_config(),
                column_order=production_power_detail_column_order(detail_view, pack_labels),
            )
        if st.button("닫기", key="close_production_power_detail_dialog", width="stretch"):
            st.session_state[table_nonce_key] = int(st.session_state.get(table_nonce_key, 0)) + 1
            st.rerun()

    _dialog()


def render_daily_inventory_detail_dialog(
    selected_row: pd.Series,
    detail_view: pd.DataFrame,
    table_nonce_key: str,
) -> None:
    item_code = clean_str(selected_row.get("_daily_item_code_base", selected_row.get("품목코드", "")))
    product_name = clean_str(selected_row.get("대표 제품명", selected_row.get("제품명", "")))
    title = f"품목코드 {item_code} 상세 - {product_name}"

    @st.dialog(title, width="large")
    def _dialog() -> None:
        st.caption(f"{item_code}에 해당하는 PACK/POWER별 재고 대응 상세 | 표시 건수: {len(detail_view):,}")
        if detail_view.empty:
            st.warning("상세 데이터가 없습니다.")
        else:
            st.dataframe(
                dataframe_for_streamlit(detail_view),
                hide_index=True,
                height=520,
                width="stretch",
                column_config=drilldown_column_config(),
                column_order=daily_inventory_detail_column_order(detail_view),
            )
        if st.button("닫기", key="close_daily_inventory_detail_dialog", width="stretch"):
            st.session_state[table_nonce_key] = int(st.session_state.get(table_nonce_key, 0)) + 1
            st.rerun()

    _dialog()


def render_daily_inventory_tab(
    daily_inventory_df: pd.DataFrame,
    code_summary: pd.DataFrame,
    sample_available_df: pd.DataFrame | None = None,
    lot_status_df: pd.DataFrame | None = None,
) -> None:
    render_panel_title(
        "일일 재고 대응",
        "매일 공유되는 재고현황표 기준으로 요청물량 외 긴급 품목과 재고부족을 확인합니다.",
    )
    if daily_inventory_df.empty:
        st.warning("일일 재고현황표를 찾지 못했거나 처리할 데이터가 없습니다.")
        return

    response_view = build_daily_inventory_response_view(daily_inventory_df, code_summary, sample_available_df, lot_status_df)
    if response_view.empty:
        st.warning("표시할 일일 재고 대응 데이터가 없습니다.")
        return

    urgent_count = int(response_view["긴급요청"].sum())
    negative_count = int((response_view["재고수량"] < 0).sum())
    request_out_count = int((response_view["대응상태"] == "요청외 긴급").sum())
    request_in_count = int(response_view["대응상태"].isin(["요청내 긴급", "요청내 재고부족"]).sum())
    shortage_qty = float(response_view["재고부족수량"].sum())
    render_metric_card_grid(
        [
            ("긴급요청 품목", f"{urgent_count:,}", "warn" if urgent_count else "normal"),
            ("요청외 긴급", f"{request_out_count:,}", "warn" if request_out_count else "normal"),
            ("요청내 부족/긴급", f"{request_in_count:,}", "warn" if request_in_count else "normal"),
            (
                "재고부족수량",
                format_int(shortage_qty),
                "warn" if shortage_qty > 0 else "normal",
                f"음수 {negative_count:,}품목" if negative_count else "",
            ),
        ]
    )

    f1, f2, f3 = st.columns([2.4, 1.7, 1.2], gap="small")
    with f1:
        query = st.text_input(
            "제품명/제품코드/POWER 검색",
            value="",
            placeholder="예: 소울브라운, 40P, -06.50",
            key="daily_inventory_query",
        )
    with f2:
        statuses = st.multiselect(
            "대응상태",
            sorted(response_view["대응상태"].dropna().astype(str).unique().tolist()),
            default=sorted(response_view["대응상태"].dropna().astype(str).unique().tolist()),
            key="daily_inventory_status",
        )
    with f3:
        important_only = st.checkbox("긴급/부족만 보기", value=True, key="daily_inventory_important_only")

    view = response_view.copy()
    if query.strip():
        view = view[daily_inventory_query_mask(view, query)].copy()
    if statuses:
        view = view[view["대응상태"].isin(statuses)].copy()
    if important_only:
        view = view[(view["긴급요청"]) | (view["재고수량"] < 0) | (view["재고부족수량"] > 0)].copy()

    hidden_daily_inventory_cols = [
        "재고표 제품명",
        "전일재고",
        "재고증감",
        "재고부족수량",
        "요청제품명",
        "판매코드 수",
        "대상품목",
        "포장부족(재고 PCS)",
        "포장 PACK",
        "미입고 PACK",
    ]
    detail_view = view.drop(columns=hidden_daily_inventory_cols, errors="ignore")
    main_view = build_daily_inventory_main_view(view)
    full_export_view = response_view.drop(columns=hidden_daily_inventory_cols, errors="ignore")

    dl_col, _ = st.columns([1.2, 4.8], gap="small")
    with dl_col:
        render_excel_download(
            "엑셀 다운로드",
            "일일_재고_대응",
            {
                "일일 재고 대응": main_view,
                "일일 재고 상세": detail_view,
                "일일 재고 전체": full_export_view,
            },
            key="download_daily_inventory_excel",
        )

    table_nonce_key = "daily_inventory_main_table_nonce"
    table_nonce = int(st.session_state.get(table_nonce_key, 0))
    selected_daily_row = render_selectable_table(
        "일일 재고 대응 테이블",
        f"품목코드 S### 기준 집계 | 표시 건수: {len(main_view):,} | 상세 건수: {len(view):,}",
        main_view,
        key=f"daily_inventory_table_{table_nonce}",
        height=560,
        column_order=[
            "대응상태",
            "품목코드",
            "대표 제품명",
            "상세 건수",
            "긴급요청 수",
            "재고수량",
            "재고부족수량",
            "요청 PACK",
            "용마입고 PACK",
            "용마입고대기 PACK",
            "포장가능재고(PCS)",
            "생산부족 PCS",
            "생산진도율",
            "최소 납기",
        ],
    )
    if selected_daily_row is not None:
        selected_item_code = clean_str(
            selected_daily_row.get("_daily_item_code_base", selected_daily_row.get("품목코드", ""))
        )
        detail_scope = detail_view[
            detail_view["품목코드"].map(daily_item_code_base) == selected_item_code
        ].copy()
        detail_scope["_daily_status_sort"] = detail_scope["대응상태"].map(daily_inventory_status_rank)
        detail_scope["_daily_pack_sort"] = detail_scope["PACK"].map(pack_sort_key)
        detail_scope["_daily_power_sort"] = pd.to_numeric(
            detail_scope["POWER"].astype(str).str.replace("-00.00", "0", regex=False),
            errors="coerce",
        ).fillna(0.0)
        detail_scope["_daily_stock_shortage_sort"] = pd.to_numeric(
            detail_scope.get("재고수량", pd.Series(0.0, index=detail_scope.index)),
            errors="coerce",
        ).fillna(0.0)
        detail_scope = detail_scope.sort_values(
            ["_daily_status_sort", "_daily_stock_shortage_sort", "_daily_pack_sort", "_daily_power_sort"],
            ascending=[True, True, True, True],
            kind="stable",
        ).drop(
            columns=[
                "_daily_status_sort",
                "_daily_pack_sort",
                "_daily_power_sort",
                "_daily_stock_shortage_sort",
            ],
            errors="ignore",
        )
        render_daily_inventory_detail_dialog(selected_daily_row, detail_scope, table_nonce_key)


def render_product_summary_tab(
    product_summary: pd.DataFrame,
    code_summary: pd.DataFrame,
    daily_inventory_df: pd.DataFrame | None = None,
    sample_available_df: pd.DataFrame | None = None,
) -> None:
    main_products, _ = split_main_sample(product_summary)
    stock_threshold_pack = float(INVENTORY_STOCK_THRESHOLD_DEFAULT)

    family_view = build_family_progress_view(main_products)
    top_shortage_view = build_top_shortage_view(product_summary, top_n=10)
    gap_top_view = build_gap_top_view(product_summary, top_n=10)
    exception_kpis, exception_detail = build_daily_exception_report_view(
        daily_inventory_df,
        code_summary,
        sample_available_df,
        max_rows=10,
    )
    urgent_summary_view = build_urgent_request_summary_view(
        daily_inventory_df,
        code_summary,
        sample_available_df,
    )

    title_col, download_col = st.columns([4.8, 1.2], gap="small", vertical_alignment="center")
    with title_col:
        render_panel_title(
            "제품 진도 현황",
            "국내 요청 물량이 생산 → 포장 → 용마 입고까지 정상적으로 진행되고 있는지 확인하고, 부족 및 지연 품목을 우선 대응하기 위한 화면입니다.",
        )
    with download_col:
        ppt_bytes = build_ppt_report(
            product_view=product_summary,
            code_summary=code_summary,
            product_names=product_summary["제품명"],
            scope_label="전체",
            daily_inventory_df=daily_inventory_df,
            sample_available_df=sample_available_df,
        )
        st.download_button(
            "PPT 보고서 다운로드",
            data=ppt_bytes,
            file_name=f"국내_제품_포장현황_운영보고서_{pd.Timestamp.now(tz='Asia/Seoul').strftime('%Y%m%d_%H%M')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            width="stretch",
            key="download_ppt_report",
        )
        render_excel_download(
            "엑셀 다운로드",
            "제품_진도_현황",
            {
                "제품 요약": product_summary,
                "미입고 TOP10": top_shortage_view,
                "본품 분류별 진도": family_view,
                "생산완료 후 미입고 TOP10": gap_top_view,
                "요청 긴급 요약": urgent_summary_view,
                "요청 긴급 상세": exception_detail,
            },
            key="download_product_progress_excel",
        )

    render_status_board(
        product_summary,
        code_summary,
        daily_inventory_df,
        sample_available_df,
        stock_threshold_pack,
    )
    render_kpi_scope_panels(code_summary)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    render_panel_title(
        "본품 분류별 진도현황",
        "제품군별 요청 PACK, 생산진도율, 용마입고율, 생산부족 PCS를 비교합니다.",
    )
    st.markdown("<div class='panel-box'>", unsafe_allow_html=True)
    render_family_progress_cards(family_view)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    render_panel_title(
        "미입고 TOP10",
        "미입고 PACK이 큰 제품의 생산·포장·입고 진도를 확인합니다.",
    )
    st.markdown("<div class='panel-box'>", unsafe_allow_html=True)
    render_top_shortage_list(top_shortage_view)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    render_panel_title(
        "생산완료 후 미입고 TOP10",
        "생산은 진행됐지만 용마 입고가 지연되는 제품을 GAP 기준으로 표시합니다.",
    )
    st.markdown("<div class='panel-box'>", unsafe_allow_html=True)
    render_gap_top_list(gap_top_view)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    urgent_sku_count = int(
        pd.to_numeric(urgent_summary_view.get("SKU 수", pd.Series(dtype=float)), errors="coerce")
        .fillna(0)
        .sum()
    )
    render_panel_title(
        "요청 긴급 요약",
        f"일일 재고표 기준 긴급요청 S코드 {len(urgent_summary_view):,}개 / SKU {urgent_sku_count:,}개를 확인합니다.",
    )
    st.markdown("<div class='panel-box drill-panel'>", unsafe_allow_html=True)
    render_urgent_request_summary_table(urgent_summary_view)
    st.markdown("</div>", unsafe_allow_html=True)


def render_production_code_tab(code_summary: pd.DataFrame) -> None:
    render_panel_title(
        "생산코드 상세",
        "P로 시작하는 생산코드 5자리 기준으로 제품군 위험도를 확인하고, 선택 시 POWER별 상세를 팝업으로 확인합니다.",
    )
    production_unit_mode = UNIT_PACK
    pack_options = available_pack_options(code_summary)
    pack_labels = PRODUCTION_CODE_PACK_LABELS
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
        pack_label=selected_pack,
        sample_scope=sample_scope,
        product_group=selected_group,
    )
    production_view = build_production_power_main_view(
        production_source,
        pack_labels=pack_labels,
        shortage_only=shortage_only,
    )
    production_detail_view = build_production_power_detail_view(
        production_source,
        pack_labels=pack_labels,
    )
    render_production_power_kpis(production_view, unit_mode=production_unit_mode)
    production_main_export = production_view[
        production_progress_column_order(production_view, pack_labels, production_unit_mode)
    ].copy()
    production_detail_export = production_detail_view[
        production_power_detail_column_order(production_detail_view, pack_labels)
    ].copy()
    dl_col, _ = st.columns([1.2, 4.8], gap="small")
    with dl_col:
        render_excel_download(
            "엑셀 다운로드",
            "생산코드_상세",
            {
                "생산코드 집계": production_main_export,
                "POWER 상세": production_detail_export,
            },
            key="download_production_code_excel",
        )

    table_nonce_key = "production_code_main_table_nonce"
    table_nonce = int(st.session_state.get(table_nonce_key, 0))
    selected_production_row = render_selectable_table(
        "생산코드 메인 테이블",
        f"P로 시작하는 생산코드 5자리 기준 집계 | 납기일, 포장부족, 생산부족 순 정렬 | 표시 건수: {len(production_view):,}",
        production_view,
        key=f"production_code_main_table_{table_nonce}",
        height=620,
        column_order=production_progress_column_order(production_view, pack_labels, production_unit_mode),
    )
    if selected_production_row is None:
        return

    selected_production = clean_str(selected_production_row.get("_production_code_prefix", selected_production_row.get("생산코드", "")))
    detail_view = build_production_power_detail_view(
        production_source,
        pack_labels=pack_labels,
        production_prefix=selected_production,
    )
    render_production_power_detail_dialog(
        selected_production_row,
        detail_view,
        pack_labels,
        table_nonce_key,
    )


def render_sales_code_tab(code_summary: pd.DataFrame) -> None:
    render_panel_title(
        "판매코드 상세",
        "출고/오더 관점에서 판매코드별 생산·포장 진도와 납기 상태를 확인합니다.",
    )
    sales_unit_mode = render_unit_selector("sales_progress_unit_mode")
    pack_options = available_pack_options(code_summary)
    power_options = available_power_options(code_summary)

    threshold_col, _ = st.columns([1.2, 4.8], gap="small")
    with threshold_col:
        stock_threshold_pack = st.number_input(
            "긴급 재고 기준(PACK)",
            min_value=0,
            value=INVENTORY_STOCK_THRESHOLD_DEFAULT,
            step=10,
            key="sales_inventory_stock_threshold_pack",
        )

    urgent_sales_base = build_sales_order_main_view(
        code_summary,
        stock_threshold_pack=float(stock_threshold_pack),
    )
    render_urgent_sales_packing_list(urgent_sales_base)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    sf1, sf2, sf3, sf4, sf5 = st.columns([1.9, 1.6, 1.6, 1.2, 1.2], gap="small")
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

    sales_source = filter_operational_code_summary(
        code_summary,
        product_query=product_query,
        production_query=production_query,
        sales_query=sales_query,
        pack_label=selected_pack,
        power_label=selected_power,
    )
    sales_view = build_sales_order_main_view(
        sales_source,
        stock_threshold_pack=float(stock_threshold_pack),
    )
    dl_col, _ = st.columns([1.2, 4.8], gap="small")
    with dl_col:
        render_excel_download(
            "엑셀 다운로드",
            "판매코드_상세",
            {
                "긴급 포장 리스트": urgent_sales_base,
                "판매코드": sales_view,
            },
            key="download_sales_code_excel",
        )
    selected_sales_row = render_selectable_table(
        "판매코드",
        f"판매코드 기준 출고/오더 상세 | 표시 건수: {len(sales_view):,}",
        sales_view.drop(columns=["power_value"], errors="ignore"),
        key="sales_code_main_table",
        height=620,
        column_order=sales_progress_column_order(sales_view, sales_unit_mode),
    )
    if selected_sales_row is None:
        return

    selected_sales = str(selected_sales_row["판매코드"])
    st.markdown(f"<div class='breadcrumb'>판매코드 <span>{escape(selected_sales)}</span></div>", unsafe_allow_html=True)
    inventory_view = build_inventory_detail_view(code_summary, selected_sales)
    render_panel_title("WMS 재고 상세", "용마WMS재고현황 기준 판매코드별 PACK 재고")
    st.markdown("<div class='panel-box drill-panel'>", unsafe_allow_html=True)
    if inventory_view.empty or set(inventory_view["매칭여부"].astype(str)) == {"미매칭"}:
        st.info("선택한 판매코드와 매칭되는 WMS 재고가 없습니다.")
    st.dataframe(inventory_view, hide_index=True, height=120, width="stretch")
    st.markdown("</div>", unsafe_allow_html=True)


def render_power_tab(code_summary: pd.DataFrame) -> None:
    render_panel_title(
        "POWER 상세",
        "렌즈 POWER 기준 요청/생산/포장/부족 현황과 하위 생산·판매코드를 확인합니다.",
    )
    power_unit_mode = render_unit_selector("power_progress_unit_mode")
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
    if power_unit_mode == UNIT_PCS:
        request_pcs = float(power_summary["요청합계(PCS)"].sum()) if not power_summary.empty else 0.0
        production_shortage_pcs = (
            float(power_summary["생산부족수량(PCS)"].sum()) if not power_summary.empty else 0.0
        )
        request_pack = float(power_summary["요청합계(PACK)"].sum()) if not power_summary.empty else 0.0
        packing_shortage_pack = float(power_summary["포장부족(PACK)"].sum()) if not power_summary.empty else 0.0
        production_progress = (
            (request_pcs - production_shortage_pcs) / request_pcs * 100.0
            if request_pcs > 0
            else 0.0
        )
        packing_progress = (
            (request_pack - packing_shortage_pack) / request_pack * 100.0
            if request_pack > 0
            else 0.0
        )
        production_progress = min(100.0, max(0.0, production_progress))
        packing_progress = min(100.0, max(0.0, packing_progress))
        render_metric_card_grid(
            [
                ("대상 도수", f"{ops_kpi['rows']:,}", "normal"),
                ("요청합계(PCS)", format_int(request_pcs), "normal"),
                ("생산부족수량(PCS)", format_int(production_shortage_pcs), "warn"),
                ("생산진도율", f"{production_progress:.1f}%", metric_progress_tone(production_progress)),
                ("포장진도율", f"{packing_progress:.1f}%", metric_progress_tone(packing_progress)),
            ]
        )
    else:
        render_metric_card_grid(
            [
                ("대상 도수", f"{ops_kpi['rows']:,}", "normal"),
                ("포장부족 도수", f"{ops_kpi['shortage_rows']:,}", "warn"),
                ("미착수 도수", f"{ops_kpi['not_started_rows']:,}", "warn"),
                ("하이파워 부족", f"{ops_kpi['high_power_shortage_rows']:,}", "warn"),
                ("포장부족(PACK) 합계", format_int(ops_kpi["shortage_qty"]), "warn"),
            ]
        )

    dl_col, _ = st.columns([1.2, 4.8], gap="small")
    with dl_col:
        render_excel_download(
            "엑셀 다운로드",
            "POWER_상세",
            {
                "POWER 요약": power_summary,
                "POWER 상세": power_detail_for_heatmap,
            },
            key="download_power_excel",
        )

    selected_power_row = render_selectable_table(
        "POWER 히트맵 상세",
        f"POWER 기준 요청/생산/포장/부족 | 표시 건수: {len(power_summary):,}",
        power_summary.drop(columns=["power_value"], errors="ignore"),
        key="power_summary_table",
        height=430,
        column_order=power_progress_column_order(power_summary, power_unit_mode),
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
        column_order=visible_columns(
            sku_detail,
            [
                "생산코드",
                "판매코드",
                "제품명",
                "PACK",
                "요청합계(PACK)",
                "포장부족(PACK)",
                "생산부족수량(PCS)",
                "납기",
            ]
            if power_unit_mode == UNIT_PACK
            else [
                "생산코드",
                "판매코드",
                "제품명",
                "PACK",
                "요청합계(PCS)",
                "생산필요수량(PCS)",
                "생산부족수량(PCS)",
                "납기",
            ],
        ),
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


def file_fingerprint(path: Path | None) -> tuple[str, int, int] | None:
    if path is None:
        return None
    stat = path.stat()
    return (str(path.resolve()), int(stat.st_mtime_ns), int(stat.st_size))


def load_dashboard_data(
    request_fingerprint: tuple[str, int, int],
    packing_fingerprint: tuple[str, int, int],
    progress_fingerprint: tuple[str, int, int] | None,
    inventory_fingerprint: tuple[str, int, int] | None,
    daily_inventory_fingerprint: tuple[str, int, int] | None,
    cache_version: int,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    request_file = Path(request_fingerprint[0])
    packing_file = Path(packing_fingerprint[0])
    progress_file = Path(progress_fingerprint[0]) if progress_fingerprint is not None else None
    inventory_file = Path(inventory_fingerprint[0]) if inventory_fingerprint is not None else None
    daily_inventory_file = Path(daily_inventory_fingerprint[0]) if daily_inventory_fingerprint is not None else None

    request_df = normalize_request(request_file)
    packing_df, yongma_df, sample_available_df = normalize_packing_workbook(packing_file)
    inventory_df = normalize_inventory(inventory_file)
    daily_inventory_df = normalize_daily_inventory_file(daily_inventory_file)
    daily_inventory_df = enrich_daily_inventory_from_wms(daily_inventory_df, inventory_df)
    product_summary, _unmatched_packing_total, code_summary = build_summaries(request_df, packing_df, yongma_df)
    code_summary = attach_inventory_to_code_summary(code_summary, inventory_df)
    progress_df, _progress_info = normalize_progress(progress_file, request_df)
    production_progress_df = filter_progress_for_production_month(progress_df)
    product_summary = enrich_product_summary(product_summary, production_progress_df)
    code_summary = attach_progress_to_code_summary(code_summary, production_progress_df)
    code_summary = attach_sample_available_to_code_summary(code_summary, sample_available_df)
    product_summary = attach_inventory_to_product_summary(product_summary, code_summary)
    lot_status_df = build_lot_receipt_status_view(packing_df, yongma_df, code_summary)
    return product_summary, code_summary, lot_status_df, daily_inventory_df, sample_available_df


def render_dashboard_nav() -> str:
    selected = st.segmented_control(
        "대시보드 메뉴",
        options=DASHBOARD_TABS,
        default=DASHBOARD_TABS[0],
        label_visibility="collapsed",
        key="dashboard_active_tab",
    )
    st.markdown("<div class='dashboard-nav-divider'></div>", unsafe_allow_html=True)
    return str(selected or DASHBOARD_TABS[0])


def main() -> None:
    render_style()
    st.title("국내 생산·포장 현황 대시보드")

    base_dir = Path.cwd()
    try:
        files = discover_source_files(base_dir)
        product_summary, code_summary, lot_status_df, daily_inventory_df, sample_available_df = load_dashboard_data(
            file_fingerprint(files.request_file),
            file_fingerprint(files.packing_file),
            file_fingerprint(files.progress_file),
            file_fingerprint(files.inventory_file),
            file_fingerprint(files.daily_inventory_file),
            DATA_CACHE_VERSION,
        )
    except DashboardConfigError as exc:
        st.error("데이터 설정 오류")
        for msg in exc.messages:
            st.write(f"- {msg}")
        st.stop()
    except Exception as exc:
        st.error(f"처리 중 오류가 발생했습니다: {exc}")
        st.stop()

    active_tab = render_dashboard_nav()
    if active_tab == "제품 진도 현황":
        render_product_summary_tab(product_summary, code_summary, daily_inventory_df, sample_available_df)
    elif active_tab == "일일 재고 대응":
        render_daily_inventory_tab(daily_inventory_df, code_summary, sample_available_df, lot_status_df)
    elif active_tab == "생산코드 상세":
        render_production_code_tab(code_summary)
    elif active_tab == "판매코드 상세":
        render_sales_code_tab(code_summary)
    elif active_tab == "POWER 상세":
        render_power_tab(code_summary)
    elif active_tab == "포장 LOT 상세":
        render_packing_lot_tab(lot_status_df)


if __name__ == "__main__":
    main()
