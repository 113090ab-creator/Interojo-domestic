from __future__ import annotations

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pandas as pd


NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

SITE_COL = "설비 사이트 코드"
PROCESS_COL = "공정 코드"
SHORT_CODE_COL = "제품코드(5자리)"
SKU_COL = "제품 코드"
PRODUCT_NAME_COL = "제품 이름"
POWER_COL = "POWER"
ORDER_COL = "수요량(PCS)"
SHORTAGE_COL = "주문 대응 수량"
MIN_DUE_COL = "최소 납기일"

SITE_FILTER = "C관"
TARGET_PROCESSES = ("[10]사출조립", "[80]누수/규격검사")
INPUT_RE = re.compile(r"^수요정보_(\d{8})\.xlsx$")


def column_index(column_ref: str) -> int:
    index = 0
    for char in column_ref:
        index = index * 26 + ord(char.upper()) - ord("A") + 1
    return index - 1


def column_letters(cell_ref: str) -> str:
    return "".join(char for char in cell_ref if char.isalpha())


def read_shared_strings(xlsx: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(xlsx.read("xl/sharedStrings.xml"))
    except KeyError:
        return []

    return [
        "".join(text.text or "" for text in item.findall(".//a:t", NS))
        for item in root.findall("a:si", NS)
    ]


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.findall(".//a:t", NS))

    value = cell.find("a:v", NS)
    if value is None:
        return ""

    text = value.text or ""
    if cell_type == "s" and text:
        return shared_strings[int(text)]
    return text


def read_xlsx_raw(path: Path) -> pd.DataFrame:
    """Read sheet1 raw values so quantity cells formatted as dates stay numeric."""
    with zipfile.ZipFile(path) as xlsx:
        shared_strings = read_shared_strings(xlsx)
        root = ET.fromstring(xlsx.read("xl/worksheets/sheet1.xml"))

    rows: list[list[str]] = []
    max_index = 0
    for row in root.findall("a:sheetData/a:row", NS):
        values_by_index: dict[int, str] = {}
        for cell in row.findall("a:c", NS):
            index = column_index(column_letters(cell.attrib["r"]))
            max_index = max(max_index, index)
            values_by_index[index] = cell_value(cell, shared_strings)
        rows.append([values_by_index.get(index, "") for index in range(max_index + 1)])

    if not rows:
        return pd.DataFrame()

    headers = rows[0]
    records = []
    for row in rows[1:]:
        if len(row) < len(headers):
            row = row + [""] * (len(headers) - len(row))
        records.append(row[: len(headers)])
    return pd.DataFrame(records, columns=headers)


def find_input_files(root: Path) -> list[tuple[str, Path]]:
    files = []
    for path in root.glob("수요정보_*.xlsx"):
        match = INPUT_RE.match(path.name)
        if match:
            files.append((match.group(1), path))
    return sorted(files)


def to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def validate_columns(path: Path, df: pd.DataFrame) -> None:
    required = {
        SITE_COL,
        PROCESS_COL,
        SHORT_CODE_COL,
        SKU_COL,
        PRODUCT_NAME_COL,
        POWER_COL,
        ORDER_COL,
        SHORTAGE_COL,
        MIN_DUE_COL,
    }
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"{path.name}: missing columns: {', '.join(sorted(missing))}")


def prepare_target_data(df: pd.DataFrame) -> pd.DataFrame:
    target = df[
        df[SITE_COL].astype(str).str.contains(SITE_FILTER, na=False)
        & df[PROCESS_COL].isin(TARGET_PROCESSES)
    ].copy()
    target[SKU_COL] = target[SKU_COL].astype(str).str.strip()
    target = target[target[SKU_COL].ne("") & target[SKU_COL].ne("총합계")]
    target[ORDER_COL] = to_number(target[ORDER_COL])
    target[SHORTAGE_COL] = to_number(target[SHORTAGE_COL])

    return (
        target.groupby([SKU_COL, PROCESS_COL], as_index=False)
        .agg(
            **{
                SITE_COL: (SITE_COL, "first"),
                SHORT_CODE_COL: (SHORT_CODE_COL, "first"),
                PRODUCT_NAME_COL: (PRODUCT_NAME_COL, "first"),
                POWER_COL: (POWER_COL, "first"),
                ORDER_COL: (ORDER_COL, "sum"),
                SHORTAGE_COL: (SHORTAGE_COL, "sum"),
                MIN_DUE_COL: (MIN_DUE_COL, "min"),
            }
        )
    )


def analyze_file(date: str, path: Path, df: pd.DataFrame, target: pd.DataFrame) -> dict[str, object]:
    records = df[df[SKU_COL].astype(str).ne("총합계")].copy()
    c_records = records[records[SITE_COL].astype(str).str.contains(SITE_FILTER, na=False)]
    process_counts = target[PROCESS_COL].value_counts()
    all_processes = ", ".join(records[PROCESS_COL].astype(str).dropna().unique()[:10])

    return {
        "파일일자": date,
        "파일명": path.name,
        "전체행수": len(records),
        "C관행수": len(c_records),
        "대상행수": len(target),
        "[10]행수": int(process_counts.get("[10]사출조립", 0)),
        "[80]행수": int(process_counts.get("[80]누수/규격검사", 0)),
        "대상여부": "대상" if not target.empty else "제외",
        "메모": "" if not target.empty else "C관 [10]/[80] 데이터 없음",
        "파일 내 공정": all_processes,
    }


def add_prefixed_columns(df: pd.DataFrame, prefix: str) -> pd.DataFrame:
    return df.rename(
        columns={
            SITE_COL: f"{prefix}_{SITE_COL}",
            SHORT_CODE_COL: f"{prefix}_{SHORT_CODE_COL}",
            PRODUCT_NAME_COL: f"{prefix}_{PRODUCT_NAME_COL}",
            POWER_COL: f"{prefix}_{POWER_COL}",
            ORDER_COL: f"{prefix}_주문수량",
            SHORTAGE_COL: f"{prefix}_부족수량",
            MIN_DUE_COL: f"{prefix}_{MIN_DUE_COL}",
        }
    )


def compare_dates(
    output_date: str,
    output_df: pd.DataFrame,
    compare_date: str,
    compare_df: pd.DataFrame,
) -> pd.DataFrame:
    output = add_prefixed_columns(output_df, "산출일")
    compare = add_prefixed_columns(compare_df, "비교일")
    merged = compare.merge(output, on=[SKU_COL, PROCESS_COL], how="outer")

    numeric_columns = [
        "비교일_주문수량",
        "산출일_주문수량",
        "비교일_부족수량",
        "산출일_부족수량",
    ]
    for column in numeric_columns:
        merged[column] = to_number(merged[column])

    merged.insert(0, "산출일", output_date)
    merged.insert(1, "비교일", compare_date)

    merged["주문증가분"] = (merged["산출일_주문수량"] - merged["비교일_주문수량"]).clip(lower=0)
    merged["주문감소분"] = (merged["비교일_주문수량"] - merged["산출일_주문수량"]).clip(lower=0)
    merged["부족감소분"] = (merged["비교일_부족수량"] - merged["산출일_부족수량"]).clip(lower=0)
    merged["부족증가분"] = (merged["산출일_부족수량"] - merged["비교일_부족수량"]).clip(lower=0)

    merged["보정전_유효생산량"] = (
        merged["비교일_부족수량"] + merged["주문증가분"] - merged["산출일_부족수량"]
    ).clip(lower=0)
    merged["유효생산량"] = (merged["보정전_유효생산량"] - merged["주문감소분"]).clip(lower=0)

    merged[SITE_COL] = merged["산출일_" + SITE_COL].combine_first(merged["비교일_" + SITE_COL])
    merged[SHORT_CODE_COL] = merged["산출일_" + SHORT_CODE_COL].combine_first(
        merged["비교일_" + SHORT_CODE_COL]
    )
    merged[PRODUCT_NAME_COL] = merged["산출일_" + PRODUCT_NAME_COL].combine_first(
        merged["비교일_" + PRODUCT_NAME_COL]
    )
    merged[POWER_COL] = merged["산출일_" + POWER_COL].combine_first(merged["비교일_" + POWER_COL])

    columns = [
        "산출일",
        "비교일",
        SITE_COL,
        PROCESS_COL,
        SHORT_CODE_COL,
        SKU_COL,
        PRODUCT_NAME_COL,
        POWER_COL,
        "비교일_주문수량",
        "산출일_주문수량",
        "주문증가분",
        "주문감소분",
        "비교일_부족수량",
        "산출일_부족수량",
        "부족감소분",
        "부족증가분",
        "보정전_유효생산량",
        "유효생산량",
    ]
    return merged[columns].sort_values(["산출일", PROCESS_COL, SKU_COL])


def summarize(details: pd.DataFrame) -> pd.DataFrame:
    if details.empty:
        return pd.DataFrame()

    return (
        details.groupby(["산출일", "비교일", PROCESS_COL], as_index=False)
        .agg(
            **{
                "SKU수": (SKU_COL, "count"),
                "비교일_주문수량": ("비교일_주문수량", "sum"),
                "산출일_주문수량": ("산출일_주문수량", "sum"),
                "주문증가분": ("주문증가분", "sum"),
                "주문감소분": ("주문감소분", "sum"),
                "비교일_부족수량": ("비교일_부족수량", "sum"),
                "산출일_부족수량": ("산출일_부족수량", "sum"),
                "부족감소분": ("부족감소분", "sum"),
                "부족증가분": ("부족증가분", "sum"),
                "보정전_유효생산량": ("보정전_유효생산량", "sum"),
                "유효생산량": ("유효생산량", "sum"),
            }
        )
        .sort_values(["산출일", PROCESS_COL])
    )


def build_adjacent_details(eligible: list[tuple[str, Path, pd.DataFrame]]) -> pd.DataFrame:
    frames = []
    for index in range(1, len(eligible)):
        compare_date, _, compare_df = eligible[index - 1]
        output_date, _, output_df = eligible[index]
        frames.append(compare_dates(output_date, output_df, compare_date, compare_df))
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def build_latest_base_details(eligible: list[tuple[str, Path, pd.DataFrame]]) -> pd.DataFrame:
    if len(eligible) < 2:
        return pd.DataFrame()

    output_date, _, output_df = eligible[-1]
    frames = [
        compare_dates(output_date, output_df, compare_date, compare_df)
        for compare_date, _, compare_df in eligible[:-1]
    ]
    return pd.concat(frames, ignore_index=True)


def main() -> None:
    root = Path(".")
    input_files = find_input_files(root)
    if not input_files:
        raise SystemExit("수요정보_YYYYMMDD.xlsx 파일이 없습니다.")

    file_analysis_rows = []
    eligible: list[tuple[str, Path, pd.DataFrame]] = []

    for date, path in input_files:
        df = read_xlsx_raw(path)
        validate_columns(path, df)
        target = prepare_target_data(df)
        file_analysis_rows.append(analyze_file(date, path, df, target))
        if not target.empty:
            eligible.append((date, path, target))

    latest_details = build_latest_base_details(eligible)
    latest_summary = summarize(latest_details)
    adjacent_details = build_adjacent_details(eligible)
    adjacent_summary = summarize(adjacent_details)
    file_analysis = pd.DataFrame(file_analysis_rows).sort_values("파일일자")

    latest_date = input_files[-1][0]
    output_path = root / f"유효생산량_일자별관리_{latest_date}.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        latest_summary.to_excel(writer, sheet_name="일자별요약", index=False)
        latest_details.to_excel(writer, sheet_name="일자별상세", index=False)
        adjacent_summary.to_excel(writer, sheet_name="직전대상일요약", index=False)
        adjacent_details.to_excel(writer, sheet_name="직전대상일상세", index=False)
        file_analysis.to_excel(writer, sheet_name="파일분석", index=False)

    print(f"output={output_path.name}")
    print(f"latest_base_date={latest_date}")
    print("대상 데이터 보유 일자:", ", ".join(date for date, _, _ in eligible) or "없음")
    if latest_summary.empty:
        print("일자별 산출 결과 없음")
    else:
        print(latest_summary.to_string(index=False))


if __name__ == "__main__":
    main()
