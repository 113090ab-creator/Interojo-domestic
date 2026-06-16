from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

import pandas as pd


NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

SITE_COL = "\uc124\ube44 \uc0ac\uc774\ud2b8 \ucf54\ub4dc"
PROCESS_COL = "\uacf5\uc815 \ucf54\ub4dc"
SHORT_CODE_COL = "\uc81c\ud488\ucf54\ub4dc(5\uc790\ub9ac)"
SKU_COL = "\uc81c\ud488 \ucf54\ub4dc"
PRODUCT_NAME_COL = "\uc81c\ud488 \uc774\ub984"
POWER_COL = "POWER"
DEMAND_COL = "\uc218\uc694\ub7c9(PCS)"
NEED_COL = "\uc8fc\ubb38 \ub300\uc751 \uc218\ub7c9"
MIN_DUE_COL = "\ucd5c\uc18c \ub0a9\uae30\uc77c"

SITE_FILTER = "C\uad00"
TARGET_PROCESSES = (
    "[10]\uc0ac\ucd9c\uc870\ub9bd",
    "[80]\ub204\uc218/\uaddc\uaca9\uac80\uc0ac",
)

DATE_RE = re.compile(r"(\d{8})")


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

    values: list[str] = []
    for item in root.findall("a:si", NS):
        values.append("".join(text.text or "" for text in item.findall(".//a:t", NS)))
    return values


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
    """Read the first worksheet without applying Excel date formatting.

    The source files format the quantity column as dates in some rows. Reading
    raw XML values preserves the actual numeric quantities.
    """
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


def parse_date(path: Path) -> str | None:
    match = DATE_RE.search(path.stem)
    return match.group(1) if match else None


def find_demand_files(root: Path) -> list[tuple[str, Path]]:
    files: list[tuple[str, Path]] = []
    for path in root.glob("*.xlsx"):
        date = parse_date(path)
        if date:
            files.append((date, path))
    return sorted(files)


def to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def prepare_process_data(path: Path) -> pd.DataFrame:
    df = read_xlsx_raw(path)
    required_cols = {
        SITE_COL,
        PROCESS_COL,
        SHORT_CODE_COL,
        SKU_COL,
        PRODUCT_NAME_COL,
        POWER_COL,
        DEMAND_COL,
        NEED_COL,
        MIN_DUE_COL,
    }
    missing = required_cols - set(df.columns)
    if missing:
        missing_text = ", ".join(sorted(missing))
        raise ValueError(f"{path.name}: missing columns: {missing_text}")

    df = df[
        df[SITE_COL].astype(str).str.contains(SITE_FILTER, na=False)
        & df[PROCESS_COL].isin(TARGET_PROCESSES)
    ].copy()
    df[DEMAND_COL] = to_number(df[DEMAND_COL])
    df[NEED_COL] = to_number(df[NEED_COL])
    df[SKU_COL] = df[SKU_COL].astype(str).str.strip()

    df = df[df[SKU_COL].ne("") & df[SKU_COL].ne("\ucd1d\ud569\uacc4")]

    return (
        df.groupby([SKU_COL, PROCESS_COL], as_index=False)
        .agg(
            **{
                SITE_COL: (SITE_COL, "first"),
                SHORT_CODE_COL: (SHORT_CODE_COL, "first"),
                PRODUCT_NAME_COL: (PRODUCT_NAME_COL, "first"),
                POWER_COL: (POWER_COL, "first"),
                DEMAND_COL: (DEMAND_COL, "sum"),
                NEED_COL: (NEED_COL, "sum"),
                MIN_DUE_COL: (MIN_DUE_COL, "min"),
            }
        )
    )


def add_prefixed_columns(df: pd.DataFrame, prefix: str) -> pd.DataFrame:
    renamed = {
        SITE_COL: f"{prefix}_{SITE_COL}",
        SHORT_CODE_COL: f"{prefix}_{SHORT_CODE_COL}",
        PRODUCT_NAME_COL: f"{prefix}_{PRODUCT_NAME_COL}",
        POWER_COL: f"{prefix}_{POWER_COL}",
        DEMAND_COL: f"{prefix}_\uc218\uc694\uc218\ub7c9",
        NEED_COL: f"{prefix}_\uacf5\uc815\ud544\uc694\uc218\ub7c9",
        MIN_DUE_COL: f"{prefix}_{MIN_DUE_COL}",
    }
    return df.rename(columns=renamed)


def compare_dates(
    base_date: str,
    base_df: pd.DataFrame,
    compare_date: str,
    compare_df: pd.DataFrame,
) -> pd.DataFrame:
    base = add_prefixed_columns(base_df, "\uae30\uc900\uc77c")
    compare = add_prefixed_columns(compare_df, "\ube44\uad50\uc77c")
    merged = compare.merge(base, on=[SKU_COL, PROCESS_COL], how="outer")

    for column in (
        "\uae30\uc900\uc77c_\uc218\uc694\uc218\ub7c9",
        "\ube44\uad50\uc77c_\uc218\uc694\uc218\ub7c9",
        "\uae30\uc900\uc77c_\uacf5\uc815\ud544\uc694\uc218\ub7c9",
        "\ube44\uad50\uc77c_\uacf5\uc815\ud544\uc694\uc218\ub7c9",
    ):
        merged[column] = to_number(merged[column])

    merged.insert(0, "\uae30\uc900\uc77c", base_date)
    merged.insert(0, "\ube44\uad50\uc77c", compare_date)

    merged["\uc218\uc694\uc99d\uac00\ubd84"] = (
        merged["\uae30\uc900\uc77c_\uc218\uc694\uc218\ub7c9"]
        - merged["\ube44\uad50\uc77c_\uc218\uc694\uc218\ub7c9"]
    ).clip(lower=0)
    merged["\uc218\uc694\uac10\uc18c\ubd84"] = (
        merged["\ube44\uad50\uc77c_\uc218\uc694\uc218\ub7c9"]
        - merged["\uae30\uc900\uc77c_\uc218\uc694\uc218\ub7c9"]
    ).clip(lower=0)

    merged["\uc0b0\uc2dd_\uc720\ud6a8\uc0dd\uc0b0\ub7c9"] = (
        merged["\ube44\uad50\uc77c_\uacf5\uc815\ud544\uc694\uc218\ub7c9"]
        + merged["\uc218\uc694\uc99d\uac00\ubd84"]
        - merged["\uae30\uc900\uc77c_\uacf5\uc815\ud544\uc694\uc218\ub7c9"]
    ).clip(lower=0)
    merged["\uc720\ud6a8\uc0dd\uc0b0\ub7c9"] = (
        merged["\uc0b0\uc2dd_\uc720\ud6a8\uc0dd\uc0b0\ub7c9"] - merged["\uc218\uc694\uac10\uc18c\ubd84"]
    ).clip(lower=0)

    merged[SITE_COL] = merged["\uae30\uc900\uc77c_" + SITE_COL].combine_first(
        merged["\ube44\uad50\uc77c_" + SITE_COL]
    )
    merged[SHORT_CODE_COL] = merged["\uae30\uc900\uc77c_" + SHORT_CODE_COL].combine_first(
        merged["\ube44\uad50\uc77c_" + SHORT_CODE_COL]
    )
    merged[PRODUCT_NAME_COL] = merged[
        "\uae30\uc900\uc77c_" + PRODUCT_NAME_COL
    ].combine_first(merged["\ube44\uad50\uc77c_" + PRODUCT_NAME_COL])
    merged[POWER_COL] = merged["\uae30\uc900\uc77c_" + POWER_COL].combine_first(
        merged["\ube44\uad50\uc77c_" + POWER_COL]
    )

    output_columns = [
        "\ube44\uad50\uc77c",
        "\uae30\uc900\uc77c",
        SITE_COL,
        PROCESS_COL,
        SHORT_CODE_COL,
        SKU_COL,
        PRODUCT_NAME_COL,
        POWER_COL,
        "\ube44\uad50\uc77c_\uc218\uc694\uc218\ub7c9",
        "\uae30\uc900\uc77c_\uc218\uc694\uc218\ub7c9",
        "\uc218\uc694\uc99d\uac00\ubd84",
        "\uc218\uc694\uac10\uc18c\ubd84",
        "\ube44\uad50\uc77c_\uacf5\uc815\ud544\uc694\uc218\ub7c9",
        "\uae30\uc900\uc77c_\uacf5\uc815\ud544\uc694\uc218\ub7c9",
        "\uc0b0\uc2dd_\uc720\ud6a8\uc0dd\uc0b0\ub7c9",
        "\uc720\ud6a8\uc0dd\uc0b0\ub7c9",
    ]
    return merged[output_columns].sort_values(["\ube44\uad50\uc77c", PROCESS_COL, SKU_COL])


def summarize(details: pd.DataFrame) -> pd.DataFrame:
    if details.empty:
        return pd.DataFrame()
    return (
        details.groupby(["\ube44\uad50\uc77c", "\uae30\uc900\uc77c", PROCESS_COL], as_index=False)
        .agg(
            **{
                "SKU\uc218": (SKU_COL, "count"),
                "\ube44\uad50\uc77c_\uc218\uc694\uc218\ub7c9": (
                    "\ube44\uad50\uc77c_\uc218\uc694\uc218\ub7c9",
                    "sum",
                ),
                "\uae30\uc900\uc77c_\uc218\uc694\uc218\ub7c9": (
                    "\uae30\uc900\uc77c_\uc218\uc694\uc218\ub7c9",
                    "sum",
                ),
                "\uc218\uc694\uc99d\uac00\ubd84": ("\uc218\uc694\uc99d\uac00\ubd84", "sum"),
                "\uc218\uc694\uac10\uc18c\ubd84": ("\uc218\uc694\uac10\uc18c\ubd84", "sum"),
                "\ube44\uad50\uc77c_\uacf5\uc815\ud544\uc694\uc218\ub7c9": (
                    "\ube44\uad50\uc77c_\uacf5\uc815\ud544\uc694\uc218\ub7c9",
                    "sum",
                ),
                "\uae30\uc900\uc77c_\uacf5\uc815\ud544\uc694\uc218\ub7c9": (
                    "\uae30\uc900\uc77c_\uacf5\uc815\ud544\uc694\uc218\ub7c9",
                    "sum",
                ),
                "\uc0b0\uc2dd_\uc720\ud6a8\uc0dd\uc0b0\ub7c9": (
                    "\uc0b0\uc2dd_\uc720\ud6a8\uc0dd\uc0b0\ub7c9",
                    "sum",
                ),
                "\uc720\ud6a8\uc0dd\uc0b0\ub7c9": ("\uc720\ud6a8\uc0dd\uc0b0\ub7c9", "sum"),
            }
        )
        .sort_values(["\ube44\uad50\uc77c", PROCESS_COL])
    )


def build_file_summary(rows: Iterable[dict[str, object]]) -> pd.DataFrame:
    return pd.DataFrame(rows).sort_values("\ud30c\uc77c\uc77c\uc790")


def main() -> None:
    root = Path(".")
    demand_files = find_demand_files(root)
    if not demand_files:
        raise SystemExit("No dated xlsx files found.")

    base_date, base_path = demand_files[-1]
    base_df = prepare_process_data(base_path)
    if base_df.empty:
        raise SystemExit(f"Base file has no target rows: {base_path.name}")

    file_summary_rows: list[dict[str, object]] = []
    detail_frames: list[pd.DataFrame] = []

    file_summary_rows.append(
        {
            "\ud30c\uc77c\uc77c\uc790": base_date,
            "\ud30c\uc77c\uba85": base_path.name,
            "\uc5ed\ud560": "\uae30\uc900\uc77c",
            "\ub300\uc0c1\ud589\uc218": len(base_df),
            "\uba54\ubaa8": "",
        }
    )

    for compare_date, compare_path in demand_files[:-1]:
        compare_df = prepare_process_data(compare_path)
        if compare_df.empty:
            file_summary_rows.append(
                {
                    "\ud30c\uc77c\uc77c\uc790": compare_date,
                    "\ud30c\uc77c\uba85": compare_path.name,
                    "\uc5ed\ud560": "\ube44\uad50\uc77c",
                    "\ub300\uc0c1\ud589\uc218": 0,
                    "\uba54\ubaa8": "C\uad00 [10]/[80] \ub300\uc0c1\ud589 \uc5c6\uc74c",
                }
            )
            continue

        file_summary_rows.append(
            {
                "\ud30c\uc77c\uc77c\uc790": compare_date,
                "\ud30c\uc77c\uba85": compare_path.name,
                "\uc5ed\ud560": "\ube44\uad50\uc77c",
                "\ub300\uc0c1\ud589\uc218": len(compare_df),
                "\uba54\ubaa8": "\uc0b0\ucd9c \ud3ec\ud568",
            }
        )
        detail_frames.append(compare_dates(base_date, base_df, compare_date, compare_df))

    details = pd.concat(detail_frames, ignore_index=True) if detail_frames else pd.DataFrame()
    summary = summarize(details)
    file_summary = build_file_summary(file_summary_rows)

    output_path = root / f"\uc720\ud6a8\uc0dd\uc0b0\ub7c9_\uae30\uc900\uc77c_{base_date}.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="\uacf5\uc815\ubcc4\uc694\uc57d", index=False)
        details.to_excel(writer, sheet_name="\uc0c1\uc138", index=False)
        file_summary.to_excel(writer, sheet_name="\ud30c\uc77c\uc694\uc57d", index=False)

    print(f"base_date={base_date}")
    print(f"base_file={base_path.name}")
    print(f"output={output_path.name}")
    if not summary.empty:
        print(summary.to_string(index=False))
    else:
        print("No comparison files with target rows.")


if __name__ == "__main__":
    main()
