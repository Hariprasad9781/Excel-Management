import os

import pandas as pd
from openpyxl import load_workbook

from models.excel_file import ExcelFile


def get_excel_file_path(excel_file: ExcelFile) -> str:
    if not os.path.exists(excel_file.file_path):
        raise FileNotFoundError("Stored file not found")

    return excel_file.file_path


def get_sheet_names(excel_file: ExcelFile) -> list[str]:
    file_path = get_excel_file_path(excel_file)

    excel = pd.ExcelFile(file_path)

    return excel.sheet_names


def get_sheet_preview(
    excel_file: ExcelFile,
    sheet_name: str,
    rows: int = 10,
) -> dict:
    file_path = get_excel_file_path(excel_file)

    excel = pd.ExcelFile(file_path)

    if sheet_name not in excel.sheet_names:
        raise ValueError(f"Sheet '{sheet_name}' not found")

    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        nrows=rows,
    )

    df = df.fillna("")

    return {
        "sheet_name": sheet_name,
        "columns": df.columns.tolist(),
        "rows": df.to_dict(orient="records"),
        "row_count": len(df),
    }


def update_cell(
    excel_file: ExcelFile,
    sheet_name: str,
    cell: str,
    value: str | int | float | bool | None,
) -> None:
    file_path = get_excel_file_path(excel_file)

    workbook = load_workbook(file_path)

    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found")

    worksheet = workbook[sheet_name]

    worksheet[cell] = value

    workbook.save(file_path)


# ============================================================
# Row Operations
# ============================================================


def add_row(
    excel_file: ExcelFile,
    sheet_name: str,
    row_data: list[str | int | float | bool | None],
) -> int:
    file_path = get_excel_file_path(excel_file)

    workbook = load_workbook(file_path)

    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found")

    worksheet = workbook[sheet_name]

    # Add the row at the end
    worksheet.append(row_data)

    row_number = worksheet.max_row

    workbook.save(file_path)

    return row_number


def update_row(
    excel_file: ExcelFile,
    sheet_name: str,
    row_number: int,
    row_data: list[str | int | float | bool | None],
) -> None:
    file_path = get_excel_file_path(excel_file)

    workbook = load_workbook(file_path)

    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found")

    worksheet = workbook[sheet_name]

    if row_number < 1 or row_number > worksheet.max_row:
        raise ValueError(f"Row {row_number} not found")

    for column_number, value in enumerate(row_data, start=1):
        worksheet.cell(
            row=row_number,
            column=column_number,
            value=value,
        )

    workbook.save(file_path)


def delete_row(
    excel_file: ExcelFile,
    sheet_name: str,
    row_number: int,
) -> None:
    file_path = get_excel_file_path(excel_file)

    workbook = load_workbook(file_path)

    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found")

    worksheet = workbook[sheet_name]

    if row_number < 1 or row_number > worksheet.max_row:
        raise ValueError(f"Row {row_number} not found")

    worksheet.delete_rows(row_number, 1)

    workbook.save(file_path)