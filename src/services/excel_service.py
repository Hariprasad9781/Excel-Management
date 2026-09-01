import os

import pandas as pd
from openpyxl import load_workbook

from models.excel_file import ExcelFile


# ============================================================
# File Helpers
# ============================================================


def get_excel_file_path(
    excel_file: ExcelFile,
) -> str:
    """
    Return the stored Excel file path.

    Raises:
        FileNotFoundError: If the stored Excel file does not exist.
    """

    if not os.path.exists(excel_file.file_path):
        raise FileNotFoundError(
            "Stored file not found"
        )

    return excel_file.file_path


def _load_workbook(
    excel_file: ExcelFile,
):
    """
    Load the Excel workbook from storage.
    """

    file_path = get_excel_file_path(
        excel_file
    )

    return load_workbook(file_path)


def _validate_sheet(
    workbook,
    sheet_name: str,
):
    """
    Validate that the requested sheet exists.
    """

    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found"
        )

    return workbook[sheet_name]


# ============================================================
# Sheet Operations
# ============================================================


def get_sheet_names(
    excel_file: ExcelFile,
) -> list[str]:
    file_path = get_excel_file_path(
        excel_file
    )

    excel = pd.ExcelFile(file_path)

    return excel.sheet_names


def create_sheet(
    excel_file: ExcelFile,
    sheet_name: str,
) -> None:
    file_path = get_excel_file_path(
        excel_file
    )

    if not sheet_name.strip():
        raise ValueError(
            "Sheet name cannot be empty"
        )

    workbook = load_workbook(file_path)

    if sheet_name in workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' already exists"
        )

    workbook.create_sheet(
        sheet_name
    )

    workbook.save(file_path)


def rename_sheet(
    excel_file: ExcelFile,
    sheet_name: str,
    new_sheet_name: str,
) -> None:
    file_path = get_excel_file_path(
        excel_file
    )

    workbook = load_workbook(file_path)

    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found"
        )

    if not new_sheet_name.strip():
        raise ValueError(
            "New sheet name cannot be empty"
        )

    if new_sheet_name in workbook.sheetnames:
        raise ValueError(
            f"Sheet '{new_sheet_name}' already exists"
        )

    worksheet = workbook[sheet_name]

    worksheet.title = new_sheet_name

    workbook.save(file_path)


def delete_sheet(
    excel_file: ExcelFile,
    sheet_name: str,
) -> None:
    file_path = get_excel_file_path(
        excel_file
    )

    workbook = load_workbook(file_path)

    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found"
        )

    if len(workbook.sheetnames) == 1:
        raise ValueError(
            "Cannot delete the only sheet in the workbook"
        )

    worksheet = workbook[sheet_name]

    workbook.remove(worksheet)

    workbook.save(file_path)


def get_sheet_preview(
    excel_file: ExcelFile,
    sheet_name: str,
    rows: int = 10,
) -> dict:
    file_path = get_excel_file_path(
        excel_file
    )

    excel = pd.ExcelFile(file_path)

    if sheet_name not in excel.sheet_names:
        raise ValueError(
            f"Sheet '{sheet_name}' not found"
        )

    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        nrows=rows,
    )

    df = df.fillna("")

    return {
        "sheet_name": sheet_name,
        "columns": df.columns.tolist(),
        "rows": df.to_dict(
            orient="records"
        ),
        "row_count": len(df),
    }


# ============================================================
# Cell Operations
# ============================================================


def update_cell(
    excel_file: ExcelFile,
    sheet_name: str,
    cell: str,
    value: str | int | float | bool | None,
) -> None:
    file_path = get_excel_file_path(
        excel_file
    )

    workbook = load_workbook(file_path)

    worksheet = _validate_sheet(
        workbook,
        sheet_name,
    )

    worksheet[cell] = value

    workbook.save(file_path)


# ============================================================
# Row Operations
# ============================================================


def add_row(
    excel_file: ExcelFile,
    sheet_name: str,
    row_data: list[
        str | int | float | bool | None
    ],
) -> int:
    file_path = get_excel_file_path(
        excel_file
    )

    workbook = load_workbook(file_path)

    worksheet = _validate_sheet(
        workbook,
        sheet_name,
    )

    # Add row at the end of the worksheet.
    worksheet.append(row_data)

    row_number = worksheet.max_row

    workbook.save(file_path)

    return row_number


def update_row(
    excel_file: ExcelFile,
    sheet_name: str,
    row_number: int,
    row_data: list[
        str | int | float | bool | None
    ],
) -> None:
    file_path = get_excel_file_path(
        excel_file
    )

    workbook = load_workbook(file_path)

    worksheet = _validate_sheet(
        workbook,
        sheet_name,
    )

    if (
        row_number < 1
        or row_number > worksheet.max_row
    ):
        raise ValueError(
            f"Row {row_number} not found"
        )

    for column_number, value in enumerate(
        row_data,
        start=1,
    ):
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
    file_path = get_excel_file_path(
        excel_file
    )

    workbook = load_workbook(file_path)

    worksheet = _validate_sheet(
        workbook,
        sheet_name,
    )

    if (
        row_number < 1
        or row_number > worksheet.max_row
    ):
        raise ValueError(
            f"Row {row_number} not found"
        )

    worksheet.delete_rows(
        row_number,
        1,
    )

    workbook.save(file_path)


# ============================================================
# Column Operations
# ============================================================


def add_column(
    excel_file: ExcelFile,
    sheet_name: str,
    column_number: int,
) -> None:
    """
    Insert a new blank column.

    Important:
    OpenPyXL may not preserve a completely empty inserted
    column as part of worksheet.max_column after saving.

    Therefore, after inserting the column, we explicitly
    create blank cells in that column. This makes the column
    persistent and allows subsequent update/delete operations
    to find it.
    """

    file_path = get_excel_file_path(
        excel_file
    )

    workbook = load_workbook(file_path)

    worksheet = _validate_sheet(
        workbook,
        sheet_name,
    )

    if column_number < 1:
        raise ValueError(
            "Column number must be greater than 0"
        )

    # A new column can be inserted anywhere from column 1
    # through one position after the current last column.
    max_column = worksheet.max_column

    if column_number > max_column + 1:
        raise ValueError(
            f"Column {column_number} is out of range. "
            f"Next available column is {max_column + 1}."
        )

    # Insert the new column.
    worksheet.insert_cols(
        column_number,
        1,
    )

    # --------------------------------------------------------
    # IMPORTANT FIX
    # --------------------------------------------------------
    #
    # Create actual blank cells in the new column.
    #
    # Without this, a completely empty inserted column may
    # disappear from worksheet dimensions after save/reload.
    #
    # Using an empty string rather than None ensures that
    # OpenPyXL creates a real cell.
    # --------------------------------------------------------

    row_count = max(
        worksheet.max_row,
        1,
    )

    for row_number in range(
        1,
        row_count + 1,
    ):
        worksheet.cell(
            row=row_number,
            column=column_number,
            value="",
        )

    workbook.save(file_path)


def update_column(
    excel_file: ExcelFile,
    sheet_name: str,
    column_number: int,
    column_name: str,
) -> None:
    """
    Rename/update the header of a column.
    """

    file_path = get_excel_file_path(
        excel_file
    )

    workbook = load_workbook(file_path)

    worksheet = _validate_sheet(
        workbook,
        sheet_name,
    )

    if column_number < 1:
        raise ValueError(
            "Column number must be greater than 0"
        )

    if column_number > worksheet.max_column:
        raise ValueError(
            f"Column {column_number} not found"
        )

    if not column_name.strip():
        raise ValueError(
            "Column name cannot be empty"
        )

    # Update the first-row header.
    worksheet.cell(
        row=1,
        column=column_number,
        value=column_name,
    )

    workbook.save(file_path)


def delete_column(
    excel_file: ExcelFile,
    sheet_name: str,
    column_number: int,
) -> None:
    """
    Delete a column from the worksheet.
    """

    file_path = get_excel_file_path(
        excel_file
    )

    workbook = load_workbook(file_path)

    worksheet = _validate_sheet(
        workbook,
        sheet_name,
    )

    if column_number < 1:
        raise ValueError(
            "Column number must be greater than 0"
        )

    if column_number > worksheet.max_column:
        raise ValueError(
            f"Column {column_number} not found"
        )

    # Delete the requested column.
    worksheet.delete_cols(
        column_number,
        1,
    )

    workbook.save(file_path)


# ============================================================
# Excel Search
# ============================================================


def search_excel(
    excel_file: ExcelFile,
    sheet_name: str,
    search_term: str,
) -> list[dict]:
    """
    Search the complete worksheet for a term.

    Returns:
        A list containing matching cells and their
        row/column information.
    """

    file_path = get_excel_file_path(
        excel_file
    )

    workbook = load_workbook(
        file_path,
        data_only=False,
    )

    worksheet = _validate_sheet(
        workbook,
        sheet_name,
    )

    if not search_term.strip():
        raise ValueError(
            "Search term cannot be empty"
        )

    results = []

    search_value = search_term.lower()

    for row in worksheet.iter_rows():

        for cell in row:

            if cell.value is None:
                continue

            cell_value = str(
                cell.value
            )

            if search_value in cell_value.lower():

                results.append(
                    {
                        "row_number": cell.row,
                        "column_number": cell.column,
                        "cell": cell.coordinate,
                        "value": cell.value,
                    }
                )

    return results