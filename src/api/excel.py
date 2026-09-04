from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from api.schemas import (
    CellUpdate,
    CellUpdateResponse,
    ColumnAddRequest,
    ColumnDeleteRequest,
    ColumnOperationResponse,
    ColumnUpdateRequest,
    ExcelFormatRequest,
    ExcelFormatResponse,
    ExcelSearchResponse,
    RowAddRequest,
    RowDeleteRequest,
    RowOperationResponse,
    RowUpdateRequest,
    SheetCreateRequest,
    SheetDeleteRequest,
    SheetOperationResponse,
    SheetPreviewResponse,
    SheetRenameRequest,
    SheetsResponse,
    WorkbookVersionsResponse,
)
from database import get_db
from dependencies import get_current_user
from models.excel_file import ExcelFile
from models.user import User
from models.workbook import Workbook
from services.excel_service import (
    add_column,
    add_row,
    create_sheet,
    delete_column,
    delete_row,
    delete_sheet,
    format_excel_range,
    get_sheet_names,
    get_sheet_preview,
    rename_sheet,
    search_excel,
    update_cell,
    update_column,
    update_row,
)

from services.file_version_service import create_file_version
from services.file_version_service import restore_file_version as restore_physical_file_version
from services.workbook_version_service import (
    get_workbook_versions,
    restore_workbook_version,
)

# ============================================================
# Router Configuration
# ============================================================

router = APIRouter(
    prefix="/files",
    tags=["Excel Data"],
)


# ============================================================
# Get User File
# ============================================================


def get_user_file(
    file_id: int,
    current_user: User,
    db: Session,
) -> ExcelFile:

    excel_file = (
        db.query(ExcelFile)
        .filter(
            ExcelFile.id == file_id,
            ExcelFile.user_id == current_user.id,
        )
        .first()
    )

    if not excel_file:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="File not found",
        )

    return excel_file


# ============================================================
# Workbook Version History
# ============================================================


@router.get(
    "/{file_id}/versions",
    response_model=WorkbookVersionsResponse,
)
def list_workbook_versions(
    file_id: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    if excel_file.workbook_id is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File is not linked to a workbook",
        )

    versions = get_workbook_versions(
        db=db,
        workbook_id=excel_file.workbook_id,
    )

    return {
        "workbook_id": excel_file.workbook_id,
        "versions": versions,
    }

@router.post(
    "/{file_id}/versions/{version_number}/restore",
)
def restore_file_version(
    file_id: int,
    version_number: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    if excel_file.workbook_id is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File is not linked to a workbook",
        )

    workbook = (
        db.query(Workbook)
        .filter(
            Workbook.id == excel_file.workbook_id,
            Workbook.owner_id == current_user.id,
            Workbook.is_deleted.is_(False),
        )
        .first()
    )

    if workbook is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Workbook not found",
        )

    try:
        new_version = restore_workbook_version(
            db=db,
            workbook=workbook,
            version_number=version_number,
            created_by=current_user.id,
        )

        restore_physical_file_version(
            db=db,
            excel_file=excel_file,
            version_number=version_number,
        )

        db.commit()
        db.refresh(new_version)

        return {
            "message": f"Workbook restored from version {version_number}",
            "workbook_id": workbook.id,
            "restored_from_version": version_number,
            "new_version": new_version.version_number,
        }
    except ValueError as exc:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=str(exc),
        )

    except Exception:
        db.rollback()
        raise


# ============================================================
# Sheet Management
# ============================================================


@router.get(
    "/{file_id}/sheets",
    response_model=SheetsResponse,
)
def list_sheets(
    file_id: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    try:
        sheets = get_sheet_names(excel_file)

    except FileNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to read Excel file: {exc}",
        )

    return {
        "file_id": excel_file.id,
        "filename": excel_file.original_filename,
        "sheets": sheets,
    }


# ============================================================
# Create Sheet
# ============================================================


@router.post(
    "/{file_id}/sheets",
    response_model=SheetOperationResponse,
    status_code=status.HTTP_201_CREATED,
)
def create_excel_sheet(
    file_id: int,
    data: SheetCreateRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    try:
        create_sheet(
            excel_file=excel_file,
            sheet_name=data.sheet_name,
        )

        version = create_file_version(
            db=db,
            excel_file=excel_file,
            created_by=current_user.id,
            change_summary=(
                f"Created sheet '{data.sheet_name}'"
            ),
        )

        db.commit()

    except FileNotFoundError:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        )

    except Exception as exc:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to create sheet: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "message": (
            f"Sheet created successfully. "
            f"Version {version.version_number} created."
        ),
    }

# ============================================================
# Rename Sheet
# ============================================================


@router.put(
    "/{file_id}/sheets",
    response_model=SheetOperationResponse,
)
def rename_excel_sheet(
    file_id: int,
    data: SheetRenameRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    try:
        rename_sheet(
            excel_file=excel_file,
            sheet_name=data.sheet_name,
            new_sheet_name=data.new_sheet_name,
        )

        version = create_file_version(
            db=db,
            excel_file=excel_file,
            created_by=current_user.id,
            change_summary=(
                f"Renamed sheet '{data.sheet_name}' "
                f"to '{data.new_sheet_name}'"
            ),
        )

        db.commit()

    except FileNotFoundError:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        )

    except Exception as exc:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to rename sheet: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "message": (
            f"Sheet renamed successfully. "
            f"Version {version.version_number} created."
        ),
    }

# ============================================================
# Delete Sheet
# ============================================================


@router.delete(
    "/{file_id}/sheets",
    response_model=SheetOperationResponse,
)
def delete_excel_sheet(
    file_id: int,
    data: SheetDeleteRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    try:
        delete_sheet(
            excel_file=excel_file,
            sheet_name=data.sheet_name,
        )

        version = create_file_version(
            db=db,
            excel_file=excel_file,
            created_by=current_user.id,
            change_summary=(
                f"Deleted sheet '{data.sheet_name}'"
            ),
        )

        db.commit()

    except FileNotFoundError:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        )

    except Exception as exc:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to delete sheet: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "message": (
            f"Sheet deleted successfully. "
            f"Version {version.version_number} created."
        ),
    }


# ============================================================
# Sheet Preview
# ============================================================


@router.get(
    "/{file_id}/preview",
    response_model=SheetPreviewResponse,
)
def preview_sheet(
    file_id: int,
    sheet_name: str,
    rows: int = 10,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    if rows < 1 or rows > 100:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Rows must be between 1 and 100",
        )

    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    try:
        preview = get_sheet_preview(
            excel_file=excel_file,
            sheet_name=sheet_name,
            rows=rows,
        )

    except FileNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        )

    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to read Excel sheet: {exc}",
        )

    return {
        "file_id": excel_file.id,
        "filename": excel_file.original_filename,
        **preview,
    }


# ============================================================
# Cell Operations
# ============================================================


@router.put(
    "/{file_id}/cell",
    response_model=CellUpdateResponse,
)
def edit_cell(
    file_id: int,
    data: CellUpdate,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    try:
        # -----------------------------------------------------
        # 1. Update the physical Excel file
        # -----------------------------------------------------
        update_cell(
            excel_file=excel_file,
            sheet_name=data.sheet_name,
            cell=data.cell,
            value=data.value,
        )

        # -----------------------------------------------------
        # 2. Create a new version from the updated Excel file
        # -----------------------------------------------------
        version = create_file_version(
            db=db,
            excel_file=excel_file,
            created_by=current_user.id,
            change_summary=(
                f"Updated cell {data.cell} "
                f"in sheet '{data.sheet_name}'"
            ),
        )

        # -----------------------------------------------------
        # 3. Commit version + workbook version number
        # -----------------------------------------------------
        db.commit()

    except FileNotFoundError:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        )

    except Exception as exc:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to update cell: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "cell": data.cell,
        "value": data.value,
        "message": (
            f"Cell updated successfully. "
            f"Version {version.version_number} created."
        ),
    }

# ============================================================
# Row Operations
# ============================================================


@router.post(
    "/{file_id}/rows",
    response_model=RowOperationResponse,
    status_code=status.HTTP_201_CREATED,
)
def add_excel_row(
    file_id: int,
    data: RowAddRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    try:
        # -----------------------------------------------------
        # 1. Add the row to the physical Excel file
        # -----------------------------------------------------
        row_number = add_row(
            excel_file=excel_file,
            sheet_name=data.sheet_name,
            row_data=data.row_data,
        )

        # -----------------------------------------------------
        # 2. Create a version from the updated Excel file
        # -----------------------------------------------------
        version = create_file_version(
            db=db,
            excel_file=excel_file,
            created_by=current_user.id,
            change_summary=(
                f"Added row {row_number} "
                f"in sheet '{data.sheet_name}'"
            ),
        )

        # -----------------------------------------------------
        # 3. Commit version + workbook version number
        # -----------------------------------------------------
        db.commit()

    except FileNotFoundError:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        )

    except Exception as exc:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to add row: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "row_number": row_number,
        "message": (
            f"Row added successfully. "
            f"Version {version.version_number} created."
        ),
    }

@router.put(
    "/{file_id}/rows",
    response_model=RowOperationResponse,
)
def update_excel_row(
    file_id: int,
    data: RowUpdateRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    try:
        # -----------------------------------------------------
        # 1. Update the row in the physical Excel file
        # -----------------------------------------------------
        update_row(
            excel_file=excel_file,
            sheet_name=data.sheet_name,
            row_number=data.row_number,
            row_data=data.row_data,
        )

        # -----------------------------------------------------
        # 2. Create a version from the updated Excel file
        # -----------------------------------------------------
        version = create_file_version(
            db=db,
            excel_file=excel_file,
            created_by=current_user.id,
            change_summary=(
                f"Updated row {data.row_number} "
                f"in sheet '{data.sheet_name}'"
            ),
        )

        # -----------------------------------------------------
        # 3. Commit version + workbook version number
        # -----------------------------------------------------
        db.commit()

    except FileNotFoundError:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        )

    except Exception as exc:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to update row: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "row_number": data.row_number,
        "message": (
            f"Row updated successfully. "
            f"Version {version.version_number} created."
        ),
    }

@router.delete(
    "/{file_id}/rows",
    response_model=RowOperationResponse,
)
def delete_excel_row(
    file_id: int,
    data: RowDeleteRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    try:
        delete_row(
            excel_file=excel_file,
            sheet_name=data.sheet_name,
            row_number=data.row_number,
        )

        version = create_file_version(
            db=db,
            excel_file=excel_file,
            created_by=current_user.id,
            change_summary=(
                f"Deleted row {data.row_number} "
                f"in sheet '{data.sheet_name}'"
            ),
        )

        db.commit()

    except FileNotFoundError:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        )

    except Exception as exc:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to delete row: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "row_number": data.row_number,
        "message": (
            f"Row deleted successfully. "
            f"Version {version.version_number} created."
        ),
    }

# ============================================================
# Column Operations
# ============================================================


@router.post(
    "/{file_id}/columns",
    response_model=ColumnOperationResponse,
    status_code=status.HTTP_201_CREATED,
)
def add_excel_column(
    file_id: int,
    data: ColumnAddRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    try:
        add_column(
            excel_file=excel_file,
            sheet_name=data.sheet_name,
            column_number=data.column_number,
        )

        version = create_file_version(
            db=db,
            excel_file=excel_file,
            created_by=current_user.id,
            change_summary=(
                f"Added column {data.column_number} "
                f"in sheet '{data.sheet_name}'"
            ),
        )

        db.commit()

    except FileNotFoundError:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        )

    except Exception as exc:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to add column: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "column_number": data.column_number,
        "message": (
            f"Column added successfully. "
            f"Version {version.version_number} created."
        ),
    }


@router.put(
    "/{file_id}/columns",
    response_model=ColumnOperationResponse,
)
def update_excel_column(
    file_id: int,
    data: ColumnUpdateRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    try:
        update_column(
            excel_file=excel_file,
            sheet_name=data.sheet_name,
            column_number=data.column_number,
            column_name=data.column_name,
        )

        version = create_file_version(
            db=db,
            excel_file=excel_file,
            created_by=current_user.id,
            change_summary=(
                f"Updated column {data.column_number} "
                f"in sheet '{data.sheet_name}'"
            ),
        )

        db.commit()

    except FileNotFoundError:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        )

    except Exception as exc:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to update column: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "column_number": data.column_number,
        "message": (
            f"Column updated successfully. "
            f"Version {version.version_number} created."
        ),
    }


@router.delete(
    "/{file_id}/columns",
    response_model=ColumnOperationResponse,
)
def delete_excel_column(
    file_id: int,
    data: ColumnDeleteRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    try:
        delete_column(
            excel_file=excel_file,
            sheet_name=data.sheet_name,
            column_number=data.column_number,
        )

        version = create_file_version(
            db=db,
            excel_file=excel_file,
            created_by=current_user.id,
            change_summary=(
                f"Deleted column {data.column_number} "
                f"in sheet '{data.sheet_name}'"
            ),
        )

        db.commit()

    except FileNotFoundError:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        )

    except Exception as exc:
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to delete column: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "column_number": data.column_number,
        "message": (
            f"Column deleted successfully. "
            f"Version {version.version_number} created."
        ),
    }

# ============================================================
# Excel Search
# ============================================================


@router.get(
    "/{file_id}/search",
    response_model=ExcelSearchResponse,
)
def search_file(
    file_id: int,
    sheet_name: str,
    search_term: str,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    try:
        results = search_excel(
            excel_file=excel_file,
            sheet_name=sheet_name,
            search_term=search_term,
        )

    except FileNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        )

    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to search Excel file: {exc}",
        )

    return {
        "file_id": excel_file.id,
        "filename": excel_file.original_filename,
        "sheet_name": sheet_name,
        "search_term": search_term,
        "results": results,
        "result_count": len(results),
    }


# ============================================================
# Excel Formatting
# ============================================================


@router.put(
    "/{file_id}/format",
    response_model=ExcelFormatResponse,
)
def format_excel_cells(
    file_id: int,
    data: ExcelFormatRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = get_user_file(
        file_id=file_id,
        current_user=current_user,
        db=db,
    )

    try:
        format_excel_range(
            excel_file=excel_file,
            sheet_name=data.sheet_name,
            cell_range=data.cell_range,
            bold=data.bold,
            italic=data.italic,
            underline=data.underline,
            font_size=data.font_size,
            font_color=data.font_color,
            fill_color=data.fill_color,
            horizontal_alignment=data.horizontal_alignment,
            vertical_alignment=data.vertical_alignment,
            number_format=data.number_format,
        )

    except FileNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        )

    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to format Excel cells: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "cell_range": data.cell_range,
        "message": "Excel formatting applied successfully",
    }