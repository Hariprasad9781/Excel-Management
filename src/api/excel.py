from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from api.schemas import (
    CellUpdate,
    CellUpdateResponse,
    ColumnAddRequest,
    ColumnDeleteRequest,
    ColumnOperationResponse,
    ColumnUpdateRequest,
    ExcelSearchResponse,
    RowAddRequest,
    RowDeleteRequest,
    RowOperationResponse,
    RowUpdateRequest,
    SheetPreviewResponse,
    SheetsResponse,
)
from database import get_db
from dependencies import get_current_user
from models.excel_file import ExcelFile
from models.user import User
from services.excel_service import (
    add_column,
    add_row,
    delete_column,
    delete_row,
    get_sheet_names,
    get_sheet_preview,
    search_excel,
    update_cell,
    update_column,
    update_row,
)

router = APIRouter(
    prefix="/files",
    tags=["Excel Data"],
)


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
            status_code=status.HTTP_404_NOT_FOUND,
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
        update_cell(
            excel_file=excel_file,
            sheet_name=data.sheet_name,
            cell=data.cell,
            value=data.value,
        )

    except FileNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=str(exc),
        )

    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to update cell: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "cell": data.cell,
        "value": data.value,
        "message": "Cell updated successfully",
    }


# -------------------------------------------------------------------
# Row Operations
# -------------------------------------------------------------------


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
        row_number = add_row(
            excel_file=excel_file,
            sheet_name=data.sheet_name,
            row_data=data.row_data,
        )

    except FileNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=str(exc),
        )

    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to add row: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "row_number": row_number,
        "message": "Row added successfully",
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
        update_row(
            excel_file=excel_file,
            sheet_name=data.sheet_name,
            row_number=data.row_number,
            row_data=data.row_data,
        )

    except FileNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=str(exc),
        )

    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to update row: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "row_number": data.row_number,
        "message": "Row updated successfully",
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

    except FileNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=str(exc),
        )

    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to delete row: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "row_number": data.row_number,
        "message": "Row deleted successfully",
    }

# -------------------------------------------------------------------
# Column Operations
# -------------------------------------------------------------------


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

    except FileNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=str(exc),
        )

    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to add column: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "column_number": data.column_number,
        "message": "Column added successfully",
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

    except FileNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=str(exc),
        )

    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to update column: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "column_number": data.column_number,
        "message": "Column updated successfully",
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

    except FileNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=str(exc),
        )

    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unable to delete column: {exc}",
        )

    return {
        "file_id": file_id,
        "sheet_name": data.sheet_name,
        "column_number": data.column_number,
        "message": "Column deleted successfully",
    }

# -------------------------------------------------------------------
# Excel Search
# -------------------------------------------------------------------


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
            status_code=status.HTTP_404_NOT_FOUND,
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