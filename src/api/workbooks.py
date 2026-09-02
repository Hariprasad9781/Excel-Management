from fastapi import APIRouter, Depends, HTTPException, status
from pydantic import BaseModel
from sqlalchemy import delete, update
from sqlalchemy.orm import Session

from database import get_db
from dependencies import get_current_user
from models.cell import Cell
from models.user import User
from models.workbook import Workbook
from models.worksheet import Worksheet
from services.workbook_version_service import create_workbook_version


router = APIRouter(
    prefix="/workbooks",
    tags=["Workbooks"],
)


# ===========================================================================
# Request Schemas
# ===========================================================================

class CellUpdateRequest(BaseModel):
    row_index: int
    column_index: int
    value: object | None = None
    data_type: str | None = None
    formula: str | None = None
    style: dict | None = None


class WorksheetCreateRequest(BaseModel):
    name: str


class WorksheetUpdateRequest(BaseModel):
    name: str


class RowInsertRequest(BaseModel):
    row_index: int


class ColumnInsertRequest(BaseModel):
    column_index: int


# ===========================================================================
# Helper Functions
# ===========================================================================

def get_user_workbook(
    workbook_id: int,
    current_user: User,
    db: Session,
):
    """
    Return a workbook belonging to the current user.
    """

    workbook = (
        db.query(Workbook)
        .filter(
            Workbook.id == workbook_id,
            Workbook.owner_id == current_user.id,
            Workbook.is_deleted.is_(False),
        )
        .first()
    )

    if not workbook:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Workbook not found",
        )

    return workbook


def get_user_worksheet(
    workbook_id: int,
    worksheet_id: int,
    current_user: User,
    db: Session,
):
    """
    Return a worksheet belonging to the current user's workbook.
    """

    workbook = get_user_workbook(
        workbook_id=workbook_id,
        current_user=current_user,
        db=db,
    )

    worksheet = (
        db.query(Worksheet)
        .filter(
            Worksheet.id == worksheet_id,
            Worksheet.workbook_id == workbook.id,
            Worksheet.is_deleted.is_(False),
        )
        .first()
    )

    if not worksheet:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Worksheet not found",
        )

    return workbook, worksheet


# ===========================================================================
# Get Workbook
# ===========================================================================

@router.get("/{workbook_id}")
def get_workbook(
    workbook_id: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """
    Get a workbook with all worksheets and cells.
    """

    workbook = get_user_workbook(
        workbook_id=workbook_id,
        current_user=current_user,
        db=db,
    )

    return {
        "id": workbook.id,
        "name": workbook.name,
        "original_filename": workbook.original_filename,
        "version": workbook.version,
        "created_at": workbook.created_at,
        "updated_at": workbook.updated_at,
        "worksheets": [
            {
                "id": worksheet.id,
                "name": worksheet.name,
                "position": worksheet.position,
                "max_row": worksheet.max_row,
                "max_column": worksheet.max_column,
                "cells": [
                    {
                        "id": cell.id,
                        "row_index": cell.row_index,
                        "column_index": cell.column_index,
                        "value": cell.value,
                        "data_type": cell.data_type,
                        "formula": cell.formula,
                        "style": cell.style,
                        "created_at": cell.created_at,
                        "updated_at": cell.updated_at,
                    }
                    for cell in worksheet.cells
                    if not worksheet.is_deleted
                ],
            }
            for worksheet in workbook.worksheets
            if not worksheet.is_deleted
        ],
    }


# ===========================================================================
# List Worksheets
# ===========================================================================

@router.get("/{workbook_id}/worksheets")
def list_worksheets(
    workbook_id: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """
    Get all worksheets belonging to a workbook.
    """

    workbook = get_user_workbook(
        workbook_id=workbook_id,
        current_user=current_user,
        db=db,
    )

    worksheets = (
        db.query(Worksheet)
        .filter(
            Worksheet.workbook_id == workbook.id,
            Worksheet.is_deleted.is_(False),
        )
        .order_by(
            Worksheet.position.asc(),
        )
        .all()
    )

    return {
        "workbook_id": workbook.id,
        "worksheets": [
            {
                "id": worksheet.id,
                "name": worksheet.name,
                "position": worksheet.position,
                "max_row": worksheet.max_row,
                "max_column": worksheet.max_column,
            }
            for worksheet in worksheets
        ],
    }


# ===========================================================================
# Create Worksheet
# ===========================================================================

@router.post(
    "/{workbook_id}/worksheets",
    status_code=status.HTTP_201_CREATED,
)
def create_worksheet(
    workbook_id: int,
    payload: WorksheetCreateRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """
    Create a new worksheet inside a workbook.
    """

    workbook = get_user_workbook(
        workbook_id=workbook_id,
        current_user=current_user,
        db=db,
    )

    # -----------------------------------------------------------------------
    # Validate name
    # -----------------------------------------------------------------------

    worksheet_name = payload.name.strip()

    if not worksheet_name:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Worksheet name is required",
        )

    if len(worksheet_name) > 255:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Worksheet name cannot exceed 255 characters",
        )

    # -----------------------------------------------------------------------
    # Check duplicate name
    # -----------------------------------------------------------------------

    existing_worksheet = (
        db.query(Worksheet)
        .filter(
            Worksheet.workbook_id == workbook.id,
            Worksheet.name == worksheet_name,
            Worksheet.is_deleted.is_(False),
        )
        .first()
    )

    if existing_worksheet:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail="A worksheet with this name already exists",
        )

    # -----------------------------------------------------------------------
    # Determine position
    # -----------------------------------------------------------------------

    last_worksheet = (
        db.query(Worksheet)
        .filter(
            Worksheet.workbook_id == workbook.id,
            Worksheet.is_deleted.is_(False),
        )
        .order_by(
            Worksheet.position.desc(),
        )
        .first()
    )

    next_position = (
        last_worksheet.position + 1
        if last_worksheet
        else 0
    )

    # -----------------------------------------------------------------------
    # Create worksheet
    # -----------------------------------------------------------------------

    worksheet = Worksheet(
        workbook_id=workbook.id,
        name=worksheet_name,
        position=next_position,
        max_row=0,
        max_column=0,
        is_deleted=False,
    )

    db.add(worksheet)

    workbook.version += 1

    try:
        db.commit()
        db.refresh(worksheet)
        db.refresh(workbook)

    except Exception:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to create worksheet",
        )

    return {
        "message": "Worksheet created successfully",
        "workbook_id": workbook.id,
        "workbook_version": workbook.version,
        "worksheet": {
            "id": worksheet.id,
            "name": worksheet.name,
            "position": worksheet.position,
            "max_row": worksheet.max_row,
            "max_column": worksheet.max_column,
        },
    }


# ===========================================================================
# Update Worksheet
# ===========================================================================

@router.put(
    "/{workbook_id}/worksheets/{worksheet_id}",
)
def update_worksheet(
    workbook_id: int,
    worksheet_id: int,
    payload: WorksheetUpdateRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """
    Rename an existing worksheet.
    """

    workbook, worksheet = get_user_worksheet(
        workbook_id=workbook_id,
        worksheet_id=worksheet_id,
        current_user=current_user,
        db=db,
    )

    # -----------------------------------------------------------------------
    # Validate name
    # -----------------------------------------------------------------------

    worksheet_name = payload.name.strip()

    if not worksheet_name:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Worksheet name is required",
        )

    if len(worksheet_name) > 255:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Worksheet name cannot exceed 255 characters",
        )

    # -----------------------------------------------------------------------
    # Check duplicate name
    # -----------------------------------------------------------------------

    existing_worksheet = (
        db.query(Worksheet)
        .filter(
            Worksheet.workbook_id == workbook.id,
            Worksheet.name == worksheet_name,
            Worksheet.id != worksheet.id,
            Worksheet.is_deleted.is_(False),
        )
        .first()
    )

    if existing_worksheet:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail="A worksheet with this name already exists",
        )

    # -----------------------------------------------------------------------
    # Update worksheet
    # -----------------------------------------------------------------------

    worksheet.name = worksheet_name

    workbook.version += 1

    try:
        db.commit()
        db.refresh(worksheet)
        db.refresh(workbook)

    except Exception:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to update worksheet",
        )

    return {
        "message": "Worksheet updated successfully",
        "workbook_id": workbook.id,
        "workbook_version": workbook.version,
        "worksheet": {
            "id": worksheet.id,
            "name": worksheet.name,
            "position": worksheet.position,
            "max_row": worksheet.max_row,
            "max_column": worksheet.max_column,
        },
    }


# ===========================================================================
# Delete Worksheet
# ===========================================================================

@router.delete(
    "/{workbook_id}/worksheets/{worksheet_id}",
)
def delete_worksheet(
    workbook_id: int,
    worksheet_id: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """
    Soft-delete a worksheet.

    The worksheet and its cells remain in PostgreSQL,
    but the worksheet is hidden from normal queries.
    """

    workbook, worksheet = get_user_worksheet(
        workbook_id=workbook_id,
        worksheet_id=worksheet_id,
        current_user=current_user,
        db=db,
    )

    # -----------------------------------------------------------------------
    # Prevent deleting the last worksheet
    # -----------------------------------------------------------------------

    active_worksheet_count = (
        db.query(Worksheet)
        .filter(
            Worksheet.workbook_id == workbook.id,
            Worksheet.is_deleted.is_(False),
        )
        .count()
    )

    if active_worksheet_count <= 1:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="A workbook must contain at least one worksheet",
        )

    # -----------------------------------------------------------------------
    # Soft delete
    # -----------------------------------------------------------------------

    worksheet.is_deleted = True

    workbook.version += 1

    try:
        db.commit()
        db.refresh(workbook)

    except Exception:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to delete worksheet",
        )

    return {
        "message": "Worksheet deleted successfully",
        "workbook_id": workbook.id,
        "workbook_version": workbook.version,
        "worksheet_id": worksheet.id,
    }


# ===========================================================================
# Get Worksheet Preview
# ===========================================================================

@router.get(
    "/{workbook_id}/worksheets/{worksheet_id}/preview",
)
def get_worksheet_preview(
    workbook_id: int,
    worksheet_id: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """
    Get worksheet data for spreadsheet preview.

    Data is loaded entirely from PostgreSQL.
    """

    workbook, worksheet = get_user_worksheet(
        workbook_id=workbook_id,
        worksheet_id=worksheet_id,
        current_user=current_user,
        db=db,
    )

    cells = sorted(
        worksheet.cells,
        key=lambda cell: (
            cell.row_index,
            cell.column_index,
        ),
    )

    return {
        "workbook": {
            "id": workbook.id,
            "name": workbook.name,
            "version": workbook.version,
        },
        "worksheet": {
            "id": worksheet.id,
            "name": worksheet.name,
            "position": worksheet.position,
            "max_row": worksheet.max_row,
            "max_column": worksheet.max_column,
        },
        "cells": [
            {
                "id": cell.id,
                "row_index": cell.row_index,
                "column_index": cell.column_index,
                "value": cell.value,
                "data_type": cell.data_type,
                "formula": cell.formula,
                "style": cell.style,
            }
            for cell in cells
        ],
    }


# ===========================================================================
# Update Cell
# ===========================================================================

@router.put(
    "/{workbook_id}/worksheets/{worksheet_id}/cells",
)
def update_cell(
    workbook_id: int,
    worksheet_id: int,
    payload: CellUpdateRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """
    Create or update a cell directly in PostgreSQL.

    Existing formatting is preserved when style is omitted.
    """

    # -----------------------------------------------------------------------
    # Validate row / column
    # -----------------------------------------------------------------------

    if payload.row_index < 1:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="row_index must be greater than or equal to 1",
        )

    if payload.column_index < 1:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="column_index must be greater than or equal to 1",
        )

    # -----------------------------------------------------------------------
    # Get workbook + worksheet
    # -----------------------------------------------------------------------

    workbook, worksheet = get_user_worksheet(
        workbook_id=workbook_id,
        worksheet_id=worksheet_id,
        current_user=current_user,
        db=db,
    )

    # -----------------------------------------------------------------------
    # Find existing cell
    # -----------------------------------------------------------------------

    cell = (
        db.query(Cell)
        .filter(
            Cell.worksheet_id == worksheet.id,
            Cell.row_index == payload.row_index,
            Cell.column_index == payload.column_index,
        )
        .first()
    )

    # -----------------------------------------------------------------------
    # Create new cell
    # -----------------------------------------------------------------------

    if cell is None:

        cell = Cell(
            worksheet_id=worksheet.id,
            row_index=payload.row_index,
            column_index=payload.column_index,
            value=payload.value,
            data_type=payload.data_type,
            formula=payload.formula,
            style=payload.style,
        )

        db.add(cell)

        worksheet.max_row = max(
            worksheet.max_row,
            payload.row_index,
        )

        worksheet.max_column = max(
            worksheet.max_column,
            payload.column_index,
        )

    # -----------------------------------------------------------------------
    # Update existing cell
    # -----------------------------------------------------------------------

    else:

        cell.value = payload.value
        cell.data_type = payload.data_type
        cell.formula = payload.formula

        # Preserve existing formatting if no style was supplied.
        if payload.style is not None:
            cell.style = payload.style

    # -----------------------------------------------------------------------
    # Create workbook version snapshot
    # -----------------------------------------------------------------------

    try:
        create_workbook_version(
            db=db,
            workbook=workbook,
            created_by=current_user.id,
            change_summary=(
                f"Updated cell "
                f"{payload.row_index},{payload.column_index}"
            ),
        )

        # -------------------------------------------------------------------
        # Commit cell change + version snapshot together
        # -------------------------------------------------------------------

        db.commit()

        db.refresh(cell)
        db.refresh(workbook)

    except Exception:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to update cell",
        )

    # -----------------------------------------------------------------------
    # Response
    # -----------------------------------------------------------------------

    return {
        "message": "Cell updated successfully",
        "workbook_id": workbook.id,
        "worksheet_id": worksheet.id,
        "workbook_version": workbook.version,
        "cell": {
            "id": cell.id,
            "row_index": cell.row_index,
            "column_index": cell.column_index,
            "value": cell.value,
            "data_type": cell.data_type,
            "formula": cell.formula,
            "style": cell.style,
            "updated_at": cell.updated_at,
        },
    }


# ===========================================================================
# Insert Row
# ===========================================================================

@router.post(
    "/{workbook_id}/worksheets/{worksheet_id}/rows",
)
def insert_row(
    workbook_id: int,
    worksheet_id: int,
    payload: RowInsertRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """
    Insert one row at the specified position.

    Existing cells at or below the inserted row are shifted down by one.
    """

    if payload.row_index < 1:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="row_index must be greater than or equal to 1",
        )

    workbook, worksheet = get_user_worksheet(
        workbook_id=workbook_id,
        worksheet_id=worksheet_id,
        current_user=current_user,
        db=db,
    )

    old_max_row = worksheet.max_row

    try:
        # Temporary offset prevents collisions with the unique
        # worksheet/row/column constraint.
        offset = max(old_max_row, 1) + 1

        shifted_result = db.execute(
            update(Cell)
            .where(
                Cell.worksheet_id == worksheet.id,
                Cell.row_index >= payload.row_index,
            )
            .values(
                row_index=Cell.row_index + offset,
            )
        )

        # Move cells from the temporary range into their final positions.
        if shifted_result.rowcount:
            db.execute(
                update(Cell)
                .where(
                    Cell.worksheet_id == worksheet.id,
                    Cell.row_index >= payload.row_index + offset,
                )
                .values(
                    row_index=Cell.row_index - offset + 1,
                )
            )

        # Update worksheet dimensions.
        if payload.row_index <= old_max_row:
            worksheet.max_row = old_max_row + 1
        else:
            worksheet.max_row = payload.row_index

        workbook.version += 1

        db.commit()
        db.refresh(worksheet)
        db.refresh(workbook)

    except Exception:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to insert row",
        )

    return {
        "message": "Row inserted successfully",
        "workbook_id": workbook.id,
        "worksheet_id": worksheet.id,
        "row_index": payload.row_index,
        "cells_shifted": shifted_result.rowcount,
        "workbook_version": workbook.version,
        "worksheet": {
            "max_row": worksheet.max_row,
            "max_column": worksheet.max_column,
        },
    }


# ===========================================================================
# Delete Row
# ===========================================================================

@router.delete(
    "/{workbook_id}/worksheets/{worksheet_id}/rows/{row_index}",
)
def delete_row(
    workbook_id: int,
    worksheet_id: int,
    row_index: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """
    Delete one row.

    Cells in the deleted row are removed.
    Cells below the deleted row are shifted up by one.
    """

    if row_index < 1:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="row_index must be greater than or equal to 1",
        )

    workbook, worksheet = get_user_worksheet(
        workbook_id=workbook_id,
        worksheet_id=worksheet_id,
        current_user=current_user,
        db=db,
    )

    if row_index > worksheet.max_row:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="row_index is outside the worksheet dimensions",
        )

    old_max_row = worksheet.max_row

    try:
        # Delete cells belonging to the selected row.
        deleted_result = db.execute(
            delete(Cell).where(
                Cell.worksheet_id == worksheet.id,
                Cell.row_index == row_index,
            )
        )

        db.flush()

        # Temporary offset prevents unique-position collisions.
        offset = max(old_max_row, 1) + 1

        shifted_result = db.execute(
            update(Cell)
            .where(
                Cell.worksheet_id == worksheet.id,
                Cell.row_index > row_index,
            )
            .values(
                row_index=Cell.row_index + offset,
            )
        )

        if shifted_result.rowcount:
            db.execute(
                update(Cell)
                .where(
                    Cell.worksheet_id == worksheet.id,
                    Cell.row_index > row_index + offset,
                )
                .values(
                    row_index=Cell.row_index - offset - 1,
                )
            )

        worksheet.max_row = max(
            0,
            old_max_row - 1,
        )

        workbook.version += 1

        db.commit()
        db.refresh(worksheet)
        db.refresh(workbook)

    except Exception:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to delete row",
        )

    return {
        "message": "Row deleted successfully",
        "workbook_id": workbook.id,
        "worksheet_id": worksheet.id,
        "deleted_row_index": row_index,
        "cells_deleted": deleted_result.rowcount,
        "cells_shifted": shifted_result.rowcount,
        "workbook_version": workbook.version,
        "worksheet": {
            "max_row": worksheet.max_row,
            "max_column": worksheet.max_column,
        },
    }


# ===========================================================================
# Insert Column
# ===========================================================================

@router.post(
    "/{workbook_id}/worksheets/{worksheet_id}/columns",
)
def insert_column(
    workbook_id: int,
    worksheet_id: int,
    payload: ColumnInsertRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """
    Insert one column at the specified position.

    Existing cells at or to the right of the inserted column
    are shifted right by one.
    """

    if payload.column_index < 1:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="column_index must be greater than or equal to 1",
        )

    workbook, worksheet = get_user_worksheet(
        workbook_id=workbook_id,
        worksheet_id=worksheet_id,
        current_user=current_user,
        db=db,
    )

    old_max_column = worksheet.max_column

    try:
        # Temporary offset prevents collisions with the unique
        # worksheet/row/column constraint.
        offset = max(old_max_column, 1) + 1

        shifted_result = db.execute(
            update(Cell)
            .where(
                Cell.worksheet_id == worksheet.id,
                Cell.column_index >= payload.column_index,
            )
            .values(
                column_index=Cell.column_index + offset,
            )
        )

        # Move cells from the temporary range into their final positions.
        if shifted_result.rowcount:
            db.execute(
                update(Cell)
                .where(
                    Cell.worksheet_id == worksheet.id,
                    Cell.column_index >= payload.column_index + offset,
                )
                .values(
                    column_index=Cell.column_index - offset + 1,
                )
            )

        # Update worksheet dimensions.
        if payload.column_index <= old_max_column:
            worksheet.max_column = old_max_column + 1
        else:
            worksheet.max_column = payload.column_index

        workbook.version += 1

        db.commit()
        db.refresh(worksheet)
        db.refresh(workbook)

    except Exception:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to insert column",
        )

    return {
        "message": "Column inserted successfully",
        "workbook_id": workbook.id,
        "worksheet_id": worksheet.id,
        "column_index": payload.column_index,
        "cells_shifted": shifted_result.rowcount,
        "workbook_version": workbook.version,
        "worksheet": {
            "max_row": worksheet.max_row,
            "max_column": worksheet.max_column,
        },
    }


# ===========================================================================
# Delete Column
# ===========================================================================

@router.delete(
    "/{workbook_id}/worksheets/{worksheet_id}/columns/{column_index}",
)
def delete_column(
    workbook_id: int,
    worksheet_id: int,
    column_index: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """
    Delete one column.

    Cells in the deleted column are removed.
    Cells to the right of the deleted column are shifted left by one.
    """

    if column_index < 1:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="column_index must be greater than or equal to 1",
        )

    workbook, worksheet = get_user_worksheet(
        workbook_id=workbook_id,
        worksheet_id=worksheet_id,
        current_user=current_user,
        db=db,
    )

    if column_index > worksheet.max_column:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="column_index is outside the worksheet dimensions",
        )

    old_max_column = worksheet.max_column

    try:
        # Delete cells belonging to the selected column.
        deleted_result = db.execute(
            delete(Cell).where(
                Cell.worksheet_id == worksheet.id,
                Cell.column_index == column_index,
            )
        )

        db.flush()

        # Temporary offset prevents unique-position collisions.
        offset = max(old_max_column, 1) + 1

        shifted_result = db.execute(
            update(Cell)
            .where(
                Cell.worksheet_id == worksheet.id,
                Cell.column_index > column_index,
            )
            .values(
                column_index=Cell.column_index + offset,
            )
        )

        if shifted_result.rowcount:
            db.execute(
                update(Cell)
                .where(
                    Cell.worksheet_id == worksheet.id,
                    Cell.column_index > column_index + offset,
                )
                .values(
                    column_index=Cell.column_index - offset - 1,
                )
            )

        worksheet.max_column = max(
            0,
            old_max_column - 1,
        )

        workbook.version += 1

        db.commit()
        db.refresh(worksheet)
        db.refresh(workbook)

    except Exception:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to delete column",
        )

    return {
        "message": "Column deleted successfully",
        "workbook_id": workbook.id,
        "worksheet_id": worksheet.id,
        "deleted_column_index": column_index,
        "cells_deleted": deleted_result.rowcount,
        "cells_shifted": shifted_result.rowcount,
        "workbook_version": workbook.version,
        "worksheet": {
            "max_row": worksheet.max_row,
            "max_column": worksheet.max_column,
        },
    }