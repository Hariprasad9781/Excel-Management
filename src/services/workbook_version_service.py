from sqlalchemy.orm import Session

from models.cell import Cell
from models.workbook import Workbook
from models.workbook_version import WorkbookVersion
from models.worksheet import Worksheet


def create_workbook_version(
    db: Session,
    workbook: Workbook,
    created_by: int,
    change_summary: str | None = None,
) -> WorkbookVersion:
    """
    Create a snapshot of the current workbook state.

    The workbook should already contain the latest changes
    before this function is called.
    """

    # -----------------------------------------------------------------------
    # Determine the next version number
    # -----------------------------------------------------------------------

    next_version = workbook.version + 1

    # -----------------------------------------------------------------------
    # Build worksheet snapshot
    # -----------------------------------------------------------------------

    worksheets_snapshot = []

    worksheets = (
        db.query(Worksheet)
        .filter(
            Worksheet.workbook_id == workbook.id,
            Worksheet.is_deleted.is_(False),
        )
        .order_by(Worksheet.position)
        .all()
    )

    for worksheet in worksheets:

        cells_snapshot = []

        cells = (
            db.query(Cell)
            .filter(
                Cell.worksheet_id == worksheet.id,
            )
            .order_by(
                Cell.row_index,
                Cell.column_index,
            )
            .all()
        )

        for cell in cells:

            cells_snapshot.append(
                {
                    "row_index": cell.row_index,
                    "column_index": cell.column_index,
                    "value": cell.value,
                    "data_type": cell.data_type,
                    "formula": cell.formula,
                    "style": cell.style,
                }
            )

        worksheets_snapshot.append(
            {
                "id": worksheet.id,
                "name": worksheet.name,
                "position": worksheet.position,
                "max_row": worksheet.max_row,
                "max_column": worksheet.max_column,
                "cells": cells_snapshot,
            }
        )

    # -----------------------------------------------------------------------
    # Create version snapshot
    # -----------------------------------------------------------------------

    snapshot = WorkbookVersion(
        workbook_id=workbook.id,
        version_number=next_version,
        created_by=created_by,
        snapshot_data={
            "workbook": {
                "id": workbook.id,
                "name": workbook.name,
                "original_filename": workbook.original_filename,
            },
            "worksheets": worksheets_snapshot,
        },
        change_summary=change_summary,
    )

    db.add(snapshot)

    # -----------------------------------------------------------------------
    # Update current workbook version
    # -----------------------------------------------------------------------

    workbook.version = next_version

    return snapshot


def get_workbook_versions(
    db: Session,
    workbook_id: int,
) -> list[WorkbookVersion]:
    """
    Get all versions for a workbook.

    Versions are returned from newest to oldest.
    """

    return (
        db.query(WorkbookVersion)
        .filter(
            WorkbookVersion.workbook_id == workbook_id,
        )
        .order_by(
            WorkbookVersion.version_number.desc(),
        )
        .all()
    )


def restore_workbook_version(
    db: Session,
    workbook: Workbook,
    version_number: int,
    created_by: int,
) -> WorkbookVersion:
    """
    Restore a workbook to a previously saved version.

    The requested version is used as the source snapshot.
    A new version is created after the restore so that the
    original version history remains unchanged.
    """

    # -----------------------------------------------------------------------
    # Find requested version
    # -----------------------------------------------------------------------

    version = (
        db.query(WorkbookVersion)
        .filter(
            WorkbookVersion.workbook_id == workbook.id,
            WorkbookVersion.version_number == version_number,
        )
        .first()
    )

    if version is None:
        raise ValueError(
            f"Version {version_number} not found for workbook {workbook.id}"
        )

    snapshot_data = version.snapshot_data

    # -----------------------------------------------------------------------
    # Validate snapshot
    # -----------------------------------------------------------------------

    if not snapshot_data:
        raise ValueError(
            f"Version {version_number} does not contain snapshot data"
        )

    worksheets_snapshot = snapshot_data.get("worksheets", [])

    # -----------------------------------------------------------------------
    # Remove current workbook worksheets
    # -----------------------------------------------------------------------

    current_worksheets = (
        db.query(Worksheet)
        .filter(
            Worksheet.workbook_id == workbook.id,
        )
        .all()
    )

    for worksheet in current_worksheets:
        db.delete(worksheet)

    db.flush()

    # -----------------------------------------------------------------------
    # Restore worksheets and cells
    # -----------------------------------------------------------------------

    for worksheet_data in worksheets_snapshot:

        worksheet = Worksheet(
            workbook_id=workbook.id,
            name=worksheet_data["name"],
            position=worksheet_data["position"],
            max_row=worksheet_data.get("max_row", 0),
            max_column=worksheet_data.get("max_column", 0),
            is_deleted=False,
        )

        db.add(worksheet)
        db.flush()

        cells_snapshot = worksheet_data.get("cells", [])

        for cell_data in cells_snapshot:

            cell = Cell(
                worksheet_id=worksheet.id,
                row_index=cell_data["row_index"],
                column_index=cell_data["column_index"],
                value=cell_data.get("value"),
                data_type=cell_data.get("data_type"),
                formula=cell_data.get("formula"),
                style=cell_data.get("style"),
            )

            db.add(cell)


    # -----------------------------------------------------------------------
    # Flush restored cells before creating the new version
    # -----------------------------------------------------------------------

    db.flush()

    # -----------------------------------------------------------------------
    # Restore workbook metadata
    # -----------------------------------------------------------------------

    workbook_data = snapshot_data.get("workbook", {})

    if "name" in workbook_data:
        workbook.name = workbook_data["name"]

    if "original_filename" in workbook_data:
        workbook.original_filename = workbook_data["original_filename"]

    # -----------------------------------------------------------------------
    # Create a NEW version representing the restore
    # -----------------------------------------------------------------------

    new_version = create_workbook_version(
        db=db,
        workbook=workbook,
        created_by=created_by,
        change_summary=f"Restored from version {version_number}",
    )

    return new_version