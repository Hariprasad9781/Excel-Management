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