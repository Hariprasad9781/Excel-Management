from sqlalchemy import text

from database import SessionLocal
from models.excel_file import ExcelFile
from models.workbook import Workbook
from models.workbook_version import WorkbookVersion
from services.file_version_service import _build_snapshot


WORKBOOK_ID = 4
EXCEL_FILE_ID = 5
USER_ID = 3


def main():
    db = SessionLocal()

    try:
        workbook = (
            db.query(Workbook)
            .filter(Workbook.id == WORKBOOK_ID)
            .first()
        )

        if workbook is None:
            raise ValueError(
                f"Workbook {WORKBOOK_ID} not found."
            )

        excel_file = (
            db.query(ExcelFile)
            .filter(ExcelFile.id == EXCEL_FILE_ID)
            .first()
        )

        if excel_file is None:
            raise ValueError(
                f"Excel file {EXCEL_FILE_ID} not found."
            )

        if excel_file.workbook_id != workbook.id:
            raise ValueError(
                "Excel file is not linked to the expected workbook."
            )

        existing = (
            db.query(WorkbookVersion)
            .filter(
                WorkbookVersion.workbook_id == workbook.id,
                WorkbookVersion.version_number == workbook.version,
            )
            .first()
        )

        if existing:
            print(
                f"Version {workbook.version} already exists."
            )
            return

        snapshot = _build_snapshot(excel_file)

        version = WorkbookVersion(
            workbook_id=workbook.id,
            version_number=workbook.version,
            created_by=USER_ID,
            snapshot_data=snapshot,
            change_summary="Baseline snapshot",
        )

        db.add(version)
        db.commit()

        print(
            f"Initial version {workbook.version} "
            f"created successfully."
        )

    except Exception:
        db.rollback()
        raise

    finally:
        db.close()


if __name__ == "__main__":
    main()