from database import SessionLocal
from models.workbook import Workbook
from services.workbook_version_service import restore_workbook_version


def test_workbook_restore():
    db = SessionLocal()

    try:
        workbook = (
            db.query(Workbook)
            .filter(Workbook.id == 4)
            .first()
        )

        if not workbook:
            print("Workbook ID 4 not found.")
            return

        print(f"Current workbook version: {workbook.version}")
        print("Restoring Version 15...")

        version = restore_workbook_version(
            db=db,
            workbook=workbook,
            version_number=15,
            created_by=workbook.owner_id,
        )

        db.commit()
        db.refresh(version)

        print("\nRestore successful!")
        print(f"New Version ID: {version.id}")
        print(f"Workbook ID: {version.workbook_id}")
        print(f"New Version Number: {version.version_number}")
        print(f"Created By: {version.created_by}")
        print(f"Change Summary: {version.change_summary}")

    except Exception:
        db.rollback()
        raise

    finally:
        db.close()


if __name__ == "__main__":
    test_workbook_restore()