from database import SessionLocal
from models.workbook import Workbook
from services.workbook_version_service import create_workbook_version


def test_workbook_version():
    db = SessionLocal()

    try:
        workbook = (
            db.query(Workbook)
            .filter(Workbook.id == 3)
            .first()
        )

        if not workbook:
            print("Workbook ID 3 not found.")
            return

        version = create_workbook_version(
            db=db,
            workbook=workbook,
            created_by=workbook.owner_id,
            change_summary="Initial workbook snapshot",
        )

        db.commit()
        db.refresh(version)

        print("Version created successfully.")
        print(f"Version ID: {version.id}")
        print(f"Workbook ID: {version.workbook_id}")
        print(f"Version Number: {version.version_number}")
        print(f"Created By: {version.created_by}")
        print(f"Change Summary: {version.change_summary}")

        print("\nSnapshot:")
        print(version.snapshot_data)

    except Exception:
        db.rollback()
        raise

    finally:
        db.close()


if __name__ == "__main__":
    test_workbook_version()