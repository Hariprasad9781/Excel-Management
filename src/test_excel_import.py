from database import SessionLocal
from services.excel_import_service import import_excel_to_database


EXCEL_FILE_PATH = r"storage/import_test.xlsx"


def test_excel_import():
    db = SessionLocal()

    try:
        workbook = import_excel_to_database(
            db=db,
            file_path=EXCEL_FILE_PATH,
            owner_id=1,
            original_filename="test_import.xlsx",
        )

        print("Excel import successful!")
        print(f"Workbook ID: {workbook.id}")
        print(f"Workbook Name: {workbook.name}")
        print(f"Worksheets: {len(workbook.worksheets)}")

        for worksheet in workbook.worksheets:
            print(
                f"  - {worksheet.name}: "
                f"{worksheet.max_row} rows x "
                f"{worksheet.max_column} columns"
            )

            cell_count = len(worksheet.cells)

            print(f"    Cells imported: {cell_count}")

    except Exception:
        db.rollback()
        raise

    finally:
        db.close()


if __name__ == "__main__":
    test_excel_import()