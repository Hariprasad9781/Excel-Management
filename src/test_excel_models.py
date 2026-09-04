from database import SessionLocal
from models import Cell, User, Workbook, Worksheet


def test_excel_models():
    db = SessionLocal()

    try:
        # Get an existing user
        user = db.query(User).first()

        if not user:
            print("No user found. Create a user first.")
            return

        # Create workbook
        workbook = Workbook(
            owner_id=user.id,
            name="Test Workbook",
            original_filename="test.xlsx",
        )

        db.add(workbook)
        db.flush()

        # Create worksheet
        worksheet = Worksheet(
            workbook_id=workbook.id,
            name="Sheet1",
            position=0,
            max_row=1,
            max_column=2,
        )

        db.add(worksheet)
        db.flush()

        # Create cells
        cell_a1 = Cell(
            worksheet_id=worksheet.id,
            row_index=1,
            column_index=1,
            value="Hello",
            data_type="string",
        )

        cell_b1 = Cell(
            worksheet_id=worksheet.id,
            row_index=1,
            column_index=2,
            value=123,
            data_type="number",
        )

        db.add_all([cell_a1, cell_b1])
        db.commit()

        print("Excel model test successful.")
        print(f"Workbook ID: {workbook.id}")
        print(f"Worksheet ID: {worksheet.id}")
        print(f"Cell A1 ID: {cell_a1.id}")
        print(f"Cell B1 ID: {cell_b1.id}")

    except Exception:
        db.rollback()
        raise

    finally:
        db.close()


if __name__ == "__main__":
    test_excel_models()