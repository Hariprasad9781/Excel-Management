from database import SessionLocal
from models import Cell, Workbook, Worksheet


def verify_import():
    db = SessionLocal()

    try:
        workbook = (
            db.query(Workbook)
            .filter(Workbook.id == 2)
            .first()
        )

        if not workbook:
            print("Workbook not found.")
            return

        print(f"Workbook: {workbook.name}")
        print(f"Original filename: {workbook.original_filename}")
        print(f"Version: {workbook.version}")
        print()

        for worksheet in workbook.worksheets:
            print(f"Worksheet: {worksheet.name}")
            print(f"Rows: {worksheet.max_row}")
            print(f"Columns: {worksheet.max_column}")
            print()

            cells = (
                db.query(Cell)
                .filter(Cell.worksheet_id == worksheet.id)
                .order_by(Cell.row_index, Cell.column_index)
                .limit(20)
                .all()
            )

            print("First 20 imported cells:")

            for cell in cells:
                print(
                    f"Row={cell.row_index}, "
                    f"Column={cell.column_index}, "
                    f"Value={cell.value!r}, "
                    f"Type={cell.data_type!r}, "
                    f"Formula={cell.formula!r}"
                )

                if cell.style:
                    print(f"  Style: {cell.style}")

            print()

    finally:
        db.close()


if __name__ == "__main__":
    verify_import()