from database import SessionLocal
from models import Cell, Workbook


def verify_import():
    db = SessionLocal()

    try:
        workbook = (
            db.query(Workbook)
            .filter(Workbook.id == 3)
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
                .all()
            )

            print(f"Total cells: {len(cells)}")
            print()

            for cell in cells:
                print(
                    f"{cell.row_index},{cell.column_index} | "
                    f"value={cell.value!r} | "
                    f"type={cell.data_type!r} | "
                    f"formula={cell.formula!r}"
                )

                if cell.style:
                    print(f"    style={cell.style}")

    finally:
        db.close()


if __name__ == "__main__":
    verify_import()