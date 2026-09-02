from database import Base, engine

from models import (
    Cell,
    Workbook,
    Worksheet,
)


def create_excel_tables():
    Base.metadata.create_all(
        bind=engine,
        tables=[
            Workbook.__table__,
            Worksheet.__table__,
            Cell.__table__,
        ],
    )

    print("Excel data tables created successfully.")


if __name__ == "__main__":
    create_excel_tables()