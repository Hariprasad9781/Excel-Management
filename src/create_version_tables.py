from database import Base, engine
from models import WorkbookVersion


def create_version_tables():
    Base.metadata.create_all(
        bind=engine,
        tables=[
            WorkbookVersion.__table__,
        ],
    )

    print("Workbook version table created successfully.")


if __name__ == "__main__":
    create_version_tables()