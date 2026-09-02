from sqlalchemy import text

from database import engine


def add_version_unique_constraint():
    with engine.begin() as connection:
        connection.execute(
            text(
                """
                ALTER TABLE workbook_versions
                ADD CONSTRAINT uq_workbook_version_number
                UNIQUE (workbook_id, version_number)
                """
            )
        )

    print("Version unique constraint added successfully.")


if __name__ == "__main__":
    add_version_unique_constraint()