from sqlalchemy import text

from database import engine


def add_snapshot_data_column():
    with engine.begin() as connection:
        connection.execute(
            text(
                """
                ALTER TABLE workbook_versions
                ADD COLUMN IF NOT EXISTS snapshot_data JSON
                """
            )
        )

    print("snapshot_data column added successfully.")


if __name__ == "__main__":
    add_snapshot_data_column()