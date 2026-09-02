from sqlalchemy import text

from database import engine


def main():
    with engine.begin() as connection:

        # -----------------------------------------------------
        # Add workbook_id column
        # -----------------------------------------------------
        connection.execute(
            text(
                """
                ALTER TABLE excel_files
                ADD COLUMN IF NOT EXISTS workbook_id INTEGER
                """
            )
        )

        # -----------------------------------------------------
        # Add index
        # -----------------------------------------------------
        connection.execute(
            text(
                """
                CREATE INDEX IF NOT EXISTS
                ix_excel_files_workbook_id
                ON excel_files (workbook_id)
                """
            )
        )

        # -----------------------------------------------------
        # Add foreign key
        # -----------------------------------------------------
        connection.execute(
            text(
                """
                DO $$
                BEGIN
                    IF NOT EXISTS (
                        SELECT 1
                        FROM pg_constraint
                        WHERE conname = 'fk_excel_files_workbook_id'
                    ) THEN
                        ALTER TABLE excel_files
                        ADD CONSTRAINT fk_excel_files_workbook_id
                        FOREIGN KEY (workbook_id)
                        REFERENCES workbooks(id)
                        ON DELETE SET NULL;
                    END IF;
                END
                $$;
                """
            )
        )

    print(
        "workbook_id column and foreign key added successfully."
    )


if __name__ == "__main__":
    main()