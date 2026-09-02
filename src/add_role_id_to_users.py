from sqlalchemy import text

from database import engine


def add_role_id_column():
    with engine.begin() as connection:
        connection.execute(
            text(
                """
                ALTER TABLE users
                ADD COLUMN IF NOT EXISTS role_id INTEGER
                """
            )
        )

        connection.execute(
            text(
                """
                CREATE INDEX IF NOT EXISTS ix_users_role_id
                ON users (role_id)
                """
            )
        )

        connection.execute(
            text(
                """
                ALTER TABLE users
                ADD CONSTRAINT fk_users_role_id_roles
                FOREIGN KEY (role_id)
                REFERENCES roles (id)
                ON DELETE SET NULL
                """
            )
        )

    print("users.role_id added successfully.")


if __name__ == "__main__":
    add_role_id_column()