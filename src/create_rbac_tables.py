from database import Base, engine

# Import all models so SQLAlchemy knows about them.
from models import (
    User,
    ExcelFile,
    Role,
    Permission,
    RolePermission,
    UserPermission,
)


def create_rbac_tables():
    Base.metadata.create_all(
        bind=engine,
        tables=[
            Role.__table__,
            Permission.__table__,
            RolePermission.__table__,
            UserPermission.__table__,
        ],
    )

    print("RBAC tables created successfully.")


if __name__ == "__main__":
    create_rbac_tables()