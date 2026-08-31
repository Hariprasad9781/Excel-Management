from database import Base, engine
from models import ExcelFile, User


def create_tables() -> None:
    Base.metadata.create_all(bind=engine)
    print("Database tables created successfully.")


if __name__ == "__main__":
    create_tables()