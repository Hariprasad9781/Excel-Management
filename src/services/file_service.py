import os
import uuid

from sqlalchemy.orm import Session

from models.excel_file import ExcelFile


STORAGE_DIR = os.path.join(
    os.path.dirname(__file__),
    "..",
    "..",
    "storage",
)

os.makedirs(STORAGE_DIR, exist_ok=True)


def save_excel_file(
    db: Session,
    user_id: int,
    original_filename: str,
    file_content: bytes,
) -> ExcelFile:

    extension = os.path.splitext(original_filename)[1].lower()

    stored_filename = f"{uuid.uuid4()}{extension}"

    file_path = os.path.join(
        STORAGE_DIR,
        stored_filename,
    )

    with open(file_path, "wb") as file:
        file.write(file_content)

    excel_file = ExcelFile(
        user_id=user_id,
        original_filename=original_filename,
        stored_filename=stored_filename,
        file_path=file_path,
        file_size=len(file_content),
    )

    db.add(excel_file)
    db.commit()
    db.refresh(excel_file)

    return excel_file