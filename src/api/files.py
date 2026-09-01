import os

from fastapi import (
    APIRouter,
    Depends,
    File,
    HTTPException,
    UploadFile,
    status,
)
from fastapi.responses import FileResponse
from sqlalchemy.orm import Session

from api.schemas import (
    FileDeleteResponse,
    FileDetailsResponse,
    FileListItem,
    FileUploadResponse,
)
from database import get_db
from dependencies import get_current_user
from models.excel_file import ExcelFile
from models.user import User
from services.file_service import save_excel_file


# ============================================================
# Router Configuration
# ============================================================

router = APIRouter(
    prefix="/files",
    tags=["Excel Files"],
)


# ============================================================
# File Configuration
# ============================================================

ALLOWED_EXTENSIONS = {
    ".xlsx",
    ".xls",
}


# ============================================================
# Upload File
# ============================================================


@router.post(
    "/upload",
    response_model=FileUploadResponse,
    status_code=status.HTTP_201_CREATED,
)
async def upload_file(
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    # --------------------------------------------------------
    # Validate filename
    # --------------------------------------------------------

    if not file.filename:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Filename is required",
        )

    # --------------------------------------------------------
    # Validate file extension
    # --------------------------------------------------------

    _, extension = os.path.splitext(
        file.filename
    )

    extension = extension.lower()

    if extension not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Only .xlsx and .xls files are allowed",
        )

    # --------------------------------------------------------
    # Read file
    # --------------------------------------------------------

    file_content = await file.read()

    # --------------------------------------------------------
    # Validate empty file
    # --------------------------------------------------------

    if not file_content:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File is empty",
        )

    # --------------------------------------------------------
    # Save file and database metadata
    # --------------------------------------------------------

    try:
        excel_file = save_excel_file(
            db=db,
            user_id=current_user.id,
            original_filename=file.filename,
            file_content=file_content,
        )

    except Exception:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to save Excel file",
        )

    # --------------------------------------------------------
    # Response
    # --------------------------------------------------------

    return {
        "id": excel_file.id,
        "filename": excel_file.original_filename,
        "stored_filename": excel_file.stored_filename,
        "file_size": excel_file.file_size,
        "message": "File uploaded successfully",
    }


# ============================================================
# List Files
# ============================================================


@router.get(
    "/",
    response_model=list[FileListItem],
)
def list_files(
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    # Get only files belonging to the logged-in user
    files = (
        db.query(ExcelFile)
        .filter(
            ExcelFile.user_id == current_user.id
        )
        .order_by(
            ExcelFile.created_at.desc()
        )
        .all()
    )

    return [
        {
            "id": file.id,
            "filename": file.original_filename,
            "stored_filename": file.stored_filename,
            "file_size": file.file_size,
            "created_at": file.created_at,
            "updated_at": file.updated_at,
        }
        for file in files
    ]


# ============================================================
# Get File Details
# ============================================================


@router.get(
    "/{file_id}",
    response_model=FileDetailsResponse,
)
def get_file_details(
    file_id: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = (
        db.query(ExcelFile)
        .filter(
            ExcelFile.id == file_id,
            ExcelFile.user_id == current_user.id,
        )
        .first()
    )

    if not excel_file:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="File not found",
        )

    return {
        "id": excel_file.id,
        "filename": excel_file.original_filename,
        "stored_filename": excel_file.stored_filename,
        "file_size": excel_file.file_size,
        "created_at": excel_file.created_at,
        "updated_at": excel_file.updated_at,
    }


# ============================================================
# Download File
# ============================================================


@router.get(
    "/{file_id}/download",
)
def download_file(
    file_id: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = (
        db.query(ExcelFile)
        .filter(
            ExcelFile.id == file_id,
            ExcelFile.user_id == current_user.id,
        )
        .first()
    )

    if not excel_file:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="File not found",
        )

    # Check that the physical file still exists
    if not os.path.exists(
        excel_file.file_path
    ):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Stored file not found",
        )

    return FileResponse(
        path=excel_file.file_path,
        filename=excel_file.original_filename,
        media_type="application/octet-stream",
    )


# ============================================================
# Delete File
# ============================================================


@router.delete(
    "/{file_id}",
    response_model=FileDeleteResponse,
)
def delete_file(
    file_id: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    excel_file = (
        db.query(ExcelFile)
        .filter(
            ExcelFile.id == file_id,
            ExcelFile.user_id == current_user.id,
        )
        .first()
    )

    if not excel_file:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="File not found",
        )

    # --------------------------------------------------------
    # Delete physical file
    # --------------------------------------------------------

    if os.path.exists(
        excel_file.file_path
    ):
        try:
            os.remove(
                excel_file.file_path
            )

        except OSError:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail="Failed to delete stored file",
            )

    # --------------------------------------------------------
    # Delete database record
    # --------------------------------------------------------

    db.delete(excel_file)
    db.commit()

    # --------------------------------------------------------
    # Response
    # --------------------------------------------------------

    return {
        "message": "File deleted successfully",
        "file_id": file_id,
    }