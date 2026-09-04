import os

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile, status
from fastapi.responses import FileResponse
from sqlalchemy.orm import Session

from api.schemas import (
    FileDeleteResponse,
    FileDetailsResponse,
    FileListItem,
    FileUploadResponse,
)
from database import get_db
from dependencies import get_current_user, require_permission
from models.excel_file import ExcelFile
from models.user import User
from services.excel_import_service import import_excel_to_database
from services.file_service import save_excel_file


router = APIRouter(
    prefix="/files",
    tags=["Excel Files"],
)


ALLOWED_EXTENSIONS = {".xlsx", ".xls"}


# ---------------------------------------------------------------------------
# Upload Excel File
# ---------------------------------------------------------------------------

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
    """
    Upload an Excel file.

    For .xlsx files:
        1. Save the original Excel file to storage.
        2. Import its workbook/sheets/cells into PostgreSQL.
        3. If database import fails, clean up the saved file and DB record.

    For .xls files:
        The original file is saved, but database import is skipped because
        the current database importer supports .xlsx only.
    """

    # -----------------------------------------------------------------------
    # Validate filename
    # -----------------------------------------------------------------------

    if not file.filename:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Filename is required",
        )

    # -----------------------------------------------------------------------
    # Validate extension
    # -----------------------------------------------------------------------

    _, extension = os.path.splitext(file.filename)
    extension = extension.lower()

    if extension not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Only .xlsx and .xls files are allowed",
        )

    # -----------------------------------------------------------------------
    # Read uploaded file
    # -----------------------------------------------------------------------

    file_content = await file.read()

    if not file_content:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File is empty",
        )

    # -----------------------------------------------------------------------
    # Save physical Excel file + metadata
    # -----------------------------------------------------------------------

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

    # -----------------------------------------------------------------------
    # Import .xlsx into PostgreSQL
    # -----------------------------------------------------------------------

    if extension == ".xlsx":

        try:
            import_excel_to_database(
                db=db,
                file_path=excel_file.file_path,
                owner_id=current_user.id,
                original_filename=file.filename,
            )

        except Exception:
            # ---------------------------------------------------------------
            # Database import failed.
            #
            # The importer may already have rolled back its transaction,
            # but the ExcelFile record was created by save_excel_file().
            # Remove the metadata record and physical file as well.
            # ---------------------------------------------------------------

            db.rollback()

            # ---------------------------------------------------------------
            # Remove physical file
            # ---------------------------------------------------------------

            if os.path.exists(excel_file.file_path):
                try:
                    os.remove(excel_file.file_path)
                except OSError:
                    pass

            # ---------------------------------------------------------------
            # Remove ExcelFile metadata record
            # ---------------------------------------------------------------

            try:
                db.delete(excel_file)
                db.commit()

            except Exception:
                db.rollback()

            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail="Excel file was saved but database import failed",
            )

    # -----------------------------------------------------------------------
    # Return upload response
    # -----------------------------------------------------------------------

    return {
        "id": excel_file.id,
        "filename": excel_file.original_filename,
        "stored_filename": excel_file.stored_filename,
        "file_size": excel_file.file_size,
        "message": "File uploaded successfully",
    }


# ---------------------------------------------------------------------------
# List User Files
# ---------------------------------------------------------------------------

@router.get(
    "/",
    response_model=list[FileListItem],
)
def list_files(
    current_user: User = Depends(
    require_permission("view_file")),
    db: Session = Depends(get_db),
    ):
    """
    Return all Excel files belonging to the current user.
    """

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
            "id": excel_file.id,
            "filename": excel_file.original_filename,

            # FIX:
            # FileListItem requires stored_filename.
            "stored_filename": excel_file.stored_filename,

            "file_size": excel_file.file_size,
            "created_at": excel_file.created_at,
            "updated_at": excel_file.updated_at,
        }
        for excel_file in files
    ]


# ---------------------------------------------------------------------------
# Get File Details
# ---------------------------------------------------------------------------

@router.get(
    "/{file_id}",
    response_model=FileDetailsResponse,
)
def get_file_details(
    file_id: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """
    Return details for a specific Excel file.

    Users can only access their own files.
    """

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
        "file_path": excel_file.file_path,
        "file_size": excel_file.file_size,
        "created_at": excel_file.created_at,
        "updated_at": excel_file.updated_at,
    }


# ---------------------------------------------------------------------------
# Download Excel File
# ---------------------------------------------------------------------------

@router.get(
    "/{file_id}/download",
)
def download_file(
    file_id: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """
    Download the original uploaded Excel file.

    Users can only download their own files.
    """

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

    # -----------------------------------------------------------------------
    # Check physical file
    # -----------------------------------------------------------------------

    if not os.path.exists(excel_file.file_path):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Physical file not found",
        )

    # -----------------------------------------------------------------------
    # Return file
    # -----------------------------------------------------------------------

    return FileResponse(
        path=excel_file.file_path,
        filename=excel_file.original_filename,
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if excel_file.original_filename.lower().endswith(".xlsx")
            else "application/vnd.ms-excel"
        ),
    )


# ---------------------------------------------------------------------------
# Delete Excel File
# ---------------------------------------------------------------------------

@router.delete(
    "/{file_id}",
    response_model=FileDeleteResponse,
)
def delete_file(
    file_id: int,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """
    Delete an Excel file.

    This removes:
        - Physical Excel file
        - ExcelFile metadata record

    Database workbook data is intentionally NOT deleted here yet.
    That cleanup will be handled when workbook/file relationships are
    fully integrated.
    """

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

    # -----------------------------------------------------------------------
    # Remove physical file
    # -----------------------------------------------------------------------

    if os.path.exists(excel_file.file_path):

        try:
            os.remove(excel_file.file_path)

        except OSError:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail="Failed to delete physical file",
            )

    # -----------------------------------------------------------------------
    # Remove database metadata
    # -----------------------------------------------------------------------

    try:
        db.delete(excel_file)
        db.commit()

    except Exception:
        db.rollback()

        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to delete file record",
        )

    # -----------------------------------------------------------------------
    # Response
    # -----------------------------------------------------------------------

    return {
        "id": file_id,
        "message": "File deleted successfully",
    }