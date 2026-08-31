from datetime import datetime

from pydantic import BaseModel, EmailStr


# ============================================================
# User Schemas
# ============================================================

class UserCreate(BaseModel):
    username: str
    email: EmailStr
    password: str


class UserResponse(BaseModel):
    id: int
    username: str
    email: EmailStr
    is_active: bool

    model_config = {
        "from_attributes": True,
    }


class UserLogin(BaseModel):
    username: str
    password: str


class TokenResponse(BaseModel):
    access_token: str
    token_type: str


# ============================================================
# File Schemas
# ============================================================

class FileUploadResponse(BaseModel):
    id: int
    filename: str
    stored_filename: str
    file_size: int
    message: str


class FileListItem(BaseModel):
    id: int
    filename: str
    stored_filename: str
    file_size: int
    created_at: datetime
    updated_at: datetime


class FileDetailsResponse(BaseModel):
    id: int
    filename: str
    stored_filename: str
    file_size: int
    created_at: datetime
    updated_at: datetime


class FileDeleteResponse(BaseModel):
    message: str
    file_id: int


# ============================================================
# Excel Data Schemas
# ============================================================

class SheetsResponse(BaseModel):
    file_id: int
    filename: str
    sheets: list[str]


class SheetPreviewResponse(BaseModel):
    file_id: int
    filename: str
    sheet_name: str
    columns: list[str]
    rows: list[dict]
    row_count: int


# ============================================================
# Cell Operations
# ============================================================

class CellUpdate(BaseModel):
    sheet_name: str
    cell: str
    value: str | int | float | bool | None


class CellUpdateResponse(BaseModel):
    file_id: int
    sheet_name: str
    cell: str
    value: str | int | float | bool | None
    message: str


# ============================================================
# Row Operations
# ============================================================

class RowAddRequest(BaseModel):
    sheet_name: str
    row_data: list[str | int | float | bool | None]


class RowUpdateRequest(BaseModel):
    sheet_name: str
    row_number: int
    row_data: list[str | int | float | bool | None]


class RowDeleteRequest(BaseModel):
    sheet_name: str
    row_number: int


class RowOperationResponse(BaseModel):
    file_id: int
    sheet_name: str
    row_number: int
    message: str