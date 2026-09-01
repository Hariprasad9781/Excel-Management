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
# Sheet Operations
# ============================================================


class SheetCreateRequest(BaseModel):
    sheet_name: str


class SheetOperationResponse(BaseModel):
    file_id: int
    sheet_name: str
    message: str


class SheetRenameRequest(BaseModel):
    sheet_name: str
    new_sheet_name: str


class SheetDeleteRequest(BaseModel):
    sheet_name: str


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


# ============================================================
# Column Operations
# ============================================================


class ColumnAddRequest(BaseModel):
    sheet_name: str
    column_number: int


class ColumnUpdateRequest(BaseModel):
    sheet_name: str
    column_number: int
    column_name: str


class ColumnDeleteRequest(BaseModel):
    sheet_name: str
    column_number: int


class ColumnOperationResponse(BaseModel):
    file_id: int
    sheet_name: str
    column_number: int
    message: str


# ============================================================
# Excel Search
# ============================================================


class ExcelSearchRequest(BaseModel):
    sheet_name: str
    search_term: str


class ExcelSearchResult(BaseModel):
    row_number: int
    column_number: int
    cell: str
    value: str | int | float | bool | None


class ExcelSearchResponse(BaseModel):
    file_id: int
    filename: str
    sheet_name: str
    search_term: str
    results: list[ExcelSearchResult]
    result_count: int


# ============================================================
# Cell / Range Formatting
# ============================================================


class ExcelFormatRequest(BaseModel):
    sheet_name: str

    # Single cell or range.
    #
    # Examples:
    #   A1
    #   B2
    #   A1:D5
    cell_range: str

    # Font formatting
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    font_size: float | None = None
    font_color: str | None = None

    # Cell background
    fill_color: str | None = None

    # Alignment
    horizontal_alignment: str | None = None
    vertical_alignment: str | None = None

    # Excel number format
    number_format: str | None = None


class ExcelFormatResponse(BaseModel):
    file_id: int
    sheet_name: str
    cell_range: str
    message: str