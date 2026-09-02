from datetime import datetime
from zoneinfo import ZoneInfo

from sqlalchemy import Boolean, ForeignKey, String
from sqlalchemy.orm import Mapped, mapped_column, relationship

from database import Base


IST = ZoneInfo("Asia/Kolkata")


class Workbook(Base):
    __tablename__ = "workbooks"

    id: Mapped[int] = mapped_column(
        primary_key=True,
        autoincrement=True,
    )

    owner_id: Mapped[int] = mapped_column(
        ForeignKey("users.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )

    name: Mapped[str] = mapped_column(
        String(255),
        nullable=False,
    )

    original_filename: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    version: Mapped[int] = mapped_column(
        default=1,
        nullable=False,
    )

    is_deleted: Mapped[bool] = mapped_column(
        Boolean,
        default=False,
        nullable=False,
        index=True,
    )

    created_at: Mapped[datetime] = mapped_column(
        default=lambda: datetime.now(IST),
        nullable=False,
    )

    updated_at: Mapped[datetime] = mapped_column(
        default=lambda: datetime.now(IST),
        onupdate=lambda: datetime.now(IST),
        nullable=False,
    )

    owner = relationship(
        "User",
        back_populates="workbooks",
    )

    worksheets = relationship(
        "Worksheet",
        back_populates="workbook",
        cascade="all, delete-orphan",
        order_by="Worksheet.position",
    )

    versions = relationship(
    "WorkbookVersion",
    back_populates="workbook",
    cascade="all, delete-orphan",
    order_by="WorkbookVersion.version_number",
    )

    excel_file = relationship(
    "ExcelFile",
    back_populates="workbook",
    uselist=False,
    )