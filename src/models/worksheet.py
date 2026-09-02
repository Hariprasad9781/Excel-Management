from datetime import datetime
from zoneinfo import ZoneInfo

from sqlalchemy import Boolean, DateTime, ForeignKey, Integer, String
from sqlalchemy.orm import Mapped, mapped_column, relationship

from database import Base

IST = ZoneInfo("Asia/Kolkata")


class Worksheet(Base):
    __tablename__ = "worksheets"

    id: Mapped[int] = mapped_column(
        primary_key=True,
        autoincrement=True,
    )

    workbook_id: Mapped[int] = mapped_column(
        ForeignKey("workbooks.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )

    name: Mapped[str] = mapped_column(
        String(255),
        nullable=False,
    )

    position: Mapped[int] = mapped_column(
        Integer,
        nullable=False,
        default=0,
    )

    max_row: Mapped[int] = mapped_column(
        Integer,
        nullable=False,
        default=0,
    )

    max_column: Mapped[int] = mapped_column(
        Integer,
        nullable=False,
        default=0,
    )

    is_deleted: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        default=False,
        index=True,
    )

    created_at: Mapped[datetime] = mapped_column(
        DateTime,
        default=lambda: datetime.now(IST),
        nullable=False,
    )

    updated_at: Mapped[datetime] = mapped_column(
        DateTime,
        default=lambda: datetime.now(IST),
        onupdate=lambda: datetime.now(IST),
        nullable=False,
    )

    workbook = relationship(
        "Workbook",
        back_populates="worksheets",
    )

    cells = relationship(
        "Cell",
        back_populates="worksheet",
        cascade="all, delete-orphan",
    )