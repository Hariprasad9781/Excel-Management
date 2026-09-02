from datetime import datetime
from zoneinfo import ZoneInfo

from sqlalchemy import DateTime, ForeignKey, Integer, JSON, String, Text, UniqueConstraint
from sqlalchemy.orm import Mapped, mapped_column, relationship

from database import Base

IST = ZoneInfo("Asia/Kolkata")


class Cell(Base):
    __tablename__ = "cells"

    id: Mapped[int] = mapped_column(
        primary_key=True,
        autoincrement=True,
    )

    worksheet_id: Mapped[int] = mapped_column(
        ForeignKey("worksheets.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )

    row_index: Mapped[int] = mapped_column(
        Integer,
        nullable=False,
    )

    column_index: Mapped[int] = mapped_column(
        Integer,
        nullable=False,
    )

    value: Mapped[object | None] = mapped_column(
        JSON,
        nullable=True,
    )

    data_type: Mapped[str | None] = mapped_column(
        String(30),
        nullable=True,
    )

    formula: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    style: Mapped[dict | None] = mapped_column(
        JSON,
        nullable=True,
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

    worksheet = relationship(
        "Worksheet",
        back_populates="cells",
    )

    __table_args__ = (
        UniqueConstraint(
            "worksheet_id",
            "row_index",
            "column_index",
            name="uq_cell_worksheet_position",
        ),
    )