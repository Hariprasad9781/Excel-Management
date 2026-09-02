from datetime import datetime
from zoneinfo import ZoneInfo

from sqlalchemy import (
    DateTime,
    ForeignKey,
    Integer,
    JSON,
    Text,
    UniqueConstraint,
)
from sqlalchemy.orm import Mapped, mapped_column, relationship

from database import Base

IST = ZoneInfo("Asia/Kolkata")


class WorkbookVersion(Base):
    __tablename__ = "workbook_versions"

    id: Mapped[int] = mapped_column(
        primary_key=True,
        autoincrement=True,
    )

    workbook_id: Mapped[int] = mapped_column(
        ForeignKey("workbooks.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )

    version_number: Mapped[int] = mapped_column(
        Integer,
        nullable=False,
    )

    created_by: Mapped[int] = mapped_column(
        ForeignKey("users.id", ondelete="SET NULL"),
        nullable=True,
        index=True,
    )

    snapshot_data: Mapped[dict] = mapped_column(
        JSON,
        nullable=False,
    )

    created_at: Mapped[datetime] = mapped_column(
        DateTime,
        default=lambda: datetime.now(IST),
        nullable=False,
    )

    change_summary: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    workbook = relationship(
        "Workbook",
        back_populates="versions",
    )

    creator = relationship(
        "User",
    )

    __table_args__ = (
        UniqueConstraint(
            "workbook_id",
            "version_number",
            name="uq_workbook_version_number",
        ),
    )